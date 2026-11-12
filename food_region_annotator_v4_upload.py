import streamlit as st
import pandas as pd
import numpy as np
from pathlib import Path
from PIL import Image
import cv2
import json
import re
import torch
import torch.nn as nn
import torchvision.transforms.functional as TF
from torchvision.models.segmentation import deeplabv3_resnet50
from scipy import ndimage

from skimage.color import rgb2lab
from skimage.segmentation import slic, find_boundaries
from streamlit_drawable_canvas import st_canvas


# ============================================================
# SETTINGS
# ============================================================

PAIRING_FILENAME = "meal_image_pairing.xlsx"
FOOD_BREAKDOWNS_FILENAME = "Food Breakdowns.xlsx"
AFTER_QC_FILENAME = "after_mask_qc.csv"

# These should point to your already-reviewed binary masks.
BEFORE_MASK_FOLDER = Path("before_segmentation_masks")
AFTER_MASK_FOLDER = Path("after_segmentation_masks")

OUTPUT_FOLDER = Path("food_region_annotations")

REGION_FOLDER = OUTPUT_FOLDER / "regions"
LABEL_FOLDER = OUTPUT_FOLDER / "labels"

REGION_FOLDER.mkdir(parents=True, exist_ok=True)
LABEL_FOLDER.mkdir(parents=True, exist_ok=True)


# ============================================================
# SLIC SETTINGS
# ============================================================

N_SEGMENTS = 100
COMPACTNESS = 10
SIGMA = 1
# ============================================================
# REGION CLEANUP SETTINGS
# ============================================================

# Connected pieces of the same SLIC label are separated
# into independent selectable regions.
#
# Tiny connected fragments below this size are absorbed
# into a neighbouring region where possible.
MIN_REGION_PIXELS = 20

# ============================================================
# DISPLAY SETTINGS
# ============================================================

# 848 x 480 becomes approximately 1272 x 720.
DISPLAY_SCALE = 1.5

DEFAULT_BRUSH_WIDTH = 18

# ============================================================
# V1 MODEL SUGGESTION SETTINGS
# ============================================================

V1_DATASET_FOLDER = Path("food_multilabel_dataset_v4")
V1_MODEL_FILE = V1_DATASET_FOLDER / "food_multilabel_v4_best.pth"
V1_CLASSES_FILE = V1_DATASET_FOLDER / "classes.json"
V1_THRESHOLD = 0.50

# A superpixel receives a class suggestion when at least this
# fraction of its pixels are predicted as that class.
V1_REGION_COVERAGE = 0.25

V1_DEVICE = torch.device(
    "cuda" if torch.cuda.is_available() else "cpu"
)


# ============================================================
# STREAMLIT
# ============================================================

st.set_page_config(
    page_title="Food Region Annotator",
    layout="wide"
)


# ============================================================
# PATH HANDLING
# Same approach as segmentation app
# ============================================================

def resolve_path(value):

    path = Path(str(value))

    if path.exists():
        return path

    path2 = Path.cwd() / path

    if path2.exists():
        return path2

    return path


# ============================================================
# SAFE MEAL ID
# ============================================================

def safe_meal_id(meal_id):

    return (
        str(meal_id)
        .replace("/", "_")
        .replace("\\", "_")
        .replace(" ", "_")
    )


# ============================================================
# NORMALISE TEXT
# Used only for matching spreadsheet food names.
# Original names are preserved for display.
# ============================================================

def normalise_food_name(value):

    if pd.isna(value):
        return ""

    value = str(value).strip().lower()

    # Normalise whitespace
    value = re.sub(
        r"\s+",
        " ",
        value
    )

    return value


# ============================================================
# LOAD PAIRING DATA
# ============================================================

@st.cache_data
def load_pairing_data():

    df = pd.read_excel(
        PAIRING_FILENAME
    )

    required_columns = [
        "Meal_ID",
        "Before_Image",
        "After_Image",
        "Meal Breakdown"
    ]

    missing = [
        column
        for column in required_columns
        if column not in df.columns
    ]

    if missing:
        raise ValueError(
            "meal_image_pairing.xlsx is missing "
            f"required columns: {missing}"
        )

    return df


# ============================================================
# LOAD FOOD BREAKDOWNS
# ============================================================

@st.cache_data
def load_food_breakdowns():

    standard = pd.read_excel(
        FOOD_BREAKDOWNS_FILENAME,
        sheet_name="Standard"
    )

    composite = pd.read_excel(
        FOOD_BREAKDOWNS_FILENAME,
        sheet_name="Composite"
    )

    required_standard = [
        "Food Item",
        "Composite?"
    ]

    required_composite = [
        "Food Item",
        "Component"
    ]

    missing_standard = [
        column
        for column in required_standard
        if column not in standard.columns
    ]

    missing_composite = [
        column
        for column in required_composite
        if column not in composite.columns
    ]

    if missing_standard:
        raise ValueError(
            "Standard sheet is missing columns: "
            f"{missing_standard}"
        )

    if missing_composite:
        raise ValueError(
            "Composite sheet is missing columns: "
            f"{missing_composite}"
        )

    standard = standard.copy()
    composite = composite.copy()

    standard["_match_name"] = (
        standard["Food Item"]
        .apply(normalise_food_name)
    )

    composite["_match_name"] = (
        composite["Food Item"]
        .apply(normalise_food_name)
    )

    return standard, composite


@st.cache_data
def load_after_qc():

    path = Path(AFTER_QC_FILENAME)

    if not path.exists():
        return set()

    qc = pd.read_csv(path)

    required = {"Meal_ID", "QC_Status"}

    if not required.issubset(qc.columns):
        raise ValueError(
            "after_mask_qc.csv must contain "
            "Meal_ID and QC_Status."
        )

    weight_only = qc[
        qc["QC_Status"]
        .astype(str)
        .str.strip()
        .str.upper()
        .eq("WEIGHT_ONLY")
    ].copy()

    return {
        str(meal_id).strip().removesuffix(".0")
        for meal_id in weight_only["Meal_ID"]
        if pd.notna(meal_id)
    }

# ============================================================
# PARSE MEAL BREAKDOWN
# ============================================================

def parse_meal_breakdown(value):

    """
    Converts:

    188g/ml (Food Item) + 20g/ml (Food Item 2)

    into:

    [
        {
            "amount": 188.0,
            "food": "Food Item"
        },
        {
            "amount": 20.0,
            "food": "Food Item 2"
        }
    ]

    Also accepts decimal amounts such as 188.0g/ml.
    """

    if pd.isna(value):
        return []

    text = str(value).strip()

    pattern = re.compile(
        r"""
        ([0-9]+(?:\.[0-9]+)?)      # amount
        \s*
        g/ml
        \s*
        \(
        (.*?)
        \)
        (?=
            \s*\+
            \s*[0-9]+(?:\.[0-9]+)?
            \s*g/ml
            |
            \s*$
        )
        """,
        re.IGNORECASE | re.VERBOSE
    )

    matches = pattern.findall(text)

    result = []

    for amount, food in matches:

        result.append(
            {
                "amount": float(amount),
                "food": food.strip()
            }
        )

    return result


# ============================================================
# INTERPRET COMPOSITE? COLUMN
# ============================================================

def is_composite_value(value):

    if pd.isna(value):
        return False

    if isinstance(value, bool):
        return value

    text = str(value).strip().lower()

    return text in {
        "yes",
        "y",
        "true",
        "1",
        "composite"
    }


# ============================================================
# FIND EXPECTED VISUAL FOODS
# ============================================================

def get_expected_foods(
    parsed_meal,
    standard_df,
    composite_df
):

    """
    For every food item in Meal Breakdown:

    - find it in Standard
    - if non-composite:
        use the Food Item itself
    - if composite:
        retrieve its Component entries

    Returns:
        expected_foods
        lookup_information
    """

    expected_foods = []
    lookup_information = []

    for meal_item in parsed_meal:

        original_food = meal_item["food"]
        amount = meal_item["amount"]

        match_name = normalise_food_name(
            original_food
        )

        standard_matches = standard_df[
            standard_df["_match_name"]
            == match_name
        ]

        # ----------------------------------------------------
        # NOT FOUND IN STANDARD
        # ----------------------------------------------------

        if len(standard_matches) == 0:

            expected_foods.append(
                original_food
            )

            lookup_information.append(
                {
                    "served_food": original_food,
                    "amount": amount,
                    "status": "Not found in Standard",
                    "components": [
                        original_food
                    ]
                }
            )

            continue


        standard_row = (
            standard_matches.iloc[0]
        )

        composite_flag = is_composite_value(
            standard_row["Composite?"]
        )


        # ----------------------------------------------------
        # STANDARD / NON-COMPOSITE FOOD
        # ----------------------------------------------------

        if not composite_flag:

            display_name = str(
                standard_row["Food Item"]
            ).strip()

            expected_foods.append(
                display_name
            )

            lookup_information.append(
                {
                    "served_food": original_food,
                    "amount": amount,
                    "status": "Standard",
                    "components": [
                        display_name
                    ]
                }
            )

            continue


        # ----------------------------------------------------
        # COMPOSITE FOOD
        # ----------------------------------------------------

        composite_matches = composite_df[
            composite_df["_match_name"]
            == match_name
        ]


        if len(composite_matches) == 0:

            # Mark it clearly rather than silently failing.
            expected_foods.append(
                original_food
            )

            lookup_information.append(
                {
                    "served_food": original_food,
                    "amount": amount,
                    "status": (
                        "Marked composite but "
                        "no breakdown found"
                    ),
                    "components": [
                        original_food
                    ]
                }
            )

            continue


        components = []

        for component in (
            composite_matches["Component"]
            .dropna()
            .tolist()
        ):

            component = str(
                component
            ).strip()

            if component:
                components.append(
                    component
                )

                expected_foods.append(
                    component
                )


        lookup_information.append(
            {
                "served_food": original_food,
                "amount": amount,
                "status": "Composite",
                "components": components
            }
        )


    # --------------------------------------------------------
    # REMOVE DUPLICATES BUT PRESERVE ORDER
    # --------------------------------------------------------

    expected_foods = list(
        dict.fromkeys(
            expected_foods
        )
    )

    return (
        expected_foods,
        lookup_information
    )


# ============================================================
# FOOD DISPLAY COLOUR
# ============================================================

def food_colour(food_name):

    """
    Generate the same display colour every time
    for a particular food label.
    """

    seed = sum(
        ord(character)
        for character in str(food_name)
    )

    rng = np.random.default_rng(
        seed
    )

    return rng.integers(
        50,
        230,
        size=3,
        dtype=np.uint8
    )


# ============================================================
# SAVE LABELS
# ============================================================

def save_labels(
    label_path,
    labels
):

    with open(
        label_path,
        "w",
        encoding="utf-8"
    ) as file:

        json.dump(
            labels,
            file,
            indent=4
        )


# ============================================================
# LOAD LABELS
# ============================================================

def load_labels(label_path):

    if not label_path.exists():
        return {}

    try:

        with open(
            label_path,
            "r",
            encoding="utf-8"
        ) as file:

            return json.load(
                file
            )

    except Exception:

        return {}


# ============================================================
# CREATE DISPLAY IMAGE
# ============================================================

def create_display_image(
    rgb,
    segments,
    food_mask,
    labels,
    selected_regions,
    show_boundaries=True,
    highlight_unassigned=True,
    highlighted_foods=None
):

    display = rgb.copy()

    # --------------------------------------------------------
    # ASSIGNED REGIONS
    # --------------------------------------------------------

    for region_string, foods in (
        labels.items()
    ):

        try:
            region_id = int(
                region_string
            )

        except ValueError:
            continue

        pixels = (
            segments == region_id
        )

        if not np.any(pixels):
            continue

        if not foods:
            continue

        colours = np.array(
            [
                food_colour(food)
                for food in foods
            ],
            dtype=np.float32
        )

        colour = colours.mean(
            axis=0
        )

        original = (
            display[pixels]
            .astype(np.float32)
        )

        # Strong colour overlay so completed
        # annotations remain obvious.
        display[pixels] = (
            0.45 * original
            +
            0.55 * colour
        ).astype(np.uint8)

    # --------------------------------------------------------
    # FOOD-SPECIFIC HIGHLIGHT
    # --------------------------------------------------------

    highlighted_foods = set(highlighted_foods or [])

    if highlighted_foods:
        # Dim everything, then make checked-food regions unmistakable.
        display = (0.30 * display.astype(np.float32)).astype(np.uint8)
        highlight_mask = np.zeros(segments.shape, dtype=bool)

        for region_string, foods in labels.items():
            try:
                region_id = int(region_string)
            except ValueError:
                continue

            if foods and highlighted_foods.intersection(foods):
                highlight_mask |= (segments == region_id)

        if np.any(highlight_mask):
            original = rgb[highlight_mask].astype(np.float32)
            cyan = np.array([0, 255, 255], dtype=np.float32)
            display[highlight_mask] = (
                0.35 * original + 0.65 * cyan
            ).astype(np.uint8)

            highlight_boundary = find_boundaries(
                highlight_mask, mode="inner"
            )
            display[highlight_boundary] = [255, 255, 255]

    # --------------------------------------------------------
    # UNASSIGNED REGIONS
    # --------------------------------------------------------

    if highlight_unassigned:

        assigned_region_ids = {
            int(region_id)
            for region_id
            in labels.keys()
        }

        all_region_ids = set(
            int(region_id)
            for region_id
            in np.unique(segments)
            if region_id != 0
        )

        unassigned_region_ids = (
            all_region_ids
            -
            assigned_region_ids
        )

        for region_id in unassigned_region_ids:

            pixels = (
                segments == region_id
            )

            if not np.any(pixels):
                continue

            original = (
                display[pixels]
                .astype(np.float32)
            )

            # Bright magenta
            magenta = np.array(
                [255, 0, 255],
                dtype=np.float32
            )

            display[pixels] = (
                0.15 * original
                +
                0.85 * magenta
            ).astype(np.uint8)

    # --------------------------------------------------------
    # CURRENT SELECTION
    # --------------------------------------------------------

    for region_id in selected_regions:

        pixels = (
            segments == region_id
        )

        if not np.any(pixels):
            continue

        original = (
            display[pixels]
            .astype(np.float32)
        )

        yellow = np.array(
            [255, 255, 0],
            dtype=np.float32
        )

        display[pixels] = (
            0.25 * original
            +
            0.75 * yellow
        ).astype(np.uint8)


    # --------------------------------------------------------
    # SUPERPIXEL BOUNDARIES
    # --------------------------------------------------------

    if show_boundaries:

        boundaries = find_boundaries(
            segments,
            mode="inner"
        )

        boundaries = (
            boundaries
            & food_mask
        )

        display[
            boundaries
        ] = [
            255,
            255,
            255
        ]


    return display

# ============================================================
# SPLIT DISCONNECTED SLIC REGIONS
# ============================================================

def split_disconnected_regions(
    segments,
    food_mask
):
    """
    Ensures that every selectable region is spatially
    connected.

    SLIC can occasionally assign the same label to separate
    pieces of food that are far apart. This function splits
    those disconnected pieces into unique region IDs.

    Background remains 0.
    """

    output = np.zeros(
        segments.shape,
        dtype=np.int32
    )

    next_region_id = 1

    original_region_ids = np.unique(
        segments
    )

    original_region_ids = (
        original_region_ids[
            original_region_ids != 0
        ]
    )

    # 8-connectivity:
    # diagonal touching pixels count as connected.
    structure = np.ones(
        (3, 3),
        dtype=np.uint8
    )

    for region_id in original_region_ids:

        region_mask = (
            segments == region_id
        )

        connected_labels, component_count = (
            ndimage.label(
                region_mask,
                structure=structure
            )
        )

        for component_id in range(
            1,
            component_count + 1
        ):

            component = (
                connected_labels
                == component_id
            )

            if not np.any(component):
                continue

            output[
                component
            ] = next_region_id

            next_region_id += 1

    output[
        ~food_mask
    ] = 0

    return output


# ============================================================
# ABSORB TINY REGIONS
# ============================================================

def absorb_tiny_regions(
    segments,
    food_mask,
    min_pixels=20
):
    """
    Attempts to absorb tiny connected regions into the
    neighbouring region with which they share the greatest
    boundary.

    This prevents tiny, nearly invisible SLIC fragments from
    becoming separate annotation targets.

    If a tiny region has no neighbouring food region, it is
    retained rather than deleted. This is important because
    isolated crumbs/residue may be real food.
    """

    output = segments.copy()

    region_ids, counts = np.unique(
        output[output > 0],
        return_counts=True
    )

    region_sizes = dict(
        zip(
            region_ids.tolist(),
            counts.tolist()
        )
    )

    tiny_regions = [
        int(region_id)
        for region_id, size
        in region_sizes.items()
        if size < min_pixels
    ]

    # Process smallest first.
    tiny_regions.sort(
        key=lambda region_id:
        region_sizes[region_id]
    )

    kernel = np.ones(
        (3, 3),
        dtype=np.uint8
    )

    for region_id in tiny_regions:

        # Region may already have been absorbed by an
        # earlier operation.
        region_mask = (
            output == region_id
        )

        if not np.any(region_mask):
            continue

        # Dilate by one pixel to find directly touching
        # neighbouring regions.
        dilated = cv2.dilate(
            region_mask.astype(np.uint8),
            kernel,
            iterations=1
        ).astype(bool)

        neighbour_pixels = (
            dilated
            &
            ~region_mask
            &
            food_mask
        )

        neighbour_ids = output[
            neighbour_pixels
        ]

        neighbour_ids = neighbour_ids[
            neighbour_ids > 0
        ]

        neighbour_ids = neighbour_ids[
            neighbour_ids != region_id
        ]

        if len(neighbour_ids) == 0:

            # Keep genuinely isolated tiny food regions.
            continue

        ids, neighbour_counts = np.unique(
            neighbour_ids,
            return_counts=True
        )

        # Absorb into the region sharing the largest
        # immediate boundary.
        target_region = int(
            ids[
                np.argmax(
                    neighbour_counts
                )
            ]
        )

        output[
            region_mask
        ] = target_region

    output[
        ~food_mask
    ] = 0

    return output


# ============================================================
# RENUMBER REGION IDs
# ============================================================

def renumber_regions(
    segments
):
    """
    Renumbers region IDs sequentially:
    1, 2, 3, 4...

    Background remains 0.
    """

    output = np.zeros(
        segments.shape,
        dtype=np.int32
    )

    region_ids = np.unique(
        segments
    )

    region_ids = region_ids[
        region_ids != 0
    ]

    for new_id, old_id in enumerate(
        region_ids,
        start=1
    ):

        output[
            segments == old_id
        ] = new_id

    return output


# ============================================================
# CLEAN NEW SLIC REGION MAP
# ============================================================

def clean_slic_regions(
    segments,
    food_mask,
    min_pixels=20
):
    """
    Complete post-processing pipeline for a newly generated
    SLIC map.

    1. Split disconnected components.
    2. Absorb tiny fragments where safe.
    3. Split again in case absorption changes connectivity.
    4. Renumber sequentially.
    """

    cleaned = split_disconnected_regions(
        segments,
        food_mask
    )

    cleaned = absorb_tiny_regions(
        cleaned,
        food_mask,
        min_pixels=min_pixels
    )

    cleaned = split_disconnected_regions(
        cleaned,
        food_mask
    )

    cleaned = renumber_regions(
        cleaned
    )

    cleaned[
        ~food_mask
    ] = 0

    return cleaned

# ============================================================
# GET REGIONS TOUCHED BY BRUSH
# ============================================================

def get_regions_from_canvas(
    canvas_image,
    segments,
    food_mask,
    original_width,
    original_height
):

    if canvas_image is None:
        return set()

    canvas_array = np.asarray(
        canvas_image
    ).astype(np.uint8)


    if (
        canvas_array.ndim != 3
        or canvas_array.shape[2] < 3
    ):
        return set()


    # --------------------------------------------------------
    # IMPORTANT:
    # We use a bright green brush and identify that exact
    # colour rather than relying on the alpha channel.
    #
    # This avoids accidentally interpreting the background
    # image as a brush stroke.
    # --------------------------------------------------------

    red = canvas_array[:, :, 0]
    green = canvas_array[:, :, 1]
    blue = canvas_array[:, :, 2]

    painted = (
        (green > 220)
        &
        (red < 80)
        &
        (blue < 80)
    )


    if not np.any(painted):
        return set()


    # --------------------------------------------------------
    # RETURN BRUSH MASK TO ORIGINAL IMAGE RESOLUTION
    # --------------------------------------------------------

    painted_original = cv2.resize(
        painted.astype(np.uint8),
        (
            original_width,
            original_height
        ),
        interpolation=cv2.INTER_NEAREST
    ).astype(bool)


    # Only existing binary-food pixels matter.
    touched_food = (
        painted_original
        & food_mask
    )


    touched_regions = np.unique(
        segments[
            touched_food
        ]
    )


    touched_regions = (
        touched_regions[
            touched_regions != 0
        ]
    )


    return {
        int(region_id)
        for region_id
        in touched_regions
    }



# ============================================================
# V1 MODEL: LOAD + PREDICT INITIAL SUPERPIXEL LABELS
# ============================================================

@st.cache_resource
def load_v1_model():
    if not V1_MODEL_FILE.exists():
        raise FileNotFoundError(f"Could not find V1 model: {V1_MODEL_FILE}")
    if not V1_CLASSES_FILE.exists():
        raise FileNotFoundError(f"Could not find V1 classes: {V1_CLASSES_FILE}")

    with open(V1_CLASSES_FILE, "r", encoding="utf-8") as f:
        class_data = json.load(f)
    classes = class_data["classes"]

    checkpoint = torch.load(
        V1_MODEL_FILE, map_location=V1_DEVICE, weights_only=False
    )
    model_height = int(checkpoint["height"])
    model_width = int(checkpoint["width"])

    model = deeplabv3_resnet50(weights=None, weights_backbone=None)
    model.classifier[4] = nn.Conv2d(256, len(classes), kernel_size=1)
    model.aux_classifier = None
    model.load_state_dict(checkpoint["model_state_dict"])
    model = model.to(V1_DEVICE)
    model.eval()

    return model, classes, model_height, model_width


def predict_v1_probabilities(rgb):
    model, classes, model_height, model_width = load_v1_model()

    pil_image = Image.fromarray(rgb)
    resized = TF.resize(
        pil_image,
        [model_height, model_width],
        interpolation=TF.InterpolationMode.BILINEAR,
    )
    tensor = TF.to_tensor(resized)
    tensor = TF.normalize(
        tensor,
        mean=[0.485, 0.456, 0.406],
        std=[0.229, 0.224, 0.225],
    )
    tensor = tensor.unsqueeze(0).to(V1_DEVICE)

    with torch.no_grad():
        logits = model(tensor)["out"]
        probabilities = torch.sigmoid(logits)[0].cpu().numpy()

    # Return probabilities at the annotation image resolution so
    # they line up exactly with the frozen SLIC region map.
    output = np.empty(
        (len(classes), rgb.shape[0], rgb.shape[1]), dtype=np.float32
    )
    for c in range(len(classes)):
        output[c] = cv2.resize(
            probabilities[c],
            (rgb.shape[1], rgb.shape[0]),
            interpolation=cv2.INTER_LINEAR,
        )

    return output, classes


def make_v1_region_suggestions(rgb, segments, expected_foods):
    if not np.any(segments > 0):
        return {}

    probabilities, model_classes = predict_v1_probabilities(rgb)

    # Match names case/whitespace-insensitively, but preserve the
    # app's expected-food spelling in the saved annotation.
    expected_lookup = {
        normalise_food_name(food): food
        for food in expected_foods
    }

    allowed = []
    for class_index, model_food in enumerate(model_classes):
        key = normalise_food_name(model_food)
        if key in expected_lookup:
            allowed.append(
                (class_index, expected_lookup[key])
            )

    suggestions = {}

    for region_id in np.unique(segments):
        region_id = int(region_id)
        if region_id == 0:
            continue

        region_pixels = segments == region_id
        region_size = int(np.sum(region_pixels))
        if region_size == 0:
            continue

        foods = []
        for class_index, display_food in allowed:
            predicted = probabilities[class_index] >= V1_THRESHOLD
            coverage = float(np.sum(predicted & region_pixels)) / region_size
            if coverage >= V1_REGION_COVERAGE:
                foods.append(display_food)

        if foods:
            suggestions[str(region_id)] = foods

    return suggestions


# ============================================================
# CHECK INPUT FILES
# ============================================================

if not Path(
    PAIRING_FILENAME
).exists():

    st.error(
        f"Could not find "
        f"'{PAIRING_FILENAME}'."
    )

    st.stop()


if not Path(
    FOOD_BREAKDOWNS_FILENAME
).exists():

    st.error(
        f"Could not find "
        f"'{FOOD_BREAKDOWNS_FILENAME}'."
    )

    st.stop()


# ============================================================
# LOAD SPREADSHEETS
# ============================================================

try:

    pairing_df = (
        load_pairing_data()
    )

    standard_df, composite_df = (
        load_food_breakdowns()
    )

except Exception as error:

    st.error(
        f"Could not load spreadsheet data:\n\n"
        f"{error}"
    )

    st.stop()

weight_only_meals = load_after_qc()

# ============================================================
# AVAILABLE MEALS
# ============================================================

pairing_df = pairing_df[
    pairing_df["Meal_ID"].notna()
].copy()


meal_ids = sorted(
    pairing_df[
        "Meal_ID"
    ].tolist()
)


if len(meal_ids) == 0:

    st.error(
        "No Meal IDs were found."
    )

    st.stop()


# ============================================================
# SESSION STATE
# ============================================================

if "meal_dropdown" not in (
    st.session_state
):

    st.session_state.meal_dropdown = (
        meal_ids[0]
    )


if "image_type" not in (
    st.session_state
):

    st.session_state.image_type = (
        "Before"
    )


if "selected_regions" not in (
    st.session_state
):

    st.session_state.selected_regions = set()


if "canvas_version" not in (
    st.session_state
):

    st.session_state.canvas_version = 0


if "active_annotation" not in (
    st.session_state
):

    st.session_state.active_annotation = None


if "labels" not in (
    st.session_state
):

    st.session_state.labels = {}


if "manual_foods" not in (
    st.session_state
):

    st.session_state.manual_foods = []


if "labels_source" not in st.session_state:
    st.session_state.labels_source = "none"


# ============================================================
# PREVIOUS / NEXT MEAL
# ============================================================

def previous_meal():

    current = (
        st.session_state.meal_dropdown
    )

    current_index = meal_ids.index(
        current
    )

    if current_index > 0:

        st.session_state.meal_dropdown = (
            meal_ids[
                current_index - 1
            ]
        )


def next_meal():

    current = (
        st.session_state.meal_dropdown
    )

    current_index = meal_ids.index(
        current
    )

    if current_index < (
        len(meal_ids) - 1
    ):

        st.session_state.meal_dropdown = (
            meal_ids[
                current_index + 1
            ]
        )

def next_unannotated_meal():

    current = st.session_state.meal_dropdown
    current_index = meal_ids.index(current)

    # Search forward first, then wrap around.
    search_order = (
        meal_ids[current_index + 1:]
        +
        meal_ids[:current_index + 1]
    )

    for meal_id in search_order:

        safe_id = safe_meal_id(meal_id)

        before_label = (
            LABEL_FOLDER
            /
            f"meal_{safe_id}_before_labels.json"
        )

        after_label = (
            LABEL_FOLDER
            /
            f"meal_{safe_id}_after_labels.json"
        )

        meal_key = (
            str(meal_id)
            .strip()
            .removesuffix(".0")
        )

        before_missing = not before_label.exists()

        # A WEIGHT_ONLY after image deliberately does not
        # require an annotation JSON.
        after_missing = (
            not after_label.exists()
            and meal_key not in weight_only_meals
        )

        if before_missing or after_missing:

            st.session_state.meal_dropdown = meal_id

            # Open the image that actually needs annotation.
            if before_missing:
                st.session_state.image_type = "Before"
            else:
                st.session_state.image_type = "After"

            return

# ============================================================
# SIDEBAR NAVIGATION
# ============================================================

st.sidebar.header(
    "Navigation"
)


selected_meal = (
    st.sidebar.selectbox(
        "Select Meal ID",
        meal_ids,
        key="meal_dropdown"
    )
)


navigation1, navigation2 = (
    st.sidebar.columns(2)
)

with navigation1:

    st.button(
        "← Previous",
        on_click=previous_meal,
        use_container_width=True
    )

with navigation2:

    st.button(
        "Next →",
        on_click=next_meal,
        use_container_width=True
    )


st.sidebar.button(
    "⏭ Next unannotated meal",
    on_click=next_unannotated_meal,
    use_container_width=True
)


# ============================================================
# BEFORE / AFTER
# ============================================================

image_type = st.sidebar.radio(
    "Image",
    [
        "Before",
        "After"
    ],
    horizontal=True,
    key="image_type"
)


# ============================================================
# CURRENT ROW
# ============================================================

selected_meal = (
    st.session_state.meal_dropdown
)


row = pairing_df[
    pairing_df["Meal_ID"]
    == selected_meal
].iloc[0]


safe_id = safe_meal_id(
    selected_meal
)

if (
    image_type == "After"
    and str(selected_meal).strip().removesuffix(".0")
        in weight_only_meals
):
    st.warning(
        "This AFTER image is marked WEIGHT_ONLY. "
        "It is excluded from image-based food annotation "
        "because remaining food cannot be assessed reliably."
    )
    st.stop()
# ============================================================
# IMAGE + MASK PATHS
# ============================================================

if image_type == "Before":

    image_column = "Before_Image"

    mask_path = (
        BEFORE_MASK_FOLDER
        /
        f"meal_{safe_id}_mask.png"
    )

    image_suffix = "before"

else:

    image_column = "After_Image"

    mask_path = (
        AFTER_MASK_FOLDER
        /
        f"meal_{safe_id}_mask.png"
    )

    image_suffix = "after"


# ============================================================
# CHECK IMAGE VALUE
# ============================================================

if pd.isna(
    row.get(
        image_column,
        np.nan
    )
):

    st.error(
        f"Meal {selected_meal} has no "
        f"{image_type.lower()} image."
    )

    st.stop()


image_path = resolve_path(
    row[image_column]
)


if not image_path.exists():

    st.error(
        f"Could not find image:\n\n"
        f"{image_path}"
    )

    st.stop()


if not mask_path.exists():

    st.error(
        f"Could not find reviewed binary mask:\n\n"
        f"{mask_path}"
    )

    st.stop()


# ============================================================
# LOAD RGB
# ============================================================

image = Image.open(
    image_path
).convert("RGB")


rgb = np.array(
    image
)


height, width = (
    rgb.shape[:2]
)


# ============================================================
# LOAD BINARY MASK
# ============================================================

mask = cv2.imread(
    str(mask_path),
    cv2.IMREAD_GRAYSCALE
)


if mask is None:

    st.error(
        f"Could not read mask:\n\n"
        f"{mask_path}"
    )

    st.stop()


if mask.shape != (
    height,
    width
):

    st.warning(
        "Mask dimensions did not match RGB. "
        "The mask has been resized using "
        "nearest-neighbour interpolation."
    )

    mask = cv2.resize(
        mask,
        (
            width,
            height
        ),
        interpolation=cv2.INTER_NEAREST
    )


food_mask = (
    mask > 127
)


# ============================================================
# ANNOTATION FILE PATHS
# ============================================================

region_path = (
    REGION_FOLDER
    /
    f"meal_{safe_id}_{image_suffix}_regions.npy"
)


label_path = (
    LABEL_FOLDER
    /
    f"meal_{safe_id}_{image_suffix}_labels.json"
)


# ============================================================
# LOAD OR GENERATE FIXED SLIC REGIONS
# ============================================================

# First check whether the reviewed binary mask contains
# any food at all.
has_food = np.any(food_mask)


if region_path.exists():

    segments = np.load(
        region_path
    )

    # Safety check in case wrong image somehow gets loaded.
    if segments.shape != (
        height,
        width
    ):

        st.error(
            "Saved SLIC region map has the wrong "
            "dimensions for this image.\n\n"
            f"Saved: {segments.shape}\n\n"
            f"Image: {(height, width)}"
        )

        st.stop()


elif not has_food:

    # --------------------------------------------------------
    # EMPTY FOOD MASK
    # --------------------------------------------------------
    # This is a valid situation, especially for AFTER images
    # where the patient has eaten all visible food.
    #
    # Do not run SLIC because there are no food pixels to
    # divide into regions.
    # --------------------------------------------------------

    segments = np.zeros(
        (height, width),
        dtype=np.int32
    )

    # Save the empty region map so we do not need to check
    # or regenerate it again next time.
    np.save(
        region_path,
        segments
    )


else:

    # --------------------------------------------------------
    # NORMAL FOOD MASK
    # --------------------------------------------------------

    image_lab = rgb2lab(
        rgb
    )

    raw_segments = slic(
        image_lab,
        n_segments=N_SEGMENTS,
        compactness=COMPACTNESS,
        sigma=SIGMA,
        mask=food_mask,
        start_label=1,
        channel_axis=-1
    )

    raw_segments[
        ~food_mask
    ] = 0


    # --------------------------------------------------------
    # CLEAN SLIC REGIONS
    # --------------------------------------------------------
    #
    # Important:
    # This happens BEFORE the region map is saved.
    #
    # Existing saved region maps are never passed through
    # this function, so previous annotations remain valid.
    # --------------------------------------------------------

    segments = clean_slic_regions(
        raw_segments,
        food_mask,
        min_pixels=MIN_REGION_PIXELS
    )


    np.save(
        region_path,
        segments
    )


region_ids = np.unique(
    segments
)


region_ids = [
    int(region_id)
    for region_id
    in region_ids
    if region_id != 0
]


# ============================================================
# SWITCH SESSION STATE WHEN IMAGE CHANGES
# ============================================================

annotation_key = (
    f"{safe_id}_{image_suffix}"
)


if (
    st.session_state.active_annotation
    != annotation_key
):

    st.session_state.active_annotation = (
        annotation_key
    )

    # Existing JSONs are always treated as reviewed/saved labels.
    # For a new image we defer V1 suggestions until after the meal
    # breakdown has been parsed below.
    if label_path.exists():
        st.session_state.labels = load_labels(label_path)
        st.session_state.labels_source = "saved"
    else:
        st.session_state.labels = {}
        st.session_state.labels_source = "pending_v1"

    st.session_state.selected_regions = (
        set()
    )

    st.session_state.manual_foods = []

    st.session_state.canvas_version += 1


# ============================================================
# PARSE MEAL BREAKDOWN
# ============================================================

meal_breakdown_raw = row.get(
    "Meal Breakdown",
    ""
)


parsed_meal = parse_meal_breakdown(
    meal_breakdown_raw
)


expected_foods, lookup_information = (
    get_expected_foods(
        parsed_meal,
        standard_df,
        composite_df
    )
)


# Include manually-added foods.
available_foods = list(
    dict.fromkeys(
        expected_foods
        +
        st.session_state.manual_foods
    )
)


# ============================================================
# INITIAL V1 SUGGESTIONS FOR NEW / UNLABELLED IMAGES
# ============================================================

if st.session_state.labels_source == "pending_v1":
    if not has_food:
        st.session_state.labels = {}
        st.session_state.labels_source = "empty"
    else:
        try:
            st.session_state.labels = make_v1_region_suggestions(
                rgb, segments, expected_foods
            )
            st.session_state.labels_source = "v1_suggestion"
        except Exception as error:
            st.session_state.labels = {}
            st.session_state.labels_source = "v1_failed"
            st.warning(
                "V1 initial prediction could not be generated. "
                f"You can still annotate manually. Error: {error}"
            )


if (
    st.session_state.labels_source
    == "v1_suggestion"
):

    if st.button(
        "✓ Prediction is correct — save annotation",
        type="primary"
    ):

        save_labels(
            label_path,
            st.session_state.labels
        )

        st.session_state.labels_source = (
            "saved"
        )

        st.success(
            "Model prediction reviewed and saved."
        )

        st.rerun()

if st.session_state.labels_source == "empty":

    st.info(
        "The reviewed binary mask contains no visible food."
    )

    if st.button(
        "✓ Confirm empty image",
        type="primary"
    ):

        save_labels(
            label_path,
            {}
        )

        st.session_state.labels_source = "saved"

        st.success(
            "Empty image reviewed and saved."
        )

        st.rerun()

# ============================================================
# CURRENT IMAGE COMPLETION
# ============================================================

assigned_count = len(st.session_state.labels)
assigned_region_ids = {int(r) for r in st.session_state.labels.keys()}
assigned_pixel_mask = np.isin(segments, list(assigned_region_ids))
total_food_pixels = int(np.sum(food_mask))
assigned_food_pixels = int(np.sum(assigned_pixel_mask & food_mask))
pixel_completion = (assigned_food_pixels / total_food_pixels) if total_food_pixels > 0 else 1.0
completion = (assigned_count / len(region_ids)) if len(region_ids) > 0 else 1.0
completion = min(completion, 1.0)

# ============================================================
# PAGE TITLE
# ============================================================

st.title(
    "Food Region Annotator"
)

top_progress1, top_progress2 = st.columns(2)
top_progress1.metric("Completed", f"{assigned_count} / {len(region_ids)} superpixels", f"{completion * 100:.1f}%")
top_progress2.metric("Food pixels completed", f"{pixel_completion * 100:.3f}%")
st.progress(completion)


if st.session_state.labels_source == "v1_suggestion":
    st.info(
        "V1 has pre-filled suggested food labels for this previously "
        "unlabelled image. Review and correct the superpixels before "
        "treating this annotation as complete. Suggestions have not "
        "been written to the label JSON just by opening the image."
    )
elif st.session_state.labels_source == "saved":
    st.caption("Loaded existing saved annotation JSON; V1 was not run.")
elif st.session_state.labels_source == "v1_failed":
    st.warning("No V1 suggestions were loaded; annotate this image manually.")


# ============================================================
# MEAL INFORMATION
# ============================================================

information = st.columns(5)


details = [
    (
        "Meal ID",
        row.get(
            "Meal_ID",
            ""
        )
    ),
    (
        "Patient",
        row.get(
            "Patient_ID",
            ""
        )
    ),
    (
        "Meal",
        row.get(
            "Meal",
            ""
        )
    ),
    (
        "Image",
        image_type
    ),
    (
        "Regions",
        len(region_ids)
    )
]


for column, (
    label,
    value
) in zip(
    information,
    details
):

    column.metric(
        label,
        str(value)
    )


# ============================================================
# MEAL BREAKDOWN
# ============================================================

st.divider()

st.subheader(
    "Meal information"
)


if parsed_meal:

    st.write(
        "**Meal Breakdown:**"
    )

    for item in parsed_meal:

        amount = item["amount"]
        food = item["food"]

        st.write(
            f"• {amount:g} g/ml — {food}"
        )

else:

    st.warning(
        "The Meal Breakdown could not be parsed."
    )

    st.code(
        str(meal_breakdown_raw)
    )


# ============================================================
# EXPECTED VISUAL FOODS
# ============================================================

st.write(
    "**Expected annotation labels:**"
)


if expected_foods:

    st.write(
        " • ".join(
            expected_foods
        )
    )

else:

    st.warning(
        "No expected food labels were generated."
    )


# ============================================================
# FOOD BREAKDOWN DETAILS
# ============================================================

with st.expander(
    "Show Food Breakdowns lookup"
):

    if lookup_information:

        for item in lookup_information:

            st.markdown(
                f"**{item['served_food']}** "
                f"({item['amount']:g} g/ml)"
            )

            st.write(
                f"Lookup: {item['status']}"
            )

            st.write(
                "Components: "
                +
                ", ".join(
                    item["components"]
                )
            )

            st.divider()

    else:

        st.caption(
            "No lookup information available."
        )


# ============================================================
# ANNOTATION STATUS
# ============================================================

unassigned_count = max(
    len(region_ids)
    -
    assigned_count,
    0
)


status1, status2, status3, status4 = (
    st.columns(4)
)


status1.metric(
    "Superpixels",
    len(region_ids)
)


status2.metric(
    "Selected",
    len(
        st.session_state.selected_regions
    )
)


status3.metric(
    "Assigned",
    assigned_count
)


status4.metric(
    "Unassigned",
    unassigned_count
)


# ============================================================
# DISPLAY CONTROLS
# ============================================================

control1, control2, control3 = st.columns(
    [1, 1, 2]
)


with control1:

    show_boundaries = st.checkbox(
        "Show superpixel boundaries",
        value=True
    )


with control2:

    highlight_unassigned = st.checkbox(
        "Highlight unassigned",
        value=False
    )


with control3:

    brush_width = st.slider(
        "Selection brush size",
        min_value=4,
        max_value=50,
        value=DEFAULT_BRUSH_WIDTH
    )


# ============================================================
# HIGHLIGHT ASSIGNED FOODS
# ============================================================

st.write("**Highlight assigned foods:**")
highlight_food_options = list(dict.fromkeys(
    available_foods + [food for foods in st.session_state.labels.values() for food in foods]
))
highlighted_foods = []

if highlight_food_options:
    highlight_columns = st.columns(min(4, len(highlight_food_options)))
    for index, food in enumerate(highlight_food_options):
        with highlight_columns[index % len(highlight_columns)]:
            if st.checkbox(
                food,
                value=False,
                key=f"highlight_food_{annotation_key}_{normalise_food_name(food)}"
            ):
                highlighted_foods.append(food)
else:
    st.caption("No assigned/available foods to highlight.")


# ============================================================
# CREATE DISPLAY
# ============================================================

display = create_display_image(
    rgb,
    segments,
    food_mask,
    st.session_state.labels,
    st.session_state.selected_regions,
    show_boundaries,
    highlight_unassigned,
    highlighted_foods
)


display_width = int(
    width * DISPLAY_SCALE
)


display_height = int(
    height * DISPLAY_SCALE
)


display_large = cv2.resize(
    display,
    (
        display_width,
        display_height
    ),
    interpolation=cv2.INTER_NEAREST
)


background_image = Image.fromarray(
    display_large
)


# ============================================================
# ANNOTATION CANVAS
# ============================================================

st.divider()

st.subheader(
    f"Annotate {image_type.lower()} image"
)


st.caption(
    "Brush across one or more superpixels. "
    "When you release the mouse, every food region "
    "touched by the brush is added to the selection."
)


# ============================================================
# ANNOTATION CANVAS + CLEAN REFERENCE IMAGE
# ============================================================

edit_column, reference_column = st.columns(
    [3,2]
)


with edit_column:

    st.markdown("**Annotation image**")

    canvas_result = st_canvas(

        fill_color="#00FF00",

        stroke_color="#00FF00",

        stroke_width=brush_width,

        background_image=background_image,

        drawing_mode="freedraw",

        update_streamlit=True,

        return_image_data=True,

        height=display_height,

        width=display_width,

        key=(
            f"food_canvas_"
            f"{annotation_key}_"
            f"{st.session_state.canvas_version}"
        )
    )


with reference_column:

    st.markdown("**Original image — reference**")

    # Completely untouched RGB image:
    # no food mask, no SLIC boundaries, no annotation colours.
    st.image(
        rgb,
        use_container_width=True
    )


# ============================================================
# PROCESS BRUSH
# ============================================================

if (
    canvas_result.image_data
    is not None
):

    touched_regions = (
        get_regions_from_canvas(
            canvas_result.image_data,
            segments,
            food_mask,
            width,
            height
        )
    )

    new_regions = (
        touched_regions
        -
        st.session_state.selected_regions
    )

    if new_regions:

        st.session_state.selected_regions.update(
            new_regions
        )


# ============================================================
# CURRENT SELECTION
# ============================================================

st.write(
    f"**Selected: "
    f"{len(st.session_state.selected_regions)} "
    f"regions**"
)


if st.session_state.selected_regions:

    st.caption(
        "Region IDs: "
        +
        ", ".join(
            str(region_id)
            for region_id
            in sorted(
                st.session_state.selected_regions
            )
        )
    )


# ============================================================
# SELECTION CONTROLS
# ============================================================

selection1, selection2 = st.columns(
    [1, 4]
)


with selection1:

    if st.button(
        "Clear selection",
        use_container_width=True
    ):

        st.session_state.selected_regions = (
            set()
        )

        st.session_state.canvas_version += 1

        st.rerun()

with selection2:

    if st.button(
        "Select all unannotated",
        use_container_width=True,
        help=(
            "Select every food superpixel that does not currently "
            "have a food assignment."
        )
    ):

        assigned_region_ids = {
            int(region_id)
            for region_id
            in st.session_state.labels.keys()
        }

        unannotated_region_ids = {
            int(region_id)
            for region_id
            in region_ids
            if int(region_id) not in assigned_region_ids
        }

        st.session_state.selected_regions = (
            unannotated_region_ids
        )

        st.session_state.canvas_version += 1

        st.rerun()

# ============================================================
# FOOD ASSIGNMENT
# ============================================================

st.divider()

st.subheader(
    "Assign food"
)


selected_foods = st.multiselect(
    "Food(s) in selected regions",
    options=available_foods,
    help=(
        "Select more than one food when the region "
        "contains a mixture, for example "
        "Sponge + Custard."
    )
)


# ============================================================
# ADD UNEXPECTED FOOD
# ============================================================

with st.expander(
    "Add unexpected food"
):

    st.write(
        "Use this only when a food is visibly present "
        "but is not included in the expected list."
    )

    new_food = st.text_input(
        "Food name",
        key="new_manual_food"
    )

    if st.button(
        "Add food to this image"
    ):

        cleaned_food = (
            new_food.strip()
        )

        if cleaned_food:

            if cleaned_food not in (
                st.session_state.manual_foods
            ):

                st.session_state.manual_foods.append(
                    cleaned_food
                )

            st.rerun()


# ============================================================
# ASSIGN / REMOVE
# ============================================================

assign1, assign2, assign3 = (
    st.columns(
        [1, 1, 2]
    )
)


with assign1:

    if st.button(
        "Assign food",
        type="primary",
        use_container_width=True
    ):

        if not (
            st.session_state.selected_regions
        ):

            st.warning(
                "Select at least one region first."
            )

        elif not selected_foods:

            st.warning(
                "Select at least one food."
            )

        else:

            for region_id in (
                st.session_state.selected_regions
            ):

                st.session_state.labels[
                    str(region_id)
                ] = list(
                    selected_foods
                )

            save_labels(
                label_path,
                st.session_state.labels
            )
            st.session_state.labels_source = "saved"

            st.session_state.selected_regions = (
                set()
            )

            st.session_state.canvas_version += 1

            st.rerun()


with assign2:

    if st.button(
        "Remove assignment",
        use_container_width=True
    ):

        for region_id in (
            st.session_state.selected_regions
        ):

            st.session_state.labels.pop(
                str(region_id),
                None
            )

        save_labels(
            label_path,
            st.session_state.labels
        )
        st.session_state.labels_source = "saved"

        st.session_state.selected_regions = (
            set()
        )

        st.session_state.canvas_version += 1

        st.rerun()


# ============================================================
# ANNOTATION SUMMARY
# ============================================================

st.divider()

st.subheader(
    "Annotation summary"
)


summary = {}


for foods in (
    st.session_state.labels.values()
):

    if not foods:
        continue

    label = " + ".join(
        sorted(
            foods
        )
    )

    summary[label] = (
        summary.get(
            label,
            0
        )
        + 1
    )


if summary:

    for label, count in sorted(
        summary.items()
    ):

        foods = label.split(
            " + "
        )

        colours = np.array(
            [
                food_colour(food)
                for food in foods
            ],
            dtype=np.float32
        )

        colour = (
            colours
            .mean(axis=0)
            .astype(int)
        )

        red, green, blue = (
            colour
        )

        st.markdown(
            f"""
            <div style="
                display:flex;
                align-items:center;
                gap:10px;
                margin-bottom:7px;
            ">

                <div style="
                    width:24px;
                    height:24px;
                    background:
                        rgb({red},{green},{blue});
                    border:
                        1px solid #888;
                    border-radius:4px;
                ">
                </div>

                <div>
                    <b>{label}</b>
                    — {count} regions
                </div>

            </div>
            """,
            unsafe_allow_html=True
        )

else:

    st.caption(
        "No food regions assigned yet."
    )

# ============================================================
# COMPLETION SUMMARY
# ============================================================

st.progress(completion)
st.write(f"**{assigned_count} / {len(region_ids)}** superpixels assigned ({completion * 100:.1f}%)")
st.write(f"**Food pixels assigned: {pixel_completion * 100:.3f}%**")


# ============================================================
# IMAGE COMPLETION STATUS
# ============================================================

if len(region_ids) == 0:

    st.success(
        f"{image_type} image contains no visible food. "
        "No food-region annotation is required."
    )

elif assigned_count >= len(region_ids):

    st.success(
        f"{image_type} image fully annotated."
    )


# ============================================================
# SAVED FILES
# ============================================================

with st.expander(
    "Saved annotation files"
):

    st.write(
        f"Image: `{image_path}`"
    )

    st.write(
        f"Binary mask: `{mask_path}`"
    )

    st.write(
        f"SLIC regions: `{region_path}`"
    )

    st.write(
        f"Food labels: `{label_path}`"
    )