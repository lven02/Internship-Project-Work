import streamlit as st
import pandas as pd
import numpy as np
from pathlib import Path
from PIL import Image
import cv2
from streamlit_drawable_canvas import st_canvas
import plotly.graph_objects as go
import torch
from torch import nn

from torchvision.transforms import functional as TF
from torchvision.transforms import InterpolationMode

from torchvision.models.segmentation import (
    deeplabv3_resnet50,
    DeepLabV3_ResNet50_Weights
)


# ============================================================
# SETTINGS - CURRENTLY SET UP FOR POST MEAL IMAGES
# ============================================================

PAIRING_FILENAME = "meal_image_pairing.xlsx"
MASTER_FOLDER = Path("Cleaned_Raw_Image_Data_2026")
MASK_FOLDER = Path("after_segmentation_masks")

MASK_FOLDER.mkdir(
    parents=True,
    exist_ok=True
)

# Your images are 848 x 480, so this keeps the canvas manageable.
DISPLAY_WIDTH = 848

MODEL_FILE = Path(
    r"C:\INSERT FILE NAME HERE\after_segmentation_dataset\after_food_segmentation_model_v5.pth"
)

MODEL_IMAGE_HEIGHT = 256
MODEL_IMAGE_WIDTH = 448

MODEL_THRESHOLD = 0.5


# ============================================================
# STREAMLIT SETTINGS
# ============================================================

st.set_page_config(
    page_title="Food Segmentation Annotation Tool",
    layout="wide"
)


# ============================================================
# LOAD SPREADSHEET
# ============================================================

@st.cache_data
def load_data():

    df = pd.read_excel(PAIRING_FILENAME)

    required_columns = [
        "Meal_ID",
        "After_Image"
    ]

    missing = [
        column
        for column in required_columns
        if column not in df.columns
    ]

    if missing:
        raise ValueError(
            f"Missing required columns: {missing}"
        )

    df = df[
        df["After_Image"].notna()
    ].copy()

    return df

@st.cache_resource
def load_segmentation_model():

    device = torch.device(
        "cuda"
        if torch.cuda.is_available()
        else "cpu"
    )

    weights = (
        DeepLabV3_ResNet50_Weights.DEFAULT
    )

    model = deeplabv3_resnet50(
        weights=weights
    )

    model.classifier[4] = nn.Conv2d(
        256,
        1,
        kernel_size=1
    )

    if model.aux_classifier is not None:

        model.aux_classifier[4] = nn.Conv2d(
            256,
            1,
            kernel_size=1
        )

    model.load_state_dict(
        torch.load(
            MODEL_FILE,
            map_location=device
        )
    )

    model = model.to(
        device
    )

    model.eval()

    return model, device

# ============================================================
# PATH HANDLING
# ============================================================

def resolve_path(value):

    path = Path(str(value))

    # Path works exactly as stored
    if path.exists():
        return path

    # Try relative to the Python script
    path2 = Path.cwd() / path

    if path2.exists():
        return path2

    return path

# ============================================================
# ML PREDICTION HANDLING
# ============================================================

def predict_food_mask(
    image,
    model,
    device,
    threshold=0.5
):

    original_width, original_height = (
        image.size
    )

    # --------------------------------------------------------
    # Resize to model input size
    # --------------------------------------------------------

    resized = TF.resize(
        image,
        [
            MODEL_IMAGE_HEIGHT,
            MODEL_IMAGE_WIDTH
        ],
        interpolation=InterpolationMode.BILINEAR
    )

    # --------------------------------------------------------
    # Convert to tensor
    # --------------------------------------------------------

    tensor = TF.to_tensor(
        resized
    )

    tensor = TF.normalize(
        tensor,
        mean=[
            0.485,
            0.456,
            0.406
        ],
        std=[
            0.229,
            0.224,
            0.225
        ]
    )

    tensor = (
        tensor
        .unsqueeze(0)
        .to(device)
    )

    # --------------------------------------------------------
    # Predict
    # --------------------------------------------------------

    with torch.no_grad():

        output = model(
            tensor
        )

        logits = output[
            "out"
        ]

        probabilities = torch.sigmoid(
            logits
        )

    probability_map = (
        probabilities
        .squeeze()
        .cpu()
        .numpy()
    )

    # --------------------------------------------------------
    # Threshold
    # --------------------------------------------------------

    predicted_mask = (
        probability_map >= threshold
    ).astype(
        np.uint8
    ) * 255

    # --------------------------------------------------------
    # Resize back to original image size
    # --------------------------------------------------------

    predicted_mask = cv2.resize(
        predicted_mask,
        (
            original_width,
            original_height
        ),
        interpolation=cv2.INTER_NEAREST
    )

    return predicted_mask

# ============================================================
# LOAD MODEL
# ============================================================

model, device = load_segmentation_model()
# ============================================================
# AUTOMATIC MASKING
# ============================================================

def automatic_mask(rgb, depth):

    """
    Creates an initial food mask using:

    1. Fixed tray region
    2. Blue plate detection for main meal food
    3. Depth-based detection for desserts/extras
       in the top-right region
    4. QR-code exclusion in bottom-right

    Returns:
        final_mask
        debug_data
    """

    image = rgb.copy()

    height, width = image.shape[:2]

    # ========================================================
    # 1. BASIC CHECKS
    # ========================================================

    if depth is None:

        depth = np.zeros(
            (height, width),
            dtype=np.float32
        )

    # Make sure depth matches RGB dimensions
    if depth.shape[:2] != (height, width):

        depth = cv2.resize(
            depth.astype(np.float32),
            (width, height),
            interpolation=cv2.INTER_NEAREST
        )

    depth = depth.astype(np.float32)


    # ========================================================
    # 2. HSV VERSION OF RGB IMAGE
    # ========================================================

    hsv = cv2.cvtColor(
        image,
        cv2.COLOR_RGB2HSV
    )


    # ========================================================
    # 3. FIXED TRAY REGION
    # ========================================================

    # Because the camera setup is controlled,
    # use a fixed region instead of detecting the red tray.

    tray_mask = np.zeros(
        (height, width),
        dtype=np.uint8
    )

    tray_x1 = int(width * 0.26)
    tray_x2 = int(width * 0.88)

    tray_y1 = int(height * 0.20)
    tray_y2 = int(height * 0.99)

    tray_mask[
        tray_y1:tray_y2,
        tray_x1:tray_x2
    ] = 255


    # ========================================================
    # 4. DETECT BLUE PIXELS
    # ========================================================

    lower_blue = np.array(
        [85, 40, 40]
    )

    upper_blue = np.array(
        [140, 255, 255]
    )

    blue_pixels = cv2.inRange(
        hsv,
        lower_blue,
        upper_blue
    )

    blue_pixels = cv2.morphologyEx(
        blue_pixels,
        cv2.MORPH_CLOSE,
        np.ones((7, 7), np.uint8)
    )

    blue_pixels = cv2.morphologyEx(
        blue_pixels,
        cv2.MORPH_OPEN,
        np.ones((5, 5), np.uint8)
    )


    # ========================================================
    # 5. FIND MAIN BLUE PLATE
    # ========================================================

    plate_contours, _ = cv2.findContours(
        blue_pixels,
        cv2.RETR_EXTERNAL,
        cv2.CHAIN_APPROX_SIMPLE
    )

    plate_mask = np.zeros(
        (height, width),
        dtype=np.uint8
    )

    plate_contour = None

    if plate_contours:

        # Prefer a large contour on the left side
        possible_plates = []

        for contour in plate_contours:

            area = cv2.contourArea(
                contour
            )

            if area < 1000:
                continue

            x, y, w, h = cv2.boundingRect(
                contour
            )

            centre_x = x + w / 2

            # Main plate is normally on left half
            if centre_x < width * 0.65:

                possible_plates.append(
                    contour
                )

        if possible_plates:

            plate_contour = max(
                possible_plates,
                key=cv2.contourArea
            )

        else:

            plate_contour = max(
                plate_contours,
                key=cv2.contourArea
            )


    if plate_contour is not None:

        cv2.drawContours(
            plate_mask,
            [plate_contour],
            -1,
            255,
            thickness=cv2.FILLED
        )


    # ========================================================
    # 6. CREATE INNER PLATE REGION
    # ========================================================

    # Erode to remove the rim/ring.

    inner_plate_mask = cv2.erode(
        plate_mask,
        np.ones((25, 25), np.uint8),
        iterations=1
    )


    # ========================================================
    # 7. FIND MAIN FOOD ON PLATE
    # ========================================================

    non_blue = cv2.bitwise_not(
        blue_pixels
    )

    main_food_candidates = cv2.bitwise_and(
        non_blue,
        inner_plate_mask
    )


    # Remove extreme darkness / glare
    value = hsv[:, :, 2]

    valid_food_pixels = (
        (value > 25)
        &
        (value < 250)
    )

    main_food_candidates[
        ~valid_food_pixels
    ] = 0


    # Clean small noise
    main_food_candidates = cv2.morphologyEx(
        main_food_candidates,
        cv2.MORPH_OPEN,
        np.ones((5, 5), np.uint8)
    )

    main_food_candidates = cv2.morphologyEx(
        main_food_candidates,
        cv2.MORPH_CLOSE,
        np.ones((9, 9), np.uint8)
    )


    # ========================================================
    # 8. REMOVE SMALL MAIN-FOOD COMPONENTS
    # ========================================================

    num_labels, labels, stats, _ = (
        cv2.connectedComponentsWithStats(
            main_food_candidates,
            connectivity=8
        )
    )

    main_food_mask = np.zeros_like(
        main_food_candidates
    )

    for label in range(
        1,
        num_labels
    ):

        area = stats[
            label,
            cv2.CC_STAT_AREA
        ]

        if area >= 250:

            main_food_mask[
                labels == label
            ] = 255


    # ========================================================
    # 9. DESSERT / EXTRAS SEARCH REGION
    # ========================================================

    dessert_zone = tray_mask.copy()

    # ========================================================
    # 10. EXCLUDE MAIN PLATE
    # ========================================================

    plate_exclusion = cv2.dilate(
        plate_mask,
        np.ones((31, 31), np.uint8),
        iterations=1
    )

    dessert_zone[
        plate_exclusion > 0
    ] = 0


    # ========================================================
    # 11. EXCLUDE QR CODE REGION
    # ========================================================

    qr_x1 = int(
        width * 0.68
    )

    qr_y1 = int(
        height * 0.62
    )

    dessert_zone[
        qr_y1:height,
        qr_x1:width
    ] = 0


    # ========================================================
    # ESTIMATE TRAY DEPTH FROM A FIXED EMPTY PATCH
    # ========================================================

    # Adjust these percentages so this rectangle sits on
    # consistently empty red tray.

    sample_x1 = int(width * 0.76)
    sample_x2 = int(width * 0.83)

    sample_y1 = int(height * 0.78)
    sample_y2 = int(height * 0.90)


    tray_sample_mask = np.zeros(
        (height, width),
        dtype=np.uint8
    )

    tray_sample_mask[
        sample_y1:sample_y2,
        sample_x1:sample_x2
    ] = 255


    valid_sample = (
        (depth > 0)
        &
        np.isfinite(depth)
        &
        (tray_sample_mask > 0)
    )


    tray_depth = None

    if np.any(valid_sample):

        tray_depth_values = depth[
            valid_sample
        ]

        tray_depth = np.median(
            tray_depth_values
        )


        # ====================================================
        # 13. FIND OBJECTS RAISED ABOVE TRAY
        # ====================================================

        # Assuming SMALLER depth means closer to camera.
        #
        # Start with 12 mm threshold.
        # Tune this if needed.

        # Smaller depth = closer to camera.

        strong_threshold = 12.0
        weak_threshold = 5.0

        strong_raised = (
            (depth > 0)
            &
            np.isfinite(depth)
            &
            (depth < tray_depth - strong_threshold)
            &
            (dessert_zone > 0)
        )

        weak_raised = (
            (depth > 0)
            &
            np.isfinite(depth)
            &
            (depth < tray_depth - weak_threshold)
            &
            (dessert_zone > 0)
        )


        strong_mask = np.zeros(
            (height, width),
            dtype=np.uint8
        )

        strong_mask[
            strong_raised
        ] = 255


        weak_mask = np.zeros(
            (height, width),
            dtype=np.uint8
        )

        weak_mask[
            weak_raised
        ] = 255

        # ========================================================
        # GROW STRONG OBJECTS INTO WEAKER DEPTH REGIONS
        # ========================================================

        strong_expanded = cv2.dilate(
            strong_mask,
            np.ones((15, 15), np.uint8),
            iterations=2
        )

        connected_weak = cv2.bitwise_and(
            weak_mask,
            strong_expanded
        )

        dessert_depth_mask = cv2.bitwise_or(
            strong_mask,
            connected_weak
        )


        # ====================================================
        # 14. CLEAN DEPTH-BASED DESSERT MASK
        # ====================================================

        dessert_depth_mask = cv2.morphologyEx(
            dessert_depth_mask,
            cv2.MORPH_CLOSE,
            np.ones((19, 19), np.uint8),
            iterations=2
        )

        dessert_depth_mask = cv2.morphologyEx(
            dessert_depth_mask,
            cv2.MORPH_OPEN,
            np.ones((3, 3), np.uint8)
        )


    # ========================================================
    # 15. KEEP REASONABLE DESSERT OBJECTS
    # ========================================================

    num_labels, labels, stats, _ = (
        cv2.connectedComponentsWithStats(
            dessert_depth_mask,
            connectivity=8
        )
    )

    dessert_mask = np.zeros_like(
        dessert_depth_mask
    )


    for label in range(
        1,
        num_labels
    ):

        area = stats[
            label,
            cv2.CC_STAT_AREA
        ]

        object_width = stats[
            label,
            cv2.CC_STAT_WIDTH
        ]

        object_height = stats[
            label,
            cv2.CC_STAT_HEIGHT
        ]


        # Ignore very small noise
        if area < 150:
            continue


        # Ignore huge accidental regions
        if area > 50000:
            continue


        aspect_ratio = (
            max(
                object_width,
                object_height
            )
            /
            max(
                min(
                    object_width,
                    object_height
                ),
                1
            )
        )


        # Reject long thin objects
        if aspect_ratio > 6:
            continue


        dessert_mask[
            labels == label
        ] = 255


    # ========================================================
    # 16. SLIGHTLY EXPAND DESSERT OBJECTS
    # ========================================================

    dessert_mask = cv2.dilate(
        dessert_mask,
        np.ones((7, 7), np.uint8),
        iterations=1
    )


    # ========================================================
    # 17. COMBINE MAIN FOOD + DESSERTS
    # ========================================================

    final_mask = cv2.bitwise_or(
        main_food_mask,
        dessert_mask
    )


    # Nothing outside tray region allowed
    final_mask[
        tray_mask == 0
    ] = 0


    # ========================================================
    # 18. DEBUG DATA
    # ========================================================

    debug_data = {

        "tray_mask":
            tray_mask,

        "blue_pixels":
            blue_pixels,

        "plate_mask":
            plate_mask,

        "inner_plate":
            inner_plate_mask,

        "main_food":
            main_food_mask,

        "dessert_zone":
            dessert_zone,

        "dessert_depth_mask":
            dessert_depth_mask,

        "dessert_mask":
            dessert_mask,

        "tray_depth":
            tray_depth
    }


    return final_mask, debug_data


# ============================================================
# OVERLAY
# ============================================================

def create_overlay(
    rgb_image,
    mask
):

    overlay = rgb_image.copy()

    food_pixels = mask > 127

    # Red overlay over predicted food.
    overlay[food_pixels] = (
        overlay[food_pixels] * 0.55
        +
        np.array([255, 0, 0]) * 0.45
    ).astype(np.uint8)

    return overlay



# ============================================================
# DEPTH
# ============================================================

def load_depth(row):

    if "After_Depth" not in row.index:
        return None

    if pd.isna(row["After_Depth"]):
        return None

    depth_path = resolve_path(
        row["After_Depth"]
    )

    if not depth_path.exists():
        return None

    try:

        return np.load(depth_path)

    except Exception:

        return None


# ============================================================
# CHECK FILES
# ============================================================

if not Path(PAIRING_FILENAME).exists():

    st.error(
        f"Could not find '{PAIRING_FILENAME}'."
    )

    st.stop()


if not MASTER_FOLDER.exists():

    st.error(
        f"Could not find '{MASTER_FOLDER}'."
    )

    st.stop()


MASK_FOLDER.mkdir(
    exist_ok=True
)


# ============================================================
# LOAD DATA
# ============================================================

try:

    df = load_data()

except Exception as error:

    st.error(
        f"Could not read spreadsheet: {error}"
    )

    st.stop()


if len(df) == 0:

    st.error(
        "No meals containing After_Image were found."
    )

    st.stop()


# ============================================================
# MEAL SELECTION
# ============================================================

meal_ids = sorted(df["Meal_ID"].tolist())


# Initialise dropdown on first run
if "meal_dropdown" not in st.session_state:
    st.session_state.meal_dropdown = meal_ids[0]


# ============================================================
# PREVIOUS / NEXT FUNCTIONS
# ============================================================

def previous_meal():

    current_index = meal_ids.index(
        st.session_state.meal_dropdown
    )

    if current_index > 0:
        st.session_state.meal_dropdown = (
            meal_ids[current_index - 1]
        )


def next_meal():

    current_index = meal_ids.index(
        st.session_state.meal_dropdown
    )

    if current_index < len(meal_ids) - 1:
        st.session_state.meal_dropdown = (
            meal_ids[current_index + 1]
        )


# ============================================================
# DROPDOWN
# ============================================================

selected_meal = st.sidebar.selectbox(
    "Select Meal ID",
    meal_ids,
    key="meal_dropdown"
)


# ============================================================
# BUTTONS
# ============================================================

col1, col2 = st.columns(2)

with col1:

    st.button(
        "← Prev.",
        on_click=previous_meal,
        use_container_width=True
    )


with col2:

    st.button(
        "Next →",
        on_click=next_meal,
        use_container_width=True
    )


# Current meal used by rest of app
selected_meal = st.session_state.meal_dropdown

# ============================================================
# GET ROW
# ============================================================

row = df[
    df["Meal_ID"] == selected_meal
].iloc[0]


# ============================================================
# LOAD IMAGE
# ============================================================

image_path = resolve_path(
    row["After_Image"]
)


if not image_path.exists():

    st.error(
        f"Could not find image:\n{image_path}"
    )

    st.stop()


image = Image.open(
    image_path
).convert("RGB")


rgb = np.array(
    image
)


height, width = rgb.shape[:2]

# ============================================================
# LOAD CORRESPONDING DEPTH DATA
# ============================================================

depth_path = resolve_path(
    row["After_Depth"]
)

if not depth_path.exists():

    st.error(
        f"Could not find depth file:\n{depth_path}"
    )

    st.stop()


depth = np.load(
    depth_path
).astype(np.float32)


# Check that RGB and depth dimensions match

if depth.shape != (height, width):

    st.warning(
        f"Depth shape {depth.shape} does not match "
        f"RGB shape {(height, width)}."
    )

    depth = cv2.resize(
        depth,
        (width, height),
        interpolation=cv2.INTER_NEAREST
    )

# ============================================================
# CREATE AUTOMATIC MASK WHEN MEAL CHANGES
# ============================================================

if st.session_state.get("active_meal") != selected_meal:

    st.session_state.active_meal = selected_meal

    # --------------------------------------------------------
    # MODEL PREDICTION
    # --------------------------------------------------------

    predicted_mask = predict_food_mask(
        image,
        model,
        device
    )

    st.session_state.auto_mask = (
        predicted_mask.copy()
    )

    st.session_state.working_mask = (
        predicted_mask.copy()
    )

    st.session_state.preview_mask = (
        predicted_mask.copy()
    )


# ============================================================
# TITLE
# ============================================================

st.title(
    "Food Segmentation Annotation Tool"
)


# ============================================================
# MEAL INFORMATION
# ============================================================

information = st.columns(5)

details = [
    (
        "Meal ID",
        row.get("Meal_ID", "")
    ),
    (
        "Patient",
        row.get("Patient_ID", "")
    ),
    (
        "Date",
        row.get("After_Time", "")
    ),
    (
        "Meal",
        row.get("Meal", "")
    ),
    (
        "After weight",
        row.get("After_Weight", "")
    )
]


for column, (label, value) in zip(
    information,
    details
):

    column.metric(
        label,
        str(value)
    )


# ============================================================
# AUTOMATIC MASK
# ============================================================

st.divider()

st.subheader(
    "1. Automatic initial segmentation"
)

st.write(
    """
The computer has produced an initial estimate of where the food is.

**Red = currently classified as food.**

This is only a starting point. You should check and correct it.
"""
)


preview1, preview2, preview3 = st.columns(3)


with preview1:

    st.image(
        rgb,
        caption="Original After image",
        use_container_width=True
    )


with preview2:

    st.image(
        st.session_state.auto_mask,
        caption="Model Prediciton",
        use_container_width=True
    )


with preview3:

    st.image(
        create_overlay(
            rgb,
            st.session_state.working_mask
        ),
        caption="Current overlay",
        use_container_width=True
    )

# ============================================================
# QUICK SAVE AUTOMATIC MASK
# ============================================================

if st.button(
    "✓ Save Automatic Mask",
    use_container_width=True,
    type="primary"
):

    mask_path = (
        MASK_FOLDER
        / f"meal_{selected_meal}_mask.png"
    )

    # Make sure mask is uint8 0/255
    mask_to_save = (
        st.session_state.working_mask > 0
    ).astype(np.uint8) * 255

    cv2.imwrite(
        str(mask_path),
        mask_to_save
    )

    st.success(
        f"Automatic mask saved for Meal {selected_meal}"
    )
# ============================================================
# EDIT MASK
# ============================================================

st.divider()

st.subheader(
    "2. Correct the segmentation"
)

st.write(
    """
Use the canvas below to correct the mask.

- **Add food** = paint white
- **Erase food** = paint black
"""
)


tool1, tool2 = st.columns(2)


with tool1:

    brush_mode = st.radio(
        "Brush mode",
        [
            "Add food",
            "Erase food"
        ],
        horizontal=True
    )


with tool2:

    brush_size = st.slider(
        "Brush size",
        min_value=2,
        max_value=80,
        value=20
    )


if st.button(
    "Reset to model prediciton"
):

    st.session_state.working_mask = (
        st.session_state.auto_mask.copy()
    )

    st.rerun()


# ============================================================
# DRAWING CANVAS
# ============================================================

scale = min(
    DISPLAY_WIDTH / width,
    1.0
)

canvas_width = int(
    width * scale
)

canvas_height = int(
    height * scale
)


# ============================================================
# CREATE RGB + MASK OVERLAY FOR EDITING
# ============================================================

editor_background = create_overlay(
    rgb,
    st.session_state.working_mask
)

editor_background = cv2.resize(
    editor_background,
    (
        canvas_width,
        canvas_height
    ),
    interpolation=cv2.INTER_LINEAR
)


# ============================================================
# CHOOSE A UNIQUE BRUSH COLOUR
# ============================================================

# These colours are deliberately very bright so that
# we can identify the user's brush strokes afterwards.

if brush_mode == "Add food":

    brush_colour = "#00FF00"   # bright green

else:

    brush_colour = "#FF00FF"   # bright magenta


# ============================================================
# DRAW DIRECTLY OVER IMAGE + MASK
# ============================================================

canvas = st_canvas(
    fill_color=brush_colour,
    stroke_color=brush_colour,
    stroke_width=brush_size,

    background_image=Image.fromarray(
        editor_background
    ),

    drawing_mode="freedraw",
    update_streamlit=True,

    height=canvas_height,
    width=canvas_width,

    return_image_data=True,

    # Changing mode gives us a clean drawing layer
    key=f"canvas_{selected_meal}_{brush_mode}"
)


# ============================================================
# APPLY USER EDITS
# ============================================================

if canvas.image_data is not None:

    canvas_data = (
        canvas.image_data
        .astype(np.uint8)
    )


    # ========================================================
    # FIND ONLY THE NEW BRUSH STROKES
    # ========================================================

    if brush_mode == "Add food":

        # Bright green pixels
        painted_pixels = (
            (canvas_data[:, :, 1] > 240)
            &
            (canvas_data[:, :, 0] < 30)
            &
            (canvas_data[:, :, 2] < 30)
        )

    else:

        # Bright magenta pixels
        painted_pixels = (
            (canvas_data[:, :, 0] > 240)
            &
            (canvas_data[:, :, 1] < 30)
            &
            (canvas_data[:, :, 2] > 240)
        )


    # Current mask at canvas resolution
    edited_mask = cv2.resize(
        st.session_state.working_mask,
        (
            canvas_width,
            canvas_height
        ),
        interpolation=cv2.INTER_NEAREST
    )


    # ========================================================
    # APPLY STROKES TO MASK
    # ========================================================

    if np.any(painted_pixels):

        if brush_mode == "Add food":

            edited_mask[
                painted_pixels
            ] = 255

        else:

            edited_mask[
                painted_pixels
            ] = 0


        # Return to original image resolution
        st.session_state.preview_mask = cv2.resize(
            edited_mask,
            (
                width,
                height
            ),
            interpolation=cv2.INTER_NEAREST
        )

    else:

        st.session_state.preview_mask = (
            st.session_state.working_mask.copy()
        )

else:

    st.session_state.preview_mask = (
        st.session_state.working_mask.copy()
    )


# ============================================================
# COMMIT EDIT
# ============================================================

if st.button(
    "Apply brush strokes"
):

    st.session_state.working_mask = (
        st.session_state.preview_mask.copy()
    )

    st.rerun()


# ============================================================
# REVIEW
# ============================================================

st.divider()

st.subheader(
    "3. Review your corrected mask"
)


review1, review2 = st.columns(2)


with review1:

    st.image(
        st.session_state.working_mask,
        caption="Final binary mask",
        use_container_width=True
    )


with review2:

    st.image(
        create_overlay(
            rgb,
            st.session_state.working_mask
        ),
        caption="Final mask over original image",
        use_container_width=True
    )


# ============================================================
# SAVE
# ============================================================

if st.button(
    "Save corrected mask",
    type="primary"
):

    safe_id = (
        str(selected_meal)
        .replace("/", "_")
        .replace("\\", "_")
    )

    save_path = (
        MASK_FOLDER
        /
        f"meal_{safe_id}_mask.png"
    )

    final_mask = np.where(
        st.session_state.working_mask > 127,
        255,
        0
    ).astype(np.uint8)


    Image.fromarray(
        final_mask
    ).save(
        save_path
    )


    st.success(
        f"Mask saved successfully:\n{save_path}"
    )



# ============================================================
# DEPTH DATA
# ============================================================

st.divider()

# ============================================================
# INTERACTIVE DEPTH MAP
# ============================================================

st.subheader("RGB and depth inspection")

col_rgb, col_depth = st.columns(2)

with col_rgb:

    st.image(
        rgb,
        caption="RGB image",
        use_container_width=True
    )

with col_depth:

    depth_display = depth.astype(
        np.float32
    ).copy()

    depth_display[
        depth_display <= 0
    ] = np.nan

    fig = go.Figure(
        data=go.Heatmap(
            z=depth_display,
            colorscale="Viridis",

            hovertemplate=(
                "X: %{x}<br>"
                "Y: %{y}<br>"
                "Depth: %{z:.1f}"
                "<extra></extra>"
            ),

            colorbar=dict(
                title="Depth"
            )
        )
    )

    fig.update_layout(
        height=400,

        margin=dict(
            l=0,
            r=0,
            t=0,
            b=0
        )
    )

    fig.update_yaxes(
        autorange="reversed",
        scaleanchor="x",
        scaleratio=1
    )

    st.plotly_chart(
        fig,
        use_container_width=True
    )


# ============================================================
# PROGRESS
# ============================================================

st.divider()

saved_masks = list(
    MASK_FOLDER.glob(
        "meal_*_mask.png"
    )
)


st.subheader(
    "Annotation progress"
)


total_meals = len(df)

completed = len(
    saved_masks
)


st.progress(
    min(
        completed / max(total_meals, 1),
        1.0
    )
)


st.write(
    f"**{completed} / {total_meals}** masks saved."
)


st.caption(
    "Aim for around 30–50 carefully checked masks initially. "
    "We will use these to train the first dedicated food segmentation model."
)
