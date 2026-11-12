import streamlit as st
import pandas as pd
from pathlib import Path
from datetime import datetime
import re


# ============================================================
# SETTINGS
# ============================================================

BASE_FOLDER = Path(r"C:\INSERT FILE NAME HERE")
MASTER_FOLDER = BASE_FOLDER / "Cleaned_Raw_Image_Data_2026"
SHEET = BASE_FOLDER / "meal_image_pairing.xlsx"
NUT_SHEET = BASE_FOLDER / "meal_nutritional_data_streamlit.xlsx"
CONSUMPTION_FILE = (
    BASE_FOLDER
    / "food_consumed_weights_master.xlsx"
)

df = pd.read_excel(SHEET)

# Nutritional values for each food/component
nut_df = pd.read_excel(
    NUT_SHEET,
    sheet_name="Food_Nutrients"
)

# Estimated consumption percentages for displayed foods
consumption_df = pd.read_excel(
    CONSUMPTION_FILE,
    sheet_name="Consumed_Foods"
)

# Estimated consumed weights of individual components
components_df = pd.read_excel(
    CONSUMPTION_FILE,
    sheet_name="Consumed_Components"
)

# ============================================================
# CONSUMPTION THRESHOLD FOR CALCS
# ============================================================

CONSUMPTION_THRESHOLD = 10.0

def adjusted_consumption(value):
    """Return the percentage used for nutritional calculations."""
    if pd.isna(value):
        return None

    value = float(value)

    if value <= CONSUMPTION_THRESHOLD:
        return 0.0

    return value


def display_consumption(value):
    """Return a formatted percentage for the Streamlit page."""
    if pd.isna(value):
        return "N/A"

    value = float(value)

    if value <= CONSUMPTION_THRESHOLD:
        return "~0%"

    return f"{value:.1f}%"

# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="Patient Image Viewer",
    page_icon="🖼️",
    layout="wide"
)


# ============================================================
# FUNCTIONS
# ============================================================

def parse_datetime_folder(folder_name):
    """
    Extract the date and time from a folder name.

    Expected format:

    yyyy-mm-dd hh;mm&####.#

    Example:

    2025-01-01 12;00&1000.0

    Only the date and time are used.
    The weight/value after '&' is ignored.
    """

    # Take everything before '&'
    date_time_part = folder_name.split("&")[0]

    # Convert the semicolon between hour and minute
    # into a colon for datetime.strptime()
    date_time_part = date_time_part.replace(";", ":")

    return datetime.strptime(
        date_time_part,
        "%Y-%m-%d %H:%M"
    )

def get_breakdown(image_path):

    image_path = normalise_image_path(image_path)

    for _, row in df.iterrows():

        before_path = normalise_image_path(
            row["Before_Image"]
        )

        after_path = normalise_image_path(
            row["After_Image"]
        )

        if image_path in (before_path, after_path):
            return row["Meal Breakdown"]

    return None


def get_meal_info(image_path):

    image_path = normalise_image_path(image_path)

    for _, row in df.iterrows():

        before_path = normalise_image_path(
            row["Before_Image"]
        )

        after_path = normalise_image_path(
            row["After_Image"]
        )

        if image_path == before_path:
            return {
                "Meal_ID": row["Meal_ID"],
                "Image_Type": "BEFORE"
            }

        if image_path == after_path:
            return {
                "Meal_ID": row["Meal_ID"],
                "Image_Type": "AFTER"
            }

    return {
        "Meal_ID": None,
        "Image_Type": None
    }


def get_food_percentages(meal_id):

    if meal_id is None:
        return []

    rows = consumption_df[
        consumption_df["Meal_ID"].astype(str)
        == str(meal_id)
    ]

    results = []

    for index, row in rows.iterrows():

        food = row["Served_Food"]
        percent = row["Percent_Eaten"]

        percent_text = display_consumption(percent)

        results.append(
            {
                "food": food,
                "percent": percent_text,
                "status": row.get(
                    "Calculation_Status",
                    ""
                ),
                "image_match": row.get(
                    "Image_Match",
                    ""
                )
            }
        )

    return results

NUTRIENTS = {
    "Energy": ("Energy (kcal) (kcal) (actual)", "kcal", 2000),
    "Protein": ("Protein (g) (actual)", "g", 50),
    "Fat": ("Fat (g) (actual)", "g", 70),
    "Saturated fat": ("Satd FA /100g fd (g) (actual)", "g", 20),
    "Carbohydrates": ("Carbohydrate (g) (actual)", "g", 260),
    "Sugars": ("Total sugars (g) (actual)", "g", 90),
    "Salt": ("Salt (g) (actual)", "g", 6),
    "Fibre": ("AOAC fibre (g) (actual)", "g", 30),
}


def normalise_food(value):
    if pd.isna(value):
        return ""
    return " ".join(str(value).strip().casefold().split())


def normalise_code(value):
    if pd.isna(value):
        return ""
    return str(value).strip().upper()


# Prepare matching columns for both spreadsheets
for table, served_col, component_col, code_col in (
    (nut_df, "Food_Description", "Food_Item", "McCance_Widdowson_Code"),
    (components_df, "Served_Food", "Component", "M&W_Code"),
):
    table["_meal"] = pd.to_numeric(table["Meal_ID"], errors="coerce")
    table["_served"] = table[served_col].map(normalise_food)
    table["_component"] = table[component_col].map(normalise_food)
    table["_code"] = table[code_col].map(normalise_code)


def calculate_consumed_nutrition(meal_id, meal_breakdown=None):

    results = {name: None for name in NUTRIENTS}

    if pd.isna(meal_id):
        return results

    meal_number = pd.to_numeric(meal_id, errors="coerce")

    served = nut_df[
        nut_df["_meal"] == meal_number
    ].copy()

    consumed = components_df[
        components_df["_meal"] == meal_number
    ].copy()

    if served.empty or consumed.empty:
        return results

    keys = ["_meal", "_served", "_component", "_code"]

    # Avoid duplicating nutrient values if matching keys repeat
    if (
        served.duplicated(keys).any()
        or consumed.duplicated(keys).any()
    ):
        return results

    merged = served.merge(
        consumed[
            keys + [
                "Component_Served_Weight",
                "Component_Consumed_Weight"
            ]
        ],
        on=keys,
        how="outer",
        validate="one_to_one",
        indicator=True
    )

    # Every component must have a matching record
    if not merged["_merge"].eq("both").all():
        return results

    original_weight = pd.to_numeric(
        merged["Weight_g"],
        errors="coerce"
    )

    served_weight = pd.to_numeric(
        merged["Component_Served_Weight"],
        errors="coerce"
    )

    eaten_weight = pd.to_numeric(
        merged["Component_Consumed_Weight"],
        errors="coerce"
    )

    # Check that weights are valid and consistent
    
    # ============================================================
    # APPLY 10% CONSUMPTION THRESHOLD
    # ============================================================

    meal_foods = consumption_df[
        pd.to_numeric(
            consumption_df["Meal_ID"], errors="coerce"
        ) == meal_number
    ].copy()

    meal_foods["_served"] = meal_foods[
        "Served_Food"
    ].map(normalise_food)

    meal_foods["Adjusted_Percent"] = meal_foods[
        "Percent_Eaten"
    ].apply(adjusted_consumption)

    # Check for conflicting duplicate food percentages
    conflicts = meal_foods.groupby("_served")[
        "Adjusted_Percent"
    ].nunique(dropna=True)

    if (conflicts > 1).any():
        return results

    food_percentages = (
        meal_foods.drop_duplicates("_served")
        .set_index("_served")["Adjusted_Percent"]
    )

    component_percentages = merged[
        "_served"
    ].map(food_percentages)

    # Identify components belonging to foods at or below 10%
    
    # Calculate consumption percentage for each component
    component_consumption_percent = (
        eaten_weight / served_weight
    ) * 100

    # Components at or below 10% contribute zero nutrients
    zero_consumption = (
        component_consumption_percent <= CONSUMPTION_THRESHOLD
    )


    # ============================================================
    # VALIDATE WEIGHTS
    # ============================================================

    if (
        original_weight.isna().any()
        or served_weight.isna().any()
        or eaten_weight.isna().any()
        or (original_weight <= 0).any()
        or (served_weight <= 0).any()
        or ((original_weight - served_weight).abs() > 0.05).any()
    ):
        return results

    # Only reject invalid consumed weights for foods that
    # are NOT being treated as zero consumption.
    invalid_consumed_weight = (
        (eaten_weight < 0)
        | (eaten_weight > served_weight + 0.01)
    )

    if (invalid_consumed_weight & ~zero_consumption).any():
        return results

    # ============================================================
    # CALCULATE CONSUMED FRACTIONS
    # ============================================================

    fraction = eaten_weight / served_weight

    # Foods at or below 10% contribute zero nutrients.
    fraction = fraction.mask(zero_consumption, 0.0)


    for name, (column, unit, reference) in NUTRIENTS.items():

        if column not in merged.columns:
            continue

        amount = pd.to_numeric(
            merged[column],
            errors="coerce"
        )

        # Do not treat missing nutritional data as zero
        # for components that were eaten
        # Treat missing nutritional values as zero
        amount = amount.fillna(0)

        # Calculate total estimated nutrient consumption
        results[name] = float(
            (amount * fraction).sum()
        )

    return results

def find_patient_images(patient_id):
    """
    Find all images for a patient across all week folders.

    Only files ending in '-image.jpg' are included.
    """

    results = []

    if not MASTER_FOLDER.exists():
        return results

    # Patient folder name
    patient_folder_name = f"patient_{patient_id}"

    # Search all week folders
    for week_folder in MASTER_FOLDER.iterdir():

        if not week_folder.is_dir():
            continue

        # Find the patient's folder inside this week
        patient_folder = week_folder / patient_folder_name

        if not patient_folder.exists():
            continue

        if not patient_folder.is_dir():
            continue

        # Search each date/time folder
        for datetime_folder in patient_folder.iterdir():

            if not datetime_folder.is_dir():
                continue

            # Extract date/time
            try:
                image_datetime = parse_datetime_folder(
                    datetime_folder.name
                )

            except ValueError:
                # Ignore folders that don't match the expected format
                continue

            # Find ONLY the required image files
            for image_path in datetime_folder.glob("*-image.jpeg"):

                meal_info = get_meal_info(
                    image_path
                )
                
                results.append(
                    {
                        "path": image_path,
                        "datetime": image_datetime,
                        "Meal_ID": meal_info["Meal_ID"],
                        "Image_Type": meal_info["Image_Type"],
                        "Meal Breakdown": get_breakdown(image_path)
                    }
                )

    # Most recent → oldest
    results.sort(
        key=lambda x: x["datetime"],
        reverse=True
    )

    return results

def get_food_options():
    """
    Build a list of searchable foods from the meal breakdowns.
    """

    foods = set()

    for breakdown in df["Meal Breakdown"].dropna():

        # Split foods separated by +
        parts = str(breakdown).split(" + ")

        for part in parts:

            # Remove leading weights, e.g. 100.0g/ml
            food = re.sub(
                r"^\s*\d+(?:\.\d+)?\s*g/ml\s*",
                "",
                part
            ).strip()

            if food:
                foods.add(food)

    return sorted(foods, key=str.casefold)


def find_food_images(food_name):
    """
    Find BEFORE and AFTER images for meals containing
    the selected food.
    """

    results = []
    seen_paths = set()

    for _, row in df.iterrows():

        breakdown = str(row.get("Meal Breakdown", ""))

        if food_name.casefold() not in breakdown.casefold():
            continue

        for image_type, column in [
            ("BEFORE", "Before_Image"),
            ("AFTER", "After_Image")
        ]:

            image_value = row.get(column)

            if pd.isna(image_value):
                continue

            image_path = Path(str(image_value))

            if not image_path.is_absolute():
                image_path = BASE_FOLDER / image_path

            if not image_path.exists():
                continue

            if str(image_path) in seen_paths:
                continue

            seen_paths.add(str(image_path))

            try:
                image_datetime = parse_datetime_folder(
                    image_path.parent.name
                )
            except ValueError:
                continue

            results.append({
                "path": image_path,
                "datetime": image_datetime,
                "Meal_ID": row["Meal_ID"],
                "Image_Type": image_type,
                "Meal Breakdown": row["Meal Breakdown"]
            })

    results.sort(
        key=lambda x: x["datetime"],
        reverse=True
    )

    return results

def normalise_image_path(path):
    if pd.isna(path):
        return ""

    path = str(path).strip().replace("\\", "/")

    # Match using the patient folder, datetime folder and filename
    parts = path.lower().split("/")

    for i, part in enumerate(parts):
        if part.startswith("patient_"):
            return "/".join(parts[i:])

    return path.lower()

# ============================================================
# HEADER
# ============================================================

st.title("Patient Image Viewer")

st.write(
    "Enter a patient ID to view all available images."
)

st.divider()


# ============================================================
# SEARCH
# ============================================================

search_mode = st.radio(
    "Search by",
    ["Patient ID", "Food"],
    horizontal=True
)

col1, col2, col3 = st.columns([4, 2, 1])

with col1:

    if search_mode == "Patient ID":

        patient_id = st.text_input(
            "Patient ID",
            placeholder="Enter patient ID"
        )

        selected_food = None

    else:

        selected_food = st.selectbox(
            "Select food",
            options=get_food_options(),
            index=None,
            placeholder="Choose a food"
        )

        patient_id = ""

with col2:

    meal_filter = st.selectbox(
        "Meal type",
        ["Both", "Lunch", "Dinner"]
    )

with col3:

    st.write("")

    search_button = st.button(
        "Search",
        use_container_width=True
    )

# ============================================================
# RESULTS
# ============================================================

if search_button:

    images = []
    search_description = ""

    if search_mode == "Patient ID":

        patient_id = patient_id.strip()

        if patient_id.lower().startswith("patient_"):
            patient_id = patient_id[8:]

        if not patient_id:

            st.warning("Please enter a patient ID.")

        else:

            images = find_patient_images(patient_id)

            search_description = f"patient {patient_id}"

    else:

        if not selected_food:

            st.warning("Please select a food.")

        else:

            images = find_food_images(selected_food)

            search_description = f"food '{selected_food}'"

    # Apply meal filter to either search method.
    if meal_filter == "Lunch":

        images = [
            image for image in images
            if 11 <= (
                image["datetime"].hour
                + image["datetime"].minute / 60
            ) < 16.5
        ]

    elif meal_filter == "Dinner":

        images = [
            image for image in images
            if 16.5 <= (
                image["datetime"].hour
                + image["datetime"].minute / 60
            ) < 20
        ]

    # Remove images that could not be matched to a meal
    images = [
        image for image in images
        if pd.notna(image["Meal_ID"])
    ]

    if search_description:

        if not images:

            st.error(
                f"No images found for {search_description}."
            )

        else:

            st.success(
                f"{len(images)} image(s) found for "
                f"{search_description}."
            )

            st.divider()

            # Group matching BEFORE/AFTER captures by Meal ID.
            # Keep the most recent capture first.
            meals = {}
            for image in images:
                meal_id = image.get("Meal_ID")
                if pd.isna(meal_id):
                    # Unknown IDs cannot safely be paired.
                    key = ("unmatched", str(image["path"]))
                else:
                    key = ("meal", str(meal_id))

                entry = meals.setdefault(key, {
                    "Meal_ID": meal_id,
                    "datetime": image["datetime"],
                    "breakdown": None,
                    "BEFORE": None,
                    "AFTER": None,
                })
                entry["datetime"] = max(entry["datetime"], image["datetime"])
                image_type = image.get("Image_Type")
                if image_type in ("BEFORE", "AFTER"):
                    existing = entry[image_type]
                    if existing is None or image["datetime"] > existing["datetime"]:
                        entry[image_type] = image
                breakdown = image.get("Meal Breakdown")
                if pd.notna(breakdown) and str(breakdown).strip():
                    entry["breakdown"] = str(breakdown)

            for entry in sorted(meals.values(), key=lambda m: m["datetime"], reverse=True):
                meal_id = entry["Meal_ID"]
                meal_label = (
                    f"Meal {int(meal_id)}" if pd.notna(meal_id) else "Unmatched image"
                )
                with st.container(border=True):
                    st.subheader(meal_label)
                    st.caption(entry["datetime"].strftime("%d %B %Y · %H:%M"))

                    left, right = st.columns(2, gap="medium")
                    for column, kind in ((left, "AFTER"), (right, "BEFORE")):
                        with column:
                            st.markdown(f"**{kind}**")
                            capture = entry[kind]
                            if capture is not None:
                                st.image(str(capture["path"]), use_container_width=True)
                            else:
                                st.info(f"No matched {kind.lower()} image available.")

                    st.markdown("**Meal description**")
                    st.write(entry["breakdown"] or "Meal description unavailable")

                    eaten_col, nutrients_col, recommendations_col = st.columns(
                        [1.25, 1, 1], gap="large"
                    )
                    with eaten_col:
                        st.markdown("**Estimated eaten %**")
                        food_percentages = get_food_percentages(meal_id)
                        if food_percentages:
                            for result in food_percentages:
                                st.markdown(f"{result['food']} — **{result['percent']}**")
                                if result["status"] and result["status"] != "CALCULATED":
                                    st.caption(str(result["status"]))
                        else:
                            st.caption("No food consumption estimates available.")

                    nutritional_results = calculate_consumed_nutrition(
                        meal_id,
                        entry["breakdown"]
                    )

                    with nutrients_col:

                        st.markdown("**Nutritional breakdown**")

                        for nutrient, (column, unit, reference) in NUTRIENTS.items():

                            value = nutritional_results[nutrient]

                            if value is None:
                                st.write(f"{nutrient}: N/A")
                            else:
                                st.write(
                                    f"{nutrient}: {value:.1f} {unit}"
                                )


                    with recommendations_col:

                        st.markdown("**% of government recommendations**")

                        for nutrient, (column, unit, reference) in NUTRIENTS.items():

                            value = nutritional_results[nutrient]

                            if value is None or reference is None:
                                st.write(f"{nutrient}: N/A")
                            else:
                                percentage = value / reference * 100

                                st.write(
                                    f"{nutrient}: {percentage:.1f}% "
                                    f"(of {reference:g} {unit})"
                                )

            st.caption(
                "Nutritional values are estimates. Consumption percentages "
                "of 10% or less are displayed as ~0% and treated as zero "
                "for nutritional calculations. Original consumption values "
                "are retained in the Excel spreadsheets."
            )
