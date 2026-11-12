#THIS SCRIPT WILL PRODUCE A EXCEL SPREADSHEET CONTAINING ALL AVAILABLE NUTRITIONAL VALUES FOR EVERY MEAL FOR THE CHOSEN NUTRIENTS FROM M&W
import pandas as pd
import os
import re

# ============================================================
# FILENAMES / SHEETS
# ============================================================

FOODCODE_FILENAME = "meal_food_analysis.xlsx"
MWNUTRIENT_FILENAME = "McCance_Widdowsons_Composition_of_Foods_Integrated_Dataset_2021..xlsx"

WEIGHT_SHEETNAME = "Original_Data"
CODE_SHEETNAME = "MW_Coding"
COMPOSITE_SHEETNAME = "Composite_Coding"


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def choose_multiple(items, prompt):
    """Choose one or more items from a numbered list."""
    for i, item in enumerate(items, start=1):
        print(f"{i}. {item}")

    while True:
        raw = input(prompt).strip()
        try:
            numbers = [int(x.strip()) for x in raw.split(",") if x.strip()]
            if numbers and all(1 <= n <= len(items) for n in numbers):
                # Preserve order but remove duplicates
                seen = set()
                result = []
                for n in numbers:
                    item = items[n - 1]
                    if item not in seen:
                        seen.add(item)
                        result.append(item)
                return result
        except ValueError:
            pass

        print("Please enter valid numbers separated by commas, e.g. 1,3,4.")


def choose_one_column(df, description):
    """Ask the user to choose a single dataframe column."""
    print(f"\n{description}\nColumns found:")
    for i, column in enumerate(df.columns, start=1):
        print(f"{i}. {column}")

    while True:
        try:
            number = int(input("\nEnter the column number: "))
            if 1 <= number <= len(df.columns):
                return df.columns[number - 1]
        except ValueError:
            pass

        print("Please enter a valid column number.")


def normalise_code(series):
    """Standardise M&W codes for safer merges."""
    return series.astype("string").str.strip()


def clean_food_text(series):
    return (
        series
        .astype("string")
        .str.strip()
        .str.lower()
    )


def find_food_code_column(df, sheet_name):
    """Use 'Food Code' automatically when present; otherwise ask."""
    if "Food Code" in df.columns:
        return "Food Code"

    return choose_one_column(
        df,
        f"Could not automatically find 'Food Code' in M&W sheet '{sheet_name}'. "
        "Select the food-code column."
    )


def format_breakdown(group, nutrient_actual_column):
    """Create a readable component-level nutrient breakdown for one meal."""
    parts = []

    for _, row in group.iterrows():
        value = row[nutrient_actual_column]
        if pd.isna(value):
            continue

        food_name = row.get("Food_Item", row.get("Food_Description", "Food"))
        parts.append(f"{food_name}: {value:.2f}")

    return " | ".join(parts)


# ============================================================
# LOAD MAIN WORKBOOK SHEETS
# ============================================================

weight_df = pd.read_excel(FOODCODE_FILENAME, sheet_name=WEIGHT_SHEETNAME)
code_df = pd.read_excel(FOODCODE_FILENAME, sheet_name=CODE_SHEETNAME)
composite_df = pd.read_excel(FOODCODE_FILENAME, sheet_name=COMPOSITE_SHEETNAME)


# ============================================================
# SELECT ORIGINAL-DATA COLUMNS NEEDED FOR FINAL OUTPUT
# ============================================================

portion_breakdown_column = choose_one_column(
    weight_df,
    "Select the column containing the Portion Breakdown."
)

patient_column = choose_one_column(
    weight_df,
    "Select the Patient ID column."
)

date_column = choose_one_column(
    weight_df,
    "Select the Date column."
)

week_column = choose_one_column(
    weight_df,
    "Select the Week column."
)

meal_type_column = choose_one_column(
    weight_df,
    "Select the Lunch/Dinner column."
)

meal_description_column = choose_one_column(
    weight_df,
    "Select the Meal Description column."
)

menu_item_column = choose_one_column(
    weight_df,
    "Select the Menu Item column."
)


# ============================================================
# SELECT FOOD-CODE COLUMN IN MW_CODING
# ============================================================

code_column = choose_one_column(
    code_df,
    "Select the column containing the assigned M&W food codes in MW_Coding."
)


# ============================================================
# SELECT MULTIPLE M&W SHEETS + NUTRIENTS FROM EACH
# ============================================================

mw_excel = pd.ExcelFile(MWNUTRIENT_FILENAME)
all_mw_sheets = mw_excel.sheet_names

print("\n============================================================")
print("M&W NUTRIENT SHEETS")
print("============================================================")

selected_sheets = choose_multiple(
    all_mw_sheets,
    "\nEnter the numbers of ALL sheets you want to use (comma-separated): "
)

print("\nSelected M&W sheets:")
for sheet in selected_sheets:
    print(f"- {sheet}")

nutrient_lookup = None
selected_nutrients = []

for sheet_name in selected_sheets:
    sheet_df = pd.read_excel(MWNUTRIENT_FILENAME, sheet_name=sheet_name)
    food_code_column = find_food_code_column(sheet_df, sheet_name)

    available_columns = [c for c in sheet_df.columns if c != food_code_column]

    print("\n============================================================")
    print(f"NUTRIENTS AVAILABLE: {sheet_name}")
    print("============================================================")

    chosen_columns = choose_multiple(
        available_columns,
        "\nEnter the numbers of the nutrient columns to keep "
        "from this sheet (comma-separated): "
    )

    # Keep only food code + requested columns
    temp = sheet_df[[food_code_column] + chosen_columns].copy()
    temp = temp.rename(columns={food_code_column: "_Food_Code"})
    temp["_Food_Code"] = normalise_code(temp["_Food_Code"])

    # If a nutrient name already exists from another sheet, prefix it with sheet name
    rename_map = {}
    for nutrient in chosen_columns:
        final_name = nutrient
        if final_name in selected_nutrients:
            final_name = f"{sheet_name} - {nutrient}"
        rename_map[nutrient] = final_name
        selected_nutrients.append(final_name)

    temp = temp.rename(columns=rename_map)

    # Ensure one row per food code before merging
    temp = temp.drop_duplicates(subset=["_Food_Code"])

    if nutrient_lookup is None:
        nutrient_lookup = temp
    else:
        nutrient_lookup = nutrient_lookup.merge(
            temp,
            on="_Food_Code",
            how="outer"
        )

if nutrient_lookup is None or not selected_nutrients:
    raise ValueError("No M&W nutrient columns were selected.")

print("\nNutrients selected across all sheets:")
for nutrient in selected_nutrients:
    print(f"- {nutrient}")


# ============================================================
# EXTRACT FOOD DESCRIPTIONS AND WEIGHTS FROM PORTION BREAKDOWN
# ============================================================

food_records = []

for _, row in weight_df.iterrows():
    meal_id = row["Meal_ID"]
    meal = row[portion_breakdown_column]

    if pd.isna(meal):
        continue

    meal = str(meal)

    portion_matches = re.finditer(
        r"([\d.]+)\s*g/ml\s*\(",
        meal
    )

    for match in portion_matches:
        weight = float(match.group(1))
        start = match.end()
        bracket_depth = 1
        position = start

        while position < len(meal) and bracket_depth > 0:
            if meal[position] == "(":
                bracket_depth += 1
            elif meal[position] == ")":
                bracket_depth -= 1
            position += 1

        if bracket_depth == 0:
            description = meal[start:position - 1]

            food_records.append({
                "Meal_ID": meal_id,
                "Food_Description": description.strip(),
                "Weight_g": weight
            })

food_items = pd.DataFrame(food_records)

if food_items.empty:
    raise ValueError("No food portions could be extracted from the selected Portion Breakdown column.")


# ============================================================
# WEIGHT CONSISTENCY CHECK
# ============================================================

weight_check = (
    food_items
    .groupby("Food_Description")["Weight_g"]
    .agg(["count", "nunique", "min", "max"])
    .reset_index()
)

variable_weights = weight_check[weight_check["nunique"] > 1]

print("\n============================================================")
print("WEIGHT CONSISTENCY CHECK")
print("============================================================")
print(f"\nTotal unique foods: {len(weight_check)}")
print(
    "Foods with a single consistent weight: "
    f"{len(weight_check[weight_check['nunique'] == 1])}"
)
print(f"Foods with multiple different weights: {len(variable_weights)}")

if len(variable_weights) > 0:
    print("\nFoods with VARIABLE weights:")
    print(variable_weights.to_string(index=False))
else:
    print("\nExcellent! Every food has a consistent weight.")


# ============================================================
# MATCH FOOD DESCRIPTIONS TO M&W CODES
# ============================================================

# The user selects which code column to use, but the food-name column remains Food_Item.
code_lookup = code_df[["Food_Item", code_column]].copy()
code_lookup = code_lookup.rename(columns={code_column: "McCance_Widdowson_Code"})

food_items["Food_Description"] = clean_food_text(food_items["Food_Description"])
code_lookup["Food_Item"] = clean_food_text(code_lookup["Food_Item"])
code_lookup["McCance_Widdowson_Code"] = normalise_code(
    code_lookup["McCance_Widdowson_Code"]
)


# ============================================================
# CHECK / EXPAND COMPOSITE FOODS
# ============================================================

composite_df["Meal_Name"] = clean_food_text(composite_df["Meal_Name"])
composite_df["Food_Item"] = clean_food_text(composite_df["Food_Item"])
composite_df["M&W_Code"] = normalise_code(composite_df["M&W_Code"])

percentage_check = (
    composite_df
    .groupby("Meal_Name")["Portion_Percentage_(%)"]
    .sum()
)

invalid_percentages = percentage_check[~percentage_check.round(6).eq(100)]

print("\n============================================================")
print("COMPOSITE PERCENTAGE CHECK")
print("============================================================")

if len(invalid_percentages) > 0:
    print("\nWARNING: These composite foods do not total 100%:")
    print(invalid_percentages)
else:
    print("\nAll composite food percentages total 100%.")

composite_foods = food_items[
    food_items["Food_Description"].isin(composite_df["Meal_Name"])
].copy()

composite_foods = composite_foods.merge(
    composite_df,
    left_on="Food_Description",
    right_on="Meal_Name",
    how="left"
)

composite_foods["New_Weight_g"] = (
    composite_foods["Weight_g"]
    * composite_foods["Portion_Percentage_(%)"]
    / 100
)


# ============================================================
# BUILD SINGLE-FOOD ROWS
# ============================================================

all_data = food_items.merge(
    code_lookup,
    left_on="Food_Description",
    right_on="Food_Item",
    how="left"
)

# For non-composite foods the displayed component is the food description itself.
all_data["Food_Item"] = all_data["Food_Description"]
all_data["McCance_Widdowson_Code"] = normalise_code(
    all_data["McCance_Widdowson_Code"]
)

all_data = all_data.merge(
    nutrient_lookup,
    left_on="McCance_Widdowson_Code",
    right_on="_Food_Code",
    how="left"
).drop(columns=["_Food_Code"])


# ============================================================
# BUILD COMPOSITE COMPONENT ROWS
# ============================================================

if not composite_foods.empty:
    composite_foods = composite_foods.merge(
        nutrient_lookup,
        left_on="M&W_Code",
        right_on="_Food_Code",
        how="left"
    ).drop(columns=["_Food_Code"])


# ============================================================
# CALCULATE ACTUAL NUTRIENT VALUES
# ============================================================

for nutrient in selected_nutrients:
    all_data[nutrient] = pd.to_numeric(all_data[nutrient], errors="coerce")
    all_data[f"{nutrient} (actual)"] = (
        all_data[nutrient]
        * all_data["Weight_g"]
        / 100
    )

    if not composite_foods.empty:
        composite_foods[nutrient] = pd.to_numeric(
            composite_foods[nutrient],
            errors="coerce"
        )
        composite_foods[f"{nutrient} (actual)"] = (
            composite_foods[nutrient]
            * composite_foods["New_Weight_g"]
            / 100
        )

# ============================================================
# CALCULATE SALT FROM SODIUM
# ============================================================

sodium_columns = [
    col for col in selected_nutrients
    if "sodium" in col.lower()
]

if sodium_columns:

    sodium_col = sodium_columns[0]

    # Sodium is in mg per 100g.
    # Salt is in g per 100g.
    all_data["Salt (g)"] = (
        all_data[sodium_col] * 2.5 / 1000
    )

    all_data["Salt (g) (actual)"] = (
        all_data["Salt (g)"]
        * all_data["Weight_g"]
        / 100
    )

    if not composite_foods.empty:

        composite_foods["Salt (g)"] = (
            composite_foods[sodium_col] * 2.5 / 1000
        )

        composite_foods["Salt (g) (actual)"] = (
            composite_foods["Salt (g)"]
            * composite_foods["New_Weight_g"]
            / 100
        )

    selected_nutrients.append("Salt (g)")

else:

    print(
        "WARNING: Sodium not selected. "
        "Salt cannot be calculated."
    )
# ============================================================
# REPLACE COMPOSITE MEAL ROWS WITH THEIR COMPONENT ROWS
# ============================================================

unique_composites = composite_foods["Food_Description"].unique() if not composite_foods.empty else []

single_foods = all_data[
    ~all_data["Food_Description"].isin(unique_composites)
].copy()

if not composite_foods.empty:
    composite_output = composite_foods.copy()
    composite_output["McCance_Widdowson_Code"] = composite_output["M&W_Code"]
    composite_output["Weight_g"] = composite_output["New_Weight_g"]

    # Keep just the columns needed downstream and align with single_foods.
    downstream_columns = [
        "Meal_ID",
        "Food_Description",
        "Food_Item",
        "Weight_g",
        "McCance_Widdowson_Code",
    ]

    for nutrient in selected_nutrients:
        downstream_columns.extend([
            nutrient,
            f"{nutrient} (actual)"
        ])

    composite_output = composite_output[downstream_columns]
    single_foods = single_foods[downstream_columns]

    all_data = pd.concat(
        [single_foods, composite_output],
        ignore_index=True
    )
else:
    all_data = single_foods.copy()


# ============================================================
# MEAL TOTAL WEIGHT
# Use original extracted portions so composite expansion cannot
# accidentally alter the meal's total weight.
# ============================================================

total_weight = (
    food_items
    .groupby("Meal_ID")["Weight_g"]
    .sum()
    .reset_index(name="total meal weight")
)


# ============================================================
# CREATE NUTRIENT TOTAL + BREAKDOWN FOR EACH MEAL
# ============================================================

meal_nutrition = all_data[["Meal_ID"]].drop_duplicates().copy()

for nutrient in selected_nutrients:
    actual_column = f"{nutrient} (actual)"

    totals = (
        all_data
        .groupby("Meal_ID")[actual_column]
        .sum(min_count=1)
        .reset_index(name=f"{nutrient} total")
    )

    breakdown = (
        all_data
        .groupby("Meal_ID", group_keys=False)
        .apply(
            lambda group: format_breakdown(group, actual_column)
        )
        .reset_index(name=f"{nutrient} breakdown")
    )

    meal_nutrition = meal_nutrition.merge(totals, on="Meal_ID", how="left")
    meal_nutrition = meal_nutrition.merge(breakdown, on="Meal_ID", how="left")


# ============================================================
# ADD PATIENT / MEAL METADATA FROM ORIGINAL_DATA
# ============================================================

metadata = weight_df[
    [
        "Meal_ID",
        patient_column,
        date_column,
        week_column,
        meal_type_column,
        meal_description_column,
        menu_item_column,
        portion_breakdown_column,
    ]
].drop_duplicates(subset=["Meal_ID"]).copy()

metadata = metadata.rename(columns={
    patient_column: "patientID",
    date_column: "date",
    week_column: "week",
    meal_type_column: "lunch or dinner",
    meal_description_column: "meal description",
    menu_item_column: "menu item",
    portion_breakdown_column: "portion breakdown",
})

meal_totals = metadata.merge(
    total_weight,
    on="Meal_ID",
    how="left"
)

meal_totals = meal_totals.merge(
    meal_nutrition,
    on="Meal_ID",
    how="left"
)


# ============================================================
# FINAL COLUMN ORDER
# ============================================================

final_columns = [
    "Meal_ID",
    "patientID",
    "date",
    "week",
    "lunch or dinner",
    "meal description",
    "menu item",
    "total meal weight",
    "portion breakdown",
]

for nutrient in selected_nutrients:
    final_columns.extend([
        f"{nutrient} total",
        f"{nutrient} breakdown",
    ])

meal_totals = meal_totals[final_columns]


# ============================================================
# ASK USER FOR OUTPUT FILE NAME
# ============================================================

output_name = input(
    "\nEnter the name for the output Excel file (without .xlsx): "
).strip()

if output_name == "":
    output_name = "meal_nutritional_totals"

if not output_name.lower().endswith(".xlsx"):
    output_name += ".xlsx"

output_path = os.path.join(
    os.path.dirname(FOODCODE_FILENAME),
    output_name
)


# ============================================================
# SAVE TO EXCEL
# ============================================================

with pd.ExcelWriter(output_path, engine="openpyxl") as writer:

    meal_totals.to_excel(
        writer,
        sheet_name="Meal_Totals",
        index=False
    )

    all_data.to_excel(
        writer,
        sheet_name="Food_Nutrients",
        index=False
    )

print("\n============================================================")
print("MEAL NUTRITIONAL TOTALS SAVED")
print("============================================================")
print(f"\nFile: {output_path}")
print(f"Number of meals: {len(meal_totals)}")

print("\nColumns included:")
for column in meal_totals.columns:
    print(f"- {column}")
