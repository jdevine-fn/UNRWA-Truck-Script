import pandas as pd
import numpy as np
import os
import shutil
import requests  # For downloading the kcal_reference.xlsx file from GitHub
from datetime import datetime
import platform
import difflib  # For fuzzy string matching

# =====================
# STEP 3: APPLY KCAL VALUES AND CALCULATE WEIGHTS
# =====================

# Get the current date in the format YYYYMMDD
current_date = datetime.now().strftime('%Y%m%d')

# Determine the user's desktop location
if platform.system() == "Darwin":  # macOS
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
elif platform.system() == "Windows":  # Windows
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
else:
    raise Exception("Unsupported operating system. This script works on macOS and Windows only.")

# Path to the folder created by previous steps
folder_name = f"UNRWA Truck Data_{current_date}"
data_dir = os.path.join(desktop, folder_name)

# Input file path (output from Step 2)
data_path = os.path.join(data_dir, "unrwa_trucks.xlsx")

# URL of the kcal_reference.xlsx file in your GitHub repository
kcal_ref_url = "https://raw.githubusercontent.com/jdevine-fn/UNRWA-Truck-Script/main/reference/kcal_reference.xlsx"

# Function to download kcal_reference.xlsx from GitHub
def download_kcal_reference(url, save_path):
    response = requests.get(url)
    with open(save_path, 'wb') as file:
        file.write(response.content)

# Download kcal_reference.xlsx into the working directory
kcal_ref_path = os.path.join(data_dir, "kcal_reference.xlsx")
download_kcal_reference(kcal_ref_url, kcal_ref_path)

# Load the processed data from Step 2 and the kcal reference file
data = pd.read_excel(data_path, sheet_name='unrwa_clean')
kcal_ref = pd.read_excel(kcal_ref_path)

# Ensure 'unit' and 'Quantity' columns are available
if 'unit' not in data.columns or 'Quantity' not in data.columns:
    raise KeyError("The required columns 'unit' and 'Quantity' are missing from the data.")

# Prepare kcal_ref for matching
kcal_ref['food_item'] = kcal_ref['food_item'].astype(str).str.strip().str.lower()

# Create a list of unique food items from kcal_ref for matching
kcal_food_items = kcal_ref['food_item'].tolist()

# Define a custom mapping dictionary for known mismatches
custom_mapping = {
    'canned white beans': 'white beans',
    'canned whit beans': 'white beans',
    'canned wihte beans': 'white beans',
    'palmera date': 'dates',
    'lentis': 'lentils',
    'lintels': 'lentils',
    'lentil soup': 'lentils',
    'red lentils': 'lentils',
    'vermicelli': 'noodles',
    'molasses': 'sugar',
    'peas and carrots': 'peas',
    'peanut butter': 'peanuts',
    'date bars': 'dates',
    'date bar': 'dates',
    'frozen peas': 'peas',
    'canned green beans with meat': 'green beans',
    'green beans with meat': 'green beans',
    'chicken broth': 'chicken soup',
    'cornflakes': 'corn',
    'cooked beans': 'beans',
    'cooked meal': 'prepared food',
    'mixed canned meal': 'prepared food',
    'food commodity': 'food items',
    'extra meat': 'meat',
    'pineapples': 'pineapple',
    'mango': 'mangoes',
    'matresses': 'mattresses',
    'matressess': 'mattresses',
    'medical supplies': None,  # Non-food item
    'medical supply': None,    # Non-food item
    'medicine': None,          # Non-food item
    'medicines': None,         # Non-food item
    # Add more mappings as needed
}

# Define a function to singularize words (simple heuristic)
def singularize(word):
    if word.endswith('s') and len(word) > 3:
        return word[:-1]
    else:
        return word

# Define a function to find the best match using fuzzy matching
def find_best_match(item, reference_list, cutoff=0.7):
    matches = difflib.get_close_matches(item, reference_list, n=1, cutoff=cutoff)
    if matches:
        return matches[0]
    else:
        return None

# List of item columns
item_columns = [col for col in data.columns if col.startswith('item_') and not col.endswith(('_kg', '_kcal', '_matched'))]

# Initialize sets to collect unmatched items
unmatched_items = set()

# Initialize total item weight and kcal columns with float dtype
for item_col in item_columns:
    data[f'{item_col}_kg'] = np.nan  # Use NaN to ensure float dtype
    data[f'{item_col}_kcal'] = np.nan

# Iterate over each row to calculate item weights and kcals
for index, row in data.iterrows():
    unit = row['unit']
    quantity = row['Quantity']
    item_count = row['item_count']
    if item_count == 0 or pd.isna(quantity) or quantity == 0:
        continue  # Skip rows with no items or zero quantity

    for item_col in item_columns:
        item = row[item_col]
        if pd.isna(item) or item == '':
            continue

        # Preprocess the item
        item_processed = singularize(str(item).strip().lower())

        # Check if item is in custom mapping
        if item_processed in custom_mapping:
            mapped_item = custom_mapping[item_processed]
            if mapped_item is None:
                unmatched_items.add(str(item))  # Explicitly unmatched (non-food)
                continue
        else:
            mapped_item = item_processed

        # First, try exact match
        if mapped_item in kcal_food_items:
            best_match = mapped_item
        else:
            # If no exact match, try fuzzy matching
            best_match = find_best_match(mapped_item, kcal_food_items)

        if best_match:
            match = kcal_ref[kcal_ref['food_item'] == best_match]
            pallet_weight = match['pallet_kg'].values[0]
            item_kcal_per_kg = match['Nutval Kcal KG'].values[0]

            # Calculate item weight based on unit
            if unit.lower() == 'pallets':
                # Calculate total pallets per item
                pallets_per_item = quantity / item_count
                item_weight_kg = pallets_per_item * pallet_weight
            elif unit.lower() == 'ton':
                item_weight_kg = (quantity * 1000) / item_count
            elif unit.lower() == 'kg':
                item_weight_kg = quantity / item_count
            else:
                item_weight_kg = np.nan  # Unknown unit

            # Calculate item kcal
            item_kcal = item_weight_kg * item_kcal_per_kg

            # Assign calculated values
            data.at[index, f'{item_col}_kg'] = item_weight_kg
            data.at[index, f'{item_col}_kcal'] = item_kcal
            data.at[index, f'{item_col}_matched'] = best_match  # Optional: track the matched item
        else:
            unmatched_items.add(str(item))  # Ensure item is a string
            data.at[index, f'{item_col}_kg'] = np.nan
            data.at[index, f'{item_col}_kcal'] = np.nan

# Sum item weights and kcals to get truck totals
data['truck_weight_kg'] = data[[f'{col}_kg' for col in item_columns]].sum(axis=1, min_count=1)
data['truck_kcal'] = data[[f'{col}_kcal' for col in item_columns]].sum(axis=1, min_count=1)

# Replace NaN values with zeros
data['truck_weight_kg'].fillna(0, inplace=True)
data['truck_kcal'].fillna(0, inplace=True)

# Save the updated data as a new sheet in the existing Excel file
with pd.ExcelWriter(data_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    data.to_excel(writer, sheet_name='unrwa_trucks_kcal', index=False)

# Archive the workbook
archive_dir = os.path.join(data_dir, "archive")
os.makedirs(archive_dir, exist_ok=True)
timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
archive_file_path = os.path.join(archive_dir, f"unrwa_trucks_{timestamp}.xlsx")
shutil.copy(data_path, archive_file_path)

# Save unmatched items to a text file for review
unmatched_items_path = os.path.join(data_dir, "unmatched_items.txt")
with open(unmatched_items_path, 'w') as f:
    # Convert all items to strings to avoid TypeError during sorting
    unmatched_items_str = [str(item) for item in unmatched_items]
    for item in sorted(unmatched_items_str):
        f.write(f"{item}\n")

print(f"Processing and saving completed. Archived as {archive_file_path}.")
print(f"Unmatched items saved to {unmatched_items_path}.")
