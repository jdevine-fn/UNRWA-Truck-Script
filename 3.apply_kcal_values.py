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
kcal_ref_url = "https://raw.githubusercontent.com/jdevine-fn/UNRWA-Truck-Script/main/kcal_reference.xlsx"

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

# Ensure required columns are available
required_columns = ['unit', 'Quantity', 'Cargo Category', 'item_count', 'Donating Country/ Organization']
missing_columns = [col for col in required_columns if col not in data.columns]
if missing_columns:
    raise KeyError(f"The required columns {missing_columns} are missing from the data.")

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
    'canned green beans with meat': 'green beans with meat',
    'chicken broth': 'chicken soup',
    'cooked beans': 'beans',
    'cooked meal': 'prepared food',
    'mixed canned meal': 'prepared food',
    'food commodity': 'food items',
    'extra meat': 'meat',
    'pineapples': 'pineapple',
    'mango': 'mangoes',
    # Add more mappings as needed
}

# Define a function to singularize words (simple heuristic)
def singularize(word):
    if word.endswith('s') and len(word) > 3:
        return word[:-1]
    else:
        return word

# List of known non-food items
non_food_items = set([
    'mats', 'tents', 'blankets', 'clothes', 'medicines', 'medicine',
    'hygiene kits', 'sanitary items', 'medical equipment', 'soap',
    'toothbrushes', 'water filters', 'jerry cans', 'tarpaulins'
])

# Define a function to find the best match using fuzzy matching
def find_best_match(item, reference_list, cutoff=0.85):
    matches = difflib.get_close_matches(item, reference_list, n=1, cutoff=cutoff)
    if matches:
        return matches[0]
    else:
        return None
    
# Exclude non-item columns that start with 'item_'
non_item_columns = ['item_count', 'item_count_matched']
# List of item columns
item_columns = [
    col for col in data.columns
    if col.startswith('item_')
    and not col.endswith(('_kg', '_kcal', '_matched'))
    and col not in non_item_columns
]

# Initialize sets to collect unmatched items and units
unmatched_items = set()
unmatched_units = set()

# Initialize total item weight and kcal columns with float dtype
for item_col in item_columns:
    data[f'{item_col}_kg'] = np.nan  # Use NaN to ensure float dtype
    data[f'{item_col}_kcal'] = np.nan

# Set default pallet weight
default_pallet_weight = 850  # in kg
# Calculate average item kcal per kg from kcal_ref
average_item_kcal_per_kg = kcal_ref['Nutval Kcal KG'].mean()

# Iterate over each row to calculate item weights and kcals
for index, row in data.iterrows():
    unit = str(row['unit']).lower()
    quantity = row['Quantity']
    item_count = row['item_count']
    cargo_category = row.get('Cargo Category', None)
    donor = row.get('Donating Country/ Organization', '')
    donor_contains_wfp = 'wfp' in str(donor).lower()
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

    # **Check if item is a known non-food item**
    if mapped_item in non_food_items:
        is_food_item = False
        data.at[index, f'{item_col}_matched'] = 'non-food'
        data.at[index, f'{item_col}_kcal'] = np.nan
        unmatched_items.add(str(item))  # Record as unmatched
        continue  # Skip to the next item

    # First, try exact match
    if mapped_item in kcal_food_items:
        best_match = mapped_item
        is_food_item = True
        match_type = 'exact'
    else:
        # If no exact match, try fuzzy matching with higher cutoff
        best_match = find_best_match(mapped_item, kcal_food_items, cutoff=0.85)
        is_food_item = best_match is not None
        if is_food_item:
            match_type = 'fuzzy'
            # Record the fuzzy matched item
            fuzzy_matched_items.append({
                'Row Index': index + 2,  # Adjust for header and zero indexing
                'Item Column': item_col,
                'Original Item': item,
                'Matched Item': best_match
            })
        else:
            match_type = 'unmatched'
            unmatched_items.add(str(item))  # Record unmatched item


        # Calculate item weight based on unit
        if unit == 'pallets':
            pallets_per_item = quantity  # Quantity is per item
            if is_food_item:
                if donor_contains_wfp:
                    # Use specific pallet weight if matched, else default
                    if best_match:
                        match = kcal_ref[kcal_ref['food_item'] == best_match]
                        if not match.empty and not np.isnan(match['pallet_kg'].values[0]):
                            pallet_weight = match['pallet_kg'].values[0]
                        else:
                            pallet_weight = default_pallet_weight
                    else:
                        pallet_weight = default_pallet_weight
                else:
                    # Donor does not include WFP, use default pallet weight
                    pallet_weight = default_pallet_weight
            else:
                pallet_weight = default_pallet_weight  # Use default for non-food items
            item_weight_kg = pallets_per_item * pallet_weight
        elif unit in ['ton', 'tons', 'mt']:
            item_weight_kg = quantity * 1000  # Quantity is per item in tons
        elif unit == 'kg':
            item_weight_kg = quantity  # Quantity is per item in kg
        elif unit == 'truck':
            item_weight_kg = 14000  # 14 MT per truck
        else:
            item_weight_kg = np.nan  # Unknown unit
            unmatched_units.add(unit)

        # Assign calculated weight
        data.at[index, f'{item_col}_kg'] = item_weight_kg

        if is_food_item:
            # Calculate item kcal
            if best_match:
                match = kcal_ref[kcal_ref['food_item'] == best_match]
                if not match.empty and not np.isnan(match['Nutval Kcal KG'].values[0]):
                    item_kcal_per_kg = match['Nutval Kcal KG'].values[0]
                else:
                    item_kcal_per_kg = average_item_kcal_per_kg
            else:
                item_kcal_per_kg = average_item_kcal_per_kg
            item_kcal = item_weight_kg * item_kcal_per_kg

            # Assign kcal and matched item
            data.at[index, f'{item_col}_kcal'] = item_kcal
            data.at[index, f'{item_col}_matched'] = best_match if best_match else 'default'
        else:
            # Non-food items have no kcal
            data.at[index, f'{item_col}_kcal'] = np.nan
            data.at[index, f'{item_col}_matched'] = 'non-food'

# Sum up all item weights to get total truck weight before modifying item weights
item_weight_cols = [f'{col}_kg' for col in item_columns]
data['truck_weight_kg'] = data[item_weight_cols].sum(axis=1, min_count=1)

# Now, set non-food item weights to NaN before calculating truck_food_kg
for item_col in item_columns:
    matched_col = f'{item_col}_matched'
    weight_col = f'{item_col}_kg'
    kcal_col = f'{item_col}_kcal'
    if matched_col in data.columns:
        is_food = data[matched_col].apply(lambda x: x != 'non-food' and pd.notna(x))
        # Set item_weight_kg and item_kcal to NaN where is_food is False
        data.loc[~is_food, weight_col] = np.nan
        data.loc[~is_food, kcal_col] = np.nan

# Sum up food item weights and kcals
food_weight_cols = [f'{col}_kg' for col in item_columns]
food_kcal_cols = [f'{col}_kcal' for col in item_columns]
data['truck_food_kg'] = data[food_weight_cols].sum(axis=1, min_count=1)
data['truck_kcal'] = data[food_kcal_cols].sum(axis=1, min_count=1)

# Replace NaN values with zeros
data['truck_food_kg'].fillna(0, inplace=True)
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
    unmatched_items_str = [str(item) for item in unmatched_items]
    for item in sorted(unmatched_items_str):
        f.write(f"{item}\n")

# Save unmatched units to a text file for review
unmatched_units_path = os.path.join(data_dir, "unmatched_units.txt")
with open(unmatched_units_path, 'w') as f:
    for unit in sorted(unmatched_units):
        f.write(f"{unit}\n")

print(f"Processing and saving completed. Archived as {archive_file_path}.")
print(f"Unmatched items saved to {unmatched_items_path}.")
print(f"Unmatched units saved to {unmatched_units_path}.")
