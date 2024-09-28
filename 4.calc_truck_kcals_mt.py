import pandas as pd
import numpy as np
import os
import shutil
from datetime import datetime
import platform

# =====================
# STEP 4: CALCULATE TRUCK KCALS & METRIC TONS
# =====================

# Get the current date in the format YYYYMMDD (to match the folder created in step 1, 2, and 3)
current_date = datetime.now().strftime('%Y%m%d')

# Determine the user's desktop location (macOS and Windows compatibility)
if platform.system() == "Darwin":  # macOS
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
elif platform.system() == "Windows":  # Windows
    desktop = os.path.join(os.environ["HOMEPATH"], "Desktop")
else:
    raise Exception("Unsupported operating system. This script works on macOS and Windows only.")

# Path to the folder created by previous steps (same date-based folder)
folder_name = f"UNRWA Truck Data_{current_date}"
data_dir = os.path.join(desktop, folder_name)

# Input file path (output from Step 3)
data_path = os.path.join(data_dir, "unrwa_trucks.xlsx")

# Load the data from the 'unrwa_trucks_kcal' sheet generated by script 3
data = pd.read_excel(data_path, sheet_name='unrwa_trucks_kcal')

# Function to clean the item text (consistent with script 3)
def clean_item_text(text):
    if pd.isna(text):
        return text
    return text.strip().replace('(', '').replace(')', '').replace('"', '')

# Function to calculate item weight based on unit and cargo description
def calculate_item_weight(unit, cargo, quantity, item_weight):
    if unit == 'Pallets':
        cargo = cargo.lower()
        if 'rte' in cargo:
            return 790 * quantity
        elif 'wheat flour' in cargo:
            return 1000 * quantity
        elif 'lns' in cargo:
            return 900 * quantity
        elif 'date bar' in cargo:
            return 540 * quantity
        else:
            return 637.5 * quantity
    elif unit == 'Truck':
        return 15000  # Assume a full truckload weight of 15000 kg
    elif unit == 'Ton':
        return 1000 * quantity
    elif pd.notna(item_weight):
        return item_weight  # Use item_weight if specified
    else:
        return 0  # Default case if no unit or weight is given

# Split 'cargo' text by '+' into 'item_' variables (compatible with script 3 logic)
cargo_split = data['cargo'].str.split('+', expand=True)

# Determine max number of items in any entry
max_items = cargo_split.shape[1]

# Create 'item_1' to 'item_N' variables based on max items and populate them
for i in range(max_items):
    data[f'item_{i+1}'] = cargo_split[i].apply(clean_item_text)

# Identify item columns and kcal columns from the updated data
item_columns = [f'item_{i+1}' for i in range(max_items)]
kcal_columns = [f'{col}_kcal' for col in item_columns]

# Calculate truck_kcal based on the revised logic
data['truck_kcal'] = 0.0  # Initialize truck_kcal column with float type
for index, row in data.iterrows():
    truck_kcal = 0
    for i in range(1, max_items + 1):
        item = row[f'item_{i}']
        if pd.isna(item) or item == '':
            continue
        item_kcal = row[f'item_{i}_kcal']
        quantity = row['Quantity']
        unit = row['unit']
        item_weight = row['truck_weight_kg']  # Assuming this is the column name for item weight
        cargo = row['cargo']
        
        # Calculate item weight based on the given unit and cargo description
        calculated_weight = calculate_item_weight(unit, cargo, quantity, item_weight)
        
        # Calculate kcal contribution of the item
        item_contribution = item_kcal * calculated_weight
        truck_kcal += item_contribution
    
    # Set the final kcal value for the truck, cast as float
    data.at[index, 'truck_kcal'] = float(truck_kcal)

# Determine truck type based on food item count
def determine_truck_type(row):
    if row['food_item_count'] > 0 and row['food_item_count'] == row['item_count']:
        return 'Food Truck'
    elif row['food_item_count'] == 0:
        return 'Non-Food Truck'
    else:
        return 'Mixed Food/Non-Food Truck'

# Add food item counts and truck types to the data
data['food_item_count'] = data[kcal_columns].gt(0).sum(axis=1)
data['item_count'] = data[item_columns].notna().sum(axis=1)
data['truck_type'] = data.apply(determine_truck_type, axis=1)

# Determine sector based on 'Donation Type'
def determine_sector(donation_type):
    if pd.isna(donation_type):
        return 'unknown'
    donation_type_lower = donation_type.lower()
    if 'private sector' in donation_type_lower:
        return 'private'
    elif 'humanitarian' in donation_type_lower:
        return 'humanitarian'
    else:
        return 'unknown'

data['sector'] = data['Donation Type'].apply(determine_sector)

# Additional calculations
data['truck_food_kg'] = data.apply(lambda row: calculate_item_weight(row['unit'], row['cargo'], row['Quantity'], row['truck_weight_kg']) if row['truck_type'] != 'Non-Food Truck' else 0, axis=1)
data['truck_food_mt'] = data['truck_food_kg'] / 1000
data['daily_kcal_food'] = data[kcal_columns].sum(axis=1)
data['daily_food_mt'] = data['truck_food_kg'] / 1000
data['daily_mt'] = data['truck_weight_kg'] / 1000
data['truck_food_ratio'] = data['truck_food_mt'] / data['daily_mt']

# Save the updated data with the original sheet name
output_file = os.path.join(data_dir, "unrwa_trucks.xlsx")
with pd.ExcelWriter(output_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    data.to_excel(writer, sheet_name='unrwa_trucks_kcal', index=False)

# Archive the workbook
archive_dir = os.path.join(data_dir, "archive")
os.makedirs(archive_dir, exist_ok=True)
timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
archive_file_path = os.path.join(archive_dir, f"unrwa_trucks_{timestamp}.xlsx")
shutil.copy(output_file, archive_file_path)

print(f"Processing and saving completed. Archived as {archive_file_path}.")
