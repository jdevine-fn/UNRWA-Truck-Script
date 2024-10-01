import pandas as pd
import numpy as np
import os
import shutil
from datetime import datetime
import platform

# =====================
# STEP 4: CALCULATE TRUCK KCALS & METRIC TONS
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

# Input file path (output from Step 3)
data_path = os.path.join(data_dir, "unrwa_trucks.xlsx")

# Load the data from the 'unrwa_trucks_kcal' sheet
data = pd.read_excel(data_path, sheet_name='unrwa_trucks_kcal')

# Identify item columns, kg columns, and kcal columns
item_columns = [col for col in data.columns if col.startswith('item_') and not col.endswith(('_kg', '_kcal', '_matched'))]
kg_columns = [f'{col}_kg' for col in item_columns]
kcal_columns = [f'{col}_kcal' for col in item_columns]

# =====================
# Add food item counts and truck types to the data
# =====================
# Count number of food items (items with kcal > 0)
data['food_item_count'] = data[kcal_columns].gt(0).sum(axis=1)

# Count total number of items
data['item_count'] = data[item_columns].notna().sum(axis=1)

# Determine truck type based on food item count
def determine_truck_type(row):
    if row['food_item_count'] > 0 and row['food_item_count'] == row['item_count']:
        return 'Food Truck'
    elif row['food_item_count'] == 0:
        return 'Non-Food Truck'
    else:
        return 'Mixed Food/Non-Food Truck'

data['truck_type'] = data.apply(determine_truck_type, axis=1)

# Determine sector based on 'Donation Type'
def determine_sector(donation_type):
    if pd.isna(donation_type):
        return 'unknown'
    donation_type_lower = str(donation_type).lower()
    if 'private sector' in donation_type_lower:
        return 'private'
    elif 'humanitarian' in donation_type_lower:
        return 'humanitarian'
    else:
        return 'unknown'

data['sector'] = data['Donation Type'].apply(determine_sector)

# =====================
# Additional calculations
# =====================
# Calculate truck food weight in metric tons
data['truck_food_mt'] = data['truck_weight_kg'] / 1000

# Calculate truck food ratio
data['truck_food_ratio'] = data['truck_food_mt'] / (data['truck_weight_kg'] / 1000)
data['truck_food_ratio'].fillna(0, inplace=True)

# Replace infinite values with zero
data['truck_food_ratio'].replace([np.inf, -np.inf], 0, inplace=True)

# =====================
# Save the updated data with the original sheet name
# =====================
output_file = os.path.join(data_dir, "unrwa_trucks.xlsx")
with pd.ExcelWriter(output_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    data.to_excel(writer, sheet_name='unrwa_trucks_kcal_mt', index=False)

# Archive the workbook
archive_dir = os.path.join(data_dir, "archive")
os.makedirs(archive_dir, exist_ok=True)
timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
archive_file_path = os.path.join(archive_dir, f'unrwa_trucks_{timestamp}.xlsx')
shutil.copy(output_file, archive_file_path)

print(f'Processing and saving completed. Archived as {archive_file_path}.')
