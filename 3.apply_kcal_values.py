import pandas as pd
import numpy as np
import os
import shutil
import requests  # For downloading the kcal_reference.xlsx file from GitHub
from datetime import datetime
import platform

# =====================
# STEP 3: APPLY KCAL VALUES
# =====================

# Get the current date in the format YYYYMMDD (to match the folder created in step 1 and step 2)
current_date = datetime.now().strftime('%Y%m%d')

# Determine the user's desktop location (macOS and Windows compatibility)
if platform.system() == "Darwin":  # macOS
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
elif platform.system() == "Windows":  # Windows
    desktop = os.path.join(os.environ["HOMEPATH"], "Desktop")
else:
    raise Exception("Unsupported operating system. This script works on macOS and Windows only.")

# Path to the folder created by step 1 and step 2 (same date-based folder)
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

# Function to clean the item text
def clean_item_text(text):
    if pd.isna(text):
        return text
    return text.strip().replace('(', '').replace(')', '').replace('"', '')

# Split 'cargo' text by '+' into 'item_' variables
cargo_split = data['cargo'].str.split('+', expand=True)

# Determine max number of items in any entry
max_items = cargo_split.shape[1]

# Create 'item_1' to 'item_N' variables based on max items and populate them
for i in range(max_items):
    data[f'item_{i+1}'] = cargo_split[i].apply(clean_item_text)

# Add kcal information
item_columns = [f'item_{i+1}' for i in range(max_items)]
item_kcal_dict = {}
NA_match_items = []

for column in item_columns:
    if column in data.columns:
        item_kcal_dict[column] = []
        for item in data[column]:
            if pd.isna(item) or item == '':
                item_kcal_dict[column].append(0)
            else:
                match = kcal_ref[kcal_ref['food_item'].str.lower() == str(item).strip().lower()]
                if not match.empty:
                    item_kcal_dict[column].append(match['food_item_kcal'].values[0])
                else:
                    item_kcal_dict[column].append(0)
                    if item not in NA_match_items:
                        NA_match_items.append(item)

# Adding kcal information back to the data
for column, kcal_list in item_kcal_dict.items():
    data[column + '_kcal'] = kcal_list

# Fill missing values in kcal columns with 0
kcal_columns = [f'{col}_kcal' for col in item_columns]
data[kcal_columns] = data[kcal_columns].fillna(0)

# Save the updated data as a new sheet in the existing Excel file
with pd.ExcelWriter(data_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    data.to_excel(writer, sheet_name='unrwa_trucks_kcal', index=False)

# Archive the workbook
archive_dir = os.path.join(data_dir, "archive")
os.makedirs(archive_dir, exist_ok=True)
timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
archive_file_path = os.path.join(archive_dir, f"unrwa_trucks_{timestamp}.xlsx")
shutil.copy(data_path, archive_file_path)

print(f"Processing and saving completed. Archived as {archive_file_path}.")
