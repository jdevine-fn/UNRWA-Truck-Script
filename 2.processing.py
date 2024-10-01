import pandas as pd
import numpy as np
import os
from datetime import datetime
import platform

# =====================
# STEP 2: PROCESS DATA
# =====================

# Get the current date in the format YYYYMMDD
current_date = datetime.now().strftime('%Y%m%d')

# Determine the user's desktop location (macOS and Windows compatibility)
if platform.system() == "Darwin":  # macOS
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
elif platform.system() == "Windows":  # Windows
    desktop = os.path.join(os.environ["HOMEPATH"], "Desktop")
else:
    raise Exception("Unsupported operating system. This script works on macOS and Windows only.")

# Path to the folder created by step 1
folder_name = f"UNRWA Truck Data_{current_date}"
data_dir = os.path.join(desktop, folder_name)

# Input file (output from Step 1)
file_path = os.path.join(data_dir, "unrwa_trucks_raw.xlsx")

# Output file path
output_file_path = os.path.join(data_dir, "unrwa_trucks.xlsx")

# Archive folder within the same directory
archive_dir = os.path.join(data_dir, "archive")
os.makedirs(archive_dir, exist_ok=True)

# Archive the existing output file if it exists
if os.path.exists(output_file_path):
    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    archive_file_path = os.path.join(archive_dir, f"unrwa_trucks_{timestamp}.xlsx")
    os.rename(output_file_path, archive_file_path)

# Load the Excel file from the "Supply Page" sheet
data = pd.read_excel(file_path, sheet_name='Supply Page')

# Debug: Print column names
print("Column names after loading the file:", data.columns)

# Strip any leading/trailing whitespace from column names
data.columns = data.columns.str.strip()

# Rename 'Units' to 'unit'
data.rename(columns={'Units': 'unit'}, inplace=True)
print("Column names after renaming 'Units' to 'unit':", data.columns)

# Convert 'Quantity' to numeric
data['Quantity'] = pd.to_numeric(data['Quantity'], errors='coerce')

# Remove observations where 'unit' is 'Pallets' and 'Quantity' is greater than 40
if 'unit' in data.columns:
    data = data[~((data['unit'].str.lower() == 'pallets') & (data['Quantity'] > 40))]
else:
    print("Error: The column 'unit' does not exist after renaming.")
    exit()

# Convert 'Donation Type' to lowercase
data['Donation Type'] = data['Donation Type'].str.lower()

# Combine 'Manifest of' and 'Description of Cargo' into 'cargo' (if needed)
# Since 'Manifest of' and 'Description of Cargo' are separate, we'll use 'Description of Cargo' for cargo description
data['cargo'] = data['Description of Cargo'].str.lower()

# Drop unnecessary columns if needed
data.drop(columns=['Description of Cargo'], inplace=True)

# Convert 'Received Date' to date format and rename to 'date'
data['date'] = pd.to_datetime(data['Received Date'], errors='coerce')
data.drop(columns=['Received Date'], inplace=True)

# Replace NaN in 'unit' with 'Unknown'
data['unit'] = data['unit'].fillna('Unknown')

# Function to clean the item text
def clean_item_text(text):
    if pd.isna(text):
        return text
    return text.strip().replace('(', '').replace(')', '').replace('"', '').strip()

# Split 'cargo' text by '+' or ';' into 'item_' variables
cargo_split = data['cargo'].str.replace('+', ';').str.split(';', expand=True)

# Determine max number of items in any entry
max_items = cargo_split.shape[1]

# Create 'item_1' to 'item_N' variables based on max items and populate them
for i in range(max_items):
    data[f'item_{i+1}'] = cargo_split[i].apply(clean_item_text)

# Remove trailing spaces and convert to lowercase for item columns
item_columns = [f'item_{i+1}' for i in range(max_items)]
for col in item_columns:
    data[col] = data[col].str.strip().str.lower()

# Count number of non-blank item_ columns per truck and store in 'item_count'
data['item_count'] = data[item_columns].notna().sum(axis=1)

# Save the processed data to the output file with a clean sheet name
data.to_excel(output_file_path, index=False, sheet_name='unrwa_clean')

print(f"Processing complete. Output saved to: {output_file_path}")
