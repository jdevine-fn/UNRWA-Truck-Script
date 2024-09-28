# Import missing platform module to resolve the error.
import platform

# Now re-run the revised script.
import pandas as pd
import numpy as np
import os
from datetime import datetime

# =====================
# STEP 2: PROCESS DATA
# =====================

# Get the current date in the format YYYYMMDD (to match the folder created in step 1)
current_date = datetime.now().strftime('%Y%m%d')

# Determine the user's desktop location (macOS and Windows compatibility)
if platform.system() == "Darwin":  # macOS
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
elif platform.system() == "Windows":  # Windows
    desktop = os.path.join(os.environ["HOMEPATH"], "Desktop")
else:
    raise Exception("Unsupported operating system. This script works on macOS and Windows only.")

# Path to the folder created by step 1 (same date-based folder)
folder_name = f"UNRWA Truck Data_{current_date}"
data_dir = os.path.join(desktop, folder_name)

# Input file (output from Step 1)
file_path = os.path.join(data_dir, "unrwa_trucks_raw.xlsx")

# Output file path (will overwrite after processing)
output_file_path = os.path.join(data_dir, "unrwa_trucks.xlsx")

# Archive folder within the same directory
archive_dir = os.path.join(data_dir, "archive")
os.makedirs(archive_dir, exist_ok=True)

# Archive the existing output file if it exists
if os.path.exists(output_file_path):
    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    archive_file_path = os.path.join(archive_dir, f"unrwa_trucks_{timestamp}.xlsx")
    os.rename(output_file_path, archive_file_path)

# Load the Excel file
data = pd.read_excel(file_path)

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
    data = data[~((data['unit'] == 'Pallets') & (data['Quantity'] > 40))]
else:
    print("Error: The column 'unit' does not exist after renaming.")
    exit()

# Convert 'Donation Type' to lowercase
data['Donation Type'] = data['Donation Type'].str.lower()

# Convert 'Description of Cargo' to lowercase and rename to 'cargo'
data['cargo'] = data['Description of Cargo'].str.lower()
data.drop(columns=['Description of Cargo'], inplace=True)

# Convert 'Received Date' to date format and rename to 'date'
data['date'] = pd.to_datetime(data['Received Date'], errors='coerce')
data.drop(columns=['Received Date'], inplace=True)

# Calculate 'truck_weight_kg' based on 'unit' and 'cargo'
def calculate_truck_weight(row):
    if row['unit'] == 'Pallets':
        if isinstance(row['cargo'], str):
            if 'ready meals' in row['cargo'] or 'ready to eat food' in row['cargo'] or 'ready-to-eat food' in row['cargo']:
                return row['Quantity'] * 790
            elif 'flour' in row['cargo']:
                return row['Quantity'] * 1000
            elif 'nutritional supplement' in row['cargo'] or 'nutritional supplements' in row['cargo']:
                return row['Quantity'] * 900
            elif 'date snacks' in row['cargo']:
                return row['Quantity'] * 540
            else:
                return row['Quantity'] * 637.5  # Default pallet weight
        else:
            return row['Quantity'] * 637.5  # Default pallet weight if 'cargo' is not a string
    elif row['unit'] == 'Ton':
        return row['Quantity'] * 1000
    else:
        return 14000  # Default truck weight

# Apply the truck weight calculation
data['truck_weight_kg'] = data.apply(calculate_truck_weight, axis=1)

# Save the processed data to the output file with a clean sheet name
data.to_excel(output_file_path, index=False, sheet_name='unrwa_clean')

print(f"Processing complete. Output saved to: {output_file_path}")
