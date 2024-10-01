import os
import pandas as pd
import numpy as np
from datetime import datetime
import openpyxl
import platform

# =====================
# STEP 5: DAILY SUMMARY CALCULATIONS
# =====================

# Get the current date in the format YYYYMMDD (to match the folder created in steps 1-4)
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

# File path (input from step 4)
file_path = os.path.join(data_dir, "unrwa_trucks.xlsx")

# Load the data from the 'unrwa_trucks_kcal' sheet generated by script 4
data = pd.read_excel(file_path, sheet_name='unrwa_trucks_kcal')

# Check if the 'Crossing' column exists (formerly 'entry')
if 'Crossing' not in data.columns:
    print("Warning: The 'Crossing' column does not exist in the dataset. Please check the data or previous processing steps.")
else:
    print("The 'Crossing' column exists.")

# Insert the check for 'truck_kcal' here
if 'truck_kcal' not in data.columns:
    raise KeyError("The 'truck_kcal' column does not exist in the dataset. Please check the data or previous processing steps.")

# Ensure 'truck_type' exists in the data
if 'truck_type' not in data.columns:
    raise KeyError("The 'truck_type' column does not exist in the dataset. Please check the data or previous processing steps.")

# Ensure the 'sector' column exists
if 'sector' not in data.columns:
    raise KeyError("The 'sector' column does not exist in the dataset. Please check the data or previous processing steps.")

# Create a new dataframe named `data_daily`
data_daily = pd.DataFrame()

# In `data_daily`, create a `date` column
data_daily['date'] = data['date']

# Ensure only numeric columns are summed
numeric_cols = data.select_dtypes(include=[np.number]).columns

# Generate daily totals in `data_daily` using `date` from `data`
data_daily_totals = data.groupby('date')[numeric_cols].sum().reset_index()

# Merge the daily totals into `data_daily`
data_daily = pd.merge(data_daily[['date']], data_daily_totals, on='date', how='left')

# Compute `count_daily_truck_mixed`
data_daily['count_daily_truck_mixed'] = data[data['truck_type'] == 'Mixed Food/Non-Food Truck'].groupby('date').size().reindex(data_daily['date'], fill_value=0).values

# Compute `count_daily_truck_food`
data_daily['count_daily_truck_food'] = data[data['truck_type'] == 'Food Truck'].groupby('date').size().reindex(data_daily['date'], fill_value=0).values

# Compute `count_daily_truck_nonfood`
data_daily['count_daily_truck_nonfood'] = data[data['truck_type'] == 'Non-Food Truck'].groupby('date').size().reindex(data_daily['date'], fill_value=0).values

# Compute `count_daily_sector_humanitarian`
data_daily['count_daily_sector_humanitarian'] = data[data['sector'] == 'humanitarian'].groupby('date').size().reindex(data_daily['date'], fill_value=0).values

# Compute `daily_kcal`
data_daily['daily_kcal'] = data.groupby('date')['truck_kcal'].sum().reindex(data_daily['date'], fill_value=0).values

# Compute `daily_food_mt`
data_daily['daily_food_mt'] = data.groupby('date')['truck_food_mt'].sum().reindex(data_daily['date'], fill_value=0).values

# Compute `daily_mt`
data_daily['daily_mt'] = data.groupby('date')['truck_weight_kg'].sum().reindex(data_daily['date'], fill_value=0).values / 1000

# Only attempt to compute 'Crossing' related fields if the 'Crossing' column exists
if 'Crossing' in data.columns:
    # Compute `entry_kerem_count`
    data_daily['entry_kerem_count'] = data[data['Crossing'] == 'Kerem Shalom'].groupby('date').size().reindex(data_daily['date'], fill_value=0).values

    # Compute `entry_rafah_count`
    data_daily['entry_rafah_count'] = data[data['Crossing'] == 'Rafah'].groupby('date').size().reindex(data_daily['date'], fill_value=0).values

# Compute `cargo_type_unrwa_food_count`
data_daily['cargo_type_unrwa_food_count'] = data[data['cargo'] == 'unrwa_food'].groupby('date').size().reindex(data_daily['date'], fill_value=0).values

# Compute `cargo_type_unrwa_nonfood_count`
data_daily['cargo_type_unrwa_nonfood_count'] = data[data['cargo'] == 'unrwa_nonfood'].groupby('date').size().reindex(data_daily['date'], fill_value=0).values

# Compute `cargo_type_unrwa_medical_count`
data_daily['cargo_type_unrwa_medical_count'] = data[data['cargo'] == 'unrwa_medical'].groupby('date').size().reindex(data_daily['date'], fill_value=0).values

# Save `data_daily` with all columns in `unrwa_trucks.xlsx` as a new sheet `unrwa_daily_entries`
with pd.ExcelWriter(file_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
    data_daily.to_excel(writer, sheet_name='unrwa_daily_entries', index=False)

print("Processing complete and data saved to 'unrwa_daily_entries' sheet.")