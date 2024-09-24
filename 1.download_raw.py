import gdown
import os
import platform
import pandas as pd
from datetime import datetime

# ==================
# DOWNLOAD DATA FROM GOOGLE DRIVE
# ==================

# Google Drive file ID for the input file (do not change)
file_id = '19oQZt7zWE29hK6Whnr9zop4gIGUValfxK14fQVHW18s'
download_url = f'https://drive.google.com/uc?id={file_id}'

# Get the current date in the format YYYYMMDD
current_date = datetime.now().strftime('%Y%m%d')

# Determine the user's desktop location (macOS and Windows compatibility)
if platform.system() == "Darwin":  # macOS
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
elif platform.system() == "Windows":  # Windows
    desktop = os.path.join(os.environ["HOMEPATH"], "Desktop")
else:
    raise Exception("Unsupported operating system. This script works on macOS and Windows only.")

# Create a folder on the desktop with the format "UNRWA Truck Data_YYYYMMDD"
folder_name = f"UNRWA Truck Data_{current_date}"
output_dir = os.path.join(desktop, folder_name)
os.makedirs(output_dir, exist_ok=True)

# Download the file to the new folder on the desktop
output_file = os.path.join(output_dir, "unwra_trucks_raw.xlsx")
gdown.download(download_url, output_file, quiet=False, use_cookies=True)

# Check if the file was downloaded successfully
try:
    raw_df = pd.read_excel(output_file)
    print(f"Raw data file successfully saved as {output_file}")
except Exception as e:
    print(f'Error reading the downloaded file: {e}')
    exit()

# ==================
# PROCESS DATA
# ==================
# (Perform any data processing here if needed)
# For this example, let's just simulate processed data by modifying the raw data
processed_df = raw_df.copy()
processed_df["Processed"] = True  # Example modification for processed data

# ==================
# SAVE BOTH RAW AND PROCESSED DATA TO A SINGLE EXCEL FILE WITH MULTIPLE SHEETS
# ==================

final_output_file = os.path.join(output_dir, f"unwra_trucks_{current_date}.xlsx")

# Use pandas ExcelWriter to save both raw and processed data to one Excel file with multiple sheets
with pd.ExcelWriter(final_output_file, engine='xlsxwriter') as writer:
    raw_df.to_excel(writer, sheet_name='Raw Data', index=False)
    processed_df.to_excel(writer, sheet_name='Processed Data', index=False)

print(f"Final Excel file with multiple sheets saved as {final_output_file}")
