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
    df = pd.read_excel(output_file)
    print(f"File successfully saved as {output_file}")
except Exception as e:
    print(f'Error reading the downloaded file: {e}')
    exit()

# ==================
# PROCESS DATA
# ==================
# (Perform any data processing here if needed)
# For example: Cleaning, filtering, etc.

# Save the processed file in the same folder on the desktop
processed_file = os.path.join(output_dir, f"unwra_trucks_processed_{current_date}.xlsx")
df.to_excel(processed_file, index=False)
print(f"Processed file saved locally as {processed_file}")

