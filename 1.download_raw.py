import gdown
import os
import platform
from datetime import datetime

# ==================
# DOWNLOAD DATA FROM GOOGLE DRIVE WITHOUT MODIFICATION
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

# Ensure the directory exists
os.makedirs(output_dir, exist_ok=True)

# Define the final output file name
output_file = os.path.join(output_dir, "unrwa_trucks_raw.xlsx")

# Download the file using gdown (without any modifications to structure or content)
gdown.download(download_url, output_file, quiet=False, use_cookies=False, verify=False)

# The file is downloaded as-is, retaining the original structure, including multiple sheets and formatting.
print(f"File successfully downloaded and saved as {output_file} without any alterations.")
