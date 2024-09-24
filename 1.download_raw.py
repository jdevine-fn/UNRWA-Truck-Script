import gdown
import os
import pandas as pd
from datetime import datetime
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive

# ==================
# DOWNLOAD DATA FROM GOOGLE DRIVE
# ==================

# Google Drive file ID for the input file (do not change)
file_id = '19oQZt7zWE29hK6Whnr9zop4gIGUValfxK14fQVHW18s'
download_url = f'https://drive.google.com/uc?id={file_id}'

# Set up the local data directory (relative to the repository location)
data_dir = os.path.join(os.getcwd(), "data")
os.makedirs(data_dir, exist_ok=True)

# Download the file to the local data directory
output_file = os.path.join(data_dir, "unwra_trucks_raw.xlsx")
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
# (Perform any data processing here as needed)
# For example: Cleaning, filtering, etc.

# Save the processed file locally
processed_dir = os.path.join(data_dir, "processed")
os.makedirs(processed_dir, exist_ok=True)

processed_file = os.path.join(processed_dir, f"unwra_trucks_processed_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx")
df.to_excel(processed_file, index=False)
print(f"Processed file saved locally as {processed_file}")

# ==================
# UPLOAD TO SPECIFIC GOOGLE DRIVE FOLDER
# ==================

# Authenticate and initialize PyDrive
gauth = GoogleAuth()
gauth.LocalWebserverAuth()  # Creates local webserver and automatically handles authentication
drive = GoogleDrive(gauth)

# Google Drive folder ID (replace this with the folder link's ID)
folder_id = '1PwHYUHx7T7Ey8MgSlXf4zT0vF2sOoLBb'

# Create a file and set the parent folder ID
gdrive_file = drive.CreateFile({'title': f'unwra_trucks_processed_{datetime.now().strftime("%Y%m%d%H%M%S")}.xlsx', 'parents': [{'id': folder_id}]})
gdrive_file.SetContentFile(processed_file)
gdrive_file.Upload()

print(f"Processed file uploaded to Google Drive folder with title: {gdrive_file['title']}")
