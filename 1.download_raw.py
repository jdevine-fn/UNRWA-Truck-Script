import gdown
import os
import pandas as pd
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
from datetime import datetime
from getpass import getpass

# ==================
# DOWNLOAD DATA FROM GOOGLE DRIVE
# ==================

# Google Drive file ID for the input file
file_id = '19oQZt7zWE29hK6Whnr9zop4gIGUValfxK14fQVHW18s'  # Update with actual file ID if needed
download_url = f'https://drive.google.com/uc?id={file_id}'

# Set up the local data directory (relative to the repository location)
data_dir = os.path.join(os.getcwd(), "data")
os.makedirs(data_dir, exist_ok=True)

# Download the file to the local data directory
output_file = os.path.join(data_dir, "unwra_trucks_raw.xlsx")
gdown.download(download_url, output_file, quiet=False, use_cookies=False)

# Check if the file was downloaded successfully
try:
    df = pd.read_excel(output_file)
    print(f"File successfully saved as {output_file}")
except Exception as e:
    print(f'Error reading the downloaded file: {e}')
    exit()

# ==================
# UPLOAD DATA TO SHAREPOINT
# ==================

# Set up SharePoint credentials and folder
sharepoint_url = "https://chemonics.sharepoint.com/sites/FEWSNET_Technical_Team"
sharepoint_folder = "/sites/FEWSNET_Technical_Team/Shared Documents/02.Markets_and_Trade/04.Reports/06.Special_reports/06. Gaza Food Supply Reports 2024/Master Data Workbooks"

# Use interactive login for SharePoint (OAuth)
username = input("Enter your SharePoint username (email): ")
password = getpass("Enter your SharePoint password: ")

# Authenticate to SharePoint using user credentials (OAuth)
ctx = ClientContext(sharepoint_url).with_credentials(UserCredential(username, password))

# Prepare the file for upload (for example, you can upload the raw file or a processed version)
sharepoint_file_name = f"unwra_trucks_processed_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx"

# Upload the file to SharePoint
with open(output_file, 'rb') as file_content:
    target_url = f"{sharepoint_folder}/{sharepoint_file_name}"
    ctx.web.get_folder_by_server_relative_url(sharepoint_folder).upload_file(sharepoint_file_name, file_content.read()).execute_query()

print(f"File successfully uploaded to SharePoint as {sharepoint_file_name}")
