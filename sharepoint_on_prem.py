import os
import requests
from requests_ntlm import HttpNtlmAuth

# SharePoint site and file details
site_url = "http://sharepoint/sites/finance"
file_relative_url = "/sites/finance/Shared Documents/report.xlsx"  # Adjust this path
local_file_path = "downloaded_report.xlsx"  # Desired local file name

# SharePoint credentials
username = "SPADMIN"
password = "brained@123"

# Create the full URL to the file
file_url = f"{site_url}{file_relative_url}"

# Initialize the NTLM authentication handler
auth = HttpNtlmAuth(username, password)

# Perform the GET request to download the file
response = requests.get(file_url, auth=auth, stream=True)

# Check if the request was successful
if response.status_code == 200:
    with open(local_file_path, 'wb') as file:
        for chunk in response.iter_content(chunk_size=8192):
            file.write(chunk)
    print(f"File downloaded successfully as '{local_file_path}'")
else:
    print(f"Failed to download file. HTTP status code: {response.status_code}")
