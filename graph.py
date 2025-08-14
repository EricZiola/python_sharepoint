# Import required libraries
from azure.identity import ClientSecretCredential
from msgraph import GraphServiceClient
from dotenv import load_dotenv
import requests
import os
import json

load_dotenv(override=True)

# -----------------------------
# Step 1: Define Azure AD credentials
# -----------------------------
TENANT_ID = os.environ.get("AZURE_TENANT_ID")
CLIENT_ID = os.environ.get("AZURE_CLIENT_ID")
CLIENT_SECRET = os.environ.get("AZURE_CLIENT_SECRET")
# -----------------------------
# Step 2: Authenticate using Microsoft Identity Platform
# -----------------------------
credential = ClientSecretCredential(
    tenant_id=TENANT_ID,
    client_id=CLIENT_ID,
    client_secret=CLIENT_SECRET
)
scopes = ["https://graph.microsoft.com/.default"]
client = GraphServiceClient(credential, scopes)
token = credential.get_token("https://graph.microsoft.com/.default")

# Step 2: Define file location
base_url = os.environ.get("SHAREPOINT_BASE_URL")
site_path = os.environ.get("SHAREPOINT_SYSTEMTWO_SITE_PATH")

# Step 3: Get site ID
site_info = requests.get(
    f"https://graph.microsoft.com/v1.0/sites/{base_url}:/{site_path}",
    headers={"Authorization": f"Bearer {token.token}"}
).json()
site_id = site_info["id"]
print("Site ID: ", site_id)


# Step 4: Get drive ID
drive_info = requests.get(
    f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive",
    headers={"Authorization": f"Bearer {token.token}"}
).json()
drive_id = drive_info["id"]
print("Drive ID: ", drive_id)


response = requests.get(
    f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root/children",
    headers={"Authorization": f"Bearer {token.token}"}
)

response_json = response.json()
print(type(response_json))

download_url = None

for item in response_json.get('value', []):
    if item.get('name') == 'test.csv':
        download_url = item.get('@microsoft.graph.downloadUrl')
        print("Download URL", download_url)
        break

if not download_url:
    print("test.csv not found.")

if download_url:
    file_response = requests.get(download_url)
    with open("./downloads/test.csv", "wb") as f:
        f.write(file_response.content)
    print("Downloaded test.csv successfully.")


# # Step 5: Get file metadata
# file_info = requests.get(
#     f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{file_path}",
#     headers={"Authorization": f"Bearer {token.token}"}
# ).json()
# # print(file_info)
