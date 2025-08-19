# Documentation: https://github.com/microsoftgraph/msgraph-sdk-python/blob/main/README.md
# Import required libraries
import asyncio
import requests
import os
import json
from azure.identity import ClientSecretCredential
from msgraph import GraphServiceClient
from dotenv import load_dotenv

load_dotenv(override=True)

# -----------------------------
# Step 1: Define Azure AD credentials
# -----------------------------
AZURE_TENANT_ID = os.environ.get("AZURE_TENANT_ID")
AZURE_CLIENT_ID = os.environ.get("AZURE_CLIENT_ID")
AZURE_CLIENT_SECRET = os.environ.get("AZURE_CLIENT_SECRET")
# -----------------------------
# Step 2: Authenticate using Microsoft Identity Platform
# -----------------------------
# Create a credential object using your Azure AD app's tenant ID, client ID, and client secret.
# This proves your app's identity to Azure.
credential = ClientSecretCredential(
    tenant_id=AZURE_TENANT_ID,
    client_id=AZURE_CLIENT_ID,
    client_secret=AZURE_CLIENT_SECRET
)

# Create a Microsoft Graph client using the credential.
# This client will be used to make API calls to Microsoft 365 services.
client = GraphServiceClient(credential)

# Request an access token object from Azure for Microsoft Graph.
# This token is needed to authenticate API requests.
token_object = credential.get_token("https://graph.microsoft.com/.default")
token = token_object.token
print(type(token_object))

# Step 2: Define file location
base_url = os.environ.get("SHAREPOINT_BASE_URL")
site_path = os.environ.get("SHAREPOINT_SYSTEMTWO_SITE_PATH")

# Step 3: Get site ID
site_info = requests.get(
    f"https://graph.microsoft.com/v1.0/sites/{base_url}:/{site_path}",
    headers={"Authorization": f"Bearer {token}"}
).json()
site_id = site_info["id"]
print("Site ID: ", site_id)


# Step 4: Get the site default document drive ID
drive_info = requests.get(
    f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive",
    headers={"Authorization": f"Bearer {token}"}
).json()
drive_id = drive_info["id"]
print("Site Drive ID: ", drive_id)

# Get root folder (Shared Documents) information
response_root = requests.get(
    f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root/children",
    headers={"Authorization": f"Bearer {token}"}
).json()
with open ("./jsons/response_root.json", "w", encoding="utf-8") as f:
    print("Writing response_root.json")
    f.write(json.dumps(response_root))

# Get child folder "Benchmarking Data" information
folder_path = "Benchmarking Data"
response_child = requests.get(
    f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{folder_path}:/children",
    headers={"Authorization": f"Bearer {token}"}
).json()
with open ("./jsons/response_child.json", "w", encoding="utf-8") as f:
    print("Writing response_child.json")
    f.write(json.dumps(response_child))

download_url = None
for item in response_root.get('value', []):
    if item.get('name') == 'test.csv':
        download_url = item.get('@microsoft.graph.downloadUrl')
        break
if not download_url:
    print("test.csv not found.")
if download_url:
    file_response = requests.get(download_url)
    with open("./downloads/test.csv", "wb") as f:
        print("Download URL response: ", type(file_response))
        print("Downloading test.txt")
        f.write(file_response.content)
    print("Downloaded test.csv successfully.")

# Step 5: Get file metadata
file_path = "test.csv"
file_info = requests.get(
    f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{file_path}",
    headers={f"Authorization": f"Bearer {token}"}
).json()
with open("./jsons/file_info.json", "w") as f:
    print("Writing file_info.json")
    f.write(json.dumps(file_info))

users = requests.get(
    "https://graph.microsoft.com/v1.0/users",
    headers={f"Authorization": f"Bearer {token}"}
).json()
with open("./jsons/entra_users.json", "w") as f:
    print("Writing users.json")
    f.write(json.dumps(users))

users_sdk = client.users.get()
print(type(users_sdk))

async def get_users():
    users = await users_sdk
    print("Returned by get(): ", type(users))
    print("users.value: ", type(users.value))
    print("users.value[0]: ", type(users.value[0]))
    for user in users.value:
        print(user.display_name)

users = asyncio.run(get_users())
print("------------------Complete------------------")