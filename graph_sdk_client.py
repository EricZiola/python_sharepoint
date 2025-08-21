"""
This module uses the Microsoft Graph SDK to interact with SharePoint and other Microsoft 365 services.

Beginner Guide:

What the SDK does:
- Makes it easy to authenticate and navigate Microsoft 365 resources (like files, folders, users, sites) using Python objects.
- Lets you list, create, update, and delete files, folders, and other resources.
- Provides metadata about files and folders (names, IDs, URLs, etc.).

What the SDK does NOT do:
- It does NOT directly download or upload file contents.
- You cannot use the SDK alone to get the actual bytes of a file.

How to download files:
- Use the SDK to find the file and get its metadata.
- Look for the '@microsoft.graph.downloadUrl' property in the file's metadata.
- Use the 'requests' library (or another HTTP client) to download the file from that URL.

How to upload files:
- Use the SDK to find the destination folder and get its metadata.
- To upload small files, you can use a simple HTTP PUT request to the appropriate Graph API endpoint (e.g., '/drive/items/{parent-id}:/{filename}:/content').
- For large files, use the Graph API's upload session, which lets you upload in chunks. The SDK can help you create the session, but you use 'requests' or another HTTP client to upload the file data.

Summary:
Use the SDK for authentication and navigation, and use 'requests' for downloading or uploading file contents.

Documentation: https://github.com/microsoftgraph/msgraph-sdk-python/blob/main/README.md
"""
# Import required libraries
import asyncio
import json
import os
import requests
from azure.identity import ClientSecretCredential
from msgraph import GraphServiceClient
from dotenv import load_dotenv

# Import .env environment variables for authentication
load_dotenv(override=True)


# Define Azure AD credentials
AZURE_TENANT_ID = os.environ.get("AZURE_TENANT_ID")
AZURE_CLIENT_ID = os.environ.get("AZURE_CLIENT_ID")
AZURE_CLIENT_SECRET = os.environ.get("AZURE_CLIENT_SECRET")

# Authenticate using Microsoft Identity Platform
# Create a credential object using the Azure AD app's tenant ID, client ID, and client secret.
# This proves your app's identity to Azure.
credential = ClientSecretCredential(
    tenant_id=AZURE_TENANT_ID,
    client_id=AZURE_CLIENT_ID,
    client_secret=AZURE_CLIENT_SECRET
)
scopes = ["https://graph.microsoft.com/.default"]

# Create a Microsoft Graph client using the credential.
# This client will be used to make API calls to Microsoft 365 services.
# Mic
client = GraphServiceClient(credential, scopes)

# Use the Microsoft Graph SDK client to get a list of users.
# client.users.get() returns a coroutine (async function) object.
# to properly use coroutine objects, collect all the asyncronous code into
# an async functions (coroutines) and use the `await` keyword to call the async functions
# all together at the end of the script using an outer `main()` function.
# Finally, only call the "event loop" one time with `asyncio.run(main())` to run all of the
# asynchronous code in the coroutine function.

# coroutine to list all users
async def get_users():
    print("client.users.get() object type:", type(client.users.get))
    print("client.users.get() returns object type: ", type(await client.users.get()))
    user_collection = await client.users.get()
    users_list = user_collection.value
    print("User in users is of type: ", type(users_list[0]))
    return users_list

# coroutine to list all drives
async def get_drives():
    drives = await client.drives.get()
    if drives and drives.value:
        return drives.value

async def get_drive_items(drive_id):
    drive_items = await client.drives.by_drive_id(drive_id).items.by_drive_item_id("root").children.get()
    if drive_items:
        return drive_items.value

# use the `main()` coroutine as a wrapper to call all asyncronous functions
# at once as `asyncio.run()` can only be called an instantiate the event loop
# once per script.
async def main():
    users = await get_users()
    if users:
        user_list = []
        for user in users:
            user_list.append(
                {
                    "name": user.display_name,
                    "email": user.mail
                }
            )
        with open("./jsons/sdk_users.json", "w") as f:
            print("Writing ./jsons/sdk_users.json")
            f.write(json.dumps(user_list))
    
    drives = await get_drives()
    if drives:
        drives_list = []
        for drive in drives:
            drives_list.append(
                {
                    "id": drive.id,
                    "name": drive.name,
                    "description": drive.description,
                    "drive_type": drive.drive_type,
                    "web_url": drive.web_url
                }
            )
        with open("./jsons/sdk_drives.json", "w") as f:
            print("Writing ./jsons/sdk/sdk_drives.json")
            f.write(json.dumps(drives_list))

    drive_items = await get_drive_items("b!RnTwTMVo2EKcymWB2WGv9jAMWmGK0WVIu2fb1Y2f14B8_8WhwEG5Qbu0Ctf16gYA")
    print("Items in drive: ", len(drive_items))
    for drive_item in drive_items:
        if drive_item.file:
            print("File: ", drive_item.name)
        if drive_item.folder:
            print("Folder: ", drive_item.name)
        if drive_item.name == "test.csv":
            print("-" * 10, "test.csv Download URL", "-" * 10)
            print(drive_item.additional_data["@microsoft.graph.downloadUrl"])
            print()
            download_url = drive_item.additional_data["@microsoft.graph.downloadUrl"]
            file_response = requests.get(download_url)
            with open("./downloads/sdk_test.csv", "wb") as f:
                print("Download URL response: ", type(file_response))
                print("Downloading test.txt")
                f.write(file_response.content)
            print("Downloaded sdk_test.csv successfully.")

asyncio.run(main())

print("\n------------------Complete------------------")