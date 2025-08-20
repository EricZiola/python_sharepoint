"""
This module uses the Microsoft Graph SDK to interact with SharePoint
Documentation: https://github.com/microsoftgraph/msgraph-sdk-python/blob/main/README.md
"""
# Import required libraries
import asyncio
import json
import os
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

# Create a Microsoft Graph client using the credential.
# This client will be used to make API calls to Microsoft 365 services.
# Mic
client = GraphServiceClient(credential)

# Use the Microsoft Graph SDK client to get a list of users.
# client.users.get() returns a coroutine (async function) object.
# to properly use coroutine objects, collect all the asyncronous code into
# an async function (coroutine) and use the `await` keyword to call the async functions
# finally, only call the "event loop" one time with `asyncio.run()` to run all of the
# asynchronous code in the coroutine function.
async def main():
    print("client.users.get() object type:", type(client.users.get))
    print("client.users.get() returns object type: ", type(await client.users.get()))
    user_collection = await client.users.get()
    users_list = user_collection.value
    print("User in users is of type: ", type(users_list[0]))
    return users_list

# coroutine objects must be run with an "event loop" that is instantiated with asyncio.run()
user_list = []
users = asyncio.run(main())
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

print("------------------Complete------------------")