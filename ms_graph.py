import os
import sys
import json
import requests
from msal import PublicClientApplication
from dotenv import load_dotenv
from bs4 import BeautifulSoup
import pandas as pd

os.environ["PYTHONUTF8"] = "1"


load_dotenv()
CLIENT_ID = os.getenv("CLIENT_ID")  # Application (client) ID
TENANT_ID = os.getenv("TENANT_ID", "common")  # or your tenant GUID / domain
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["User.Read", "Mail.Read"]

if not CLIENT_ID:
    print("CLIENT_ID is missing. Put it in your .env as CLIENT_ID=...")
    sys.exit(1)


app = PublicClientApplication(client_id=CLIENT_ID, authority=AUTHORITY)

# 1) Try cached token

result = None
accounts = app.get_accounts()
if accounts:
    result = app.acquire_token_silent(SCOPES, account=accounts[0])


# 2) Interactive login (loopback port)

if not result:
    try:
        print("Opening browser for interactive login on http://localhost:8000 ...")
        result = app.acquire_token_interactive(scopes=SCOPES, port=8000)
    except TypeError:
        result = app.acquire_token_interactive(scopes=SCOPES)


# 3) Device Code Flow fallback

if not result or "access_token" not in result:
    print("Interactive login failed or returned no token. Falling back to Device Code Flow...")
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        print("Failed to create device flow:", json.dumps(flow, indent=2))
        sys.exit(1)
    print("\nDEVICE CODE FLOW")
    print(flow["message"])
    result = app.acquire_token_by_device_flow(flow)


if "access_token" not in result:
    print("Failed to acquire token:", json.dumps(result, indent=2))
    sys.exit(1)

access_token = result["access_token"]
headers = {"Authorization": f"Bearer {access_token}"}


sender_email = "kfupm-stu@kfupm.edu.sa"  # student affairs email change this if needed
url = f"https://graph.microsoft.com/v1.0/me/messages?$top=50&$filter=from/emailAddress/address eq '{sender_email}'"
resp = requests.get(url, headers=headers)

if not resp.ok:
    print("Graph API error:", resp.status_code, resp.text)
    sys.exit(1)

data = resp.json()



rows = []
for msg in data.get("value", []):
    sender = (msg.get("from") or {}).get("emailAddress", {}).get("address", "")
    subject = msg.get("subject", "")
    received = msg.get("receivedDateTime", "")
    body_html = msg.get("body", {}).get("content", "")

    # Parse HTML to find all images
    soup = BeautifulSoup(body_html, "html.parser")
    images = [img.get("src") for img in soup.find_all("img") if img.get("src")]

    # Take only the second image if it exists
    second_image = images[1] if len(images) > 1 else ""

    rows.append({
        "Sender": sender,
        "Subject": subject,
        "Time Sent": received,
        "Image": second_image  # only second image
    })

# Save to Excel

df = pd.DataFrame(rows)
#df.to_excel("emails_with_second_image.xlsx", index=False, engine="openpyxl")
#print("Saved emails to emails_with_second_image.xlsx")
