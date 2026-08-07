#!/usr/bin/env python3
"""List the SharePoint _recovery/ folder to confirm uploads landed and print
their webUrls."""
import os, json, urllib.request, urllib.parse

TENANT_ID = os.environ.get("MS_TENANT_ID", "b2a05ba0-a5ea-4518-8560-c9d2b631798d")
CLIENT_ID = os.environ.get("MS_CLIENT_ID", "d4aec1ec-50cc-46ee-b17e-dddedb03513b")
REFRESH_TOKEN = os.environ["APATTISON_MS_REFRESH_TOKEN"]
SITE_ID = "prudentgrowthnc.sharepoint.com,135b93a5-2bc7-4ca8-affe-e10b7296d932,10aeddee-8db3-42e9-a9cc-0c32b2592d34"
DRIVE_ID = "b!pZNbE8crqEyv_uELcpbZMu7drhCzjelCqcwMMrJZLTSYJfXQhQnGRo9Sea2xcjR0"
SCOPE = "https://graph.microsoft.com/.default offline_access"

def get_token():
    data = urllib.parse.urlencode({
        "client_id": CLIENT_ID, "refresh_token": REFRESH_TOKEN,
        "grant_type": "refresh_token", "scope": SCOPE,
    }).encode()
    req = urllib.request.Request(
        f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token",
        data=data, method="POST")
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.loads(r.read())["access_token"]

def main():
    token = get_token()
    folder = "General/_Prudent Growth Operations, LLC/LEASING & ASSET MANAGEMENT/WORKING DOCS/_recovery"
    path = urllib.parse.quote(folder, safe="/")
    url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drives/{DRIVE_ID}/root:/{path}:/children"
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {token}"})
    with urllib.request.urlopen(req, timeout=60) as r:
        data = json.loads(r.read())
    items = data.get("value", [])
    print(f"Found {len(items)} items in _recovery/:\n")
    for it in items:
        name = it.get("name")
        size = it.get("size")
        modified = it.get("lastModifiedDateTime")
        web = it.get("webUrl")
        print(f"  {name}")
        print(f"    size={size:,} modified={modified}")
        print(f"    {web}")
        print()

if __name__ == "__main__":
    main()
