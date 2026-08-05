"""
One-shot resolver — takes the SharePoint folder sharing URL and prints
the driveId + parent folder path in a format the rebuild.py script can use.

Run via workflow_dispatch: gh workflow run "Snap Shot — resolve folder"
"""
import base64
import os
import sys

import requests

MS_TENANT_ID = "b2a05ba0-a5ea-4518-8560-c9d2b631798d"
MS_CLIENT_ID = "d4aec1ec-50cc-46ee-b17e-dddedb03513b"
MS_REFRESH_TOKEN = os.environ["APATTISON_MS_REFRESH_TOKEN"].strip()

SHARING_URL = "https://prudentgrowthnc.sharepoint.com/:f:/s/DBMigration/IgDIiDV_fAJOQ7j1Jb_C0cRdAcpKF7xIgqwXrjDSeb1Skqs?e=DOQgHv"


def get_access_token() -> str:
    r = requests.post(
        f"https://login.microsoftonline.com/{MS_TENANT_ID}/oauth2/v2.0/token",
        data={
            "client_id": MS_CLIENT_ID,
            "refresh_token": MS_REFRESH_TOKEN,
            "grant_type": "refresh_token",
            "scope": "Files.ReadWrite.All Sites.ReadWrite.All offline_access",
        },
        timeout=30,
    )
    r.raise_for_status()
    return r.json()["access_token"]


def encode_share_url(url: str) -> str:
    return "u!" + base64.urlsafe_b64encode(url.encode()).decode().rstrip("=")


def main():
    token = get_access_token()
    encoded = encode_share_url(SHARING_URL)

    r = requests.get(
        f"https://graph.microsoft.com/v1.0/shares/{encoded}/driveItem",
        headers={"Authorization": f"Bearer {token}"},
        timeout=30,
    )
    if r.status_code != 200:
        print(f"ERROR resolving share: {r.status_code} {r.text[:500]}", file=sys.stderr)
        sys.exit(1)

    item = r.json()
    print("=" * 70)
    print("RESOLVED FOLDER")
    print("=" * 70)
    print(f"Folder name       : {item.get('name')}")
    print(f"Folder id         : {item.get('id')}")
    print(f"Drive id          : {item.get('parentReference', {}).get('driveId')}")
    print(f"Drive type        : {item.get('parentReference', {}).get('driveType')}")
    print(f"Site id           : {item.get('parentReference', {}).get('siteId')}")
    print(f"Parent path       : {item.get('parentReference', {}).get('path')}")
    print(f"Full webUrl       : {item.get('webUrl')}")
    print()

    parent_path = item.get("parentReference", {}).get("path", "")
    if parent_path.startswith("/drive/root:"):
        server_relative = parent_path[len("/drive/root:"):]
    else:
        server_relative = parent_path
    folder_path = f"{server_relative}/{item.get('name')}".lstrip("/")
    print("=" * 70)
    print("USE THESE VALUES IN THE WORKFLOW / rebuild.py")
    print("=" * 70)
    print(f"SHAREPOINT_DRIVE_ID    = {item.get('parentReference', {}).get('driveId')}")
    print(f"SHAREPOINT_SITE_ID     = {item.get('parentReference', {}).get('siteId')}")
    print(f"SHAREPOINT_FOLDER_PATH = {folder_path}")

    print()
    print("Also listing folder contents to confirm it's the right one:")
    r2 = requests.get(
        f"https://graph.microsoft.com/v1.0/drives/{item.get('parentReference', {}).get('driveId')}"
        f"/items/{item.get('id')}/children?$select=name,size,folder,file&$top=50",
        headers={"Authorization": f"Bearer {token}"},
        timeout=30,
    )
    if r2.ok:
        for c in r2.json().get("value", []):
            kind = "DIR" if c.get("folder") else "FILE"
            print(f"  [{kind}] {c.get('name')}")


if __name__ == "__main__":
    main()
