"""
One-shot cleanup — delete the misplaced Leasing-Snap-Shot.xlsx from the
old SharePoint location. Run once after fixing the folder path, then
delete this file + its workflow.
"""
import os
import sys

import requests

MS_TENANT_ID = "b2a05ba0-a5ea-4518-8560-c9d2b631798d"
MS_CLIENT_ID = "d4aec1ec-50cc-46ee-b17e-dddedb03513b"
MS_REFRESH_TOKEN = os.environ["APATTISON_MS_REFRESH_TOKEN"].strip()

DRIVE_ID = "b!pZNbE8crqEyv_uELcpbZMu7drhCzjelCqcwMMrJZLTSYJfXQhQnGRo9Sea2xcjR0"
OLD_FILE_PATH = "General/Brain Snap Shot/Leasing-Snap-Shot.xlsx"


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


def main():
    token = get_access_token()
    url = (
        f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}"
        f"/root:/{OLD_FILE_PATH}"
    )
    print(f"Looking up: {url}")
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=30)
    if r.status_code == 404:
        print("Nothing to delete — old file not found (already gone or moved).")
        return
    r.raise_for_status()
    item = r.json()
    print(f"Found: {item.get('name')} ({item.get('size')} bytes, id={item.get('id')})")

    del_url = f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/items/{item['id']}"
    r2 = requests.delete(del_url, headers={"Authorization": f"Bearer {token}"}, timeout=30)
    if r2.status_code in (200, 204):
        print("DELETED old file.")
    else:
        print(f"Delete failed: {r2.status_code} {r2.text[:300]}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
