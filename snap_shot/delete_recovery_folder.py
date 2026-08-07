#!/usr/bin/env python3
"""One-off cleanup: delete the SharePoint _recovery/ folder created during the
2026-08-06 REM Comment wipe investigation. The three RECOVERY snapshots
(v01/v02/v04) served their diagnostic purpose; the workbook has since fully
rebuilt from live data twice with the L2 unit-normalization fix (commit
d5106e8) verified end-to-end.

Deletes the entire _recovery/ folder (folder + contents) via a single
Graph DELETE. Idempotent: prints and exits 0 if the folder is already gone.

Requires APATTISON_MS_REFRESH_TOKEN. Same auth pattern as
snap_shot/list_recovery_folder.py.
"""
import os, json, urllib.request, urllib.parse, urllib.error, sys

TENANT_ID = os.environ.get("MS_TENANT_ID", "b2a05ba0-a5ea-4518-8560-c9d2b631798d")
CLIENT_ID = os.environ.get("MS_CLIENT_ID", "d4aec1ec-50cc-46ee-b17e-dddedb03513b")
REFRESH_TOKEN = os.environ["APATTISON_MS_REFRESH_TOKEN"]
SITE_ID = "prudentgrowthnc.sharepoint.com,135b93a5-2bc7-4ca8-affe-e10b7296d932,10aeddee-8db3-42e9-a9cc-0c32b2592d34"
DRIVE_ID = "b!pZNbE8crqEyv_uELcpbZMu7drhCzjelCqcwMMrJZLTSYJfXQhQnGRo9Sea2xcjR0"
SCOPE = "https://graph.microsoft.com/.default offline_access"

FOLDER_PATH = "General/_Prudent Growth Operations, LLC/LEASING & ASSET MANAGEMENT/WORKING DOCS/_recovery"


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


def get_folder(token):
    path = urllib.parse.quote(FOLDER_PATH, safe="/")
    url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drives/{DRIVE_ID}/root:/{path}"
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {token}"})
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            return json.loads(r.read())
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return None
        raise


def list_children(token, folder_id):
    url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drives/{DRIVE_ID}/items/{folder_id}/children"
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {token}"})
    with urllib.request.urlopen(req, timeout=60) as r:
        data = json.loads(r.read())
    return data.get("value", [])


def delete_item(token, item_id):
    url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drives/{DRIVE_ID}/items/{item_id}"
    req = urllib.request.Request(
        url, method="DELETE", headers={"Authorization": f"Bearer {token}"})
    with urllib.request.urlopen(req, timeout=60) as r:
        return r.status


def main():
    token = get_token()
    folder = get_folder(token)
    if folder is None:
        print(f"_recovery/ folder is already gone. Nothing to do.")
        return 0

    folder_id = folder["id"]
    print(f"Found _recovery/ (id={folder_id})")

    # Preview contents so the workflow log has an audit trail.
    try:
        children = list_children(token, folder_id)
        print(f"Contents to be deleted ({len(children)} items):")
        for c in children:
            name = c.get("name")
            size = c.get("size", 0)
            print(f"  - {name}  ({size:,} bytes)")
    except Exception as e:
        print(f"(could not list children before delete: {e})")

    print("\nDeleting folder recursively...")
    status = delete_item(token, folder_id)
    print(f"Delete returned HTTP {status}")

    # Verify.
    verify = get_folder(token)
    if verify is None:
        print("Verified: _recovery/ no longer exists.")
        return 0
    print("WARNING: folder still exists after delete!", file=sys.stderr)
    return 1


if __name__ == "__main__":
    sys.exit(main())
