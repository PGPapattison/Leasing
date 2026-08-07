#!/usr/bin/env python3
"""One-off recovery script: list SharePoint version history for the Snap Shot
workbook and download the previous version (before the overnight rebuild that
wiped REM Comments). Uploads the recovered file as a GitHub Actions artifact.

Graph reference: GET /drives/{drive-id}/items/{item-id}/versions
                 GET /drives/{drive-id}/items/{item-id}/versions/{ver-id}/content
"""
import os
import sys
import json
import time
import urllib.request
import urllib.parse

TENANT_ID = os.environ.get("MS_TENANT_ID", "b2a05ba0-a5ea-4518-8560-c9d2b631798d")
CLIENT_ID = os.environ.get("MS_CLIENT_ID", "d4aec1ec-50cc-46ee-b17e-dddedb03513b")
REFRESH_TOKEN = os.environ["APATTISON_MS_REFRESH_TOKEN"]
SITE_ID = "prudentgrowthnc.sharepoint.com,135b93a5-2bc7-4ca8-affe-e10b7296d932,10aeddee-8db3-42e9-a9cc-0c32b2592d34"
DRIVE_ID = "b!pZNbE8crqEyv_uELcpbZMu7drhCzjelCqcwMMrJZLTSYJfXQhQnGRo9Sea2xcjR0"
ITEM_ID = "01JHOGWCBJX3IJNDMFHZCJ3FMWF76FWP4D"
SCOPE = "https://graph.microsoft.com/.default offline_access"


def get_token():
    data = urllib.parse.urlencode({
        "client_id": CLIENT_ID,
        "refresh_token": REFRESH_TOKEN,
        "grant_type": "refresh_token",
        "scope": SCOPE,
    }).encode()
    req = urllib.request.Request(
        f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token",
        data=data, method="POST",
    )
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.loads(r.read())["access_token"]


def graph_get(token, url):
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {token}"})
    with urllib.request.urlopen(req, timeout=60) as r:
        return r


def main():
    token = get_token()
    print("[ok] Got Graph token")

    # List versions
    url = f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/items/{ITEM_ID}/versions"
    r = graph_get(token, url)
    data = json.loads(r.read())
    versions = data.get("value", [])
    print(f"[ok] Found {len(versions)} versions")
    print()
    for i, v in enumerate(versions):
        vid = v.get("id")
        modified = v.get("lastModifiedDateTime")
        who = v.get("lastModifiedBy", {}).get("user", {}).get("displayName", "?")
        size = v.get("size", "?")
        print(f"  [{i}] id={vid}  modified={modified}  by={who}  size={size}")
    print()

    if not versions:
        print("[fatal] No versions returned")
        sys.exit(1)

    # Save the version manifest as an artifact for cross-reference
    os.makedirs("recovery_out", exist_ok=True)
    with open("recovery_out/versions.json", "w") as f:
        json.dump(versions, f, indent=2, default=str)

    # Download the top 5 most recent versions — we can pick the right one after
    for i, v in enumerate(versions[:6]):
        vid = v.get("id")
        modified = v.get("lastModifiedDateTime", "unknown").replace(":", "").replace("-", "")
        safe_ts = modified.replace("T", "_").split(".")[0].rstrip("Z")
        dl_url = f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/items/{ITEM_ID}/versions/{vid}/content"
        try:
            r = graph_get(token, dl_url)
            content = r.read()
            fname = f"recovery_out/snap_v{i:02d}_{safe_ts}.xlsx"
            with open(fname, "wb") as f:
                f.write(content)
            print(f"[ok] Downloaded v{i} → {fname} ({len(content):,} bytes)")
        except Exception as e:
            print(f"[warn] Failed to download v{i}: {e}")
        time.sleep(1)


if __name__ == "__main__":
    main()
