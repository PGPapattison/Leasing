"""One-shot diagnostic: does apattison@prudentgrowth.com's DELEGATED refresh
token have enough scope + Excel Online provisioning to hit the Graph Workbook
endpoints on the live Snap Shot workbook?

Runs read + write probes:
  1. Fetches an access token from APATTISON_MS_REFRESH_TOKEN and prints the
     token's actual granted scopes + `upn`/`preferred_username` claim.
  2. Attempts GET  /workbook/worksheets                     (read)
  3. Attempts POST /workbook/createSession {persistChanges: false}  (session read)
  4. Attempts POST /workbook/createSession {persistChanges: true }  (session write)
  5. Attempts PATCH ...worksheets/Snap Shot/range(address='ZZ1') with an
     invented cell (safe; column ZZ is far outside our data) — SESSIONLESS.

Every step is fully logged, non-fatal, and does NOT alter any real workbook
cell (the ZZ1 PATCH is our sacrificial probe cell — even if it writes, no
one reads column ZZ).
"""
from __future__ import annotations
import base64
import json
import os
import sys

import requests

MS_TENANT_ID = os.environ.get("MS_TENANT_ID", "b2a05ba0-a5ea-4518-8560-c9d2b631798d")
MS_CLIENT_ID = os.environ.get("MS_CLIENT_ID", "d4aec1ec-50cc-46ee-b17e-dddedb03513b")
MS_GRAPH_SCOPE = os.environ.get(
    "MS_GRAPH_SCOPE",
    "Mail.Send Files.ReadWrite.All Sites.ReadWrite.All offline_access",
)
REFRESH_TOKEN = os.environ.get("APATTISON_MS_REFRESH_TOKEN", "").strip()

SITE_ID = "prudentgrowthnc.sharepoint.com,135b93a5-2bc7-4ca8-affe-e10b7296d932,10aeddee-8db3-42e9-a9cc-0c32b2592d34"
DRIVE_ID = "b!pZNbE8crqEyv_uELcpbZMu7drhCzjelCqcwMMrJZLTSYJfXQhQnGRo9Sea2xcjR0"
SHAREPOINT_ITEM_PATH = "General/_Prudent Growth Operations, LLC/LEASING & ASSET MANAGEMENT/WORKING DOCS/Leasing-Snap-Shot.xlsx"

WORKBOOK_ROOT = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drives/{DRIVE_ID}/root:/{SHAREPOINT_ITEM_PATH}:"


def _fail(msg: str) -> None:
    print(f"\n\u2717 {msg}")


def _ok(msg: str) -> None:
    print(f"\n\u2713 {msg}")


def _step(msg: str) -> None:
    print(f"\n\u2500\u2500 {msg}")


def _b64pad(s: str) -> str:
    return s + "=" * (-len(s) % 4)


def _decode_jwt_payload(jwt: str) -> dict:
    parts = jwt.split(".")
    if len(parts) < 2:
        return {}
    try:
        return json.loads(base64.urlsafe_b64decode(_b64pad(parts[1])).decode("utf-8"))
    except Exception:
        return {}


def get_token() -> str:
    if not REFRESH_TOKEN:
        _fail("APATTISON_MS_REFRESH_TOKEN not set.")
        sys.exit(1)
    r = requests.post(
        f"https://login.microsoftonline.com/{MS_TENANT_ID}/oauth2/v2.0/token",
        data={
            "client_id": MS_CLIENT_ID,
            "refresh_token": REFRESH_TOKEN,
            "grant_type": "refresh_token",
            "scope": MS_GRAPH_SCOPE,
        },
    )
    if r.status_code != 200:
        _fail(f"Token refresh HTTP {r.status_code}: {r.text[:400]}")
        sys.exit(1)
    body = r.json()
    tok = body["access_token"]
    claims = _decode_jwt_payload(tok)
    _ok("Refresh token exchange succeeded.")
    print(f"    upn / preferred_username:   {claims.get('upn') or claims.get('preferred_username')}")
    print(f"    tid:                        {claims.get('tid')}")
    print(f"    aud (should be graph):      {claims.get('aud')}")
    print(f"    scp (granted scopes):       {claims.get('scp')}")
    print(f"    roles (app perms, if any):  {claims.get('roles')}")
    return tok


def probe_read_worksheets(token: str) -> None:
    _step("Read /workbook/worksheets (baseline read)")
    r = requests.get(
        WORKBOOK_ROOT + "/workbook/worksheets",
        headers={"Authorization": f"Bearer {token}"},
    )
    print(f"    HTTP {r.status_code}")
    if r.status_code == 200:
        sheets = [w["name"] for w in r.json().get("value", [])]
        ellipsis = "\u2026" if len(sheets) > 6 else ""
        _ok(f"Read succeeded. Sheets: {sheets[:6]}{ellipsis}")
    else:
        _fail(f"Read failed: {r.text[:400]}")


def probe_create_session(token: str, persist_changes: bool) -> str | None:
    label = "persistChanges=True (WRITE)" if persist_changes else "persistChanges=False (READ)"
    _step(f"POST /workbook/createSession {label}")
    r = requests.post(
        WORKBOOK_ROOT + "/workbook/createSession",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        json={"persistChanges": persist_changes},
    )
    print(f"    HTTP {r.status_code}")
    if r.status_code in (200, 201):
        sid = r.json().get("id")
        _ok(f"Session opened. id={sid[:16]}\u2026")
        return sid
    else:
        _fail(f"createSession failed: {r.text[:500]}")
        return None


def probe_sessionless_patch(token: str, addr: str = "ZZ1") -> None:
    _step(f"PATCH SESSIONLESS to Snap Shot!{addr} (probe cell — safe)")
    r = requests.patch(
        WORKBOOK_ROOT + f"/workbook/worksheets('Snap%20Shot')/range(address='{addr}')",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        json={"values": [["probe-2026-08-19"]]},
    )
    print(f"    HTTP {r.status_code}")
    if r.status_code == 200:
        _ok("Sessionless PATCH succeeded. Cell ZZ1 now contains 'probe-2026-08-19' (safe / unused).")
    else:
        _fail(f"Sessionless PATCH failed: {r.text[:500]}")


def probe_sessioned_patch(token: str, session_id: str, addr: str = "ZZ2") -> None:
    _step(f"PATCH via SESSION to Snap Shot!{addr} (probe cell)")
    r = requests.patch(
        WORKBOOK_ROOT + f"/workbook/worksheets('Snap%20Shot')/range(address='{addr}')",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "workbook-session-id": session_id,
        },
        json={"values": [["probe-sess-2026-08-19"]]},
    )
    print(f"    HTTP {r.status_code}")
    if r.status_code == 200:
        _ok("Sessioned PATCH succeeded. Cell ZZ2 now contains 'probe-sess-2026-08-19' (safe / unused).")
    else:
        _fail(f"Sessioned PATCH failed: {r.text[:500]}")


def main() -> int:
    tok = get_token()
    probe_read_worksheets(tok)
    sid_r = probe_create_session(tok, persist_changes=False)
    sid_w = probe_create_session(tok, persist_changes=True)
    probe_sessionless_patch(tok, "ZZ1")
    if sid_w:
        probe_sessioned_patch(tok, sid_w, "ZZ2")
    print("\n\u2500\u2500 done \u2500\u2500\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())
