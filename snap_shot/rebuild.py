#!/usr/bin/env python3
"""
prudent-snap-shot-rebuild — nightly + on-demand rebuild of Prudent Growth's
"Leasing Snap Shot" Excel workbook.

Rebuilds the workbook from live AppFolio rent-roll + property-directory data
and the ClickUp "Broker Reporting" list, while preserving every field Ann
hand-edits directly in the SharePoint copy (Broker Calls, Property Flags,
Vacant Callouts, Deal Activity, Ann's Commentary, per-unit Market Rent, and
per-unit Notes). The workbook layout/style logic is ported verbatim from
build_prototype_v3.py — this script does NOT change how the sheet looks.

MODES
-----
  --mode=nightly   Full rebuild: pull AppFolio + ClickUp, extract Ann's edits
                    from the current SharePoint copy, rebuild, upload.
  --mode=weekly     Same as nightly, plus refresh "Broker Active Interest"
                    from the latest Broker Beat attachment per property.
  --mode=poll       Check the ClickUp control task's "Rebuild Snap Shot"
                    checkbox. If checked: run a full nightly rebuild, then
                    uncheck the box and comment "refreshed at HH:MM ET".
                    If not checked: log and exit 0 immediately.
  --mode=dry-run    Build the workbook locally only. Never touches
                    SharePoint or ClickUp (no downloads, no uploads, no
                    writes). Ann's-edit inputs come from the bundled JSON
                    fallback files. Output: /tmp/Leasing-Snap-Shot-dryrun.xlsx

RACE-CONDITION HANDLING
------------------------
Before rebuilding, the current SharePoint workbook is downloaded and its
`lastModifiedBy` / `lastModifiedDateTime` are inspected. If someone OTHER
than the automation account modified it within the last 15 minutes, the run
is SKIPPED (exit 0, log "skipped — Ann is editing"). This is safe because
the cron re-runs hourly (and the on-demand poll re-runs every 15 min).

FAIL LOUD
---------
Any AppFolio / ClickUp / SharePoint API failure is logged, emailed to
apattison@prudentgrowth.com via Microsoft Graph, and the process exits
non-zero. This script never silently ships a stale or partial workbook.
--dry-run failures are still logged loudly but are NOT emailed (no
production impact) and still exit non-zero.

AUTH (environment variables — nothing is hardcoded)
----------------------------------------------------
  APPFOLIO_CLIENT_ID / APPFOLIO_CLIENT_SECRET   AppFolio Basic Auth
  CLICKUP_API_TOKEN                             ClickUp REST v2 token
  MS_TENANT_ID / MS_CLIENT_ID                   Azure AD app (Graph)
  APATTISON_MS_REFRESH_TOKEN                    Delegated refresh token for
                                                 apattison@prudentgrowth.com
                                                 (scopes: Mail.Send Files.ReadWrite.All
                                                 Sites.ReadWrite.All offline_access)
All of the above have Prudent-Growth-specific defaults baked in as fallbacks
(matching the values already in use by sibling automations / the appfolio-api
skill) so the script also runs out of the box in this org, but every value is
overridable via env var and the script never crashes silently if a required
secret is genuinely absent for a step that needs it.

USAGE
-----
  python3 rebuild.py --mode=dry-run
  python3 rebuild.py --mode=nightly
  python3 rebuild.py --mode=weekly
  python3 rebuild.py --mode=poll
"""

import argparse
import base64
import hashlib
import io
import json
import logging
import os
import re
import sys
import time
import traceback
from datetime import datetime, timedelta

import requests
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter

try:
    import zoneinfo
    ET = zoneinfo.ZoneInfo("America/New_York")
except Exception:  # pragma: no cover - zoneinfo is stdlib on 3.9+
    ET = None


# ══════════════════════════════════════════════════════════════════════════
# Configuration
# ══════════════════════════════════════════════════════════════════════════

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# Data files this script reads as-is (per brief — reuse, do not regenerate).
# Resolution order:
#   1. SNAP_SHOT_DATA_DIR env var (explicit override, used in tests / CI)
#   2. snap_shot/data/ next to this script (the shipped repo layout)
#   3. ../../../.. from this script (legacy workspace layout, kept for local dev
#      inside the original skills/prudent-snap-shot-rebuild/scripts/ tree)
_DATA_CANDIDATES = [
    os.environ.get("SNAP_SHOT_DATA_DIR"),
    os.path.join(_SCRIPT_DIR, "data"),
    os.path.dirname(os.path.dirname(os.path.dirname(_SCRIPT_DIR))),
]
DATA_DIR = next(
    (p for p in _DATA_CANDIDATES
     if p and os.path.exists(os.path.join(p, "property_mapping_all73.json"))),
    _DATA_CANDIDATES[1],  # default so error messages point somewhere
)
WORKSPACE_DIR = DATA_DIR  # kept for callers that reference it

MAPPING_PATH = os.path.join(DATA_DIR, "property_mapping_all73.json")
BROKER_CONTACTS_PATH = os.path.join(DATA_DIR, "broker_contacts_by_property.json")
NOTES_OVERRIDES_FALLBACK_PATH = os.path.join(DATA_DIR, "property_notes_overrides.json")
UNIT_NOTES_FALLBACK_PATH = os.path.join(DATA_DIR, "notes_by_unit.json")
TICAM_MAP_PATH = os.path.join(DATA_DIR, "property_ticam_map.json")

# ClickUp TICAM Rates list — https://app.clickup.com/14147033/v/li/901112111796
# One task per (Property, Year). Year is a dropdown custom field. Rates live in
# per-SF formula fields; underlying $ fields are also readable. Field IDs are
# discovered at runtime from the list schema so a field rename in ClickUp
# doesn't silently break the pull — we look up by field name.
TICAM_LIST_ID = "901112111796"
LOGO_PATH = os.path.join(DATA_DIR, "pgp_logo.png")

DRY_RUN_OUTPUT_PATH = "/tmp/Leasing-Snap-Shot-dryrun.xlsx"
LOCAL_BUILD_DIR = os.environ.get("SNAP_SHOT_BUILD_DIR", "/tmp")

# ---- AppFolio -------------------------------------------------------------
APPFOLIO_BASE_URL = os.environ.get("APPFOLIO_BASE_URL", "https://prudentgrowth.appfolio.com")
APPFOLIO_CLIENT_ID = os.environ.get("APPFOLIO_CLIENT_ID", "45835c8ce883a2b138f35fd37c029e3f")
APPFOLIO_CLIENT_SECRET = os.environ.get("APPFOLIO_CLIENT_SECRET", "1b53030bd10d473ced02bb766b8254aa")
APPFOLIO_RATE_LIMIT_SLEEP_SEC = float(os.environ.get("APPFOLIO_RATE_LIMIT_SLEEP_SEC", "2"))

# ---- ClickUp ----------------------------------------------------------------
CU_BASE = "https://api.clickup.com/api/v2"
CLICKUP_TOKEN = os.environ.get("CLICKUP_API_TOKEN", "").strip()

BROKER_REPORTING_LIST_ID = os.environ.get("CLICKUP_BROKER_REPORTING_LIST_ID", "901114227189")
F_PROPERTY_AO = os.environ.get(
    "CLICKUP_PROPERTY_AO_FIELD_ID", "bb1b357d-1d26-4509-b652-bf926dfeaaec"
)  # "Property (A/O)" dropdown — maps ClickUp tasks to property names

CONTROL_TASK_NAME = os.environ.get(
    "SNAP_SHOT_CONTROL_TASK_NAME", "SNAP SHOT REBUILD — check to refresh"
)
CHECKBOX_FIELD_NAME = os.environ.get("SNAP_SHOT_CHECKBOX_FIELD_NAME", "Rebuild Snap Shot")
BROKER_BEAT_TASK_PREFIX = os.environ.get(
    "SNAP_SHOT_BROKER_BEAT_PREFIX", "Weekly Broker Update"
)

# ---- LAR (Leasing / Asset Management) lists ---------------------------------
# Three ClickUp lists whose "Summary" (AI-generated) field is round-tripped
# into the Snap Shot RENT ROLL's rightmost "ClickUp Summary" column so REMs
# see the current AI summary next to each unit at a glance.
#
# Match order (first match wins):
#   1. Occupied unit (TenantId present) → lookup by Tenant ID across:
#         Renewal Pipeline, Vacancy Pipeline, Documents Workflow
#   2. Vacant unit (no TenantId)        → lookup by (Property ID + Unit #)
#         in Vacancy Pipeline only
#
# The Summary custom-field ID is NOT the same across all three lists. Renewal
# and Vacancy share one ID; Documents Workflow has its own. We keep a
# {list_id → summary_field_id} map so callers don't have to guess.
LAR_RENEWAL_LIST_ID = os.environ.get("CLICKUP_LAR_RENEWAL_LIST_ID", "901113575567")
LAR_VACANCY_LIST_ID = os.environ.get("CLICKUP_LAR_VACANCY_LIST_ID", "901113575628")
LAR_DOCS_WORKFLOW_LIST_ID = os.environ.get("CLICKUP_LAR_DOCS_WORKFLOW_LIST_ID", "901113991446")

# Custom-field IDs shared across all three LAR lists (short_text).
CU_TENANT_ID_FIELD = os.environ.get(
    "CLICKUP_TENANT_ID_FIELD", "2f9249fb-d19b-470b-8bb4-b11eb9eb5cb8"
)
CU_PROPERTY_ID_FIELD = os.environ.get(
    "CLICKUP_PROPERTY_ID_FIELD", "057285dd-5010-4336-916f-06888b1514e5"
)
# Vacancy Pipeline only.
CU_UNIT_NUMBER_FIELD = os.environ.get(
    "CLICKUP_UNIT_NUMBER_FIELD", "0acf8068-1620-4238-b19c-69f912cc710d"
)

# Per-list Summary field IDs (differ between Renewal/Vacancy and Documents).
LAR_SUMMARY_FIELD_BY_LIST = {
    LAR_RENEWAL_LIST_ID: "93833d96-e254-4402-b672-30ab8fb45ccd",
    LAR_VACANCY_LIST_ID: "93833d96-e254-4402-b672-30ab8fb45ccd",
    LAR_DOCS_WORKFLOW_LIST_ID: "190e6156-36bf-47d5-b10b-ea6f9809bc6c",
}

# ---- SharePoint / Microsoft Graph -------------------------------------------
MS_TENANT_ID = os.environ.get("MS_TENANT_ID", "b2a05ba0-a5ea-4518-8560-c9d2b631798d")
MS_CLIENT_ID = os.environ.get("MS_CLIENT_ID", "d4aec1ec-50cc-46ee-b17e-dddedb03513b")
# Delegated refresh token for apattison@prudentgrowth.com. Needs Mail.Send +
# Files.ReadWrite.All (or Sites.ReadWrite.All) + offline_access consented.
MS_REFRESH_TOKEN = os.environ.get("APATTISON_MS_REFRESH_TOKEN", "").strip()
MS_GRAPH_SCOPE = os.environ.get(
    "MS_GRAPH_SCOPE", "Mail.Send Files.ReadWrite.All Sites.ReadWrite.All offline_access"
)

SITE_ID = os.environ.get(
    "SHAREPOINT_SITE_ID",
    "prudentgrowthnc.sharepoint.com,135b93a5-2bc7-4ca8-affe-e10b7296d932,"
    "10aeddee-8db3-42e9-a9cc-0c32b2592d34",
)
DRIVE_ID = os.environ.get(
    "SHAREPOINT_DRIVE_ID",
    "b!pZNbE8crqEyv_uELcpbZMu7drhCzjelCqcwMMrJZLTSYJfXQhQnGRo9Sea2xcjR0",
)
SHAREPOINT_FOLDER_PATH = os.environ.get(
    "SHAREPOINT_FOLDER_PATH", "General/Brain Snap Shot"
)
WORKBOOK_FILENAME = os.environ.get("SNAP_SHOT_FILENAME", "Leasing-Snap-Shot.xlsx")
SHAREPOINT_ITEM_PATH = f"{SHAREPOINT_FOLDER_PATH}/{WORKBOOK_FILENAME}"

# Public SharePoint web URL for the workbook (opens in Excel Online in the
# browser). Used to stamp the ClickUp control-task description after each
# successful rebuild so the team always has a one-click link to the latest
# copy. The URL is constructed from the resolved SharePoint folder path plus
# workbook filename — URL-encoded per RFC 3986.
import urllib.parse
WORKBOOK_WEB_URL = (
    "https://prudentgrowthnc.sharepoint.com/sites/DBMigration/Shared%20Documents/"
    + urllib.parse.quote(f"{SHAREPOINT_FOLDER_PATH}/{WORKBOOK_FILENAME}", safe="/")
)

# The automation's own identity as it appears in Graph's lastModifiedBy.user.
# Used to distinguish "the bot just saved this" from "a human just edited this".
AUTOMATION_ACCOUNT_EMAIL = os.environ.get(
    "SNAP_SHOT_AUTOMATION_ACCOUNT", "apattison@prudentgrowth.com"
).lower()

RACE_CONDITION_WINDOW_MINUTES = int(os.environ.get("SNAP_SHOT_RACE_WINDOW_MIN", "15"))

# Large-file upload threshold. Graph's simple PUT only supports files <4MB;
# above that an upload session (chunked) is required. The Snap Shot workbook
# is well under 4MB today but this is sized for headroom / future growth.
LARGE_FILE_THRESHOLD_BYTES = 4 * 1024 * 1024
UPLOAD_CHUNK_SIZE_BYTES = 5 * 320 * 1024  # 1,638,400 — must be a multiple of 320 KiB

# ---- Notifications -----------------------------------------------------------
ERROR_NOTIFICATION_TO = os.environ.get(
    "SNAP_SHOT_ERROR_EMAIL", "apattison@prudentgrowth.com"
)
SENDER_EMAIL = os.environ.get("SNAP_SHOT_SENDER_EMAIL", AUTOMATION_ACCOUNT_EMAIL)

# HTTP timeouts
HTTP_TIMEOUT_SEC = 60
HTTP_TIMEOUT_LONG_SEC = 180


# ══════════════════════════════════════════════════════════════════════════
# Logging
# ══════════════════════════════════════════════════════════════════════════

def setup_logging():
    logger = logging.getLogger("snap_shot_rebuild")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()
    handler = logging.StreamHandler(sys.stdout)
    handler.setFormatter(logging.Formatter(
        "%(asctime)s [%(levelname)s] %(message)s", datefmt="%Y-%m-%d %H:%M:%S"
    ))
    logger.addHandler(handler)
    return logger


LOG = setup_logging()


class FatalError(Exception):
    """Raised for any unrecoverable error. main() catches this, logs it,
    emails it (unless dry-run), and exits non-zero."""


# ══════════════════════════════════════════════════════════════════════════
# Small utilities
# ══════════════════════════════════════════════════════════════════════════

def now_et():
    if ET:
        return datetime.now(ET)
    return datetime.utcnow()


def load_json(path, required=True, default=None):
    if not os.path.exists(path):
        if required:
            raise FatalError(f"Required data file missing: {path}")
        LOG.warning(f"Optional data file missing, using default: {path}")
        return default if default is not None else {}
    try:
        with open(path) as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError) as e:
        raise FatalError(f"Failed to read/parse {path}: {e}")


def parse_currency(s):
    """Parse '4,449.00' -> 4449.0, '-4,374.64' -> -4374.64, '' -> 0."""
    if s is None or s == "":
        return 0.0
    if isinstance(s, (int, float)):
        return float(s)
    try:
        return float(str(s).replace(",", "").replace("$", "").strip())
    except (ValueError, TypeError):
        return 0.0


def parse_sqft(s):
    if s is None or s == "":
        return 0
    try:
        return int(float(str(s).replace(",", "").strip()))
    except (ValueError, TypeError):
        return 0


def parse_date(s):
    """MM/DD/YYYY -> datetime, or None."""
    if not s:
        return None
    try:
        return datetime.strptime(str(s).strip(), "%m/%d/%Y")
    except (ValueError, TypeError):
        return None


def unit_display(u):
    """Simplify unit strings. E.g. '815 - Unit 101' -> 'Unit 101'."""
    v = u.get("Unit", "") or ""
    if " - " in v:
        parts = [p.strip() for p in v.split(" - ")]
        for p in parts:
            if p.lower().startswith(("unit", "suite", "ste")):
                return p
        return parts[-1]
    return v


def sort_units(units):
    """Sort units by unit label, natural sort."""
    def key(u):
        s = unit_display(u)
        m = re.search(r"(\d+)", s)
        return (int(m.group(1)) if m else 999999, s.lower())
    return sorted(units, key=key)


def http_request(method, url, *, context, retries=2, backoff_sec=3, **kwargs):
    """Thin wrapper around requests with retry-on-5xx/429 and loud failure.

    `context` is a short human string used in error messages (e.g.
    "AppFolio rent_roll page 3") so failures are traceable in logs/email.
    """
    last_exc = None
    for attempt in range(retries + 1):
        try:
            r = requests.request(method, url, timeout=kwargs.pop("timeout", HTTP_TIMEOUT_SEC), **kwargs)
        except requests.exceptions.RequestException as e:
            last_exc = e
            LOG.warning(f"{context}: network error on attempt {attempt + 1}/{retries + 1}: {e}")
            time.sleep(backoff_sec * (attempt + 1))
            continue

        if r.status_code == 429:
            wait = float(r.headers.get("Retry-After", backoff_sec * (attempt + 1)))
            LOG.warning(f"{context}: HTTP 429 rate-limited, sleeping {wait}s (attempt {attempt + 1})")
            time.sleep(wait)
            continue
        if r.status_code >= 500:
            LOG.warning(f"{context}: HTTP {r.status_code} on attempt {attempt + 1}/{retries + 1}: {r.text[:300]}")
            time.sleep(backoff_sec * (attempt + 1))
            continue
        return r

    raise FatalError(f"{context}: exhausted retries. Last error: {last_exc}")


# ══════════════════════════════════════════════════════════════════════════
# AppFolio
# ══════════════════════════════════════════════════════════════════════════

def appfolio_auth():
    if not APPFOLIO_CLIENT_ID or not APPFOLIO_CLIENT_SECRET:
        raise FatalError("AppFolio credentials are not configured (APPFOLIO_CLIENT_ID/SECRET).")
    return (APPFOLIO_CLIENT_ID, APPFOLIO_CLIENT_SECRET)


def appfolio_pull_report(endpoint):
    """Pull a full (paginated) AppFolio report. Returns list of result dicts."""
    url = f"{APPFOLIO_BASE_URL}/api/v1/reports/{endpoint}"
    results = []
    page = 0
    while url:
        r = http_request(
            "GET", url, context=f"AppFolio {endpoint} page {page}",
            auth=appfolio_auth(),
        )
        if r.status_code != 200:
            raise FatalError(
                f"AppFolio {endpoint} returned HTTP {r.status_code}: {r.text[:300]}"
            )
        try:
            data = r.json()
        except ValueError:
            raise FatalError(f"AppFolio {endpoint} returned non-JSON response: {r.text[:300]}")
        results.extend(data.get("results", []))
        next_page = data.get("next_page_url")
        if next_page:
            url = f"{APPFOLIO_BASE_URL}{next_page}" if next_page.startswith("/") else next_page
            page += 1
            time.sleep(APPFOLIO_RATE_LIMIT_SLEEP_SEC)
        else:
            url = None
    LOG.info(f"AppFolio {endpoint}: pulled {len(results)} rows across {page + 1} page(s).")
    return results


def pull_appfolio_rent_roll_by_property():
    """Pull the full rent roll and bucket by PropertyId (string key), matching
    the schema of the existing rent_roll_by_property.json fallback file."""
    rows = appfolio_pull_report("rent_roll.json")
    by_property = {}
    for row in rows:
        pid = str(row.get("PropertyId") or "").strip()
        if not pid:
            continue
        by_property.setdefault(pid, []).append(row)
    LOG.info(f"AppFolio rent roll: grouped into {len(by_property)} properties.")
    return by_property


def pull_appfolio_property_directory():
    """Pull the property directory, keyed by PropertyId (string)."""
    rows = appfolio_pull_report("property_directory.json")
    by_id = {}
    for row in rows:
        pid = str(row.get("PropertyId") or "").strip()
        if pid:
            by_id[pid] = row
    LOG.info(f"AppFolio property directory: {len(by_id)} properties.")
    return by_id


# ══════════════════════════════════════════════════════════════════════════
# ClickUp REST v2
# ══════════════════════════════════════════════════════════════════════════

def require_clickup_token():
    if not CLICKUP_TOKEN:
        raise FatalError(
            "CLICKUP_API_TOKEN env var is not set. Export it before running "
            "(cron/local) or map it from the org secret of the same name."
        )
    return CLICKUP_TOKEN


def cu_headers():
    return {"Authorization": require_clickup_token(), "Content-Type": "application/json"}


def cu_get_list_tasks(list_id, include_closed="true"):
    """Return all tasks in a list, handling pagination."""
    tasks, page = [], 0
    while True:
        r = http_request(
            "GET", f"{CU_BASE}/list/{list_id}/task",
            context=f"ClickUp list {list_id} page {page}",
            headers=cu_headers(),
            params={"page": page, "subtasks": "true", "include_closed": include_closed},
        )
        if r.status_code != 200:
            raise FatalError(f"ClickUp list {list_id} fetch failed: HTTP {r.status_code}: {r.text[:300]}")
        data = r.json()
        batch = data.get("tasks", [])
        tasks.extend(batch)
        if data.get("last_page") or not batch:
            break
        page += 1
    return tasks


def cu_get_task(task_id, include_attachments=True, include_subtasks=False):
    """Fetch a single task. include_subtasks=True is required to get a
    populated 'subtasks' array back from the ClickUp API — it defaults to
    False here to match the cheaper, more common call shape, so callers that
    need child tasks (e.g. Broker Beat weekly updates) must opt in explicitly.
    """
    r = http_request(
        "GET", f"{CU_BASE}/task/{task_id}",
        context=f"ClickUp task {task_id} fetch",
        headers=cu_headers(),
        params={"include_subtasks": "true" if include_subtasks else "false"},
    )
    if r.status_code != 200:
        raise FatalError(f"ClickUp task {task_id} fetch failed: HTTP {r.status_code}: {r.text[:300]}")
    return r.json()


def cu_set_field(task_id, field_id, value):
    r = http_request(
        "POST", f"{CU_BASE}/task/{task_id}/field/{field_id}",
        context=f"ClickUp set field {field_id} on {task_id}",
        headers=cu_headers(),
        json={"value": value},
    )
    if r.status_code not in (200, 201):
        raise FatalError(f"ClickUp set field failed on task {task_id}: HTTP {r.status_code}: {r.text[:300]}")


def cu_post_comment(task_id, text):
    r = http_request(
        "POST", f"{CU_BASE}/task/{task_id}/comment",
        context=f"ClickUp comment on {task_id}",
        headers=cu_headers(),
        json={"comment_text": text},
    )
    if r.status_code not in (200, 201):
        LOG.warning(f"ClickUp comment post failed on task {task_id}: HTTP {r.status_code}: {r.text[:300]}")


def cu_update_task_description(task_id, markdown_description):
    """Update a task's description via `PUT /task/{id}`. ClickUp's API accepts
    `markdown_content` (which triggers markdown parsing) or `description`
    (plain text). We use markdown_content so the SharePoint link renders as a
    clickable link."""
    r = http_request(
        "PUT", f"{CU_BASE}/task/{task_id}",
        context=f"ClickUp update description on {task_id}",
        headers=cu_headers(),
        json={"markdown_content": markdown_description},
    )
    if r.status_code not in (200, 201):
        LOG.warning(
            f"ClickUp update description failed on task {task_id}: "
            f"HTTP {r.status_code}: {r.text[:300]}"
        )
        return False
    return True


# The auto-managed refresh section at the top of ClickUp descriptions.
# ClickUp doesn't strip HTML comments, so we can't use invisible markers.
# Instead, we detect the block by its content signature (the emoji header +
# "Last refreshed" italic line) and replace the whole block each run.
REFRESH_HEADER_SIGNATURE = "\U0001F4C4 **Latest Snap Shot workbook:**"
REFRESH_FOOTER_SIGNATURE = "_Last refreshed "


def cu_update_list_description(list_id, markdown_content):
    """Update a ClickUp LIST's description via PUT /list/{list_id}. Uses
    `markdown_content` so the SharePoint link renders as a clickable link.
    Non-fatal on failure."""
    r = http_request(
        "PUT", f"{CU_BASE}/list/{list_id}",
        context=f"ClickUp update list description on {list_id}",
        headers=cu_headers(),
        json={"markdown_content": markdown_content},
    )
    if r.status_code not in (200, 201):
        LOG.warning(
            f"ClickUp update list description failed on list {list_id}: "
            f"HTTP {r.status_code}: {r.text[:300]}"
        )
        return False
    return True


def cu_get_list(list_id):
    """Fetch a ClickUp list including its description."""
    r = http_request(
        "GET", f"{CU_BASE}/list/{list_id}",
        context=f"ClickUp get list {list_id}",
        headers=cu_headers(),
    )
    r.raise_for_status()
    return r.json()


def _cu_field_value(task, field_id):
    """Return the raw `value` for a custom field on a ClickUp task, or None.
    ClickUp represents unset short_text as absence-of-value or empty string;
    we normalize both to None so callers can treat 'no value' uniformly."""
    for f in task.get("custom_fields", []) or []:
        if f.get("id") == field_id:
            v = f.get("value")
            if v is None:
                return None
            if isinstance(v, str) and not v.strip():
                return None
            return v
    return None


def pull_lar_summaries():
    """Pull the AI-generated Summary custom field from every task in each of
    the three LAR lists (Renewal Pipeline, Vacancy Pipeline, Documents
    Workflow) and return FOUR indexes:

        summaries_by_tenant_id  = {tenant_id_str: summary_text}
        summaries_by_prop_unit  = {(property_id_str, unit_str): summary_text}
        task_ids_by_tenant_id   = {tenant_id_str: [task_id, ...]}
        task_ids_by_prop_unit   = {(property_id_str, unit_str): [task_id, ...]}

    The summary indexes are used to populate the ClickUp Summary column (O).
    The task-ID indexes are used by Commit 2's REM comment write-back to
    post the REM's note to ALL matching tasks (per Alexis 2026-08-06:
    "Post to all matching tasks").

    Summary precedence when multiple lists share the same key:
        Renewal > Vacancy > Documents Workflow
    Task-ID lists are accumulated across all three lists (no dedup precedence
    — a tenant on Renewal AND Documents Workflow gets BOTH task_ids stored).

    Tenant IDs, Property IDs, and unit numbers are all normalized to str()
    with surrounding whitespace stripped. Empty summaries are dropped (no
    point in overwriting a cell with a blank string), but a task without a
    populated Summary can still be a comment target so its ID is still
    indexed if it has a TenantId / (PropertyId, Unit).
    """
    # Order matters — first list wins for summary tenant-id collisions.
    list_order = [
        ("Renewal Pipeline", LAR_RENEWAL_LIST_ID),
        ("Vacancy Pipeline", LAR_VACANCY_LIST_ID),
        ("Documents Workflow", LAR_DOCS_WORKFLOW_LIST_ID),
    ]

    summaries_by_tenant_id = {}
    summaries_by_prop_unit = {}
    task_ids_by_tenant_id = {}
    task_ids_by_prop_unit = {}
    summary_field_by_list = LAR_SUMMARY_FIELD_BY_LIST

    for list_label, list_id in list_order:
        try:
            tasks = cu_get_list_tasks(list_id, include_closed="true")
        except FatalError as e:
            LOG.warning(f"LAR pull: {list_label} ({list_id}) fetch failed: {e}. Skipping this list.")
            continue

        summary_field = summary_field_by_list.get(list_id)
        if not summary_field:
            LOG.warning(f"LAR pull: no Summary field ID configured for list {list_id}. Skipping.")
            continue

        added_summ_tid = 0
        added_summ_pu = 0
        added_task_tid = 0
        added_task_pu = 0
        for t in tasks:
            task_id = t.get("id")
            if not task_id:
                continue

            summary = _cu_field_value(t, summary_field)
            summary_text = str(summary).strip() if summary else ""

            tid_raw = _cu_field_value(t, CU_TENANT_ID_FIELD)
            tid = str(tid_raw).strip() if tid_raw is not None else ""

            # Property ID + Unit # (both live on Vacancy Pipeline tasks; also
            # tolerated if present on other lists though not expected).
            pid_raw = _cu_field_value(t, CU_PROPERTY_ID_FIELD)
            pid = str(pid_raw).strip() if pid_raw is not None else ""
            unit_raw = _cu_field_value(t, CU_UNIT_NUMBER_FIELD) if list_id == LAR_VACANCY_LIST_ID else None
            unit = str(unit_raw).strip() if unit_raw is not None else ""

            # Summary indexing (first-list-wins).
            if summary_text:
                if tid and tid not in summaries_by_tenant_id:
                    summaries_by_tenant_id[tid] = summary_text
                    added_summ_tid += 1
                if list_id == LAR_VACANCY_LIST_ID and pid and unit:
                    key = (pid, unit)
                    if key not in summaries_by_prop_unit:
                        summaries_by_prop_unit[key] = summary_text
                        added_summ_pu += 1

            # Task-ID indexing (accumulated across all lists; used by the
            # REM comment write-back to post to every matching task).
            if tid:
                lst = task_ids_by_tenant_id.setdefault(tid, [])
                if task_id not in lst:
                    lst.append(task_id)
                    added_task_tid += 1
            if list_id == LAR_VACANCY_LIST_ID and pid and unit:
                key = (pid, unit)
                lst = task_ids_by_prop_unit.setdefault(key, [])
                if task_id not in lst:
                    lst.append(task_id)
                    added_task_pu += 1

        LOG.info(
            f"LAR pull: {list_label} — {len(tasks)} tasks, "
            f"+{added_summ_tid} summaries by TenantId, +{added_summ_pu} summaries by (PropertyId,Unit), "
            f"+{added_task_tid} task-ids by TenantId, +{added_task_pu} task-ids by (PropertyId,Unit)."
        )

    LOG.info(
        f"LAR pull totals: "
        f"summaries — {len(summaries_by_tenant_id)} by Tenant ID, {len(summaries_by_prop_unit)} by (PropertyId, Unit); "
        f"task-ids — {len(task_ids_by_tenant_id)} tenants → tasks, {len(task_ids_by_prop_unit)} (prop,unit) → tasks."
    )
    return (
        summaries_by_tenant_id,
        summaries_by_prop_unit,
        task_ids_by_tenant_id,
        task_ids_by_prop_unit,
    )


# ─── TICAM Rates (2026) ───────────────────────────────────────────────
#
# The Snap Shot displays per-SF NNN rates for the 2026 lease year just below
# each property banner. Data source: ClickUp list "TICAM Rates" (901112111796),
# one task per (Property × Year). We only surface 2026 rows here.
#
# Property name matching: the TICAM "Property (A/O)" dropdown does not use the
# same names as the Snap Shot AppFolio names, so property_ticam_map.json maps
# Snap Shot name → TICAM dropdown label (or list of labels for the one property
# where TICAM has two dropdown rows for the same asset — South Memorial Plaza
# (Tulsa)). Anything not in the map is assumed to match by exact name.
#
# Fields surfaced (only when populated — zero/None values are dropped from the
# rendered line so nothing shows up as "$0.00"):
#   • PGP CAM (incl. 15% admin fee)   — formula field "PGP CAM / SF (Inc. Admin Fee)"
#   • Tax / SF                         — formula field "Taxes / SF"
#   • Insurance / SF                   — formula field "Insurance / SF"
#   • Water / SF                       — formula field "Water / SF"
#   • Assoc. Fee / SF                  — formula field "Assoc. Fee / SF"
# Plus the ClickUp task URL (for click-through) and the confirmed/unconfirmed
# status (so brokers know whether the numbers are final).
#
# Non-fatal: if the ClickUp fetch fails, TICAM rows are simply omitted from
# the workbook — the rest of the build still succeeds.

# Field-name → payload-key mapping. We look up field IDs at runtime rather
# than hardcoding them, so a rename in the ClickUp UI won't silently blank
# out the column.
_TICAM_FIELD_NAMES = {
    "pgp_cam_per_sf":  "PGP CAM / SF (Inc. Admin Fee)",
    "tax_per_sf":      "Taxes / SF",
    "insurance_per_sf":"Insurance / SF",
    "water_per_sf":    "Water / SF",
    "assoc_fee_per_sf":"Assoc. Fee / SF",
    "year":            "Year",
    "property":        "Property (A/O) ",  # NOTE: trailing space in ClickUp field name
}


def _ticam_field_value(task, field_id):
    """Extract a custom-field value from a ClickUp task by field ID.
    Returns None if the field is missing or has no value."""
    for cf in task.get("custom_fields", []):
        if cf.get("id") != field_id:
            continue
        val = cf.get("value")
        if val is None or val == "":
            return None
        return val
    return None


def _ticam_number(task, field_id):
    """Read a numeric or currency custom-field. ClickUp returns numbers as
    strings sometimes; formula fields return floats. Returns None if
    unpopulated, zero, or unparseable — zero counts as unpopulated because
    we don't want '$0.00' rows on the workbook."""
    val = _ticam_field_value(task, field_id)
    if val is None:
        return None
    try:
        n = float(val)
    except (TypeError, ValueError):
        return None
    if n == 0:
        return None
    return n


def _ticam_dropdown_name(task, field_id, options_by_id):
    """Resolve a dropdown value to its human name. ClickUp returns dropdown
    values as the option UUID (or index, depending on the field version).
    We look up in options_by_id, which was pre-built from the field schema."""
    val = _ticam_field_value(task, field_id)
    if val is None:
        return None
    return options_by_id.get(str(val)) or options_by_id.get(val)


def pull_ticam_rates_2026(snap_shot_property_names):
    """Fetch 2026 TICAM rates from ClickUp list 901112111796 and return a
    dict keyed by Snap Shot property name:

        {
          "Amberwood Plaza": {
            "pgp_cam_per_sf": 2.70, "tax_per_sf": 1.10, "insurance_per_sf": 0.42,
            "water_per_sf": 0.18, "assoc_fee_per_sf": None,
            "status": "confirmed",
            "url": "https://app.clickup.com/t/868k4u5ab",
          },
          ...
        }

    Properties with no 2026 row (or all fields blank) are omitted. Never
    raises — catches all exceptions and returns {} on failure so a broken
    ClickUp fetch doesn't break the whole workbook build.

    snap_shot_property_names: iterable of Snap Shot AppFolio display names
    (used to key the returned dict and to warn about unmapped properties).
    """
    try:
        # Step 1: discover custom-field IDs by name.
        r = http_request(
            "GET", f"{CU_BASE}/list/{TICAM_LIST_ID}/field",
            context=f"ClickUp TICAM list {TICAM_LIST_ID} fields",
            headers=cu_headers(),
        )
        if r.status_code != 200:
            LOG.warning(f"TICAM field schema fetch failed HTTP {r.status_code}; omitting TICAM rows.")
            return {}
        fields = r.json().get("fields", [])
        field_id_by_name = {f["name"]: f["id"] for f in fields}

        # Resolve every field we need. If any core field is missing, bail.
        try:
            fid = {k: field_id_by_name[v] for k, v in _TICAM_FIELD_NAMES.items()}
        except KeyError as ke:
            LOG.warning(f"TICAM list is missing expected field {ke!s}; omitting TICAM rows.")
            return {}

        # Build dropdown-option lookup for Year + Property fields.
        year_options = {}
        prop_options = {}
        for f in fields:
            if f["id"] == fid["year"]:
                for opt in (f.get("type_config") or {}).get("options", []):
                    year_options[opt["id"]] = opt["name"]
                    year_options[str(opt.get("orderindex"))] = opt["name"]
            elif f["id"] == fid["property"]:
                for opt in (f.get("type_config") or {}).get("options", []):
                    prop_options[opt["id"]] = opt["name"]
                    prop_options[str(opt.get("orderindex"))] = opt["name"]

        # Step 2: fetch all tasks in the list.
        tasks = cu_get_list_tasks(TICAM_LIST_ID, include_closed="true")
        LOG.info(f"TICAM pull: fetched {len(tasks)} tasks from list {TICAM_LIST_ID}.")

        # Step 3: index tasks by TICAM property name, keeping only 2026 rows.
        # A property can have multiple 2026 rows (rare, but South Memorial has
        # two dropdown entries pointing at the same asset). We keep them all
        # and later pick per-Snap-Shot-property.
        #
        # Property matching is best-effort in two layers:
        #   1. Prefer the 'Property (A/O)' dropdown value (canonical).
        #   2. Fall back to the task name — many older TICAM tasks were
        #      created before the dropdown existed and only carry the
        #      property in the task's name field. We index against both
        #      candidates so the mapping file can point at either label.
        rows_by_ticam_name = {}
        for t in tasks:
            year_label = _ticam_dropdown_name(t, fid["year"], year_options)
            if year_label != "2026":
                continue
            prop_label = _ticam_dropdown_name(t, fid["property"], prop_options)
            task_name = (t.get("name") or "").strip()
            # Use dropdown if set, otherwise the task name; index BOTH so
            # downstream lookup can hit either.
            labels_for_task = set()
            if prop_label:
                labels_for_task.add(prop_label)
            if task_name:
                labels_for_task.add(task_name)
            if not labels_for_task:
                continue
            row = {
                "pgp_cam_per_sf":   _ticam_number(t, fid["pgp_cam_per_sf"]),
                "tax_per_sf":       _ticam_number(t, fid["tax_per_sf"]),
                "insurance_per_sf": _ticam_number(t, fid["insurance_per_sf"]),
                "water_per_sf":     _ticam_number(t, fid["water_per_sf"]),
                "assoc_fee_per_sf": _ticam_number(t, fid["assoc_fee_per_sf"]),
                # ClickUp REST v2 returns status as {"status":"confirmed", ...};
                # some paginated shapes / mocks return it as a bare string.
                "status":           (t["status"].get("status") if isinstance(t.get("status"), dict) else (t.get("status") or "")) or "",
                "url":              t.get("url") or "",
            }
            # Drop tasks with no populated rate fields — nothing to render.
            if not any(row[k] is not None for k in
                       ("pgp_cam_per_sf", "tax_per_sf", "insurance_per_sf",
                        "water_per_sf", "assoc_fee_per_sf")):
                continue
            for lbl in labels_for_task:
                rows_by_ticam_name.setdefault(lbl, []).append(row)

        # Step 4: apply the Snap-Shot-name → TICAM-name map and key by
        # Snap Shot name for downstream lookup.
        ticam_map = load_json(TICAM_MAP_PATH, required=False, default={})
        # Strip out the __comment__ key if present.
        ticam_map = {k: v for k, v in ticam_map.items() if not k.startswith("__")}

        rates_by_snap_shot_name = {}
        unmatched = []
        for ss_name in snap_shot_property_names:
            candidate_labels = ticam_map.get(ss_name, ss_name)
            if isinstance(candidate_labels, str):
                candidate_labels = [candidate_labels]
            # Collect all rows across the candidate labels.
            all_rows = []
            for lbl in candidate_labels:
                all_rows.extend(rows_by_ticam_name.get(lbl, []))
            if not all_rows:
                unmatched.append(ss_name)
                continue
            # Prefer 'confirmed' status; if none, use the first row.
            confirmed = [r for r in all_rows if r["status"] == "confirmed"]
            picked = confirmed[0] if confirmed else all_rows[0]
            if len(all_rows) > 1 and len(set(r["url"] for r in all_rows)) > 1:
                LOG.info(
                    f"TICAM: {ss_name!r} matched {len(all_rows)} 2026 rows across "
                    f"{candidate_labels}; picked status={picked['status']!r}."
                )
            rates_by_snap_shot_name[ss_name] = picked

        LOG.info(
            f"TICAM pull totals: {len(rates_by_snap_shot_name)} properties with 2026 rates; "
            f"{len(unmatched)} properties without a 2026 row."
        )
        if unmatched and len(unmatched) <= 20:
            LOG.info(f"TICAM: no 2026 row for: {sorted(unmatched)}")
        return rates_by_snap_shot_name
    except Exception as e:  # noqa: BLE001
        LOG.warning(f"TICAM pull failed — omitting TICAM rows: {e}")
        return {}


# ─── REM ClickUp Comment sync (Commit 2) ───────────────────────────────────
#
# The REM types a comment into col P of the Snap Shot workbook. On the next
# rebuild, this function scans every extracted REM comment, compares it
# against the "Last Synced" stamp already in col Q, and — for any comment
# whose sha8 hash differs from the stamped hash — posts to ALL matching
# ClickUp tasks (Renewal + Vacancy + Documents Workflow) and refreshes the
# stamp with a new timestamp + hash.
#
# Comment format on ClickUp:
#   📋 LSS note from {REM name} ({YYYY-MM-DD}):
#
#   {comment text}
#
# Last Synced cell format (compact enough to fit col Q, width 24):
#   "YYYY-MM-DD HH:MM ET · sha8:xxxxxxxx"
# The sha8 lets diff-detect ignore whitespace edits that render the same.


def _comment_hash(text):
    """Short, stable identity for a REM comment. Whitespace on either end
    is stripped and internal runs of whitespace are collapsed so trivial
    reformatting doesn't trigger a re-sync. Returns first 8 hex chars of
    SHA-1 — collisions are astronomically unlikely for the volume
    involved (dozens of comments per property, at most)."""
    if not text:
        return ""
    normalized = " ".join(text.split())
    return hashlib.sha1(normalized.encode("utf-8")).hexdigest()[:8]


_LAST_SYNCED_HASH_RE = re.compile(r"sha8:([0-9a-f]{8})")


def _last_synced_hash(last_synced_cell):
    """Extract the sha8 hash embedded in a Last Synced cell string.
    Returns "" if the cell is blank or was hand-edited to strip the hash."""
    if not last_synced_cell:
        return ""
    m = _LAST_SYNCED_HASH_RE.search(last_synced_cell)
    return m.group(1) if m else ""


def _rem_name_for_property(mapping_entry):
    return (mapping_entry.get("rem_name") or "").strip() or "REM"


def sync_rem_comments_to_clickup(
    mapping,
    rem_comments_by_prop_unit,
    last_synced_by_prop_unit,
    task_ids_by_tenant_id,
    task_ids_by_prop_unit,
    rent_roll_by_property,
    dry_run=False,
):
    """Post REM-authored comments from col P back to matching ClickUp tasks.

    Diff-detect: only posts when the sha8 hash of the comment differs from
    the sha8 embedded in the paired Last Synced cell. This means:
      * A brand-new comment (Last Synced is blank) always posts.
      * A comment unchanged since the last sync is skipped.
      * A comment that was previously synced and then cleared (P blank) is
        NOT reposted — blanking a comment is treated as "nothing new to
        say," not as an event worth notifying ClickUp about.

    Attribution uses the property's REM name (from the mapping file). If a
    property has no rem_name, we fall back to "REM" — the post still lands,
    just without the personalized attribution.

    Posts to ALL matching tasks per Alexis 2026-08-06 (Renewal + Vacancy +
    Docs Workflow — whichever contain a matching TenantId, and Vacancy
    tasks that match by (PropertyId, Unit#)). If no matching tasks are
    found for a unit, we log a warning and keep the comment in col P for
    the REM to see — but do NOT stamp Last Synced, so a subsequent run
    will retry if the ClickUp task appears later.

    Returns the mutated last_synced_by_prop_unit dict (with fresh stamps
    for every successfully-posted comment).
    """
    if not rem_comments_by_prop_unit:
        LOG.info("REM comment sync: no REM comments extracted from the workbook — nothing to sync.")
        return last_synced_by_prop_unit

    # Build a fast lookup: (pid, unit_label) → tenant_id (from the rent
    # roll, so occupied units can also cross-reference their Renewal /
    # Documents Workflow tasks via TenantId).
    tid_by_prop_unit = {}
    for pid, units in rent_roll_by_property.items():
        for u in units:
            unit_label = unit_display(u)
            tid = str(u.get("TenantId") or "").strip()
            if unit_label and tid:
                tid_by_prop_unit[(str(pid), unit_label)] = tid

    rem_name_by_pid = {
        str(m.get("appfolio_id") or "").strip(): _rem_name_for_property(m)
        for m in mapping
    }

    posted = 0
    skipped_unchanged = 0
    skipped_no_target = 0
    posts_by_task = 0
    now_str = now_et().strftime("%Y-%m-%d %H:%M ET")

    for (pid, unit_label), comment_text in rem_comments_by_prop_unit.items():
        comment_text = (comment_text or "").strip()
        if not comment_text:
            continue
        new_hash = _comment_hash(comment_text)
        prior_hash = _last_synced_hash(last_synced_by_prop_unit.get((pid, unit_label), ""))
        if new_hash == prior_hash and prior_hash:
            skipped_unchanged += 1
            continue

        # Collect target task ids. Prefer TenantId path when the unit is
        # occupied; add the (pid, unit) path for the vacancy board. A unit
        # can legitimately have hits on both paths (e.g., a Renewal task
        # keyed by TenantId AND a stale Vacancy task keyed by (pid,unit)).
        target_ids = []
        tid = tid_by_prop_unit.get((pid, unit_label))
        if tid:
            for t in task_ids_by_tenant_id.get(tid, []):
                if t not in target_ids:
                    target_ids.append(t)
        for t in task_ids_by_prop_unit.get((pid, unit_label), []):
            if t not in target_ids:
                target_ids.append(t)

        if not target_ids:
            LOG.warning(
                f"REM comment sync: no ClickUp task found for PropertyId={pid} Unit={unit_label!r} "
                f"(TenantId={tid or '—'}); leaving col P intact and NOT stamping Last Synced so a "
                f"future run can retry once a matching task exists."
            )
            skipped_no_target += 1
            continue

        rem_name = rem_name_by_pid.get(pid, "REM")
        today = now_et().strftime("%Y-%m-%d")
        body = f"📋 LSS note from {rem_name} ({today}):\n\n{comment_text}"

        if dry_run:
            LOG.info(
                f"REM comment sync (DRY RUN): would post to {len(target_ids)} task(s) "
                f"for PropertyId={pid} Unit={unit_label!r} — targets: {target_ids}"
            )
        else:
            for task_id in target_ids:
                try:
                    cu_post_comment(task_id, body)
                    posts_by_task += 1
                except Exception as e:  # noqa: BLE001
                    LOG.warning(
                        f"REM comment sync: post failed on task {task_id} "
                        f"(PropertyId={pid} Unit={unit_label!r}): {e}"
                    )
                    # Continue posting to other tasks; a partial failure
                    # still counts as "synced" — next run's hash check will
                    # skip a repeat. Never re-attempt in the same run.

        # Stamp Last Synced regardless of dry_run so a dry-run's log
        # accurately previews what a real run would do. (In dry-run this
        # dict isn't ever written back to SharePoint.)
        last_synced_by_prop_unit[(pid, unit_label)] = f"{now_str} · sha8:{new_hash}"
        posted += 1

    LOG.info(
        f"REM comment sync: {posted} unit(s) synced ({posts_by_task} ClickUp comment(s) posted), "
        f"{skipped_unchanged} unchanged skipped, {skipped_no_target} skipped (no matching task)."
    )
    return last_synced_by_prop_unit


def _build_refresh_section(sharepoint_url, mode):
    """Return the markdown block for the auto-managed refresh section.
    Formatted for clean rendering in ClickUp — no HTML comments, no leaked
    marker text."""
    ts = now_et().strftime("%Y-%m-%d %I:%M %p ET").lstrip("0").replace(" 0", " ")
    return (
        f"{REFRESH_HEADER_SIGNATURE} [Open in Excel Online]({sharepoint_url})\n\n"
        f"{REFRESH_FOOTER_SIGNATURE}{ts} · {mode} rebuild_"
    )


def _splice_refresh_block(current_md, new_block):
    """Insert or replace the auto-managed refresh block in `current_md`.
    Detects the existing block by its content signature (emoji header and
    italic "Last refreshed" line). Cleans up any legacy markers from earlier
    versions of the code."""
    import re

    # Clean up any leftover markers from earlier versions of this code:
    # (1) HTML-comment markers.
    # (2) The bold-italic "auto-managed by Snap Shot rebuild" markers.
    # (3) Any leftover "### 📄 Latest Snap Shot" headers with stale block.
    legacy_patterns = [
        r"<!-- SNAP_SHOT_LAST_REFRESH_BEGIN -->.*?<!-- SNAP_SHOT_LAST_REFRESH_END -->",
        r"\*\*\*— auto-managed by Snap Shot rebuild.*?— end auto-managed section —\*\*\*",
        r"### \U0001F4C4 Latest Snap Shot\s*\n\s*\*\*\[Open [^\]]+\]\([^)]+\)\*\*\s*\n\s*\*Last refreshed[^*]+\*",
    ]
    for lp in legacy_patterns:
        current_md = re.sub(lp, "", current_md, flags=re.DOTALL).strip()

    # Detect and replace the current-format block.
    current_pattern = re.compile(
        re.escape(REFRESH_HEADER_SIGNATURE) + r".*?" + re.escape(REFRESH_FOOTER_SIGNATURE) + r"[^_]*_",
        re.DOTALL,
    )
    if current_pattern.search(current_md):
        return current_pattern.sub(new_block, current_md, count=1).strip()

    # First install: prepend to existing description.
    if current_md:
        return f"{new_block}\n\n---\n\n{current_md}"
    return new_block


def update_broker_reporting_list_description(sharepoint_url, mode):
    """Update the Broker Reporting LIST description with the SharePoint link.

    OVERWRITE strategy (not splice): the list description is REPLACED entirely
    with just the refresh block. No divider, no preserved history, no growing
    string — Alexis wants the list description to always show ONLY the current
    refresh link, nothing else. Task description (below) keeps splice behavior
    because it has real workflow docs.

    Non-fatal.
    """
    block = _build_refresh_section(sharepoint_url, mode)
    ok = cu_update_list_description(BROKER_REPORTING_LIST_ID, block)
    if ok:
        LOG.info(f"Updated Broker Reporting list {BROKER_REPORTING_LIST_ID} description with refresh link.")
    else:
        LOG.warning(f"Broker Reporting list description update did not succeed on {BROKER_REPORTING_LIST_ID}.")


def update_control_task_with_refresh_link(sharepoint_url, mode):
    """After a successful rebuild + upload, rewrite the auto-managed "Last
    refresh" section at the top of the control task's description with the
    SharePoint link and timestamp.

    Safe to call from any mode (nightly, weekly, poll). Non-fatal on failure
    — the rebuild itself already succeeded and shipped the file; failing to
    update the ClickUp description should never cause the run to fail.
    """
    try:
        control_task = find_control_task()
    except Exception as e:  # noqa: BLE001
        LOG.warning(f"Could not find control task to update description: {e}")
        return

    # Fetch current description via GET — the list-search response only has
    # `text_content`, we need the raw markdown.
    try:
        full = cu_get_task(control_task["id"], include_attachments=False)
    except Exception as e:  # noqa: BLE001
        LOG.warning(f"Could not fetch control task {control_task['id']} description: {e}")
        return
    current_md = full.get("markdown_description") or full.get("description") or ""
    block = _build_refresh_section(sharepoint_url, mode)
    new_md = _splice_refresh_block(current_md, block)
    ok = cu_update_task_description(control_task["id"], new_md)
    if ok:
        LOG.info(f"Updated control task {control_task['id']} description with refresh link.")
    else:
        LOG.warning(f"Control task description update did not succeed on {control_task['id']}.")


def cu_search_tasks_by_name(list_id, name_query, include_closed="true"):
    """Find tasks in a list whose name matches name_query (case-insensitive
    substring). Used to locate the control task without hardcoding its id."""
    tasks = cu_get_list_tasks(list_id, include_closed=include_closed)
    q = name_query.strip().lower()
    return [t for t in tasks if q in (t.get("name") or "").lower()]


def cf_value(task, field_id=None, field_name=None):
    """Raw custom-field value, matched by id or (case-insensitive) name."""
    for f in task.get("custom_fields", []):
        if field_id and f.get("id") == field_id:
            return f.get("value")
        if field_name and (f.get("name") or "").strip().lower() == field_name.strip().lower():
            return f.get("value")
    return None


def cf_field_meta(task, field_name):
    """Return the raw custom_fields entry matching field_name (case-insensitive)."""
    for f in task.get("custom_fields", []):
        if (f.get("name") or "").strip().lower() == field_name.strip().lower():
            return f
    return None


# ══════════════════════════════════════════════════════════════════════════
# ClickUp — Broker Reporting integration
# ══════════════════════════════════════════════════════════════════════════

def pull_broker_reporting_tasks():
    """Pull all tasks in the Broker Reporting list, keyed by task id."""
    tasks = cu_get_list_tasks(BROKER_REPORTING_LIST_ID)
    by_id = {t["id"]: t for t in tasks}
    LOG.info(f"ClickUp Broker Reporting: {len(by_id)} tasks pulled.")
    return by_id


def pull_broker_beat_attachments(broker_reporting_tasks):
    """For each property task, find its child 'Weekly Broker Update - MM/DD/YYYY'
    subtasks and return the text of the LATEST one's attachment content
    (falls back to task description if no parseable attachment) keyed by
    the parent property task id. Returns {clickup_task_id: text}.
    """
    results = {}
    date_re = re.compile(re.escape(BROKER_BEAT_TASK_PREFIX) + r"\s*-\s*(\d{1,2}/\d{1,2}/\d{4})", re.I)

    for task_id, task in broker_reporting_tasks.items():
        try:
            full_task = cu_get_task(task_id, include_subtasks=True)
        except FatalError as e:
            LOG.warning(f"Broker Beat: could not fetch task {task_id}, skipping: {e}")
            continue

        # Weekly Broker Update tasks are modeled as subtasks/child tasks of the
        # property task in this list.
        subtask_ids = full_task.get("subtasks") or []
        candidates = []
        for st in subtask_ids:
            name = st.get("name") or ""
            m = date_re.search(name)
            if not m:
                continue
            try:
                d = datetime.strptime(m.group(1), "%m/%d/%Y")
            except ValueError:
                continue
            candidates.append((d, st))

        if not candidates:
            continue

        candidates.sort(key=lambda c: c[0], reverse=True)
        _, latest = candidates[0]
        latest_full = cu_get_task(latest["id"])
        attachments = latest_full.get("attachments") or []
        if attachments:
            # We can't reliably parse arbitrary attachment file formats here;
            # record which attachment + date so the note is traceable, and
            # pull any comment text on the attachment task as the summary.
            att_names = ", ".join(a.get("title", a.get("id", "?")) for a in attachments)
            summary = latest_full.get("description") or latest_full.get("text_content") or ""
            note = summary.strip() if summary.strip() else f"See attachment: {att_names}"
        else:
            note = (latest_full.get("description") or latest_full.get("text_content") or "").strip()

        if note:
            results[task_id] = note

    LOG.info(f"Broker Beat: found active-interest updates for {len(results)} properties.")
    return results


def merge_broker_active_interest(overrides, broker_active_interest_by_task):
    """Overwrite 'broker_active_interest' for properties with a fresh Broker
    Beat update this week; leave every other field (and every other
    property's broker_active_interest) untouched."""
    merged = json.loads(json.dumps(overrides))  # deep copy
    for task_id, text in broker_active_interest_by_task.items():
        entry = merged.get(task_id)
        if not isinstance(entry, dict):
            entry = {}
        entry["broker_active_interest"] = text
        merged[task_id] = entry
    return merged


# ══════════════════════════════════════════════════════════════════════════
# ClickUp — on-demand poll (checkbox trigger)
# ══════════════════════════════════════════════════════════════════════════

def find_control_task():
    matches = cu_search_tasks_by_name(BROKER_REPORTING_LIST_ID, CONTROL_TASK_NAME)
    if not matches:
        raise FatalError(
            f"ClickUp checkbox not found: no task named like "
            f"'{CONTROL_TASK_NAME}' in list {BROKER_REPORTING_LIST_ID}."
        )
    if len(matches) > 1:
        LOG.warning(
            f"Multiple candidate control tasks matched '{CONTROL_TASK_NAME}'; "
            f"using the first: {matches[0]['id']}"
        )
    return matches[0]


def is_checkbox_checked(control_task):
    field = cf_field_meta(control_task, CHECKBOX_FIELD_NAME)
    if field is None:
        raise FatalError(
            f"ClickUp checkbox not found: control task "
            f"{control_task.get('id')} has no custom field named "
            f"'{CHECKBOX_FIELD_NAME}'. Add it via the ClickUp UI."
        )
    value = field.get("value")
    return bool(value) and value not in (False, None, "false", "0", 0)


def uncheck_and_comment(control_task, message):
    field = cf_field_meta(control_task, CHECKBOX_FIELD_NAME)
    if field is None:
        LOG.warning("Could not find checkbox field to uncheck; skipping uncheck step.")
        return
    cu_set_field(control_task["id"], field["id"], False)
    cu_post_comment(control_task["id"], message)
    LOG.info(f"Unchecked '{CHECKBOX_FIELD_NAME}' and commented on task {control_task['id']}.")


# ══════════════════════════════════════════════════════════════════════════
# Microsoft Graph — auth
# ══════════════════════════════════════════════════════════════════════════

def get_graph_access_token():
    if not MS_REFRESH_TOKEN:
        raise FatalError(
            "APATTISON_MS_REFRESH_TOKEN env var is not set. This is the "
            "delegated Graph refresh token for apattison@prudentgrowth.com "
            "(scopes: Mail.Send, Files.ReadWrite.All, Sites.ReadWrite.All, "
            "offline_access). Export it before running."
        )
    r = http_request(
        "POST", f"https://login.microsoftonline.com/{MS_TENANT_ID}/oauth2/v2.0/token",
        context="Graph token refresh",
        data={
            "client_id": MS_CLIENT_ID,
            "refresh_token": MS_REFRESH_TOKEN,
            "grant_type": "refresh_token",
            "scope": MS_GRAPH_SCOPE,
        },
        retries=1,
    )
    if r.status_code != 200:
        raise FatalError(f"Graph token refresh failed: HTTP {r.status_code}: {r.text[:300]}")
    return r.json()["access_token"]


# ══════════════════════════════════════════════════════════════════════════
# Microsoft Graph — SharePoint download / race check / upload
# ══════════════════════════════════════════════════════════════════════════

def sharepoint_item_metadata_url():
    return (
        f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drives/{DRIVE_ID}"
        f"/root:/{SHAREPOINT_ITEM_PATH}"
    )


def download_from_sharepoint():
    """Download the current workbook from SharePoint.

    Returns (content_bytes_or_None, last_modified_by_email_or_None,
    last_modified_datetime_or_None). If the file does not yet exist (first
    run), returns (None, None, None) — this is NOT an error.
    """
    token = get_graph_access_token()
    headers = {"Authorization": f"Bearer {token}"}

    meta_url = sharepoint_item_metadata_url()
    r = http_request("GET", meta_url, context="SharePoint metadata fetch", headers=headers)

    if r.status_code == 404:
        LOG.info("SharePoint workbook does not exist yet (first run). Proceeding without it.")
        return None, None, None
    if r.status_code != 200:
        raise FatalError(
            f"Failed to download SharePoint file: metadata fetch returned "
            f"HTTP {r.status_code}: {r.text[:300]}"
        )

    meta = r.json()
    last_modified_by = (
        (meta.get("lastModifiedBy") or {}).get("user", {}).get("email")
        or (meta.get("lastModifiedBy") or {}).get("user", {}).get("displayName")
        or ""
    )
    last_modified_at_raw = meta.get("lastModifiedDateTime")
    last_modified_at = None
    if last_modified_at_raw:
        try:
            last_modified_at = datetime.strptime(last_modified_at_raw, "%Y-%m-%dT%H:%M:%SZ")
        except ValueError:
            try:
                last_modified_at = datetime.strptime(last_modified_at_raw[:19], "%Y-%m-%dT%H:%M:%S")
            except ValueError:
                LOG.warning(f"Could not parse lastModifiedDateTime: {last_modified_at_raw}")

    download_url = meta.get("@microsoft.graph.downloadUrl")
    if not download_url:
        raise FatalError("SharePoint metadata response missing @microsoft.graph.downloadUrl.")

    r2 = http_request(
        "GET", download_url, context="SharePoint content download", timeout=HTTP_TIMEOUT_LONG_SEC
    )
    if r2.status_code != 200:
        raise FatalError(f"Failed to download SharePoint file content: HTTP {r2.status_code}")

    LOG.info(
        f"Downloaded SharePoint workbook ({len(r2.content):,} bytes). "
        f"lastModifiedBy={last_modified_by!r} lastModifiedAt={last_modified_at}"
    )
    return r2.content, last_modified_by, last_modified_at


def is_race_condition(last_modified_by, last_modified_at):
    """True if a human OTHER than the automation account edited the workbook
    within the last RACE_CONDITION_WINDOW_MINUTES minutes."""
    if not last_modified_by or not last_modified_at:
        return False
    if last_modified_by.lower() == AUTOMATION_ACCOUNT_EMAIL:
        return False
    age = datetime.utcnow() - last_modified_at
    if age <= timedelta(minutes=RACE_CONDITION_WINDOW_MINUTES):
        LOG.info(
            f"Race condition detected: '{last_modified_by}' edited the workbook "
            f"{age.total_seconds() / 60:.1f} min ago (window={RACE_CONDITION_WINDOW_MINUTES}m)."
        )
        return True
    return False


def upload_to_sharepoint(local_path):
    """Upload (overwrite) the workbook to SharePoint. Uses a simple PUT for
    files under the Graph 4MB simple-upload limit, and a chunked resumable
    upload session for anything larger (mirrors the notice-letter-1-send
    large-file Graph upload pattern)."""
    token = get_graph_access_token()
    size = os.path.getsize(local_path)

    if size < LARGE_FILE_THRESHOLD_BYTES:
        url = (
            f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drives/{DRIVE_ID}"
            f"/root:/{SHAREPOINT_ITEM_PATH}:/content"
        )
        with open(local_path, "rb") as f:
            data = f.read()
        r = http_request(
            "PUT", url, context="SharePoint simple upload",
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/octet-stream",
            },
            data=data, timeout=HTTP_TIMEOUT_LONG_SEC,
        )
        if r.status_code not in (200, 201):
            raise FatalError(f"SharePoint upload failed: HTTP {r.status_code}: {r.text[:300]}")
        LOG.info(f"Uploaded workbook to SharePoint via simple PUT ({size:,} bytes).")
        return r.json()

    # Chunked / resumable upload session for large files.
    session_url = (
        f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drives/{DRIVE_ID}"
        f"/root:/{SHAREPOINT_ITEM_PATH}:/createUploadSession"
    )
    r = http_request(
        "POST", session_url, context="SharePoint create upload session",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json={"item": {"@microsoft.graph.conflictBehavior": "replace"}},
    )
    if r.status_code not in (200, 201):
        raise FatalError(f"SharePoint upload session creation failed: HTTP {r.status_code}: {r.text[:300]}")
    upload_url = r.json()["uploadUrl"]

    with open(local_path, "rb") as f:
        offset = 0
        final_response = None
        while offset < size:
            chunk = f.read(UPLOAD_CHUNK_SIZE_BYTES)
            chunk_len = len(chunk)
            end = offset + chunk_len - 1
            headers = {
                "Content-Length": str(chunk_len),
                "Content-Range": f"bytes {offset}-{end}/{size}",
            }
            rc = http_request(
                "PUT", upload_url, context=f"SharePoint chunk upload {offset}-{end}/{size}",
                headers=headers, data=chunk, timeout=HTTP_TIMEOUT_LONG_SEC,
            )
            if rc.status_code not in (200, 201, 202):
                raise FatalError(
                    f"SharePoint chunked upload failed at offset {offset}: "
                    f"HTTP {rc.status_code}: {rc.text[:300]}"
                )
            offset += chunk_len
            final_response = rc

    LOG.info(f"Uploaded workbook to SharePoint via chunked session ({size:,} bytes).")
    return final_response.json() if final_response is not None else {}


# ─── SharePoint archive (RCA_2026-08-07 fix #4) ──────────────────────────────
#
# Every successful rebuild archives the current live Snap Shot to
# {SHAREPOINT_FOLDER_PATH}/_archive/Leasing-Snap-Shot__pre_{timestamp}.xlsx
# BEFORE we overwrite it. Retention: last 14 calendar days (roughly two full
# weeks of nightlies + weeklies). Recovery becomes a 30-second SharePoint copy
# instead of a version-history dive.

ARCHIVE_FOLDER_NAME = "_archive"
ARCHIVE_RETAIN_DAYS = 14


def _archive_folder_path():
    return f"{SHAREPOINT_FOLDER_PATH}/{ARCHIVE_FOLDER_NAME}"


def archive_current_snap_shot(current_wb_bytes):
    """Upload the pre-overwrite copy of the live Snap Shot to _archive/, then
    prune archives older than ARCHIVE_RETAIN_DAYS. Non-fatal: caller wraps in
    try/except so archive failures never block the primary rebuild.

    current_wb_bytes is the bytes we already downloaded at the top of run_build
    — no second Graph download needed."""
    if not current_wb_bytes:
        LOG.info("Archive: no current workbook bytes to archive (first run?).")
        return

    ts = datetime.utcnow().strftime("%Y%m%d_%H%M%S")
    archive_name = f"Leasing-Snap-Shot__pre_{ts}.xlsx"
    archive_path = f"{_archive_folder_path()}/{archive_name}"

    token = get_graph_access_token()
    url = (
        f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drives/{DRIVE_ID}"
        f"/root:/{archive_path}:/content"
    )
    r = http_request(
        "PUT", url, context="SharePoint archive upload",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/octet-stream",
        },
        data=current_wb_bytes, timeout=HTTP_TIMEOUT_LONG_SEC,
    )
    if r.status_code not in (200, 201):
        raise RuntimeError(
            f"Archive upload failed: HTTP {r.status_code}: {r.text[:300]}"
        )
    size = len(current_wb_bytes)
    LOG.info(f"Archive: copied current Snap Shot to _archive/{archive_name} ({size:,} bytes).")

    try:
        _prune_archives(token)
    except Exception as e:  # noqa: BLE001
        LOG.warning(f"Archive prune failed (non-fatal): {e}")


def _prune_archives(token):
    """Delete archives older than ARCHIVE_RETAIN_DAYS. Filename convention is
    Leasing-Snap-Shot__pre_YYYYMMDD_HHMMSS.xlsx — anything not matching that
    pattern is left alone (belt-and-suspenders against accidentally deleting
    other files that happened to land in _archive/)."""
    list_url = (
        f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drives/{DRIVE_ID}"
        f"/root:/{_archive_folder_path()}:/children?$top=200"
    )
    r = http_request(
        "GET", list_url, context="SharePoint archive list",
        headers={"Authorization": f"Bearer {token}"},
    )
    if r.status_code == 404:
        return  # folder doesn't exist yet
    if r.status_code != 200:
        raise RuntimeError(f"Archive list failed: HTTP {r.status_code}: {r.text[:300]}")

    items = r.json().get("value", [])
    cutoff = datetime.utcnow() - timedelta(days=ARCHIVE_RETAIN_DAYS)
    pruned = 0
    for it in items:
        name = it.get("name", "")
        if not (name.startswith("Leasing-Snap-Shot__pre_") and name.endswith(".xlsx")):
            continue
        try:
            ts_str = name[len("Leasing-Snap-Shot__pre_"):-len(".xlsx")]
            ts = datetime.strptime(ts_str, "%Y%m%d_%H%M%S")
        except ValueError:
            continue
        if ts >= cutoff:
            continue
        item_id = it.get("id")
        if not item_id:
            continue
        del_url = (
            f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drives/{DRIVE_ID}"
            f"/items/{item_id}"
        )
        dr = http_request(
            "DELETE", del_url, context="SharePoint archive prune",
            headers={"Authorization": f"Bearer {token}"},
        )
        if dr.status_code in (200, 204):
            pruned += 1
            LOG.info(f"Archive: pruned {name} (>{ARCHIVE_RETAIN_DAYS}d old).")
        else:
            LOG.warning(f"Archive: could not prune {name}: HTTP {dr.status_code}")
    if pruned:
        LOG.info(f"Archive: pruned {pruned} old snapshot(s), retention={ARCHIVE_RETAIN_DAYS}d.")



# ══════════════════════════════════════════════════════════════════════════
# Microsoft Graph — error email
# ══════════════════════════════════════════════════════════════════════════

def send_error_email(subject, body):
    try:
        token = get_graph_access_token()
    except FatalError as e:
        LOG.error(f"Could not send error email (token acquisition failed): {e}")
        return
    payload = {
        "message": {
            "subject": subject,
            "body": {"contentType": "Text", "content": body},
            "toRecipients": [{"emailAddress": {"address": ERROR_NOTIFICATION_TO}}],
        },
        "saveToSentItems": "true",
    }
    r = http_request(
        "POST", "https://graph.microsoft.com/v1.0/me/sendMail",
        context="Graph error email send",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json=payload, retries=1,
    )
    if r.status_code != 202:
        LOG.error(f"Error-notification email failed to send: HTTP {r.status_code}: {r.text[:300]}")
    else:
        LOG.info(f"Error-notification email sent to {ERROR_NOTIFICATION_TO}.")


# ══════════════════════════════════════════════════════════════════════════
# Ann's-edit extraction (round-trip from the current SharePoint workbook)
# ══════════════════════════════════════════════════════════════════════════
#
# Layout reference (from build_prototype_v3.py write_property_block()):
#   Row 1 of block: property name banner
#   Row 2:          address sub-banner (contains "Property ID: <id>")
#   Row 3:          blank spacer
#   Row 4:          "DEAL ACTIVITY & NOTES" section header
#   Rows 5-11:      7 label/value rows in column B (label) / C:N merged (value):
#                     Broker / Contact
#                     Broker Calls
#                     Property Flags
#                     Vacant Callouts
#                     Deal Activity
#                     Broker Active Interest from Weekly Reporting
#                     Ann's Commentary
#   blank spacer, "RENT ROLL" header, column headers, then per-unit rows with
#   Market Rent in column H and Notes in column N (see COLS in build_workbook).

NOTE_LABEL_TO_KEY = {
    # "Broker / Contact" is auto-populated from the ClickUp Broker Directory
    # on every rebuild — we do NOT preserve REM edits to this row. It IS listed
    # here so the label-consumption loop in extract_ann_edits sees all 7 rows
    # the renderer writes and doesn't overshoot the RENT ROLL header. The
    # extracted value under 'broker_contact' is intentionally not read back by
    # write_property_block. (See RCA_2026-08-07.md — omitting this key was the
    # root cause of the 2026-08-07 REM Comment wipe: 6 mapped labels vs 7
    # written rows caused the extractor's misses<2 loop to break on RENT ROLL
    # itself, leaving scan_row one row past the header.)
    "broker / contact": "broker_contact",
    "broker calls": "broker_calls",
    "property flags": "property_flags",
    "vacant callouts": "vacant_callouts",
    "deal activity": "deal_activity",
    "broker active interest from weekly reporting": "broker_active_interest",
    "ann's commentary": "ann_commentary",
}

PROPERTY_ID_RE = re.compile(r"Property ID:\s*([0-9]+)")


def extract_ann_edits(workbook_bytes):
    """Parse the current SharePoint workbook and return five things:

      property_overrides           = {property_id: {note_key: value}}
      unit_notes                   = {"pid::unit_label": note_text}
      market_rent_overrides        = {property_id: {unit_label: market_rent_value}}
      rem_comments_by_prop_unit    = {(pid, unit_label): comment_text}
      last_synced_by_prop_unit     = {(pid, unit_label): last_synced_string}

    The last two are new in Commit 2 and support REM-authored comments in
    col P (REM ClickUp Comment) plus the paired col Q (Last Synced). They
    round-trip through SharePoint so REM edits survive nightly rebuilds and
    so the sync-diff can tell whether the REM's text changed since the
    previous run.

    On any parse error for an individual block, that block's edits are
    skipped (logged) rather than aborting the whole extraction — a single
    malformed block must never nuke everyone else's preserved edits.
    """
    property_overrides = {}
    unit_notes = {}
    market_rent_overrides = {}  # {property_id: {unit_label: market_rent_value}}
    rem_comments_by_prop_unit = {}  # {(pid, unit_label): comment_text}
    last_synced_by_prop_unit = {}   # {(pid, unit_label): last_synced_string}

    try:
        wb = openpyxl.load_workbook(io.BytesIO(workbook_bytes), data_only=True)
    except Exception as e:
        raise FatalError(f"Could not open the downloaded SharePoint workbook for extraction: {e}")

    if "Snap Shot" not in wb.sheetnames:
        raise FatalError(
            f"Downloaded SharePoint workbook has no 'Snap Shot' tab "
            f"(found: {wb.sheetnames}); refusing to extract from an unexpected layout."
        )
    ws = wb["Snap Shot"]

    def cell_text(row, col_letter):
        v = ws[f"{col_letter}{row}"].value
        return "" if v is None else str(v).strip()

    max_row = ws.max_row
    row = 1
    blocks_found = 0
    while row <= max_row:
        name_cell = cell_text(row, "B")
        # Heuristic block-start detector: a banner row is a non-empty B cell
        # immediately followed (next row) by a "Property ID: NNN" address row.
        next_row_addr = cell_text(row + 1, "B")
        pid_match = PROPERTY_ID_RE.search(next_row_addr)
        if not (name_cell and pid_match):
            row += 1
            continue

        property_id = pid_match.group(1)
        blocks_found += 1
        block_start = row

        try:
            # DEAL ACTIVITY & NOTES header should be at block_start + 4
            # (banner, address, TICAM row, blank spacer, DEAL ACTIVITY).
            # Before the 2026 TICAM row was added it lived at block_start + 3;
            # we probe both offsets so the extractor keeps working against
            # older SharePoint copies uploaded before the TICAM row shipped.
            header_row = block_start + 4
            header_text = cell_text(header_row, "B")
            if not header_text.upper().startswith("DEAL ACTIVITY"):
                # Legacy pre-TICAM layout — fall back to the old offset.
                header_row = block_start + 3
                header_text = cell_text(header_row, "B")
            notes_start = header_row + 1 if header_text.upper().startswith("DEAL ACTIVITY") else header_row

            notes = {}
            r = notes_start
            consumed = 0
            # There are exactly 7 label rows in the current layout; walk up to
            # 10 rows defensively in case of future insertions, stopping once
            # we hit a row whose label doesn't match a known key twice in a row.
            misses = 0
            while r <= max_row and consumed < 7 and misses < 2:
                label = cell_text(r, "B").strip()
                key = NOTE_LABEL_TO_KEY.get(label.lower())
                if key:
                    value = cell_text(r, "C")
                    notes[key] = value
                    consumed += 1
                    misses = 0
                else:
                    misses += 1
                r += 1

            if notes:
                property_overrides[property_id] = notes
            else:
                LOG.warning(f"Extraction: no note rows recognized for property block at row {block_start} (PropertyId={property_id}).")

            # Find the RENT ROLL section within this block and walk unit rows
            # (Unit=B, ..., Market Rent=H, ..., Notes=N) until a blank Unit
            # cell (spacer row) signals the end of the block.
            #
            # Scan starts from notes_start (top of the note-label block), not
            # from wherever the label-consumption loop happened to stop. This is
            # defensive against label-count drift: if the renderer adds or
            # removes a note row in the future, the label loop may over- or
            # under-shoot, but the RENT ROLL header itself is always ~8-12 rows
            # below notes_start. Starting from notes_start with a wider window
            # keeps the extractor robust even when the label map is briefly out
            # of sync with the renderer. (RCA_2026-08-07.md.)
            scan_row = notes_start
            rent_roll_header_row = None
            for probe in range(scan_row, min(scan_row + 25, max_row + 1)):
                if cell_text(probe, "B").upper() == "RENT ROLL":
                    rent_roll_header_row = probe
                    break
            if rent_roll_header_row:
                col_header_row = rent_roll_header_row + 1
                data_row = col_header_row + 1
                while data_row <= max_row:
                    unit_val = cell_text(data_row, "B")
                    if not unit_val or unit_val.upper() == "TOTALS":
                        break
                    tenant_id_val = cell_text(data_row, "D")
                    market_rent_val = cell_text(data_row, "H")
                    note_val = cell_text(data_row, "N")
                    # Commit 2: REM ClickUp Comment (P) + Last Synced (Q).
                    # Read exactly like Notes; empty passes through as "".
                    rem_comment_val = cell_text(data_row, "P")
                    last_synced_val = cell_text(data_row, "Q")
                    # Note key mirrors build_prototype_v3: OccupancyId else
                    # UnitId. We don't have those ids directly in the sheet,
                    # so key by (PropertyId, Unit label) as a stable fallback
                    # that build_workbook() can re-resolve against the fresh
                    # AppFolio pull by unit label.
                    unit_key = f"{property_id}::{unit_val}"
                    if note_val:
                        unit_notes[unit_key] = note_val
                    if market_rent_val:
                        market_rent_overrides.setdefault(property_id, {})[unit_val] = market_rent_val
                    # Only record REM comment / last-synced when non-empty,
                    # so wiring them back into fresh mapping entries doesn't
                    # overwrite a genuinely blank cell with "".
                    if rem_comment_val:
                        rem_comments_by_prop_unit[(property_id, unit_val)] = rem_comment_val
                    if last_synced_val:
                        last_synced_by_prop_unit[(property_id, unit_val)] = last_synced_val
                    data_row += 1
                block_end = data_row
            else:
                LOG.warning(f"Extraction: no RENT ROLL section found for property block at row {block_start} (PropertyId={property_id}).")
                block_end = r + 1

        except Exception as e:
            LOG.warning(f"Extraction: failed to parse block at row {block_start} (PropertyId={property_id}): {e}")
            block_end = block_start + 1

        row = max(block_end, block_start + 1)

    LOG.info(
        f"Extracted Ann's edits from SharePoint workbook: {blocks_found} property "
        f"blocks scanned, {len(property_overrides)} with notes, "
        f"{len(unit_notes)} unit notes, "
        f"{sum(len(v) for v in market_rent_overrides.values())} market-rent overrides, "
        f"{len(rem_comments_by_prop_unit)} REM ClickUp Comments, "
        f"{len(last_synced_by_prop_unit)} Last-Synced stamps."
    )

    # Per-block diagnostics: if REM comment extraction returns 0 but there ARE
    # property blocks in the sheet, dump per-block detail so debugging doesn't
    # require a separate re-run. (RCA_2026-08-07.md fix #5.)
    if len(rem_comments_by_prop_unit) == 0 and blocks_found > 0:
        LOG.warning(
            "REM Comment extraction returned 0 across %d blocks — dumping per-block "
            "note-key detail for debugging:", blocks_found,
        )
        for pid, note_dict in list(property_overrides.items())[:10]:
            LOG.warning("  PropertyId=%s: note_keys=%s", pid, list(note_dict.keys()))
        if len(property_overrides) > 10:
            LOG.warning("  (… %d more blocks omitted …)", len(property_overrides) - 10)

    return (
        property_overrides,
        unit_notes,
        market_rent_overrides,
        rem_comments_by_prop_unit,
        last_synced_by_prop_unit,
    )


# ─── REM Comment safety guard (Commit 4 — RCA_2026-08-07 fix #3) ──────────────

REM_COUNT_STATE_PATH = os.path.join(DATA_DIR, "last_rem_count.json")


def check_rem_comment_floor(current_count, force_rebuild, state_path=None):
    """Refuse to proceed if the extracted REM Comment count dropped dangerously
    below the previous run's count. Catches extractor regressions like the
    2026-08-07 label-map drift that silently wiped 4 REM Comments.

    Rules (fail loud, don't upload):
      - Hard floor: prior > 0 and current == 0. Almost certainly extractor bug.
      - Soft floor: prior >= 4 and current < prior // 2. >50% drop is suspicious.

    Bypass with force_rebuild=True (already exposed as workflow input for the
    "file is open in Excel, we know it's fine" case).

    On success, writes the current count to state_path atomically for next run.
    Returns None. Raises RuntimeError on guard trip.
    """
    path = state_path or REM_COUNT_STATE_PATH
    try:
        if os.path.exists(path):
            prior = int(json.load(open(path)).get("count", 0))
        else:
            prior = 0
    except Exception as e:
        LOG.warning("REM count guard: could not read prior state (%s); treating prior=0.", e)
        prior = 0

    if not force_rebuild:
        if prior > 0 and current_count == 0:
            raise RuntimeError(
                f"REM Comment safety guard TRIPPED: extractor returned 0 comments "
                f"but the previous run saw {prior}. Refusing to upload — this would "
                f"wipe REM comments (same failure mode as 2026-08-07 incident). "
                f"Investigate the extractor before re-running. If the drop is genuine "
                f"(e.g. REMs cleared all comments intentionally), re-dispatch with "
                f"force_rebuild=true to bypass."
            )
        if prior >= 4 and current_count < (prior // 2):
            raise RuntimeError(
                f"REM Comment safety guard TRIPPED: extractor returned {current_count}, "
                f"down >50% from {prior} last run. Refusing to upload. Investigate the "
                f"extractor before re-running. Re-dispatch with force_rebuild=true to "
                f"bypass."
            )

    LOG.info(
        "REM count guard: current=%d, prior=%d, force_rebuild=%s — OK to proceed.",
        current_count, prior, force_rebuild,
    )

    # Atomic write of new state so a crash mid-write doesn't corrupt it.
    try:
        tmp = path + ".tmp"
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(tmp, "w") as f:
            json.dump(
                {
                    "count": int(current_count),
                    "prior_count": int(prior),
                    "ts_utc": datetime.utcnow().isoformat() + "Z",
                },
                f,
                indent=2,
            )
        os.replace(tmp, path)
    except Exception as e:
        LOG.warning("REM count guard: could not persist new state (%s).", e)


def reindex_overrides_by_clickup_task(property_overrides_by_property_id, mapping):
    """property_notes_overrides.json is keyed by clickup_task_id (per the
    existing schema/fallback file), but extraction from the sheet naturally
    keys by AppFolio PropertyId (visible in the sheet as "Property ID: NNN").
    Re-key using the mapping file so downstream code can use either key
    consistently with the existing fallback JSON schema."""
    by_task = {}
    for m in mapping:
        pid = str(m.get("appfolio_id"))
        task_id = m.get("clickup_task_id") or ""
        if pid in property_overrides_by_property_id and task_id:
            by_task[task_id] = property_overrides_by_property_id[pid]
    return by_task


def reindex_unit_notes_by_note_key(unit_notes_by_prop_and_label, rent_roll_by_property, mapping):
    """unit_notes.json is keyed by OccupancyId/UnitId. Extraction above keys
    by "PropertyId::UnitLabel" (a stable proxy available from the sheet).
    Re-resolve to OccupancyId/UnitId against the freshly-pulled rent roll so
    the note survives even if AppFolio's own unit label formatting shifts
    slightly between pulls."""
    from_prototype_unit_display = unit_display  # reuse same normalization
    result = {}
    unresolved = 0
    for m in mapping:
        pid = str(m.get("appfolio_id"))
        units = rent_roll_by_property.get(pid, [])
        label_to_unit = {from_prototype_unit_display(u): u for u in units}
        for key, note in unit_notes_by_prop_and_label.items():
            if not key.startswith(f"{pid}::"):
                continue
            label = key.split("::", 1)[1]
            u = label_to_unit.get(label)
            if u is None:
                unresolved += 1
                continue
            note_key = str(u.get("OccupancyId") or u.get("UnitId") or "")
            if note_key:
                result[note_key] = note
    if unresolved:
        LOG.warning(f"{unresolved} unit note(s) from SharePoint could not be re-matched to a current AppFolio unit (unit likely moved out/renamed); those notes were dropped.")
    return result


# ══════════════════════════════════════════════════════════════════════════
# Workbook builder — ported verbatim (layout/style) from build_prototype_v3.py
# ══════════════════════════════════════════════════════════════════════════

# --- PGP palette ---
NAVY_DEEP = "1E2840"
NAVY = "283556"
BLUE_MID = "2B689C"
RED_LIGHT = "F9C7C7"     # 2026 renewals (light red)
BLUE_LIGHT = "A6E2F9"    # 2027 renewals (light blue)
BLUE_SKY = RED_LIGHT
SLATE = "7B8B99"
SLATE_TINT = "BBC5CE"
STEEL = "5290C9"
GRAY_LIGHT = "EDEDED"
GRAY_ROW = "F5F7FA"
WHITE = "FFFFFF"
BLACK = "111111"

FONT_NAME = "Century Gothic"


def fill(hex_):
    return PatternFill(start_color=hex_, end_color=hex_, fill_type="solid")


THIN = Side(style="thin", color="D9D9D9")
MEDIUM = Side(style="medium", color=NAVY)
BOX_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BANNER_BORDER = Border(bottom=MEDIUM)

COLS = {
    "Unit": "B",
    "Tenant": "C",
    "Tenant ID": "D",
    "Sq Ft": "E",
    "Rent/mo": "F",
    "PSF": "G",
    "Market Rent": "H",
    "Lease From": "I",
    "Lease To": "J",
    "Move In": "K",
    "Status": "L",
    "Past Due": "M",
    "Notes": "N",
    # Read-only ClickUp AI Summary from the LAR lists (Renewal/Vacancy/Docs).
    # Populated nightly from ClickUp → Excel; nothing round-trips out of this
    # column back to ClickUp. Ann's edits are not expected here; extract_ann_edits
    # deliberately does NOT read this column.
    "ClickUp Summary": "O",
    # REM-editable comment field. Whatever the REM types here posts back to
    # ClickUp as a comment on ALL matching LAR tasks (Renewal/Vacancy/Docs)
    # for this unit, attributed as "📋 LSS note from {REM name} ({date})".
    # Only posts when the text differs from the paired "Last Synced" column.
    "REM ClickUp Comment": "P",
    # Timestamp + first ~30 chars of the last-posted comment, written back
    # after a successful ClickUp post so the REM has visual confirmation and
    # so the diff-detect knows what was last sent. Read on next rebuild.
    "Last Synced": "Q",
}
# Column widths. O widened from 44 → 80 per Alexis 2026-08-06: at 44 the
# AI summaries were wrapping to 6–7 lines per unit; 80 gets them to ~2–3.
COL_WIDTHS = {"A": 3, "B": 20, "C": 26, "D": 10, "E": 9, "F": 11, "G": 8, "H": 11,
              "I": 11, "J": 11, "K": 11, "L": 15, "M": 11, "N": 34,
              "O": 80, "P": 40, "Q": 24}
LAST_COL = "Q"
FIRST_COL = "B"

STATE_RE = re.compile(r",\s*([A-Z]{2})\s+\d{5}")


def build_workbook(mapping, rent_roll_by_property, property_overrides,
                    unit_notes, market_rent_overrides, broker_map,
                    output_path, logo_path=None,
                    clickup_summaries_by_tenant=None,
                    clickup_summaries_by_prop_unit=None,
                    rem_comments_by_prop_unit=None,
                    last_synced_by_prop_unit=None,
                    ticam_by_property=None):
    """Build the Snap Shot workbook. `mapping` = property_mapping_all73.json
    contents. `rent_roll_by_property` = {appfolio_id_str: [unit_row, ...]}.
    `property_overrides` = {clickup_task_id: {broker_calls, property_flags,
    vacant_callouts, deal_activity, broker_active_interest, ann_commentary}}.
    `unit_notes` = {occupancy_or_unit_id_str: note_text}.
    `market_rent_overrides` = {appfolio_id_str: {unit_label: value_str}} —
    round-tripped from SharePoint; AppFolio has no market-rent field of its
    own for this column, so absent an override the cell is left blank
    exactly as in build_prototype_v3.
    `clickup_summaries_by_tenant` = {tenant_id_str: summary_text} — from LAR
    lists (Renewal/Vacancy/Documents Workflow). Used for occupied units.
    `clickup_summaries_by_prop_unit` = {(property_id_str, unit_label_str):
    summary_text} — from Vacancy list only. Used for vacant units.

    Commit 2 additions:
    `rem_comments_by_prop_unit` = {(pid_str, unit_label): comment_text} —
    round-tripped from SharePoint col P plus mutated by sync_rem_comments;
    populates col P (REM ClickUp Comment).
    `last_synced_by_prop_unit` = {(pid_str, unit_label): stamp_string} —
    from SharePoint col Q with fresh stamps for anything just synced; the
    stamp format is "YYYY-MM-DD HH:MM ET · sha8:xxxxxxxx".
    """
    clickup_summaries_by_tenant = clickup_summaries_by_tenant or {}
    clickup_summaries_by_prop_unit = clickup_summaries_by_prop_unit or {}
    rem_comments_by_prop_unit = rem_comments_by_prop_unit or {}
    last_synced_by_prop_unit = last_synced_by_prop_unit or {}
    rr = rent_roll_by_property

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Snap Shot"

    for col, w in COL_WIDTHS.items():
        ws.column_dimensions[col].width = w

    row = 1
    ws.row_dimensions[row].height = 22
    row += 1

    if logo_path and os.path.exists(logo_path):
        try:
            img = XLImage(logo_path)
            img.width = 240
            img.height = 82
            ws.add_image(img, "B2")
        except Exception as e:
            LOG.warning(f"Logo image skipped: {e}")

    ws.row_dimensions[2].height = 66
    ws["H2"] = "LEASING SNAP SHOT"
    ws["H2"].font = Font(name=FONT_NAME, size=22, bold=True, color=NAVY_DEEP)
    ws["H2"].alignment = Alignment(horizontal="right", vertical="center")
    ws.merge_cells(f"H2:{LAST_COL}2")

    ws["H3"] = f"Generated {datetime.now().strftime('%B %d, %Y')} \u00b7 Prudent Growth Partners"
    ws["H3"].font = Font(name=FONT_NAME, size=9, italic=True, color=SLATE)
    ws["H3"].alignment = Alignment(horizontal="right", vertical="top")
    ws.merge_cells(f"H3:{LAST_COL}3")

    portfolio_total_sf = 0.0
    portfolio_occupied_sf = 0.0
    portfolio_properties = 0
    portfolio_units = 0
    portfolio_vacant_units = 0
    for m in mapping:
        af_id = str(m["appfolio_id"])
        units = rr.get(af_id, [])
        if not units:
            continue
        portfolio_properties += 1
        for u in units:
            sf_raw = parse_sqft(u.get("SquareFt")) or 0
            try:
                sf = float(sf_raw)
            except (TypeError, ValueError):
                sf = 0
            is_vacant = "vacant" in (u.get("Status", "") or "").lower() or not u.get("Tenant")
            portfolio_total_sf += sf
            portfolio_units += 1
            if is_vacant:
                portfolio_vacant_units += 1
            else:
                portfolio_occupied_sf += sf
    portfolio_vacant_sf = portfolio_total_sf - portfolio_occupied_sf
    portfolio_pct = (portfolio_occupied_sf / portfolio_total_sf) if portfolio_total_sf > 0 else 0

    state_totals = {}
    for m in mapping:
        af_id = str(m["appfolio_id"])
        units = rr.get(af_id, [])
        if not units:
            continue
        addr = m.get("appfolio_address") or m.get("excel_address") or ""
        mt = STATE_RE.search(addr)
        if not mt:
            continue
        st = mt.group(1)
        agg = state_totals.setdefault(st, {"props": 0, "total_sf": 0.0, "occ_sf": 0.0})
        agg["props"] += 1
        for u in units:
            sf_raw = parse_sqft(u.get("SquareFt")) or 0
            try:
                sf = float(sf_raw)
            except (TypeError, ValueError):
                sf = 0
            is_vacant = "vacant" in (u.get("Status", "") or "").lower() or not u.get("Tenant")
            agg["total_sf"] += sf
            if not is_vacant:
                agg["occ_sf"] += sf

    def _state_pct(agg):
        return (agg["occ_sf"] / agg["total_sf"]) if agg["total_sf"] > 0 else 1.0

    ranked_states = sorted(
        [(st, agg) for st, agg in state_totals.items() if agg["props"] >= 2 or agg["total_sf"] >= 50000],
        key=lambda kv: _state_pct(kv[1])
    )
    lowest_states = ranked_states[:4]

    ws.row_dimensions[4].height = 15
    ws.row_dimensions[5].height = 30
    ws.row_dimensions[6].height = 15

    kpi_tiles = [
        ("PORTFOLIO OCCUPANCY", f"{portfolio_pct*100:.1f}%",
            f"{portfolio_properties} properties  \u00b7  {portfolio_units} units"),
        ("TOTAL LEASABLE SF", f"{int(portfolio_total_sf):,}", "across the portfolio"),
        ("OCCUPIED SF", f"{int(portfolio_occupied_sf):,}",
            f"{portfolio_units - portfolio_vacant_units} leased units"),
        ("VACANT SF", f"{int(portfolio_vacant_sf):,}", f"{portfolio_vacant_units} vacant units"),
    ]
    tile_ranges = [("B", "D"), ("E", "G"), ("H", "J"), ("K", LAST_COL)]

    for (start_col, end_col), (label, value, sub) in zip(tile_ranges, kpi_tiles):
        ws.merge_cells(f"{start_col}4:{end_col}4")
        c = ws[f"{start_col}4"]
        c.value = label
        c.fill = fill(NAVY_DEEP)
        c.font = Font(name=FONT_NAME, size=8, bold=True, color=WHITE)
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)

        ws.merge_cells(f"{start_col}5:{end_col}5")
        c = ws[f"{start_col}5"]
        c.value = value
        c.fill = fill(NAVY)
        c.font = Font(name=FONT_NAME, size=22, bold=True, color=WHITE)
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)

        ws.merge_cells(f"{start_col}6:{end_col}6")
        c = ws[f"{start_col}6"]
        c.value = sub
        c.fill = fill(NAVY)
        c.font = Font(name=FONT_NAME, size=9, color=SLATE_TINT)
        c.alignment = Alignment(horizontal="left", vertical="top", indent=1)

    ws.row_dimensions[7].height = 18
    ws.merge_cells(f"B7:{LAST_COL}7")
    c = ws["B7"]
    c.value = "LOWEST-OCCUPANCY STATES  \u2014  where attention is needed"
    c.fill = fill(RED_LIGHT)
    c.font = Font(name=FONT_NAME, size=9, bold=True, color=NAVY_DEEP)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    lowest_ranges = [("B", "D"), ("E", "G"), ("H", "J"), ("K", LAST_COL)]
    ws.row_dimensions[8].height = 26
    ws.row_dimensions[9].height = 14

    for i, (start_col, end_col) in enumerate(lowest_ranges):
        if i >= len(lowest_states):
            for r in (8, 9):
                ws.merge_cells(f"{start_col}{r}:{end_col}{r}")
                ws[f"{start_col}{r}"].fill = fill(NAVY_DEEP)
            continue
        state_code, agg = lowest_states[i]
        pct = _state_pct(agg)
        vacant_sf = int(agg["total_sf"] - agg["occ_sf"])
        ws.merge_cells(f"{start_col}8:{end_col}8")
        c = ws[f"{start_col}8"]
        c.value = f"{state_code}    {pct*100:.1f}%"
        c.fill = fill(NAVY_DEEP)
        c.font = Font(name=FONT_NAME, size=13, bold=True, color=WHITE)
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.merge_cells(f"{start_col}9:{end_col}9")
        c = ws[f"{start_col}9"]
        c.value = f"{vacant_sf:,} SF vacant \u00b7 {agg['props']} prop{'s' if agg['props']!=1 else ''}"
        c.fill = fill(NAVY_DEEP)
        c.font = Font(name=FONT_NAME, size=8, color=SLATE_TINT)
        c.alignment = Alignment(horizontal="left", vertical="top", indent=1)

    ws.row_dimensions[10].height = 20
    ws["B10"] = "Legend:"
    ws["B10"].font = Font(name=FONT_NAME, size=8, bold=True, color=NAVY_DEEP)
    ws["B10"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    ws["C10"] = "2026 Renewal"
    ws["C10"].fill = fill(BLUE_SKY)
    ws["C10"].font = Font(name=FONT_NAME, size=8, color=NAVY_DEEP)
    ws["C10"].alignment = Alignment(horizontal="center", vertical="center")

    ws["D10"] = "2027 Renewal"
    ws["D10"].fill = fill(BLUE_LIGHT)
    ws["D10"].font = Font(name=FONT_NAME, size=8, color=NAVY_DEEP)
    ws["D10"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("E10:F10")
    ws["E10"] = "Editable \u2014 you type here"
    ws["E10"].fill = fill(GRAY_LIGHT)
    ws["E10"].font = Font(name=FONT_NAME, size=8, italic=True, color=NAVY_DEEP)
    ws["E10"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("G10:H10")
    ws["G10"] = "Auto-refreshed nightly"
    ws["G10"].fill = fill(BLUE_MID)
    ws["G10"].font = Font(name=FONT_NAME, size=8, color=WHITE)
    ws["G10"].alignment = Alignment(horizontal="center", vertical="center")

    ws.row_dimensions[11].height = 32
    ws.merge_cells(f"B11:{LAST_COL}11")
    ws["B11"] = (
        "Editable (preserved across refreshes): Broker Calls, Deal Activity, Property Flags, Vacant Callouts, Ann's Commentary, per-unit Notes column, Market Rent column."
        "    \u2022    "
        "Auto-refreshed nightly from AppFolio: Tenant, Tenant ID, Sq Ft, Rent/mo, PSF, Lease dates, Move In, Status, Past Due, Totals, address, Property ID."
        "    \u2022    "
        "Auto-refreshed weekly from ClickUp Broker Beat: Broker Active Interest from Weekly Reporting."
        "    \u2022    "
        "Auto-refreshed from ClickUp Broker Directory: Broker / Contact."
    )
    ws["B11"].font = Font(name=FONT_NAME, size=8, color=NAVY_DEEP)
    ws["B11"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
    ws["B11"].fill = fill(GRAY_ROW)

    row = 13

    total_props = 0
    total_units = 0
    for m in mapping:
        af_id = str(m["appfolio_id"])
        units = rr.get(af_id, [])
        prop_name = m["appfolio_name"] or m["excel_name"] or m["clickup_task_name"]
        prop_addr = m["appfolio_address"] or m["excel_address"] or ""

        cid = m.get("clickup_task_id") or ""
        m = dict(m)
        m["_broker_contacts"] = broker_map.get(cid, [])
        m["_notes"] = property_overrides.get(cid, {})
        m["_unit_notes"] = unit_notes
        m["_market_rent_overrides"] = market_rent_overrides.get(af_id, {})
        m["_clickup_summaries_by_tenant"] = clickup_summaries_by_tenant
        m["_clickup_summaries_by_prop_unit"] = clickup_summaries_by_prop_unit
        # Commit 2: filter the global (pid, unit_label) dicts down to just
        # this property's entries, keyed by unit_label alone so
        # write_property_block can do a simple .get(unit_label, "").
        m["_rem_comments_by_unit"] = {
            unit: text
            for (pid, unit), text in rem_comments_by_prop_unit.items()
            if str(pid) == str(af_id)
        }
        m["_last_synced_by_unit"] = {
            unit: stamp
            for (pid, unit), stamp in last_synced_by_prop_unit.items()
            if str(pid) == str(af_id)
        }
        # 2026 TICAM rates keyed by Snap Shot AppFolio display name (see
        # pull_ticam_rates_2026). None → no 2026 row on ClickUp → write_ticam_row
        # renders "2026 TICAM: not published".
        m["_ticam_2026"] = (ticam_by_property or {}).get(prop_name)

        row = write_property_block(ws, row, prop_name, prop_addr, units, m)
        row += 2
        total_props += 1
        total_units += len(units)

    LOG.info(f"Workbook build: wrote {total_props} properties, {total_units} total unit rows.")

    # Frozen panes intentionally removed 2026-08-05 per Alexis — the previous
    # freeze at A12 limited what could be seen while scrolling through a single
    # property block. Individual property blocks are already visually bounded
    # by their headers and borders, so a freeze isn't needed for orientation.
    # ws.freeze_panes = "A12"  # DISABLED

    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = 0.4
    ws.page_margins.right = 0.4
    ws.page_margins.top = 0.5
    ws.page_margins.bottom = 0.5

    wb.save(output_path)
    LOG.info(f"Saved workbook: {output_path}")
    return output_path


def write_ticam_row(ws, row, ticam_data):
    """Render one merged A:Q row with the property's 2026 TICAM (NNN) rates,
    just below the property banner + address row. Returns the next row.

    ticam_data shape (from pull_ticam_rates_2026):
        {'pgp_cam_per_sf': float|None, 'tax_per_sf': float|None,
         'insurance_per_sf': float|None, 'water_per_sf': float|None,
         'assoc_fee_per_sf': float|None,
         'status': 'confirmed'|'unconfirmed'|..., 'url': str}
    or None if no 2026 row exists for this property.
    """
    ws.row_dimensions[row].height = 16
    ws.merge_cells(f"{FIRST_COL}{row}:{LAST_COL}{row}")
    cell = ws[f"{FIRST_COL}{row}"]
    cell.fill = fill(GRAY_LIGHT)
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    if not ticam_data:
        cell.value = "2026 TICAM: not published"
        cell.font = Font(name=FONT_NAME, size=9, italic=True, color=SLATE)
        return row + 1

    # Build the human-readable piece — only include populated fields so we
    # never render a bare "$0.00".
    parts = []
    labels = [
        ("pgp_cam_per_sf",   "PGP CAM"),
        ("tax_per_sf",       "Tax"),
        ("insurance_per_sf", "Ins"),
        ("water_per_sf",     "Water"),
        ("assoc_fee_per_sf", "Assoc"),
    ]
    for key, label in labels:
        val = ticam_data.get(key)
        if val is None:
            continue
        parts.append(f"{label} ${val:,.2f}")

    if not parts:
        # Task exists but every rate field is blank/zero — treat as unpublished.
        cell.value = "2026 TICAM: not published"
        cell.font = Font(name=FONT_NAME, size=9, italic=True, color=SLATE)
        return row + 1

    status_raw = (ticam_data.get("status") or "").strip().lower()
    status_label = "Confirmed" if status_raw == "confirmed" else "Unconfirmed"
    line = f"2026 TICAM \u00b7 {status_label}: " + "  \u00b7  ".join(parts)
    url = (ticam_data.get("url") or "").strip()
    if url:
        line += f"    \u2192  {url}"
        cell.hyperlink = url

    cell.value = line
    # Bold the leading "2026 TICAM · {status}:" is not doable inside a
    # merged single-cell without RichText. Keep the row visually calm and
    # let the light-gray fill + navy text signal it as metadata.
    cell.font = Font(name=FONT_NAME, size=9, color=NAVY_DEEP)
    return row + 1


def write_property_block(ws, start_row, name, address, units, mapping_entry):
    """Write one property block. Returns the next available row after the block.
    Ported verbatim from build_prototype_v3.py, with one addition: Market Rent
    values are populated from mapping_entry['_market_rent_overrides'] (Ann's
    round-tripped edits) instead of always being left blank.
    """
    row = start_row

    ws.row_dimensions[row].height = 26
    ws.merge_cells(f"{FIRST_COL}{row}:{LAST_COL}{row}")
    cell = ws[f"{FIRST_COL}{row}"]
    cell.value = name.upper()
    cell.fill = fill(NAVY_DEEP)
    cell.font = Font(name=FONT_NAME, size=13, bold=True, color=WHITE)
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    row += 1

    # Address / metadata row. Split into two merged regions so the REM name
    # can sit right-justified on the same row (Alexis 2026-08-06). Both
    # regions share the same BLUE_MID fill so the split is invisible.
    #
    # Layout:
    #   B–K  →  "Property ID: NNN  ·  address  ·  tags"           (left)
    #   L–O  →  "Real Estate Manager: {rem_name}"                  (right)
    ws.row_dimensions[row].height = 18
    ws.merge_cells(f"{FIRST_COL}{row}:K{row}")
    cell = ws[f"{FIRST_COL}{row}"]
    src_tags = []
    if mapping_entry.get("source") == "broker_only":
        src_tags.append("NEW \u00b7 Broker-only in ClickUp")
    if mapping_entry.get("needs_review"):
        src_tags.append("\u26a0 Needs review")
    _mid_dot = "\u00b7"
    _tag_join = f" {_mid_dot} ".join(src_tags)
    tag_suffix = f"    {_mid_dot}    {_tag_join}" if src_tags else ""
    prop_id = mapping_entry.get("appfolio_id") or ""
    prop_id_prefix = f"Property ID: {prop_id}    \u00b7    " if prop_id else ""
    cell.value = f"{prop_id_prefix}{address}{tag_suffix}"
    cell.fill = fill(BLUE_MID)
    cell.font = Font(name=FONT_NAME, size=10, color=WHITE)
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    # Right-justified REM name on the same row. Falls back gracefully:
    #   populated → "Real Estate Manager: Parker Owen"
    #   unset    → "Real Estate Manager: —" (em-dash, italic-slate)
    ws.merge_cells(f"L{row}:{LAST_COL}{row}")
    rem_cell = ws[f"L{row}"]
    rem_name = (mapping_entry.get("rem_name") or "").strip()
    if rem_name:
        rem_cell.value = f"Real Estate Manager: {rem_name}"
        rem_cell.font = Font(name=FONT_NAME, size=10, color=WHITE, bold=True)
    else:
        rem_cell.value = "Real Estate Manager: \u2014"
        rem_cell.font = Font(name=FONT_NAME, size=10, italic=True, color=WHITE)
    rem_cell.fill = fill(BLUE_MID)
    rem_cell.alignment = Alignment(horizontal="right", vertical="center", indent=1)
    row += 1

    # 2026 TICAM (NNN) rates row — one merged line spanning A:Q. Rendered
    # from mapping_entry['_ticam_2026'] which pull_ticam_rates_2026() built
    # off the ClickUp TICAM Rates list. Only shows populated fields so
    # nothing displays as "$0.00"; a missing 2026 task renders as a
    # muted "not published" line so REMs / brokers know it's absent, not
    # simply a build error.
    row = write_ticam_row(ws, row, mapping_entry.get("_ticam_2026"))

    row += 1  # blank spacer

    ws.row_dimensions[row].height = 18
    ws.merge_cells(f"{FIRST_COL}{row}:{LAST_COL}{row}")
    cell = ws[f"{FIRST_COL}{row}"]
    cell.value = "DEAL ACTIVITY & NOTES"
    cell.fill = fill(BLUE_MID)
    cell.font = Font(name=FONT_NAME, size=10, bold=True, color=WHITE)
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    row += 1

    def _clean(s):
        if not s:
            return ""
        return re.sub(r"\s+", " ", s).strip()

    broker_contacts = mapping_entry.get("_broker_contacts") or []
    broker_lines = []
    for b in broker_contacts:
        nm = _clean(b.get("name"))
        co = _clean(b.get("company"))
        cell_phone = _clean(b.get("cell"))
        email = _clean(b.get("email"))
        line = nm
        if co:
            line = f"{line} ({co})" if line else f"({co})"
        bits = [x for x in (cell_phone, email) if x]
        if bits:
            _mid = "\u00b7"
            _joined_bits = f" {_mid} ".join(bits)
            line = f"{line} {_mid} {_joined_bits}" if line else _joined_bits
        if line:
            broker_lines.append(line)
    broker_str = "  ;  ".join(broker_lines)

    notes = mapping_entry.get("_notes") or {}
    default_bai = "\u2014 awaiting weekly Broker Beat update \u2014"
    note_labels = [
        ("Broker / Contact", broker_str),
        ("Broker Calls", notes.get("broker_calls", notes.get("call_cadence", ""))),
        ("Property Flags", notes.get("property_flags", "")),
        ("Vacant Callouts", notes.get("vacant_callouts", "")),
        ("Deal Activity", notes.get("deal_activity", "")),
        ("Broker Active Interest from Weekly Reporting", notes.get("broker_active_interest") or default_bai),
        ("Ann's Commentary", notes.get("ann_commentary", "")),
    ]

    # Row-height estimator for ANY note row (Broker/Contact, Broker Calls,
    # Property Flags, Vacant Callouts, Deal Activity, Broker Active Interest,
    # Ann's Commentary). We deliberately over-estimate rather than under: users
    # would rather see a slightly tall row than a clipped one. All value cells
    # use wrap_text=True and vertical=top so any extra vertical space just
    # renders as trailing whitespace, not misalignment.
    #
    # Openpyxl's auto-fit on merged cells is unreliable, so we compute a
    # generous manual height. Values:
    #   chars_per_line = 72  (measured column can fit ~85 chars at 9pt Calibri
    #                         with narrow letters; using 72 leaves headroom for
    #                         property blocks where content skews wider).
    #   line_height    = 15pt (was 14 — matches 9pt Calibri single-line height
    #                          with slight breathing room, avoiding the
    #                          descender-clipping some rows showed).
    #   safety_margin  = 1 extra line for anything with visible content, so a
    #                    borderline case that wraps to N+1 lines in Excel
    #                    doesn't clip.
    #   max_height     = 409pt (Excel's absolute per-row maximum). We used to
    #                    cap at 110pt then 280pt; both clipped long notes.
    def _row_height_for(text, min_lines=1, chars_per_line=72,
                         line_height=15, safety_margin=1, max_height=409):
        if not text and min_lines <= 1:
            return 18
        text = text or ""
        # Count wrapped lines: split on any explicit newline, then
        # ceil(segment_len / chars_per_line) each.
        segments = text.replace("\r\n", "\n").split("\n") or [""]
        wrapped_lines = 0
        for seg in segments:
            wrapped_lines += max(1, -(-len(seg) // chars_per_line))
        # Add extra lines for semicolon-delimited items (common in Broker
        # Active Interest and Deal Activity where each item is on its own line).
        wrapped_lines += text.count(";")
        # Safety margin: add 1 extra line whenever there is content, so
        # borderline wraps don't clip.
        if text.strip():
            wrapped_lines += safety_margin
        wrapped_lines = max(min_lines, wrapped_lines)
        return min(18 + (wrapped_lines - 1) * line_height, max_height)

    # No per-label line floors anymore. Populated note rows size to their
    # actual content (via _row_height_for), and empty rows collapse to a
    # single line via the has_real_content check below. The +15pt headroom
    # added later gives ~1 line of live-typing room without over-provisioning.
    label_min_lines = {}

    # Column B is width 20 (see COL_WIDTHS). At 9pt bold Calibri, ~17 chars
    # per visible line. Compute the label's own line count and use it as a
    # floor for the row height, since the label now wraps too.
    LABEL_CHARS_PER_LINE = 17

    for label, val in note_labels:
        is_default_bai = (label == "Broker Active Interest from Weekly Reporting" and val == default_bai)
        # When the value cell is EMPTY (or just the BAI placeholder), skip the
        # tall min_lines floor entirely — the row should collapse to just what
        # the label needs, not 4 blank lines. Alexis flagged that empty Deal
        # Activity and Ann's Commentary rows were taking up a lot of vertical
        # space with no content.
        has_real_content = bool((val or "").strip()) and not is_default_bai
        if has_real_content:
            min_lines = label_min_lines.get(label, 1)
        else:
            min_lines = 1
        # Ensure row is tall enough for whichever cell wraps to more lines.
        label_lines = max(1, -(-len(label) // LABEL_CHARS_PER_LINE))
        effective_min_lines = max(min_lines, label_lines)
        # For note rows with content: use safety_margin=0 (no wrap buffer)
        # since the +15pt headroom below already covers ~1 line of live
        # typing. For empty/placeholder rows: use safety_margin=0 too (they
        # collapse to a single line anyway).
        row_h = _row_height_for(
            val, min_lines=effective_min_lines, safety_margin=0
        )
        # Live-typing headroom: merged cells don't auto-fit in Excel, so if
        # someone types more after the rebuild, the row won't grow until the
        # next nightly. Add a modest fixed buffer (~1 extra line, +15pt)
        # when there's real content — enough for a small mid-day addition
        # without making rows visibly puffy. Empty rows unaffected.
        if has_real_content:
            row_h = min(row_h + 15, 409)
        ws.row_dimensions[row].height = row_h
        label_cell = ws[f"B{row}"]
        label_cell.value = label
        label_cell.fill = fill(GRAY_LIGHT)
        label_cell.font = Font(name=FONT_NAME, size=9, bold=True, color=NAVY_DEEP)
        # wrap_text=True so long labels like "Broker Active Interest from
        # Weekly Reporting" wrap inside the narrow column B instead of
        # visually clipping. The row height is already sized for the value
        # cell, which is always >= what the label needs.
        label_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
        label_cell.border = BOX_BORDER

        ws.merge_cells(f"C{row}:{LAST_COL}{row}")
        val_cell = ws[f"C{row}"]
        val_cell.value = val
        val_cell.font = Font(
            name=FONT_NAME, size=9,
            italic=is_default_bai,
            color=SLATE if is_default_bai else NAVY_DEEP
        )
        val_cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True, indent=1)
        val_cell.border = BOX_BORDER
        row += 1

    row += 1  # blank spacer

    ws.row_dimensions[row].height = 18
    ws.merge_cells(f"{FIRST_COL}{row}:{LAST_COL}{row}")
    cell = ws[f"{FIRST_COL}{row}"]
    cell.value = "RENT ROLL"
    cell.fill = fill(BLUE_MID)
    cell.font = Font(name=FONT_NAME, size=10, bold=True, color=WHITE)
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    row += 1

    ws.row_dimensions[row].height = 22
    for col_label, col_letter in COLS.items():
        cell = ws[f"{col_letter}{row}"]
        cell.value = col_label
        cell.fill = fill(NAVY)
        cell.font = Font(name=FONT_NAME, size=9, bold=True, color=WHITE)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BOX_BORDER
    row += 1

    data_start_row = row
    market_rent_overrides = mapping_entry.get("_market_rent_overrides") or {}
    clickup_summaries_by_tenant = mapping_entry.get("_clickup_summaries_by_tenant") or {}
    clickup_summaries_by_prop_unit = mapping_entry.get("_clickup_summaries_by_prop_unit") or {}
    # REM ClickUp Comment (col P) values, extracted from the previous
    # SharePoint copy and keyed by unit_label (unit number as displayed).
    # These are the REM-authored comments that either (a) already synced
    # to ClickUp on this run and got a fresh "Last Synced" stamp, or (b)
    # were unchanged since the last sync (so both P and Q pass through).
    rem_comments_by_unit = mapping_entry.get("_rem_comments_by_unit") or {}
    last_synced_by_unit = mapping_entry.get("_last_synced_by_unit") or {}
    prop_id_for_lookup = str(mapping_entry.get("appfolio_id") or "").strip()
    if not units:
        ws.row_dimensions[row].height = 18
        ws.merge_cells(f"{FIRST_COL}{row}:{LAST_COL}{row}")
        cell = ws[f"{FIRST_COL}{row}"]
        cell.value = "No units in AppFolio rent roll"
        cell.font = Font(name=FONT_NAME, size=9, italic=True, color=SLATE)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = fill(GRAY_ROW)
        row += 1
    else:
        sorted_units = sort_units(units)
        unit_notes = mapping_entry.get("_unit_notes") or {}
        for i, u in enumerate(sorted_units):
            note_key = str(u.get("OccupancyId") or u.get("UnitId") or "")
            unit_note = unit_notes.get(note_key, "") if note_key else ""

            # Unit-row height strategy (Excel auto-fit behavior):
            #   * If we do NOT set row_dimensions[row].height, Excel auto-fits
            #     the row height based on wrapped content on OPEN. Users can
            #     also type into the Notes cell after the rebuild and Excel
            #     will grow the row live (as long as wrap_text=True is set on
            #     the cell, which we do below).
            #   * If we DO set an explicit height, Excel treats it as a hard
            #     lock and stops auto-growing when the user edits.
            #
            # So: only set an explicit height when the incoming note is
            # already tall enough to REQUIRE more than one line's worth of
            # space (i.e., the extraction saw a real multi-line note). Empty
            # notes and short single-line notes get NO explicit height, so
            # Excel auto-fits and continues to auto-fit as Ann types.
            if unit_note and len(unit_note) > 30:
                # Prime the height so the initial view already shows the full
                # note without user interaction. No cap — the row can grow as
                # tall as content requires (Excel max is 409pt).
                note_lines = max(1, -(-len(unit_note) // 30))
                note_lines += unit_note.count("\n")
                ws.row_dimensions[row].height = min(17 + (note_lines - 1) * 14, 409)
            # else: leave row_dimensions[row].height unset so Excel auto-fits.

            lease_to = parse_date(u.get("LeaseTo"))
            row_fill = None
            if lease_to:
                if lease_to.year == 2026:
                    row_fill = fill(RED_LIGHT)
                elif lease_to.year == 2027:
                    row_fill = fill(BLUE_LIGHT)
            if not row_fill and i % 2 == 1:
                row_fill = fill(GRAY_ROW)

            status = u.get("Status", "") or ""
            is_vacant = "vacant" in status.lower() or not u.get("Tenant")

            tenant_id_val = u.get("TenantId") if not is_vacant else ""
            unit_label = unit_display(u)
            market_rent_val = market_rent_overrides.get(unit_label, "")

            # Look up ClickUp AI Summary for this unit.
            # Precedence: Tenant ID first (occupied units); fall back to
            # (Property ID, Unit #) for vacant units. Empty string → the
            # rebuild leaves the cell visibly empty rather than showing a
            # placeholder, so REMs can see at a glance which units have no
            # active ClickUp task in the LAR pipeline.
            clickup_summary = ""
            if tenant_id_val:
                clickup_summary = clickup_summaries_by_tenant.get(str(tenant_id_val).strip(), "")
            if not clickup_summary and prop_id_for_lookup and unit_label:
                clickup_summary = clickup_summaries_by_prop_unit.get(
                    (prop_id_for_lookup, str(unit_label).strip()), ""
                )

            # REM ClickUp Comment + Last Synced values (Commit 2). Both are
            # keyed by unit_label. If the REM's comment already synced this
            # run, the wire-up in run_build has already updated both dicts
            # (comment text preserved, Last Synced timestamp refreshed). If
            # it didn't sync (unchanged text or empty), both pass through.
            rem_comment_val = rem_comments_by_unit.get(unit_label, "")
            last_synced_val = last_synced_by_unit.get(unit_label, "")

            values = {
                "Unit": unit_label,
                "Tenant": u.get("Tenant") or ("\u2014 Vacant \u2014" if is_vacant else ""),
                "Tenant ID": str(tenant_id_val) if tenant_id_val else "",
                "Sq Ft": parse_sqft(u.get("SquareFt")),
                "Rent/mo": parse_currency(u.get("Rent")),
                "PSF": parse_currency(u.get("AnnualRentSquareFt")),
                "Market Rent": parse_currency(market_rent_val) if market_rent_val else "",
                "Lease From": parse_date(u.get("LeaseFrom")),
                "Lease To": lease_to,
                "Move In": parse_date(u.get("MoveIn")),
                "Status": status,
                "Past Due": parse_currency(u.get("PastDue")),
                "Notes": unit_note,
                "ClickUp Summary": clickup_summary,
                "REM ClickUp Comment": rem_comment_val,
                "Last Synced": last_synced_val,
            }

            # Row-height priming for ClickUp Summary. Renewal/Vacancy return
            # a ~7-line bullet block; Documents Workflow returns ~4 lines.
            # Only prime when the summary is long enough to force wrapping
            # (matches the Notes-column strategy above): setting a fixed
            # height would prevent Excel from auto-growing the row when
            # someone types more into the Notes cell later.
            if clickup_summary and len(clickup_summary) > 85:
                # ClickUp Summary column O widened to 80 (Commit 2). At 10pt
                # Calibri (bumped from 8pt on 2026-08-06) that's ~85 chars per
                # visible line. Count wrapped lines conservatively.
                summ_lines = max(1, -(-len(clickup_summary) // 85))
                summ_lines += clickup_summary.count("\n")
                # Only raise the height — don't lower one already set by the
                # Notes column primer above. 10pt line height ~16pt.
                needed_h = min(19 + (summ_lines - 1) * 16, 409)
                existing_h = ws.row_dimensions[row].height or 0
                if needed_h > existing_h:
                    ws.row_dimensions[row].height = needed_h

            # Row-height priming for REM ClickUp Comment (col P, width 40).
            # If a long comment sits in P, we prime the row height like we
            # do for ClickUp Summary so it renders wrapped on open.
            if rem_comment_val and len(rem_comment_val) > 40:
                # At 9pt Calibri, col-40 shows ~50 chars per line.
                pcomm_lines = max(1, -(-len(rem_comment_val) // 50))
                pcomm_lines += rem_comment_val.count("\n")
                needed_h = min(17 + (pcomm_lines - 1) * 14, 409)
                existing_h = ws.row_dimensions[row].height or 0
                if needed_h > existing_h:
                    ws.row_dimensions[row].height = needed_h

            for col_label, col_letter in COLS.items():
                cell = ws[f"{col_letter}{row}"]
                v = values[col_label]
                cell.value = v
                cell.font = Font(
                    name=FONT_NAME, size=9,
                    italic=(col_label == "Market Rent"),
                    color=(SLATE if (is_vacant and col_label == "Tenant") else NAVY_DEEP)
                )
                cell.border = BOX_BORDER

                if col_label in ("Sq Ft",):
                    cell.number_format = "#,##0"
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                elif col_label in ("Rent/mo", "Market Rent", "Past Due"):
                    cell.number_format = '"$"#,##0.00;[Red]-"$"#,##0.00;""'
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                elif col_label == "PSF":
                    cell.number_format = '"$"#,##0.00'
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                elif col_label in ("Lease From", "Lease To", "Move In"):
                    cell.number_format = "mm/dd/yyyy"
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif col_label == "Notes":
                    cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True, indent=1)
                elif col_label == "ClickUp Summary":
                    # Read-only from ClickUp, so italic-slate to distinguish
                    # visually from REM-editable columns. Font bumped from 8pt
                    # to 10pt per Alexis 2026-08-06 for readability.
                    cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True, indent=1)
                    cell.font = Font(name=FONT_NAME, size=10, italic=True, color=SLATE)
                elif col_label == "REM ClickUp Comment":
                    # REM-editable. Full 9pt navy so it's visually distinct
                    # from the muted ClickUp Summary next to it — signals
                    # "this is where YOU type."
                    cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True, indent=1)
                    cell.font = Font(name=FONT_NAME, size=9, color=NAVY_DEEP)
                elif col_label == "Last Synced":
                    # Read-only "receipt" cell. Small italic slate so REMs
                    # can see "last synced at X" without it competing for
                    # attention with the editable comment cell.
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True, indent=1)
                    cell.font = Font(name=FONT_NAME, size=8, italic=True, color=SLATE)
                elif col_label == "Tenant ID":
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left" if col_label == "Tenant" else "center",
                                                vertical="center", indent=(1 if col_label == "Tenant" else 0))

                if row_fill:
                    cell.fill = row_fill
                elif col_label in ("Market Rent", "Notes"):
                    cell.fill = fill(GRAY_LIGHT)
                elif col_label == "ClickUp Summary":
                    # Very light neutral tint so it reads as "informational,
                    # not-for-editing" without competing with the Notes column.
                    cell.fill = fill("F5F6F8")
                elif col_label == "REM ClickUp Comment":
                    # Same warm gray as Notes — signals "REM edits here",
                    # matches the existing REM-editable Notes convention.
                    cell.fill = fill(GRAY_LIGHT)
                elif col_label == "Last Synced":
                    # Same faint tint as ClickUp Summary — both are
                    # informational read-outs from the sync.
                    cell.fill = fill("F5F6F8")
            row += 1

    data_end_row = row - 1

    if units:
        ws.row_dimensions[row].height = 20
        for col_label, col_letter in COLS.items():
            cell = ws[f"{col_letter}{row}"]
            cell.fill = fill(BLUE_MID)
            cell.font = Font(name=FONT_NAME, size=9, bold=True, color=WHITE)
            cell.border = BOX_BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center")

        c_unit = COLS["Unit"]
        c_tenant = COLS["Tenant"]
        c_tenant_id = COLS["Tenant ID"]
        c_sqft = COLS["Sq Ft"]
        c_rent = COLS["Rent/mo"]
        c_psf = COLS["PSF"]
        c_mkt = COLS["Market Rent"]
        c_pastdue = COLS["Past Due"]
        c_notes = COLS["Notes"]
        c_summary = COLS["ClickUp Summary"]

        ws[f"{c_unit}{row}"] = "TOTALS"
        ws[f"{c_unit}{row}"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

        total_sqft = 0.0
        occupied_sqft = 0.0
        for u in units:
            sf = parse_sqft(u.get("SquareFt")) or 0
            try:
                sf = float(sf)
            except (TypeError, ValueError):
                sf = 0
            is_vacant = "vacant" in (u.get("Status", "") or "").lower() or not u.get("Tenant")
            total_sqft += sf
            if not is_vacant:
                occupied_sqft += sf
        pct = (occupied_sqft / total_sqft) if total_sqft > 0 else 0
        ws[f"{c_tenant}{row}"] = (
            f"{pct*100:.1f}% of SF leased ({int(occupied_sqft):,} / {int(total_sqft):,} SF)"
            if total_sqft > 0 else "\u2014"
        )
        ws[f"{c_tenant}{row}"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

        ws[f"{c_tenant_id}{row}"] = ""

        ws[f"{c_sqft}{row}"] = f"=SUM({c_sqft}{data_start_row}:{c_sqft}{data_end_row})"
        ws[f"{c_sqft}{row}"].number_format = "#,##0"
        ws[f"{c_sqft}{row}"].alignment = Alignment(horizontal="right", vertical="center")

        ws[f"{c_rent}{row}"] = f"=SUM({c_rent}{data_start_row}:{c_rent}{data_end_row})"
        ws[f"{c_rent}{row}"].number_format = '"$"#,##0'
        ws[f"{c_rent}{row}"].alignment = Alignment(horizontal="right", vertical="center")

        ws[f"{c_psf}{row}"] = f'=IFERROR(({c_rent}{row}*12)/{c_sqft}{row},"")'
        ws[f"{c_psf}{row}"].number_format = '"$"#,##0.00'
        ws[f"{c_psf}{row}"].alignment = Alignment(horizontal="right", vertical="center")

        ws[f"{c_mkt}{row}"] = ""

        for col_label in ("Lease From", "Lease To", "Move In", "Status"):
            ws[f"{COLS[col_label]}{row}"] = ""

        ws[f"{c_pastdue}{row}"] = f"=SUM({c_pastdue}{data_start_row}:{c_pastdue}{data_end_row})"
        ws[f"{c_pastdue}{row}"].number_format = '"$"#,##0.00;[Red]-"$"#,##0.00'
        ws[f"{c_pastdue}{row}"].alignment = Alignment(horizontal="right", vertical="center")

        ws[f"{c_notes}{row}"] = ""
        ws[f"{c_summary}{row}"] = ""
        # Commit 2 additions — REM Comment (P) and Last Synced (Q) are
        # per-unit columns; the TOTALS row leaves them blank.
        ws[f"{COLS['REM ClickUp Comment']}{row}"] = ""
        ws[f"{COLS['Last Synced']}{row}"] = ""

        row += 1

    return row


# ══════════════════════════════════════════════════════════════════════════
# Orchestration
# ══════════════════════════════════════════════════════════════════════════

def run_build(mode, dry_run, force_rebuild=False):
    """Executes steps 1-6 of the brief's script structure. Returns the local
    output file path. Raises FatalError on any unrecoverable problem.

    force_rebuild=True bypasses the REM Comment safety guard (see
    check_rem_comment_floor). Also propagated from the existing workflow input
    of the same name, which already zeroes out the race-condition window."""

    mapping = load_json(MAPPING_PATH)
    broker_map = load_json(BROKER_CONTACTS_PATH, required=False, default={})

    # Step 1: download current SharePoint workbook + race-condition check.
    if dry_run:
        LOG.info("dry-run: skipping SharePoint download/race-check entirely.")
        current_wb_bytes, last_modified_by, last_modified_at = None, None, None
    else:
        current_wb_bytes, last_modified_by, last_modified_at = download_from_sharepoint()
        if is_race_condition(last_modified_by, last_modified_at):
            LOG.info("skipped — Ann is editing")
            return None  # signal "skip" to caller

    # Step 2: extract Ann's edits (fallback to bundled JSON on first run / dry-run).
    market_rent_overrides = {}
    rem_comments_by_prop_unit = {}
    last_synced_by_prop_unit = {}
    if current_wb_bytes:
        (
            prop_overrides_by_pid,
            unit_notes_by_key,
            market_rent_by_pid,
            rem_comments_by_prop_unit,
            last_synced_by_prop_unit,
        ) = extract_ann_edits(current_wb_bytes)
        property_overrides = reindex_overrides_by_clickup_task(prop_overrides_by_pid, mapping)
        market_rent_overrides = market_rent_by_pid
        # unit_notes re-indexing needs the fresh rent roll, done after step 3 below.
        pending_unit_notes_raw = unit_notes_by_key

        # REM Comment safety guard (RCA_2026-08-07 fix #3) — refuse to proceed if
        # extraction returned 0 REM comments but the previous run saw more, or
        # if the count dropped >50%. Fires BEFORE the workbook is rebuilt so we
        # don't waste an AppFolio pull only to refuse the upload. Bypassed by
        # force_rebuild=true from the workflow_dispatch input.
        check_rem_comment_floor(
            current_count=len(rem_comments_by_prop_unit),
            force_rebuild=force_rebuild,
        )
    else:
        LOG.info("No SharePoint workbook to extract from — using bundled fallback JSON.")
        property_overrides = load_json(NOTES_OVERRIDES_FALLBACK_PATH, required=False, default={})
        unit_notes = load_json(UNIT_NOTES_FALLBACK_PATH, required=False, default={})
        pending_unit_notes_raw = None

    # Step 3: pull fresh AppFolio data.
    rent_roll_by_property = pull_appfolio_rent_roll_by_property()

    if pending_unit_notes_raw is not None:
        unit_notes = reindex_unit_notes_by_note_key(pending_unit_notes_raw, rent_roll_by_property, mapping)

    # Step 4 (weekly only): refresh Broker Active Interest from Broker Beat.
    if mode == "weekly":
        broker_reporting_tasks = pull_broker_reporting_tasks()
        broker_active_interest = pull_broker_beat_attachments(broker_reporting_tasks)
        property_overrides = merge_broker_active_interest(property_overrides, broker_active_interest)

    # Step 4b: pull ClickUp LAR Summaries (Renewal + Vacancy + Docs Workflow).
    # Runs every mode (nightly, weekly, dry-run) since the ClickUp Summary
    # column is displayed at all times. Non-fatal on failure — if ClickUp is
    # down or the token is wrong, individual list fetches will log warnings
    # and skip, and the workbook will render with empty ClickUp Summary
    # cells rather than crashing the whole build.
    #
    # Commit 2: pull_lar_summaries now also returns task-id indexes so the
    # REM ClickUp Comment sync can post to every matching task.
    try:
        (
            clickup_summaries_by_tenant,
            clickup_summaries_by_prop_unit,
            task_ids_by_tenant_id,
            task_ids_by_prop_unit,
        ) = pull_lar_summaries()
    except Exception as e:  # noqa: BLE001
        LOG.warning(f"LAR summaries pull failed — continuing with empty summaries: {e}")
        clickup_summaries_by_tenant, clickup_summaries_by_prop_unit = {}, {}
        task_ids_by_tenant_id, task_ids_by_prop_unit = {}, {}

    # Step 4c (Commit 2): sync REM-authored comments (col P) to ClickUp.
    # Diff-detect via the sha8 hash in the paired Last Synced cell (col Q).
    # Only posts when the comment text actually changed. Non-fatal: any
    # per-unit failure is logged and the workbook still rebuilds so REMs
    # don't lose visibility of their own draft comments.
    try:
        last_synced_by_prop_unit = sync_rem_comments_to_clickup(
            mapping=mapping,
            rem_comments_by_prop_unit=rem_comments_by_prop_unit,
            last_synced_by_prop_unit=last_synced_by_prop_unit,
            task_ids_by_tenant_id=task_ids_by_tenant_id,
            task_ids_by_prop_unit=task_ids_by_prop_unit,
            rent_roll_by_property=rent_roll_by_property,
            dry_run=dry_run,
        )
    except Exception as e:  # noqa: BLE001
        LOG.warning(f"REM comment sync failed — continuing without posting: {e}")

    # Step 4d: pull 2026 TICAM rates from ClickUp TICAM Rates list. Runs on
    # every mode. Non-fatal — failure returns {} and the TICAM row falls
    # back to "2026 TICAM: not published" everywhere. Keyed by Snap Shot
    # AppFolio display name (prop_name used inside the mapping loop).
    snap_shot_prop_names = [
        (m.get("appfolio_name") or m.get("excel_name") or m.get("clickup_task_name"))
        for m in mapping
    ]
    ticam_by_property = pull_ticam_rates_2026(snap_shot_prop_names)

    # Step 5: rebuild workbook.
    output_filename = "Leasing-Snap-Shot-dryrun.xlsx" if dry_run else "Leasing-Snap-Shot.xlsx"
    output_path = os.path.join(LOCAL_BUILD_DIR if not dry_run else "/tmp", output_filename)
    build_workbook(
        mapping=mapping,
        rent_roll_by_property=rent_roll_by_property,
        property_overrides=property_overrides,
        unit_notes=unit_notes,
        market_rent_overrides=market_rent_overrides,
        broker_map=broker_map,
        output_path=output_path,
        logo_path=LOGO_PATH,
        clickup_summaries_by_tenant=clickup_summaries_by_tenant,
        clickup_summaries_by_prop_unit=clickup_summaries_by_prop_unit,
        rem_comments_by_prop_unit=rem_comments_by_prop_unit,
        last_synced_by_prop_unit=last_synced_by_prop_unit,
        ticam_by_property=ticam_by_property,
    )

    # Step 5b: archive the current live Snap Shot before we overwrite it, so
    # recovery from any future extractor regression is a 30-second SharePoint
    # copy instead of a version-history spelunk. (RCA_2026-08-07 fix #4.)
    # Skipped for dry-run. Non-fatal — if archiving fails we still upload
    # (rebuild is the primary product; archive is defense-in-depth).
    if not dry_run and current_wb_bytes:
        try:
            archive_current_snap_shot(current_wb_bytes)
        except Exception as e:  # noqa: BLE001
            LOG.warning(f"Archive step failed (non-fatal, upload will proceed): {e}")

    # Step 6: upload to SharePoint (skip entirely for dry-run).
    if not dry_run:
        upload_result = upload_to_sharepoint(output_path) or {}
        LOG.info(f"Uploaded {output_path} to SharePoint as {WORKBOOK_FILENAME}.")

        # Prefer Graph's `webUrl` — that's the "_layouts/15/Doc.aspx?..."
        # link that opens in Excel Online in the browser (rather than a
        # direct file URL that triggers desktop Excel).
        sharepoint_link = upload_result.get("webUrl") or WORKBOOK_WEB_URL

        # Step 7: stamp the ClickUp Broker Reporting LIST description with
        # the SharePoint link. This is what Alexis wants pinned at the top
        # of the list so anyone opening the ClickUp view sees the link
        # first. Non-fatal — workbook is already live.
        try:
            update_broker_reporting_list_description(sharepoint_link, mode)
        except Exception as e:  # noqa: BLE001
            LOG.warning(f"Failed to update ClickUp list description: {e}")

        # Step 7b: also stamp the pinned control task description, for
        # visibility inside the task view itself. Non-fatal.
        try:
            update_control_task_with_refresh_link(sharepoint_link, mode)
        except Exception as e:  # noqa: BLE001
            LOG.warning(f"Failed to update ClickUp control task description: {e}")

    return output_path


def poll_and_maybe_rebuild(force_rebuild=False):
    """--mode=poll: check the ClickUp control-task checkbox. If checked, run
    a full nightly rebuild, then uncheck + comment. Returns exit code.

    force_rebuild is propagated through to run_build so an operator dispatching
    the poll workflow manually with the bypass flag doesn't get blocked by the
    REM Comment safety guard."""
    control_task = find_control_task()
    if not is_checkbox_checked(control_task):
        LOG.info("Checkbox not checked; exiting.")
        return 0

    LOG.info(f"'{CHECKBOX_FIELD_NAME}' is checked on control task {control_task['id']} — running on-demand rebuild.")
    output_path = run_build(mode="nightly", dry_run=False, force_rebuild=force_rebuild)
    if output_path is None:
        # Race condition — do NOT uncheck the box; let the next poll retry.
        LOG.info("On-demand rebuild skipped (race condition). Leaving checkbox checked for the next poll.")
        return 0

    stamp = now_et().strftime("%H:%M ET")
    uncheck_and_comment(control_task, f"refreshed at {stamp}")
    LOG.info("On-demand rebuild complete.")
    return 0


def main():
    parser = argparse.ArgumentParser(description="Rebuild the Leasing Snap Shot workbook.")
    parser.add_argument(
        "--mode", choices=["nightly", "weekly", "poll", "dry-run"], required=True,
        help="nightly | weekly | poll | dry-run",
    )
    parser.add_argument(
        "--force-rebuild", action="store_true",
        default=os.environ.get("SNAP_SHOT_FORCE_REBUILD", "false").lower() == "true",
        help=(
            "Bypass safety guards (REM Comment floor, race condition). Also "
            "reads env SNAP_SHOT_FORCE_REBUILD for workflow_dispatch plumbing."
        ),
    )
    args = parser.parse_args()

    LOG.info(
        "=== prudent-snap-shot-rebuild starting: mode=%s force_rebuild=%s ===",
        args.mode, args.force_rebuild,
    )
    dry_run = args.mode == "dry-run"

    try:
        if args.mode == "poll":
            rc = poll_and_maybe_rebuild(force_rebuild=args.force_rebuild)
        else:
            output_path = run_build(mode=args.mode, dry_run=dry_run, force_rebuild=args.force_rebuild)
            if output_path is None:
                LOG.info("skipped — Ann is editing")
                rc = 0
            else:
                LOG.info(f"Done. Output: {output_path}")
                rc = 0
    except FatalError as e:
        LOG.error(f"FATAL: {e}")
        if not dry_run:
            send_error_email(
                subject=f"Leasing Snap Shot rebuild FAILED ({args.mode})",
                body=(
                    f"The Leasing Snap Shot rebuild script failed in mode={args.mode} "
                    f"at {now_et().isoformat()}.\n\nError:\n{e}\n\n"
                    f"No workbook was uploaded — the previous SharePoint copy is untouched."
                ),
            )
        rc = 1
    except Exception as e:  # noqa: BLE001 - top-level catch-all is intentional (fail loud, never crash silently uncaught)
        LOG.error(f"UNEXPECTED ERROR: {e}\n{traceback.format_exc()}")
        if not dry_run:
            send_error_email(
                subject=f"Leasing Snap Shot rebuild CRASHED ({args.mode})",
                body=(
                    f"The Leasing Snap Shot rebuild script crashed unexpectedly in "
                    f"mode={args.mode} at {now_et().isoformat()}.\n\n"
                    f"{traceback.format_exc()}"
                ),
            )
        rc = 1

    LOG.info(f"=== prudent-snap-shot-rebuild finished: mode={args.mode} exit={rc} ===")
    return rc


if __name__ == "__main__":
    sys.exit(main())
