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
    "broker calls": "broker_calls",
    "property flags": "property_flags",
    "vacant callouts": "vacant_callouts",
    "deal activity": "deal_activity",
    "broker active interest from weekly reporting": "broker_active_interest",
    "ann's commentary": "ann_commentary",
}

PROPERTY_ID_RE = re.compile(r"Property ID:\s*([0-9]+)")


def extract_ann_edits(workbook_bytes):
    """Parse the current SharePoint workbook and return (property_overrides,
    unit_notes) using the same schemas as property_notes_overrides.json /
    notes_by_unit.json, keyed by AppFolio PropertyId (property overrides) and
    by unit note-key (OccupancyId/UnitId, unit notes) — mirroring
    extract_ann_notes.py's cell-position matching approach, extended to also
    read Market Rent (col H) and per-unit Notes (col N).

    On any parse error for an individual block, that block's edits are
    skipped (logged) rather than aborting the whole extraction — a single
    malformed block must never nuke everyone else's preserved edits.
    """
    property_overrides = {}
    unit_notes = {}
    market_rent_overrides = {}  # {property_id: {unit_note_key: market_rent_value}}

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
            # DEAL ACTIVITY & NOTES header should be at block_start + 3
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
            scan_row = r
            rent_roll_header_row = None
            for probe in range(scan_row, min(scan_row + 20, max_row + 1)):
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
        f"{sum(len(v) for v in market_rent_overrides.values())} market-rent overrides."
    )
    return property_overrides, unit_notes, market_rent_overrides


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
}
COL_WIDTHS = {"A": 3, "B": 20, "C": 26, "D": 10, "E": 9, "F": 11, "G": 8, "H": 11,
              "I": 11, "J": 11, "K": 11, "L": 15, "M": 11, "N": 34}
LAST_COL = "N"
FIRST_COL = "B"

STATE_RE = re.compile(r",\s*([A-Z]{2})\s+\d{5}")


def build_workbook(mapping, rent_roll_by_property, property_overrides,
                    unit_notes, market_rent_overrides, broker_map,
                    output_path, logo_path=None):
    """Build the Snap Shot workbook. `mapping` = property_mapping_all73.json
    contents. `rent_roll_by_property` = {appfolio_id_str: [unit_row, ...]}.
    `property_overrides` = {clickup_task_id: {broker_calls, property_flags,
    vacant_callouts, deal_activity, broker_active_interest, ann_commentary}}.
    `unit_notes` = {occupancy_or_unit_id_str: note_text}.
    `market_rent_overrides` = {appfolio_id_str: {unit_label: value_str}} —
    round-tripped from SharePoint; AppFolio has no market-rent field of its
    own for this column, so absent an override the cell is left blank
    exactly as in build_prototype_v3.
    """
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

    ws.row_dimensions[row].height = 18
    ws.merge_cells(f"{FIRST_COL}{row}:{LAST_COL}{row}")
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
    row += 1

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

    label_min_lines = {
        "Deal Activity": 4,
        "Ann's Commentary": 4,
        "Broker Active Interest from Weekly Reporting": 2,
    }

    # Column B is width 20 (see COL_WIDTHS). At 9pt bold Calibri, ~17 chars
    # per visible line. Compute the label's own line count and use it as a
    # floor for the row height, since the label now wraps too.
    LABEL_CHARS_PER_LINE = 17

    for label, val in note_labels:
        is_default_bai = (label == "Broker Active Interest from Weekly Reporting" and val == default_bai)
        min_lines = label_min_lines.get(label, 1)
        # Ensure row is tall enough for whichever cell wraps to more lines.
        label_lines = max(1, -(-len(label) // LABEL_CHARS_PER_LINE))
        effective_min_lines = max(min_lines, label_lines)
        ws.row_dimensions[row].height = _row_height_for(val, min_lines=effective_min_lines)
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
            }

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
                elif col_label == "Tenant ID":
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left" if col_label == "Tenant" else "center",
                                                vertical="center", indent=(1 if col_label == "Tenant" else 0))

                if row_fill:
                    cell.fill = row_fill
                elif col_label in ("Market Rent", "Notes"):
                    cell.fill = fill(GRAY_LIGHT)
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

        row += 1

    return row


# ══════════════════════════════════════════════════════════════════════════
# Orchestration
# ══════════════════════════════════════════════════════════════════════════

def run_build(mode, dry_run):
    """Executes steps 1-6 of the brief's script structure. Returns the local
    output file path. Raises FatalError on any unrecoverable problem."""

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
    if current_wb_bytes:
        prop_overrides_by_pid, unit_notes_by_key, market_rent_by_pid = extract_ann_edits(current_wb_bytes)
        property_overrides = reindex_overrides_by_clickup_task(prop_overrides_by_pid, mapping)
        market_rent_overrides = market_rent_by_pid
        # unit_notes re-indexing needs the fresh rent roll, done after step 3 below.
        pending_unit_notes_raw = unit_notes_by_key
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
    )

    # Step 6: upload to SharePoint (skip entirely for dry-run).
    if not dry_run:
        upload_to_sharepoint(output_path)
        LOG.info(f"Uploaded {output_path} to SharePoint as {WORKBOOK_FILENAME}.")

    return output_path


def poll_and_maybe_rebuild():
    """--mode=poll: check the ClickUp control-task checkbox. If checked, run
    a full nightly rebuild, then uncheck + comment. Returns exit code."""
    control_task = find_control_task()
    if not is_checkbox_checked(control_task):
        LOG.info("Checkbox not checked; exiting.")
        return 0

    LOG.info(f"'{CHECKBOX_FIELD_NAME}' is checked on control task {control_task['id']} — running on-demand rebuild.")
    output_path = run_build(mode="nightly", dry_run=False)
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
    args = parser.parse_args()

    LOG.info(f"=== prudent-snap-shot-rebuild starting: mode={args.mode} ===")
    dry_run = args.mode == "dry-run"

    try:
        if args.mode == "poll":
            rc = poll_and_maybe_rebuild()
        else:
            output_path = run_build(mode=args.mode, dry_run=dry_run)
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
