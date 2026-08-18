"""One-off catch-up for the 5 REM comments orphaned by the 2026-08-12 poll run.

Context (2026-08-18)
- The 2026-08-12 comment-sync run left 6 REM col-P comments unposted because
  no matching ClickUp task existed at post-time (either the task hadn't been
  created yet, or the (pid, unit) lookup missed a hand-typed unit label with
  space-padded hyphens like "181 - B1").
- Alexis's assistant has since located 5 of the 6 tasks. The 6th (Clickety
  Clack Vape and Gifts, TenantId 1348) is still unmatched and NOT touched
  by this script.
- The permanent normalizer fix (this same commit) prevents future
  hyphen-padding orphans, but this one-off handles the 5 already-orphaned
  comments so they finally land on their tasks.

Two-part fix per (pid, unit_label):
  1. If Property ID or Tenant ID is blank on the ClickUp task, set them.
     Skips the write if the field already has a value (see WIRE_UP below);
     force-overwrite is only done for Connectivity Source per user
     confirmation 2026-08-18.
  2. Read the current col-P comment for that (pid, unit_label) from
     SharePoint, post it to the target task with the "📋 LSS note from
     {REM} ({date}):" preamble, then stamp col Q with
     "YYYY-MM-DD HH:MM ET · sha8:<hash>" — same format the regular
     comment-sync uses, so the next run's dedup skips them.

Ships with the same lock-safe pattern the regular comment-sync uses: if the
SharePoint upload comes back HTTP 423 (workbook open in Excel), we DELETE
the ClickUp comments we just posted so nothing gets stamped and the next
successful run can re-attempt idempotently.

Dry-run by default. Pass --apply to actually mutate ClickUp + SharePoint.
"""
from __future__ import annotations

import argparse
import io
import os
import sys

import openpyxl

# Add snap_shot to path so we can import from rebuild.py
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from rebuild import (  # noqa: E402
    CU_PROPERTY_ID_FIELD,
    CU_TENANT_ID_FIELD,
    LOCAL_BUILD_DIR,
    LOG,
    MAPPING_PATH,
    WORKBOOK_WEB_URL,
    WorkbookLockedError,
    _comment_hash,
    _rem_name_for_property,
    cu_delete_comment,
    cu_post_comment,
    cu_set_field,
    download_from_sharepoint,
    extract_ann_edits,
    is_race_condition,
    load_json,
    map_last_synced_addresses,
    now_et,
    upload_to_sharepoint,
)


# The 5 orphans to wire up. Each row is:
#   pid           — AppFolio Property ID (matches col P key)
#   unit_label    — Unit label as it appears in col P key (workbook-side)
#   task_id       — ClickUp task id (verified 2026-08-18)
#   set_pid       — Property ID value to write, or None to leave alone
#   set_tid       — Tenant ID value to write, or None to leave alone
#   force_tid     — if True, overwrite Tenant ID even if already set
#
# 6th orphan (Clickety Clack Vape and Gifts, TenantId 1348 / PropertyId 57
# / Unit 1396) is intentionally omitted — Alexis's assistant could not
# locate a matching ClickUp task, so we leave col Q blank and let the next
# scheduled comment-sync retry once someone creates that task.
WIRE_UP = [
    {
        # Cumberland Flex | Warehouse Cages — task already stamped
        # PID=230 / TID=1405, so both writes are no-ops. Included so we
        # post the comment + stamp col Q.
        "pid": "230",
        "unit_label": "Warehouse Cages",
        "task_id": "868kt2tzu",
        "set_pid": None,
        "set_tid": None,
        "force_tid": False,
    },
    {
        # Lakeview Village - 181 B-1 — WingStop (Vacancy Pipeline).
        # PID is already 54, TID is (correctly) empty since the unit is
        # vacant. Task matches by (pid, unit) with the normalizer fix
        # from the same commit; nothing to write, just post + stamp.
        "pid": "54",
        "unit_label": "181-B1",
        "task_id": "868gyguen",
        "set_pid": None,
        "set_tid": None,
        "force_tid": False,
    },
    {
        # Savannah Crossing - Connectivity Source, LLC - Unit B2.
        # Task's Tenant ID is currently 1479, but Alexis confirmed
        # 2026-08-18 to overwrite to 1271 (per her assistant's mapping).
        # PID already 211.
        "pid": "211",
        "unit_label": "B2",
        "task_id": "868kcmfen",
        "set_pid": None,
        "set_tid": 1271,
        "force_tid": True,
    },
    {
        # Somerset Shoppes - Donald Bell - Unit 7794. Both fields blank.
        "pid": "282",
        "unit_label": "Unit 7794",
        "task_id": "868kdk09t",
        "set_pid": 282,
        "set_tid": 2032,
        "force_tid": False,
    },
    {
        # Somerset Shoppes - Beau Reinmiller Enterprises - Unit 7746-7754.
        # Both fields blank.
        "pid": "282",
        "unit_label": "Unit 7746-7754",
        "task_id": "868kdjzey",
        "set_pid": 282,
        "set_tid": 2025,
        "force_tid": False,
    },
]


def _cu_get_field_value(task_json, field_id):
    """Return the current value of a custom field on a ClickUp task JSON dict."""
    for f in (task_json or {}).get("custom_fields") or []:
        if f.get("id") == field_id:
            v = f.get("value")
            if v is None:
                return ""
            return str(v).strip()
    return ""


def _cu_get_task(task_id):
    """Small wrapper mirroring cu_get_task_by_id but without needing subtasks."""
    import requests
    from rebuild import CU_BASE, cu_headers
    r = requests.get(f"{CU_BASE}/task/{task_id}", headers=cu_headers(), timeout=30)
    if r.status_code != 200:
        raise RuntimeError(f"cu_get_task({task_id}) failed HTTP {r.status_code}: {r.text[:200]}")
    return r.json()


def _find_col_p_addresses(wb_bytes):
    """Return {(pid, unit_label): "P<row>"} — the P-column cell address
    of each unit's REM Comment cell. Mirrors map_last_synced_addresses
    which handles the Q column.

    Kept local (not exported to rebuild.py) since this is a one-off:
    the regular flow already reads P via extract_ann_edits and doesn't
    need per-cell addresses.
    """
    from rebuild import FatalError, PROPERTY_ID_RE
    p_addresses = {}
    try:
        wb = openpyxl.load_workbook(io.BytesIO(wb_bytes), data_only=True)
    except Exception as e:
        raise FatalError(f"Could not open workbook to map P cells: {e}")

    if "Snap Shot" not in wb.sheetnames:
        raise FatalError(f"Workbook has no 'Snap Shot' tab (found: {wb.sheetnames}).")
    ws = wb["Snap Shot"]

    def cell_text(row, col_letter):
        v = ws[f"{col_letter}{row}"].value
        return "" if v is None else str(v).strip()

    max_row = ws.max_row
    row = 1
    while row <= max_row:
        name_cell = cell_text(row, "B")
        next_row_addr = cell_text(row + 1, "B")
        pid_match = PROPERTY_ID_RE.search(next_row_addr)
        if not (name_cell and pid_match):
            row += 1
            continue

        property_id = pid_match.group(1)
        block_start = row
        try:
            header_row = block_start + 4
            header_text = cell_text(header_row, "B")
            if not header_text.upper().startswith("DEAL ACTIVITY"):
                header_row = block_start + 3
                header_text = cell_text(header_row, "B")
            notes_start = header_row + 1 if header_text.upper().startswith("DEAL ACTIVITY") else header_row

            rent_roll_header_row = None
            for probe in range(notes_start, min(notes_start + 25, max_row + 1)):
                if cell_text(probe, "B").upper() == "RENT ROLL":
                    rent_roll_header_row = probe
                    break
            if not rent_roll_header_row:
                row = block_start + 1
                continue

            data_row = rent_roll_header_row + 2
            while data_row <= max_row:
                unit_val = cell_text(data_row, "B")
                if not unit_val or unit_val.upper() == "TOTALS":
                    break
                p_addresses[(property_id, unit_val)] = f"P{data_row}"
                data_row += 1
            row = data_row
        except Exception as e:  # noqa: BLE001
            LOG.warning(f"P-cell address map: failed at row {block_start} (pid={property_id}): {e}")
            row = block_start + 1

    return p_addresses


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--apply", action="store_true",
        help="Actually mutate ClickUp + SharePoint. Default is dry-run.",
    )
    args = parser.parse_args()

    LOG.info(
        "=== one_off_wire_orphans_2026_08_18: %s ===",
        "APPLY (will mutate)" if args.apply else "DRY-RUN (no writes)",
    )

    mapping = load_json(MAPPING_PATH)
    rem_name_by_pid = {
        str(m.get("appfolio_id") or "").strip(): _rem_name_for_property(m)
        for m in mapping
    }

    # Step 1: download workbook + race-guard.
    wb_bytes, last_modified_by, last_modified_at = download_from_sharepoint()
    if is_race_condition(last_modified_by, last_modified_at):
        LOG.warning(
            "SKIP: workbook was modified within the race-guard window. "
            "Wait a few minutes and try again."
        )
        return 0
    LOG.info(f"Downloaded workbook. Last edited by {last_modified_by!r} at {last_modified_at}.")

    # Step 2: extract col P + col Q.
    (
        _prop_overrides,
        _unit_notes,
        _market_rent_overrides,
        rem_comments_by_prop_unit,
        last_synced_by_prop_unit,
    ) = extract_ann_edits(wb_bytes)

    LOG.info(
        f"Extracted {len(rem_comments_by_prop_unit)} REM comment(s), "
        f"{len(last_synced_by_prop_unit)} Last Synced stamp(s)."
    )

    # Step 3: locate col P and col Q cell addresses for our 5 rows.
    p_addresses = _find_col_p_addresses(wb_bytes)
    q_addresses = map_last_synced_addresses(wb_bytes)
    LOG.info(
        f"P-address map: {len(p_addresses)} cell(s); "
        f"Q-address map: {len(q_addresses)} cell(s)."
    )

    # Step 4: for each wire-up row, resolve col-P text + verify targets.
    plan = []
    for w in WIRE_UP:
        pid, unit_label, task_id = w["pid"], w["unit_label"], w["task_id"]
        key = (pid, unit_label)
        comment_text = (rem_comments_by_prop_unit.get(key) or "").strip()
        p_addr = p_addresses.get(key, "")
        q_addr = q_addresses.get(key, "")
        rem_name = rem_name_by_pid.get(pid, "REM")
        row = {
            "pid": pid,
            "unit_label": unit_label,
            "task_id": task_id,
            "rem_name": rem_name,
            "comment": comment_text,
            "p_addr": p_addr,
            "q_addr": q_addr,
            "set_pid": w["set_pid"],
            "set_tid": w["set_tid"],
            "force_tid": w["force_tid"],
        }
        plan.append(row)
        LOG.info(
            f"  • pid={pid} unit={unit_label!r} task={task_id} "
            f"P={p_addr or '?'} Q={q_addr or '?'} "
            f"set_pid={w['set_pid']} set_tid={w['set_tid']} force_tid={w['force_tid']}"
        )
        if not comment_text:
            LOG.warning(
                f"    ! col P is BLANK for {key!r} — nothing to post. "
                f"Will still write set_pid/set_tid if requested."
            )
        else:
            preview = comment_text[:200].replace("\n", " ")
            LOG.info(f"    comment ({len(comment_text)} chars): {preview}")

    # Sanity: all 5 must have a valid col-Q address. Missing Q = layout drift.
    missing_q = [f"pid={r['pid']} unit={r['unit_label']!r}" for r in plan if not r["q_addr"]]
    if missing_q:
        LOG.error(
            "ABORT: col-Q address missing for: %s. Layout drift — bail before mutating.",
            missing_q,
        )
        return 1

    if not args.apply:
        LOG.info("Dry-run complete. Re-run with --apply to mutate ClickUp + SharePoint.")
        return 0

    # Step 5 (apply): write custom fields + post comments. Track comment
    # IDs so we can DELETE them if the SharePoint upload fails.
    posted_comment_ids = []
    for row in plan:
        task_id = row["task_id"]

        # 5a. Write Property ID if requested.
        if row["set_pid"] is not None:
            try:
                task_json = _cu_get_task(task_id)
                current_pid = _cu_get_field_value(task_json, CU_PROPERTY_ID_FIELD)
                if current_pid and str(current_pid) == str(row["set_pid"]):
                    LOG.info(f"  = task {task_id} Property ID already {current_pid} — skip.")
                else:
                    cu_set_field(task_id, CU_PROPERTY_ID_FIELD, row["set_pid"])
                    LOG.info(f"  ✓ task {task_id} Property ID {current_pid or '(empty)'} → {row['set_pid']}")
            except Exception as e:  # noqa: BLE001
                LOG.warning(f"  ! Property ID write failed on {task_id}: {e}")

        # 5b. Write Tenant ID if requested.
        if row["set_tid"] is not None:
            try:
                task_json = _cu_get_task(task_id)
                current_tid = _cu_get_field_value(task_json, CU_TENANT_ID_FIELD)
                if current_tid and not row["force_tid"] and str(current_tid) == str(row["set_tid"]):
                    LOG.info(f"  = task {task_id} Tenant ID already {current_tid} — skip.")
                elif current_tid and not row["force_tid"]:
                    LOG.warning(
                        f"  ! task {task_id} Tenant ID already set to {current_tid} "
                        f"(force_tid=False) — leaving as-is."
                    )
                else:
                    cu_set_field(task_id, CU_TENANT_ID_FIELD, row["set_tid"])
                    LOG.info(f"  ✓ task {task_id} Tenant ID {current_tid or '(empty)'} → {row['set_tid']}")
            except Exception as e:  # noqa: BLE001
                LOG.warning(f"  ! Tenant ID write failed on {task_id}: {e}")

        # 5c. Post the REM comment (if col P has text).
        if not row["comment"]:
            LOG.info(f"  · task {task_id}: col P blank, nothing to post.")
            continue

        today = now_et().strftime("%Y-%m-%d")
        body = f"📋 LSS note from {row['rem_name']} ({today}):\n\n{row['comment']}"
        try:
            comment_id = cu_post_comment(task_id, body)
            if comment_id:
                posted_comment_ids.append(comment_id)
                LOG.info(f"  ✓ posted comment {comment_id} to task {task_id}")
            else:
                LOG.warning(f"  ! post to {task_id} returned no comment_id (see prior log).")
        except Exception as e:  # noqa: BLE001
            LOG.warning(f"  ! post to {task_id} raised: {e}")

    # Step 6: patch col Q for each row we posted (or would post — comment_text
    # blank rows don't need a stamp).
    wb = openpyxl.load_workbook(io.BytesIO(wb_bytes))
    ws = wb["Snap Shot"]
    now_str = now_et().strftime("%Y-%m-%d %H:%M ET")

    stamped = 0
    for row in plan:
        if not row["comment"]:
            continue  # nothing was posted → nothing to stamp
        stamp = f"{now_str} · sha8:{_comment_hash(row['comment'])}"
        addr = row["q_addr"]
        ws[addr].value = stamp
        LOG.info(f"  ✓ stamped {addr} = {stamp!r}")
        stamped += 1

    if stamped == 0:
        LOG.info("No col-Q stamps needed. Skipping upload.")
        return 0

    output_path = os.path.join(LOCAL_BUILD_DIR, "Leasing-Snap-Shot.xlsx")
    os.makedirs(LOCAL_BUILD_DIR, exist_ok=True)
    wb.save(output_path)
    LOG.info(f"Saved patched workbook to {output_path}.")

    # Step 7: upload with the lock-safe rollback path.
    try:
        upload_to_sharepoint(output_path)
    except WorkbookLockedError as e:
        LOG.warning(
            f"Upload rejected as HTTP 423 (workbook currently open in Excel): {e}"
        )
        LOG.warning(
            f"Rolling back {len(posted_comment_ids)} ClickUp comment(s) so the next "
            f"successful run can post cleanly."
        )
        for cid in posted_comment_ids:
            cu_delete_comment(cid)
        return 2

    LOG.info(f"Uploaded patched workbook — {stamped} stamp(s) recorded.")
    LOG.info(f"Workbook: {WORKBOOK_WEB_URL}")
    LOG.info(f"Posted {len(posted_comment_ids)} ClickUp comment(s).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
