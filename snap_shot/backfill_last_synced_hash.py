#!/usr/bin/env python3
"""One-time backfill: copy sha8 stamps from Snap Shot col Q to the ClickUp
'Last Synced Hash' custom field on every matching LAR task.

Why
---
The Snap Shot REM-comment sync is being migrated from workbook col Q
(which requires WAC access to write and is currently 403'd for us) to a
ClickUp custom field. Before we flip the reader logic, we need every LAR
task that already has a col Q stamp to carry the same sha8 in its new
ClickUp field — otherwise the first run of the refactored code will see
"no prior hash" for every unit and post duplicates.

What it does
------------
1. Downloads the current live Snap Shot workbook from SharePoint.
2. Extracts col P (REM comment) + col Q (Last Synced stamp) for every unit
   via extract_ann_edits() — same reader the refactor will replace.
3. Pulls AppFolio rent-roll for TenantId lookup by (pid, unit_label).
4. Pulls every LAR task across Renewal / Vacancy / Docs Workflow, indexes
   task_ids by tenant_id and by (pid, unit_label). Same index the
   nightly rebuild uses.
5. For each workbook row with a Last Synced sha8, matches it to task ids
   via TenantId (occupied units) + (PropertyId, Unit#) (vacancy tasks),
   and writes the sha8 into the ClickUp Last Synced Hash field.

Idempotent
----------
Reads the current ClickUp field value first — skips writes that already
match. Safe to re-run.

Modes
-----
--dry-run (default)  Print the planned writes; do NOT touch ClickUp.
--apply              Perform the writes.

Env
---
CLICKUP_API_TOKEN, APPFOLIO_CLIENT_ID/SECRET, MS_CLIENT_ID/SECRET/TENANT_ID
(same as rebuild.py). All resolved through the rebuild module.

Field
-----
Last Synced Hash on all 3 LAR lists — same field id
  eddb4ff7-5a48-4919-a5bc-1663045c663b
"""
from __future__ import annotations

import argparse
import logging
import os
import sys

# Import all needed helpers from rebuild.py so we stay in lockstep with
# the exact matching logic the refactored reader will use.
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import rebuild as sb  # noqa: E402

LOG = logging.getLogger("backfill")

CU_LAST_SYNCED_HASH_FIELD = os.environ.get(
    "CLICKUP_LAST_SYNCED_HASH_FIELD", "eddb4ff7-5a48-4919-a5bc-1663045c663b"
)


def _cu_field_value(task, field_id):
    return sb._cu_field_value(task, field_id)


def collect_current_ls_hashes():
    """Return { task_id: current_last_synced_hash_field_value } across all
    3 LAR lists. Uses the same cu_get_list_tasks pager rebuild.py uses so
    field payloads match exactly. Absent / blank fields are stored as ''
    so downstream code can compare equality cleanly."""
    out = {}
    for label, list_id in [
        ("Renewal Pipeline", sb.LAR_RENEWAL_LIST_ID),
        ("Vacancy Pipeline", sb.LAR_VACANCY_LIST_ID),
        ("Documents Workflow", sb.LAR_DOCS_WORKFLOW_LIST_ID),
    ]:
        try:
            tasks = sb.cu_get_list_tasks(list_id, include_closed="true")
        except sb.FatalError as e:
            LOG.warning(f"Skipping {label} ({list_id}): {e}")
            continue
        for t in tasks:
            tid = t.get("id")
            if not tid:
                continue
            v = _cu_field_value(t, CU_LAST_SYNCED_HASH_FIELD)
            out[tid] = "" if v is None else str(v).strip()
        LOG.info(f"{label}: {len(tasks)} tasks scanned.")
    return out


def build_target_task_index(task_ids_by_tenant_id, task_ids_by_prop_unit,
                             rent_roll_by_property):
    """Return two mappings that mirror the ones sync_rem_comments_to_clickup
    already builds, so the matching key is identical to what the
    refactored reader will use.

    Returns (tid_by_prop_unit, targets_for) where
      tid_by_prop_unit: (pid_str, unit_norm) -> tenant_id_str
      targets_for(pid, unit_label) -> list of task_ids across all 3 lists
    """
    tid_by_prop_unit = {}
    for pid, units in rent_roll_by_property.items():
        for u in units:
            unit_norm = sb._norm_unit_label(sb.unit_display(u))
            tid = str(u.get("TenantId") or "").strip()
            if unit_norm and tid:
                tid_by_prop_unit[(str(pid), unit_norm)] = tid

    def targets_for(pid, unit_label):
        pid_str = str(pid)
        unit_norm = sb._norm_unit_label(unit_label)
        target_ids = []
        tid = tid_by_prop_unit.get((pid_str, unit_norm))
        if tid:
            for t in task_ids_by_tenant_id.get(tid, []):
                if t not in target_ids:
                    target_ids.append(t)
        for t in task_ids_by_prop_unit.get((pid_str, unit_norm), []):
            if t not in target_ids:
                target_ids.append(t)
        return target_ids, tid or ""

    return tid_by_prop_unit, targets_for


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true",
                     help="Perform writes. Default is dry-run.")
    ap.add_argument("--verbose", action="store_true")
    args = ap.parse_args()

    level = logging.DEBUG if args.verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    )
    # Quiet rebuild's own logger to INFO — its DEBUG is very chatty.
    logging.getLogger("rebuild").setLevel(logging.INFO)

    dry_run = not args.apply
    LOG.info(f"Backfill mode: {'DRY-RUN' if dry_run else 'APPLY'}")

    # Step 1: SharePoint pull
    LOG.info("Step 1: downloading current Snap Shot workbook from SharePoint...")
    wb_bytes, _mod_by, _mod_at = sb.download_from_sharepoint()
    if not wb_bytes:
        LOG.error("No SharePoint workbook returned — cannot backfill.")
        sys.exit(2)
    LOG.info(f"Downloaded {len(wb_bytes)} bytes.")

    # Step 2: extract col P + col Q
    LOG.info("Step 2: extracting REM comments + Last Synced stamps from workbook...")
    (
        _prop_overrides,
        _unit_notes,
        _market_rent,
        rem_comments_by_prop_unit,
        last_synced_by_prop_unit,
    ) = sb.extract_ann_edits(wb_bytes)
    LOG.info(
        f"Extracted {len(rem_comments_by_prop_unit)} REM comments and "
        f"{len(last_synced_by_prop_unit)} Last Synced stamps."
    )

    # Convert col Q strings to sha8 hashes; drop anything without a hash.
    hash_by_prop_unit = {}
    for (pid, unit_label), stamp in last_synced_by_prop_unit.items():
        h = sb._last_synced_hash(stamp)
        if h:
            hash_by_prop_unit[(str(pid), unit_label)] = h
    LOG.info(f"Parsed {len(hash_by_prop_unit)} sha8 stamps out of {len(last_synced_by_prop_unit)} Last Synced cells.")

    if not hash_by_prop_unit:
        LOG.warning("No sha8 stamps found in col Q — nothing to backfill.")
        return

    # Step 3: rent roll for TenantId lookup
    LOG.info("Step 3: pulling AppFolio rent-roll for TenantId lookup...")
    rent_roll_by_property = sb.pull_appfolio_rent_roll_by_property()

    # Step 4: LAR task index
    LOG.info("Step 4: pulling LAR tasks (Renewal + Vacancy + Docs Workflow)...")
    (
        _summ_tid, _summ_pu,
        task_ids_by_tenant_id,
        task_ids_by_prop_unit,
    ) = sb.pull_lar_summaries()

    _tidmap, targets_for = build_target_task_index(
        task_ids_by_tenant_id, task_ids_by_prop_unit, rent_roll_by_property
    )

    # Step 5: current field values
    LOG.info("Step 5: reading current Last Synced Hash values from ClickUp...")
    current_by_task = collect_current_ls_hashes()

    # Step 6: plan the writes
    plan = []          # list of (task_id, pid, unit_label, hash_to_write, current, tid)
    unmatched = []     # workbook rows with no target task
    already_ok = 0
    for (pid, unit_label), sha8 in hash_by_prop_unit.items():
        target_ids, tid = targets_for(pid, unit_label)
        if not target_ids:
            unmatched.append((pid, unit_label, sha8, tid))
            continue
        for task_id in target_ids:
            current = current_by_task.get(task_id, "")
            if current == sha8:
                already_ok += 1
                continue
            plan.append((task_id, pid, unit_label, sha8, current, tid))

    LOG.info("=" * 70)
    LOG.info("BACKFILL PLAN SUMMARY")
    LOG.info("=" * 70)
    LOG.info(f"  Rows with sha8 stamps in col Q:      {len(hash_by_prop_unit)}")
    LOG.info(f"  Task-writes already matching:        {already_ok}")
    LOG.info(f"  Task-writes to perform:              {len(plan)}")
    LOG.info(f"  Rows unmatched (no LAR task found):  {len(unmatched)}")

    if unmatched and args.verbose:
        LOG.info("Unmatched rows (first 30):")
        for pid, unit_label, sha8, tid in unmatched[:30]:
            LOG.info(f"  pid={pid} unit={unit_label!r} sha8={sha8} tenant_id={tid or '—'}")

    if not plan:
        LOG.info("Nothing to write — backfill is a no-op. Safe to enable refactor.")
        return

    if args.verbose:
        LOG.info("Planned writes (first 30):")
        for task_id, pid, unit_label, sha8, current, tid in plan[:30]:
            LOG.info(
                f"  {task_id}  pid={pid} unit={unit_label!r} tenant_id={tid or '—'} "
                f"current={current!r} -> {sha8}"
            )

    if dry_run:
        LOG.info("DRY-RUN: no writes performed. Re-run with --apply to execute.")
        return

    # Step 7: apply
    LOG.info("Applying writes...")
    ok = 0
    fail = 0
    for task_id, pid, unit_label, sha8, _current, _tid in plan:
        try:
            sb.cu_set_field(task_id, CU_LAST_SYNCED_HASH_FIELD, sha8)
            ok += 1
            if ok % 25 == 0:
                LOG.info(f"  progress: {ok}/{len(plan)} written")
        except Exception as e:  # noqa: BLE001
            fail += 1
            LOG.warning(f"  FAIL {task_id} pid={pid} unit={unit_label!r}: {e}")
    LOG.info(f"Done. ok={ok} fail={fail} of {len(plan)}")


if __name__ == "__main__":
    main()
