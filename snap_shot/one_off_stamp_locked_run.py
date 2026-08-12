"""One-off surgical fix for the 2026-08-12 10:42 AM ET locked-workbook run
(GitHub Actions run 31608260834).

That run posted 10 REM comments to ClickUp but the 30-second-later
SharePoint upload was rejected with HTTP 423 (workbook open in Excel),
so col Q was never stamped for those 10 units. Without a stamp, the
next successful comment-sync run's sha8 dedup would post duplicates.

This script:
  1. Downloads the current SharePoint workbook (respects race-guard).
  2. Extracts REM comments (col P) + Last Synced (col Q).
  3. Filters to units where the col-P sha8 differs from any sha8 in Q
     AND the unit has a matching ClickUp task (i.e. same predicate as
     the poll run's comment-sync — those are the 10 comments we posted
     but never stamped).
  4. Stamps ONLY those Q cells with the current time + sha8. Does NOT
     re-post to ClickUp — the 10 comments are already on ClickUp tasks.
  5. Uploads the patched workbook, honoring the new HTTP-423 rollback
     path (if it locks again, no ClickUp changes happen, so nothing to
     roll back — this script never posts).

Dry-run mode (default) shows what it WOULD stamp without touching
SharePoint. Pass --apply to actually stamp + upload.

Safe to re-run: idempotent by construction. If the stamps already
landed (e.g. via a successful comment-sync run that beat us to it),
this script sees sha8(P)==sha8(Q) for those units and skips them.
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
    LOCAL_BUILD_DIR,
    LOG,
    MAPPING_PATH,
    WORKBOOK_WEB_URL,
    WorkbookLockedError,
    _comment_hash,
    _last_synced_hash,
    _norm_unit_label,
    download_from_sharepoint,
    extract_ann_edits,
    is_race_condition,
    load_json,
    map_last_synced_addresses,
    now_et,
    pull_appfolio_rent_roll_by_property,
    pull_lar_summaries,
    unit_display,
    upload_to_sharepoint,
)


def find_pending_units_with_target(
    rem_comments_by_prop_unit,
    last_synced_by_prop_unit,
    task_ids_by_tenant_id,
    task_ids_by_prop_unit,
    rent_roll_by_property,
):
    """Mirror the exact filter sync_rem_comments_to_clickup uses.
    Returns the (pid, unit_label) keys for units that (a) have a col-P
    comment whose sha8 differs from col-Q AND (b) have at least one
    matching ClickUp task — i.e. the ones that would have been posted.

    Filter identical to sync_rem_comments_to_clickup, so we can't
    accidentally stamp a unit whose comment WASN'T posted to ClickUp.
    """
    # Build (pid, unit_norm) -> tenant_id lookup, same as the sync function.
    tid_by_prop_unit = {}
    for pid, units in rent_roll_by_property.items():
        for u in units:
            unit_norm = _norm_unit_label(unit_display(u))
            tid = str(u.get("TenantId") or "").strip()
            if unit_norm and tid:
                tid_by_prop_unit[(str(pid), unit_norm)] = tid

    pending_with_target = []
    for (pid, unit_label), comment_text in rem_comments_by_prop_unit.items():
        comment_text = (comment_text or "").strip()
        if not comment_text:
            continue
        new_hash = _comment_hash(comment_text)
        prior_hash = _last_synced_hash(last_synced_by_prop_unit.get((pid, unit_label), ""))
        if new_hash == prior_hash and prior_hash:
            continue  # already synced, skip

        # Look for a matching ClickUp task, same predicate as sync fn.
        unit_norm = _norm_unit_label(unit_label)
        tid = tid_by_prop_unit.get((pid, unit_norm))
        target_ids = []
        if tid:
            for t in task_ids_by_tenant_id.get(tid, []):
                if t not in target_ids:
                    target_ids.append(t)
        for t in task_ids_by_prop_unit.get((pid, unit_norm), []):
            if t not in target_ids:
                target_ids.append(t)

        if target_ids:
            pending_with_target.append({
                "pid": str(pid),
                "unit_label": unit_label,
                "tenant_id": tid or "",
                "target_task_count": len(target_ids),
                "comment": comment_text,
                "new_hash": new_hash,
                "prior_stamp": last_synced_by_prop_unit.get((pid, unit_label), ""),
            })

    return pending_with_target


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--apply", action="store_true",
        help="Actually patch + upload the workbook. Default is dry-run (report only).",
    )
    args = parser.parse_args()

    LOG.info(
        "=== one_off_stamp_locked_run: %s ===",
        "APPLY (will upload)" if args.apply else "DRY-RUN (no upload)",
    )

    _mapping = load_json(MAPPING_PATH)

    # Step 1: download workbook + race-guard.
    wb_bytes, last_modified_by, last_modified_at = download_from_sharepoint()
    if is_race_condition(last_modified_by, last_modified_at):
        LOG.warning(
            "SKIP: workbook was modified within the race-guard window. "
            "Wait a few minutes and try again."
        )
        return 0
    LOG.info(f"Downloaded workbook. Last edited by {last_modified_by!r} at {last_modified_at}.")

    # Step 2: extract.
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

    # Step 3: pull the same reference data comment-sync uses.
    rent_roll_by_property = pull_appfolio_rent_roll_by_property()
    (
        _summaries_by_tenant,
        _summaries_by_prop_unit,
        task_ids_by_tenant_id,
        task_ids_by_prop_unit,
    ) = pull_lar_summaries()

    # Step 4: find the units that would have been posted.
    to_stamp = find_pending_units_with_target(
        rem_comments_by_prop_unit,
        last_synced_by_prop_unit,
        task_ids_by_tenant_id,
        task_ids_by_prop_unit,
        rent_roll_by_property,
    )

    LOG.info(f"Found {len(to_stamp)} unit(s) needing surgical stamp.")
    for i, u in enumerate(to_stamp, 1):
        LOG.info(
            f"  [{i}] PropertyId={u['pid']} Unit={u['unit_label']!r} "
            f"tenant={u['tenant_id'] or '-'} tasks={u['target_task_count']} "
            f"new_sha8={u['new_hash']} prior_stamp={u['prior_stamp']!r}"
        )
        # Show first 200 chars of the comment for a sanity check.
        preview = (u["comment"] or "")[:200].replace("\n", " ")
        LOG.info(f"      comment: {preview}")

    if not to_stamp:
        LOG.info("Nothing to stamp — either the fix already landed, or none of the "
                 "10 comments from the locked run are still pending. Exiting cleanly.")
        return 0

    if not args.apply:
        LOG.info("Dry-run: not touching the workbook. Re-run with --apply to stamp + upload.")
        return 0

    # Step 5: map (pid, unit_label) -> Q cell address, patch, save, upload.
    q_addresses = map_last_synced_addresses(wb_bytes)
    LOG.info(f"Last-Synced address map: {len(q_addresses)} unit cell(s) located.")

    wb = openpyxl.load_workbook(io.BytesIO(wb_bytes))
    ws = wb["Snap Shot"]

    now_str = now_et().strftime("%Y-%m-%d %H:%M ET")
    patched = 0
    missing = 0
    for u in to_stamp:
        key = (u["pid"], u["unit_label"])
        addr = q_addresses.get(key)
        if not addr:
            missing += 1
            LOG.warning(f"  ! no Q-cell address for {key!r} — will retry on next comment-sync run.")
            continue
        stamp = f"{now_str} · sha8:{u['new_hash']}"
        ws[addr].value = stamp
        LOG.info(f"  ✓ stamped {addr} = {stamp!r}")
        patched += 1

    if patched == 0:
        LOG.info("Nothing actually stamped. Skipping upload.")
        return 0

    output_path = os.path.join(LOCAL_BUILD_DIR, "Leasing-Snap-Shot.xlsx")
    os.makedirs(LOCAL_BUILD_DIR, exist_ok=True)
    wb.save(output_path)
    LOG.info(f"Saved patched workbook to {output_path}.")

    try:
        upload_to_sharepoint(output_path)
    except WorkbookLockedError as e:
        LOG.warning(
            f"Upload rejected as HTTP 423 (workbook currently open in Excel): {e}\n"
            f"No ClickUp changes happened; nothing to roll back. Try again in a few "
            f"minutes when the file is closed."
        )
        return 2

    LOG.info(f"Uploaded patched workbook — {patched} stamp(s) recorded; {missing} skipped.")
    LOG.info(f"Workbook: {WORKBOOK_WEB_URL}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
