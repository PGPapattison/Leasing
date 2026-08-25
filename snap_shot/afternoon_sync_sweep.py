#!/usr/bin/env python3
"""Snap Shot — Afternoon sync-health sweep.

Runs at 5 PM ET on weekdays (schedule owned by the recurring Perplexity task
that spawns this session — not this repo's GitHub Actions). Scans the
current SharePoint copy of Leasing-Snap-Shot.xlsx and reports any REM
comment (col P) whose matching Last Synced stamp (col Q) is missing or
older than a configurable stale threshold (default: 4 hours).

Signal-to-noise choice ("Option 1" per Alexis 2026-08-11):
- Only units with text in col P are considered. Blank comments are
  ignored — a stamp is only expected when there is something to sync.
- A unit is FLAGGED if col P has text AND either:
    a) col Q is blank, OR
    b) col Q's timestamp parses successfully but is older than
       STALE_HOURS hours vs. now_et(), OR
    c) col Q is present but does not parse to a timestamp (corrupt).
- Comments with sha8 in Q but freshly synced within the window are NOT
  flagged even if they were typed earlier today — the stamp is what
  matters for the automation health signal.

Outputs:
- Sends an email to ERROR_NOTIFICATION_TO (apattison@prudentgrowth.com)
  IFF any flagged units are found. Silent otherwise.
- Also prints a summary to stdout for the recurring task's log.

Reuses helpers from rebuild.py — SharePoint download, extract_ann_edits,
now_et, send_error_email (via get_graph_access_token / http_request).

Usage:
    python3 snap_shot/afternoon_sync_sweep.py
    python3 snap_shot/afternoon_sync_sweep.py --dry-run  # no email
    python3 snap_shot/afternoon_sync_sweep.py --stale-hours=2  # tighter
"""
from __future__ import annotations

import argparse
import os
import re
import sys
from datetime import datetime, timedelta

# Reuse rebuild.py helpers rather than duplicating auth / download / extract.
_THIS_DIR = os.path.dirname(os.path.abspath(__file__))
if _THIS_DIR not in sys.path:
    sys.path.insert(0, _THIS_DIR)

from rebuild import (  # noqa: E402
    ET,
    ERROR_NOTIFICATION_TO,
    LOG,
    WORKBOOK_WEB_URL,
    download_from_sharepoint,
    extract_ann_edits,
    get_graph_access_token,
    http_request,
    now_et,
)


DEFAULT_STALE_HOURS = 4

# Matches the stamp format run_comment_sync writes:
#   "2026-08-11 15:22 ET · sha8:07e57d1c"
# The date + time portion is what we parse.
_STAMP_TS_RE = re.compile(r"^\s*(\d{4}-\d{2}-\d{2})\s+(\d{2}:\d{2})\s+ET\b")


def _parse_stamp_ts(stamp: str):
    """Parse the ET timestamp out of a Last Synced cell. Returns a
    timezone-aware datetime in ET, or None if the cell doesn't parse."""
    if not stamp:
        return None
    m = _STAMP_TS_RE.match(stamp)
    if not m:
        return None
    date_str, time_str = m.group(1), m.group(2)
    try:
        naive = datetime.strptime(f"{date_str} {time_str}", "%Y-%m-%d %H:%M")
    except ValueError:
        return None
    return naive.replace(tzinfo=ET)


def scan_workbook(workbook_bytes: bytes, stale_hours: int):
    """Return (flagged, healthy_count, stats) where flagged is a list of
    dicts describing each problem unit. Never raises on a per-unit parse
    error — we log and move on so one bad cell doesn't hide the rest.

    Each flagged entry has:
      pid, unit, comment (truncated), stamp_raw, reason, age_hours
    """
    # Post-2026-08-25 col P retirement: extract_ann_edits returns a 4-tuple
    # and `rem_comments_by_prop_unit` is always empty (col P is gone from the
    # workbook schema). The sweep still runs so the 5 PM workflow doesn't
    # need to be edited, but it will always report 0 flagged comments.
    (_prop, _unit_notes, _mkt,
     rem_comments_by_prop_unit) = extract_ann_edits(workbook_bytes)
    last_synced_by_prop_unit = {}

    now = now_et()
    stale_cutoff = now - timedelta(hours=stale_hours)

    flagged = []
    healthy = 0

    for key, comment_text in rem_comments_by_prop_unit.items():
        comment = (comment_text or "").strip()
        if not comment:
            # Blank / whitespace-only cell — not a real comment, skip.
            continue
        pid, unit = key
        stamp_raw = (last_synced_by_prop_unit.get(key) or "").strip()

        # Case A: no stamp at all.
        if not stamp_raw:
            flagged.append({
                "pid": pid, "unit": unit,
                "comment": comment,
                "stamp_raw": "",
                "reason": "no Last Synced stamp — comment never posted",
                "age_hours": None,
            })
            continue

        # Case B: stamp doesn't parse — corrupt or hand-edited.
        stamp_ts = _parse_stamp_ts(stamp_raw)
        if stamp_ts is None:
            flagged.append({
                "pid": pid, "unit": unit,
                "comment": comment,
                "stamp_raw": stamp_raw,
                "reason": "Last Synced cell does not parse to a timestamp",
                "age_hours": None,
            })
            continue

        # Case C: stamp is older than stale threshold.
        age = now - stamp_ts
        age_hours = age.total_seconds() / 3600.0
        if stamp_ts < stale_cutoff:
            flagged.append({
                "pid": pid, "unit": unit,
                "comment": comment,
                "stamp_raw": stamp_raw,
                "reason": f"Last Synced stamp is {age_hours:.1f}h old (>{stale_hours}h threshold)",
                "age_hours": age_hours,
            })
            continue

        # Otherwise: healthy.
        healthy += 1

    stats = {
        "total_comments": sum(
            1 for v in rem_comments_by_prop_unit.values() if (v or "").strip()
        ),
        "flagged": len(flagged),
        "healthy": healthy,
        "stale_hours": stale_hours,
        "now_et": now.strftime("%Y-%m-%d %H:%M ET"),
    }
    return flagged, healthy, stats


def _format_email_body(flagged, stats, workbook_link):
    lines = []
    lines.append(f"Snap Shot afternoon sync-health sweep — {stats['now_et']}")
    lines.append("")
    lines.append(
        f"{stats['flagged']} of {stats['total_comments']} col-P comment(s) "
        f"have no Last Synced stamp or a stamp older than {stats['stale_hours']}h."
    )
    lines.append("")
    lines.append(f"Workbook: {workbook_link}")
    lines.append("")
    lines.append(
        "How to troubleshoot:\n"
        "  1. Open the workbook — check col Q for each unit below.\n"
        "  2. If Q is BLANK: the comment never made it to ClickUp. Likely\n"
        "     causes: no matching ClickUp task exists for (PropertyId, Unit#)\n"
        "     or by Tenant ID; the race guard tripped on every fire today;\n"
        "     ClickUp auth was expired.\n"
        "  3. If Q has an OLD stamp but the comment text has been edited\n"
        "     since: the sha8 dedup check thinks it's the same text as before.\n"
        "     Re-save the workbook to force a re-hash.\n"
        "  4. If Q is CORRUPT (unparseable): someone hand-edited the cell.\n"
        "     Clear it and let the next comment-sync fire re-stamp it.\n"
        "  5. Force a rebuild via ClickUp 'Rebuild Snap Shot' checkbox to\n"
        "     resolve everything at once."
    )
    lines.append("")
    lines.append("─" * 60)

    for i, u in enumerate(flagged, 1):
        lines.append("")
        lines.append(f"[{i}] Property ID {u['pid']} · Unit {u['unit']}")
        lines.append(f"    Reason: {u['reason']}")
        if u["stamp_raw"]:
            lines.append(f"    Col Q raw: {u['stamp_raw']!r}")
        else:
            lines.append("    Col Q: (blank)")
        comment = u["comment"]
        if len(comment) > 400:
            comment = comment[:400] + " …"
        lines.append("    Col P comment:")
        for cline in comment.splitlines() or [""]:
            lines.append(f"      | {cline}")

    lines.append("")
    lines.append("─" * 60)
    lines.append(
        "This email is sent by the Snap Shot afternoon sync-health sweep at\n"
        "5 PM ET on weekdays. It is silent if every col-P comment has a\n"
        "fresh Last Synced stamp. Owner: Alexis (apattison@prudentgrowth.com)."
    )
    return "\n".join(lines)


def send_sweep_email(flagged, stats, workbook_link, dry_run):
    subject = (
        f"Snap Shot sync-health: {stats['flagged']} stale comment(s) "
        f"at {stats['now_et']}"
    )
    body = _format_email_body(flagged, stats, workbook_link)

    if dry_run:
        LOG.info(f"[dry-run] would email subject={subject!r}")
        LOG.info(f"[dry-run] body:\n{body}")
        return

    try:
        token = get_graph_access_token()
    except Exception as e:  # noqa: BLE001
        LOG.error(f"Could not send sync-health email (token acquisition failed): {e}")
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
        context="Sync-health sweep email send",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json=payload, retries=1,
    )
    if r.status_code != 202:
        LOG.error(
            f"Sync-health email failed to send: HTTP {r.status_code}: {r.text[:300]}"
        )
    else:
        LOG.info(f"Sync-health email sent to {ERROR_NOTIFICATION_TO}.")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--stale-hours", type=int, default=DEFAULT_STALE_HOURS,
        help=(
            "How many hours old a Last Synced stamp must be before it counts "
            f"as stale. Default {DEFAULT_STALE_HOURS}."
        ),
    )
    parser.add_argument(
        "--dry-run", action="store_true",
        help="Do everything except send the email.",
    )
    args = parser.parse_args()

    LOG.info(
        f"=== snap-shot afternoon sync-health sweep starting "
        f"stale_hours={args.stale_hours} dry_run={args.dry_run} ==="
    )

    # 1. Download.
    try:
        wb_bytes, last_mod_by, last_mod_at = download_from_sharepoint()
    except Exception as e:  # noqa: BLE001
        LOG.error(f"SharePoint download failed: {e}")
        # Send a failure notice so silence isn't ambiguous.
        try:
            token = get_graph_access_token()
            payload = {
                "message": {
                    "subject": "Snap Shot sync-health sweep FAILED to download workbook",
                    "body": {
                        "contentType": "Text",
                        "content": (
                            f"The 5 PM sync-health sweep could not download "
                            f"the workbook from SharePoint.\n\n{e}\n\n"
                            "This does NOT mean comment-sync is broken — the "
                            "monitor itself failed. Check the recurring task "
                            "log for details."
                        ),
                    },
                    "toRecipients": [
                        {"emailAddress": {"address": ERROR_NOTIFICATION_TO}}
                    ],
                },
                "saveToSentItems": "true",
            }
            http_request(
                "POST", "https://graph.microsoft.com/v1.0/me/sendMail",
                context="Sync-health sweep failure email",
                headers={
                    "Authorization": f"Bearer {token}",
                    "Content-Type": "application/json",
                },
                json=payload, retries=1,
            )
        except Exception as inner:  # noqa: BLE001
            LOG.error(f"Also failed to send failure email: {inner}")
        return 2

    LOG.info(
        f"SharePoint workbook downloaded. "
        f"Last modified by {last_mod_by} at {last_mod_at}."
    )

    # 2. Scan.
    flagged, healthy, stats = scan_workbook(wb_bytes, stale_hours=args.stale_hours)
    LOG.info(
        f"Scan complete: {stats['total_comments']} col-P comment(s), "
        f"{stats['flagged']} flagged, {stats['healthy']} healthy."
    )

    # 3. Email if anything to report.
    if not flagged:
        LOG.info("All comments have fresh stamps — nothing to report. Exiting silent.")
        return 0

    send_sweep_email(flagged, stats, WORKBOOK_WEB_URL, dry_run=args.dry_run)
    return 0


if __name__ == "__main__":
    sys.exit(main())
