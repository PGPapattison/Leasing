"""Tests for the Vacancy Pipeline prospect subtask rollup (v1.9).

These tests lock in the specific behaviors called out in the v1.9 brief:

  1. Prospect name extraction handles the standard
     "Property | Unit: X | Prospect Y" format AND ad-hoc titles.
  2. Comment tail formatting truncates to 80 chars with an ellipsis and
     collapses whitespace.
  3. Missing / empty comment lists render as status-only lines.
  4. Hide-list filtering drops closed/dead statuses.
  5. Rollup block is empty when no prospects are active (so col O renders
     identically to today for that unit).
  6. Rollup block header is "Prospects (N):" with a bullet-line body.
  7. Status is Title-Cased and preserves slashes (e.g. "loi under review"
     -> "Loi Under Review", "lease/expansion executed" -> "Lease/Expansion Executed").
"""
import os
import sys
import unittest
from datetime import datetime, timezone


HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))

# Import module under test.  Suppress the noisy top-of-file logging so pytest
# output stays clean.
import logging
logging.getLogger().setLevel(logging.CRITICAL)

import rebuild  # noqa: E402
from rebuild import (  # noqa: E402
    _extract_prospect_name,
    _extract_status_title,
    _format_comment_tail,
    _sort_prospect_subtasks,
    format_prospect_rollup_block,
    VACANCY_PROSPECT_HIDE_STATUSES,
)


try:
    import zoneinfo
    ET = zoneinfo.ZoneInfo("America/New_York")
except Exception:  # pragma: no cover
    ET = timezone.utc


class ProspectNameExtraction(unittest.TestCase):
    def test_property_unit_prospect_format(self):
        # Real live example from list 901113575628.
        self.assertEqual(
            _extract_prospect_name("Piedmont Parkway | Unit: Suite 105 | Pathway Autism Services"),
            "Pathway Autism Services",
        )

    def test_explicit_prospect_label(self):
        self.assertEqual(
            _extract_prospect_name("Somerset Plaza | Unit: 155 | Prospect: Acme Dental"),
            "Acme Dental",
        )

    def test_ad_hoc_title_verbatim(self):
        # No pipes → verbatim.
        self.assertEqual(
            _extract_prospect_name("FU with John on Restaurant LOI"),
            "FU with John on Restaurant LOI",
        )

    def test_empty_input(self):
        self.assertEqual(_extract_prospect_name(""), "")
        self.assertEqual(_extract_prospect_name(None), "")


class StatusTitleCasing(unittest.TestCase):
    def test_title_case_multiword(self):
        self.assertEqual(
            _extract_status_title({"status": {"status": "loi under review"}}),
            "Loi Under Review",
        )

    def test_preserves_slash(self):
        # v1.9 brief-side status name.
        self.assertEqual(
            _extract_status_title({"status": {"status": "lease/expansion executed"}}),
            "Lease/Expansion Executed",
        )

    def test_empty(self):
        self.assertEqual(_extract_status_title({}), "")
        self.assertEqual(_extract_status_title({"status": None}), "")


class CommentTailFormat(unittest.TestCase):
    def test_no_comments_returns_empty(self):
        self.assertEqual(_format_comment_tail([], ET), "")

    def test_short_comment_no_truncation(self):
        # 2026-09-02 12:00 ET → ms
        dt = datetime(2026, 9, 2, 12, 0, tzinfo=ET)
        ms = int(dt.timestamp() * 1000)
        comments = [{"comment_text": "Sent lease v2", "date": str(ms)}]
        tail = _format_comment_tail(comments, ET)
        self.assertIn('"Sent lease v2"', tail)
        self.assertIn("9/2", tail)

    def test_long_comment_truncated_with_ellipsis(self):
        dt = datetime(2026, 9, 2, 12, 0, tzinfo=ET)
        ms = int(dt.timestamp() * 1000)
        long_text = "x" * 200
        comments = [{"comment_text": long_text, "date": str(ms)}]
        tail = _format_comment_tail(comments, ET)
        # The text inside quotes must be no longer than 80 chars total
        # (79 visible + ellipsis).
        inside = tail.split('"')[1]
        self.assertLessEqual(len(inside), 80)
        self.assertTrue(inside.endswith("\u2026"))

    def test_multiline_whitespace_collapsed(self):
        dt = datetime(2026, 9, 2, 12, 0, tzinfo=ET)
        ms = int(dt.timestamp() * 1000)
        comments = [{"comment_text": "line one\n\n  line two\ttabbed", "date": str(ms)}]
        tail = _format_comment_tail(comments, ET)
        self.assertIn("line one line two tabbed", tail)
        self.assertNotIn("\n", tail)
        self.assertNotIn("\t", tail)

    def test_newest_first_wins(self):
        # ClickUp returns newest-first; formatter should always use [0].
        dt = datetime(2026, 9, 2, 12, 0, tzinfo=ET)
        ms = int(dt.timestamp() * 1000)
        comments = [
            {"comment_text": "NEWEST", "date": str(ms)},
            {"comment_text": "OLDER", "date": str(ms - 86400000)},
        ]
        self.assertIn("NEWEST", _format_comment_tail(comments, ET))


class SortProspectSubtasks(unittest.TestCase):
    def test_most_recent_first(self):
        subs = [
            {"id": "a", "date_updated": "1000"},
            {"id": "b", "date_updated": "3000"},
            {"id": "c", "date_updated": "2000"},
        ]
        sorted_subs = _sort_prospect_subtasks(subs)
        self.assertEqual([s["id"] for s in sorted_subs], ["b", "c", "a"])


class HideListFiltering(unittest.TestCase):
    def test_brief_specified_and_live_statuses_present(self):
        # Both the brief's generic names AND the actual live names should
        # be in the hide-list (verified via cu_get_list_statuses on
        # 2026-09-08).
        for s in [
            "closed", "closed lost", "lost", "dead",
            "lease executed", "lease-expansion executed", "complete", "completed",
            "lease/expansion executed", "dead/lost deal/completed",
        ]:
            self.assertIn(s, VACANCY_PROSPECT_HIDE_STATUSES, f"missing hide: {s}")


class RollupBlockFormatting(unittest.TestCase):
    def test_empty_rows_returns_empty(self):
        self.assertEqual(format_prospect_rollup_block([]), "")

    def test_single_prospect_block(self):
        rows = [{
            "prospect": "Pathway Autism Services",
            "status": "Loi Under Review",
            "comment": '(9/2: "Sent lease v2")',
            "line": '\u2022 Pathway Autism Services \u2014 Loi Under Review (9/2: "Sent lease v2")',
        }]
        out = format_prospect_rollup_block(rows)
        self.assertTrue(out.startswith("Prospects (1):\n"))
        self.assertIn("Pathway Autism Services", out)
        self.assertIn("Loi Under Review", out)

    def test_multi_prospect_block(self):
        rows = [
            {"line": "\u2022 A \u2014 X"},
            {"line": "\u2022 B \u2014 Y"},
            {"line": "\u2022 C \u2014 Z"},
        ]
        out = format_prospect_rollup_block(rows)
        self.assertTrue(out.startswith("Prospects (3):\n"))
        self.assertEqual(out.count("\u2022"), 3)


class PullVacancyProspectSubtasks(unittest.TestCase):
    """Behavior test for pull_vacancy_prospect_subtasks using a fake in-memory
    ClickUp response. Exercises parent/subtask indexing, hide-list, missing
    parent-key drop, and per-subtask comment fetching.
    """

    def test_pull_end_to_end(self):
        # Fake parents and subtasks. Parent A has PropertyId + Unit.
        # Parent B lacks a Unit so its child should be dropped (no key).
        parent_a = {
            "id": "PA",
            "parent": None,
            "custom_fields": [
                {"id": rebuild.CU_PROPERTY_ID_FIELD, "value": "279"},
                {"id": rebuild.CU_UNIT_NUMBER_FIELD, "value": "Suite 105"},
            ],
            "status": {"status": "available"},
            "name": "Piedmont Parkway | Unit: Suite 105",
        }
        parent_b = {
            "id": "PB",
            "parent": None,
            "custom_fields": [
                {"id": rebuild.CU_PROPERTY_ID_FIELD, "value": "300"},
                {"id": rebuild.CU_UNIT_NUMBER_FIELD, "value": ""},  # blank
            ],
            "status": {"status": "available"},
            "name": "Some Place",
        }
        sub_a1 = {
            "id": "SA1",
            "parent": "PA",
            "status": {"status": "loi under review"},
            "name": "Piedmont Parkway | Unit: Suite 105 | Pathway Autism Services",
            "date_updated": "3000",
        }
        sub_a2 = {
            "id": "SA2",
            "parent": "PA",
            "status": {"status": "lease/expansion executed"},  # hidden
            "name": "Piedmont Parkway | Unit: Suite 105 | Won Deal",
            "date_updated": "4000",
        }
        sub_b1 = {
            "id": "SB1",
            "parent": "PB",
            "status": {"status": "prospects engaged"},
            "name": "Some Place | Prospect Y",
            "date_updated": "2000",
        }

        rebuild.cu_get_list_tasks = lambda list_id, include_closed="true": (
            [parent_a, parent_b, sub_a1, sub_a2, sub_b1]
        )
        rebuild.cu_get_task_comments = lambda tid: []

        try:
            out = rebuild.pull_vacancy_prospect_subtasks()
        finally:
            # Do NOT restore \u2014 subsequent tests use their own fakes and this
            # module's real cu_* calls need a token anyway (never called in
            # tests). Skipping restore keeps the surface simple.
            pass

        # Parent A / Suite 105 should be present with one active prospect.
        # Parent B should be dropped entirely (no unit).
        key = ("279", rebuild._norm_unit_label("Suite 105"))
        self.assertIn(key, out)
        self.assertEqual(len(out[key]), 1)
        self.assertEqual(out[key][0]["prospect"], "Pathway Autism Services")
        self.assertEqual(out[key][0]["status"], "Loi Under Review")
        # SB1 dropped (parent had no unit).
        self.assertNotIn(("300", ""), out)
        for k in out:
            self.assertNotEqual(k[0], "300")


if __name__ == "__main__":  # pragma: no cover
    unittest.main()
