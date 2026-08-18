"""Tests for _norm_unit_label — the ClickUp ↔ AppFolio unit-label matcher.

The 2026-08-18 discovery: ClickUp tasks created by hand often use spaces
around hyphens (e.g. '181 - B1') while AppFolio's unit_display returns the
same unit as '181-B1'. Without normalization the (pid, unit_norm) lookup in
sync_rem_comments_to_clickup silently misses every such unit, so REM
comments in col P orphan and never post to ClickUp.

This test suite locks in the collapse-hyphen-whitespace behavior so the
same class of bug can't come back.
"""
import os
import sys
import unittest


# Make snap_shot/ importable when running from the repo root or from tests/.
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))

from rebuild import _norm_unit_label  # noqa: E402


class TestNormUnitLabel(unittest.TestCase):
    """Snap Shot unit-label normalization must survive real ClickUp inputs."""

    def test_bare_unit_number(self):
        self.assertEqual(_norm_unit_label("155"), "155")
        self.assertEqual(_norm_unit_label("Unit 155"), "155")
        self.assertEqual(_norm_unit_label("#155"), "155")
        self.assertEqual(_norm_unit_label("Suite 155"), "155")

    def test_hyphen_no_whitespace(self):
        self.assertEqual(_norm_unit_label("181-B1"), "181-b1")
        self.assertEqual(_norm_unit_label("Unit 181-B1"), "181-b1")

    def test_hyphen_with_whitespace_padding(self):
        """2026-08-18 WingStop bug: '181 - B1' must match '181-B1'."""
        self.assertEqual(_norm_unit_label("181 - B1"), "181-b1")
        self.assertEqual(_norm_unit_label("Unit 181 - B1"), "181-b1")
        self.assertEqual(_norm_unit_label("7746 - 7754"), "7746-7754")

    def test_stacked_hyphens(self):
        self.assertEqual(_norm_unit_label("a - b - c"), "a-b-c")

    def test_asymmetric_hyphen_padding(self):
        self.assertEqual(_norm_unit_label("181- B1"), "181-b1")
        self.assertEqual(_norm_unit_label("181 -B1"), "181-b1")

    def test_no_hyphen_spaces_preserved(self):
        """Absent a hyphen, internal single space is preserved (still lowercased)."""
        self.assertEqual(_norm_unit_label("Warehouse Cages"), "warehouse cages")
        self.assertEqual(_norm_unit_label("B 2"), "b 2")

    def test_empty_and_whitespace(self):
        self.assertEqual(_norm_unit_label(""), "")
        self.assertEqual(_norm_unit_label(None), "")
        self.assertEqual(_norm_unit_label("   "), "")
        self.assertEqual(_norm_unit_label("Unit"), "")

    def test_prefix_variants(self):
        self.assertEqual(_norm_unit_label("STE 300A"), "300a")
        self.assertEqual(_norm_unit_label("unit.155"), "155")
        self.assertEqual(_norm_unit_label("  Unit  B-A  "), "b-a")


if __name__ == "__main__":
    unittest.main()
