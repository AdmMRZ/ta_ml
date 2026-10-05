from __future__ import annotations

import csv
import os
import tempfile
import unittest
from collections import Counter
from pathlib import Path

from ta_ml.cleaning import build_clean_workbook
from ta_ml.panel import build_model_panel

ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = Path(os.environ.get("TA_ML_DATA_DIR", ROOT / "data"))


class ModelPanelTests(unittest.TestCase):
    def test_builds_temporally_safe_pairs_from_real_workbooks(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            temp = Path(temporary_directory)
            clean_path = temp / "clean.xlsx"
            panel_path = temp / "panel.csv"
            build_clean_workbook(source_dir=DATA_DIR, output_path=clean_path)
            summary = build_model_panel(clean_path=clean_path, output_path=panel_path)

            with panel_path.open(newline="", encoding="utf-8") as handle:
                rows = list(csv.DictReader(handle))

            self.assertEqual(summary.rows, 3085)
            self.assertEqual(len(rows), 3085)
            self.assertEqual(
                Counter(row["split"] for row in rows),
                {"development": 2468, "test": 617},
            )
            self.assertEqual(
                {row["split"] for row in rows if row["target_year"] == "2023"},
                {"development"},
            )
            self.assertEqual(
                len({(r["corridor"], r["tower_key"], r["target_year"]) for r in rows}), 3085
            )
            self.assertTrue(all(int(r["target_year"]) == int(r["feature_year"]) + 1 for r in rows))
            self.assertTrue(all(r["elevation_m"] and r["tower_height_m"] for r in rows))
            self.assertIn("grounding_leg_a_ohm_snapshot_2026", rows[0])
            self.assertIn("tla_active_asof", rows[0])
            first_target = [r for r in rows if r["target_year"] == "2021"]
            self.assertEqual(
                Counter(r["count_2020_coordinate_match_status"] for r in first_target),
                {"exact_coordinate_reassigned": 609, "unmatched": 8},
            )
            self.assertEqual(
                sum(bool(r["count_lag1"]) for r in first_target),
                609,
            )
            self.assertEqual(
                sum(not r["count_lag2"] for r in rows if r["target_year"] == "2022"),
                8,
            )
            self.assertIn("grounding_d_matches_sheet4", rows[0])


if __name__ == "__main__":
    unittest.main()
