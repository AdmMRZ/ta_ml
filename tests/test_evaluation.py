from __future__ import annotations

import csv
import json
import os
import tempfile
import unittest
from pathlib import Path

import sklearn

from ta_ml.cleaning import build_clean_workbook
from ta_ml.evaluation import evaluate_models
from ta_ml.panel import build_model_panel

ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = Path(os.environ.get("TA_ML_DATA_DIR", ROOT / "data"))


class EvaluationTests(unittest.TestCase):
    def test_evaluates_real_panel_and_writes_auditable_outputs(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            directory = Path(temporary_directory)
            clean = directory / "clean.xlsx"
            panel = directory / "panel.csv"
            build_clean_workbook(source_dir=DATA_DIR, output_path=clean)
            build_model_panel(clean_path=clean, output_path=panel)
            result = evaluate_models(
                panel_path=panel,
                pln_path=DATA_DIR / "workbook_mitigasi_r22g14_numer - SHARE.xlsx",
                output_dir=directory / "evaluation",
            )

            report = json.loads(result.report_path.read_text(encoding="utf-8"))
            with result.predictions_path.open(newline="", encoding="utf-8") as handle:
                predictions = list(csv.DictReader(handle))

            self.assertEqual(len(predictions), 617)
            self.assertEqual({r["target_year"] for r in predictions}, {"2025"})
            self.assertIn(
                report["selected_model"], {"poisson_regressor", "hist_gradient_boosting_poisson"}
            )
            self.assertEqual(
                report["selection_folds"],
                [
                    {"train_target_years": [2021, 2022], "validation_target_year": 2023},
                    {
                        "train_target_years": [2021, 2022, 2023],
                        "validation_target_year": 2024,
                    },
                ],
            )
            self.assertEqual(report["final_fit_target_years"], [2021, 2022, 2023, 2024])
            self.assertEqual(report["final_fit_rows"], 2468)
            self.assertEqual(report["final_test_year"], 2025)
            self.assertEqual(report["software_versions"]["scikit_learn"], sklearn.__version__)
            self.assertIn("without_grounding", report["validation_candidates"])
            self.assertIn("snapshot_2026_sensitivity", report)
            self.assertEqual(
                report["snapshot_2026_sensitivity"]["excluded_raw_columns"],
                ["grounding_leg_d_ohm_snapshot_2026", "grounding_total_ohm_snapshot_2026"],
            )
            self.assertNotIn(
                "count_2020_coordinate_candidate",
                sum(report["feature_group_columns"].values(), []),
            )
            self.assertEqual(report["data_quality_2020"]["reassigned_by_exact_gps"], 609)
            self.assertEqual(report["data_quality_2020"]["unresolved_missing_count"], 8)
            self.assertNotIn("pln_rank_comparison", report)
            self.assertEqual(report["pln_alignment"]["eligible_towers"], 615)
            self.assertEqual(report["pln_alignment"]["excluded_towers"], 2)
            self.assertIn("not an independent 2025 forecast", report["pln_alignment"]["warning"])
            sample = next(row for row in predictions if row["tower_key"] == "KASONGAN - SUDAN #186")
            self.assertEqual(float(sample["observed_ng_2025"]), 8.31)
            self.assertAlmostEqual(float(sample["p_strike_annual_pln"]), 0.584682, places=5)
            self.assertEqual(sample["pln_ng_matches_observed_2025"], "True")
            self.assertTrue(
                all(float(row["prediction_p_strike_annual"]) >= 0 for row in predictions)
            )
            self.assertEqual(len(report["test_by_corridor"]), 3)


if __name__ == "__main__":
    unittest.main()
