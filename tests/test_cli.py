from __future__ import annotations

import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = Path(os.environ.get("TA_ML_DATA_DIR", ROOT / "data"))


class PipelineCommandTests(unittest.TestCase):
    def test_one_command_builds_the_three_agreed_outputs(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            output_dir = Path(temporary_directory)
            completed = subprocess.run(
                [
                    sys.executable,
                    "-m",
                    "ta_ml.cli",
                    "--data-dir",
                    str(DATA_DIR),
                    "--output-dir",
                    str(output_dir),
                ],
                cwd=ROOT,
                capture_output=True,
                text=True,
                check=False,
            )
            self.assertEqual(completed.returncode, 0, completed.stderr + completed.stdout)
            self.assertTrue((output_dir / "ta_ml_clean_v2.xlsx").is_file())
            self.assertTrue((output_dir / "ta_ml_model_panel.csv").is_file())
            self.assertTrue((output_dir / "evaluation_report.json").is_file())
            self.assertTrue((output_dir / "test_2025_predictions.csv").is_file())


if __name__ == "__main__":
    unittest.main()
