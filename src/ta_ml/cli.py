"""Single-command entry point for local TA data and model artifacts."""

from __future__ import annotations

import argparse
from pathlib import Path

from .cleaning import DataValidationError
from .pipeline import run_pipeline


def main() -> int:
    parser = argparse.ArgumentParser(description="Build and evaluate the local TA ML pipeline")
    parser.add_argument("--data-dir", type=Path, default=Path("data"))
    parser.add_argument("--output-dir", type=Path, default=Path("data/output"))
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    try:
        result = run_pipeline(
            data_dir=args.data_dir,
            output_dir=args.output_dir,
            overwrite=args.overwrite,
        )
    except (DataValidationError, FileExistsError, KeyError) as error:
        parser.exit(1, f"Pipeline stopped: {error}\n")

    print(f"Clean workbook: {result.clean_workbook}")
    print(f"Model panel: {result.model_panel}")
    print(f"Evaluation report: {result.evaluation_report}")
    print(f"Test predictions: {result.test_predictions}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
