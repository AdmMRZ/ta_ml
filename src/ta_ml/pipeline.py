"""One public interface for building auditable local analysis artifacts."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from .cleaning import RAW_GEOMETRY_FILENAME, build_clean_workbook
from .evaluation import evaluate_models
from .panel import build_model_panel


@dataclass(frozen=True)
class PipelineArtifacts:
    clean_workbook: Path
    model_panel: Path
    evaluation_report: Path
    test_predictions: Path


def run_pipeline(*, data_dir: Path, output_dir: Path, overwrite: bool = False) -> PipelineArtifacts:
    """Build the four local outputs without modifying any source workbook."""

    data_dir, output_dir = data_dir.resolve(), output_dir.resolve()
    clean_path = output_dir / "ta_ml_clean_v2.xlsx"
    panel_path = output_dir / "ta_ml_model_panel.csv"
    clean = build_clean_workbook(source_dir=data_dir, output_path=clean_path, overwrite=overwrite)
    panel = build_model_panel(
        clean_path=clean.output_path, output_path=panel_path, overwrite=overwrite
    )
    evaluation = evaluate_models(
        panel_path=panel.output_path,
        pln_path=data_dir / RAW_GEOMETRY_FILENAME,
        output_dir=output_dir,
        overwrite=overwrite,
    )
    return PipelineArtifacts(
        clean_workbook=clean.output_path,
        model_panel=panel.output_path,
        evaluation_report=evaluation.report_path,
        test_predictions=evaluation.predictions_path,
    )
