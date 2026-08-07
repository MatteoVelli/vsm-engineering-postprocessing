from __future__ import annotations

import json
from pathlib import Path

import pytest

from vsm_postprocessing.doctor import DoctorCheck, run_doctor
from vsm_postprocessing.errors import ConfigurationError, PipelineError
from vsm_postprocessing.pipeline_engine import _prepare_output_root, load_pipeline_config, run_pipeline
from vsm_postprocessing.utils import atomic_write_text

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SOURCE_WORKBOOK = PROJECT_ROOT / "reference_files" / "Sprayer_Caiman_SP_9300Kg_Hybrid_Gen80kW_30kph_74Ht_4000KgAQ_57-4pcSOC_5-80_1C2G_02.xlsx"


def _write_pipeline_config(tmp_path: Path, *, input_inside_output: bool = False) -> Path:
    output_dir = tmp_path / "output"
    source_dir = output_dir if input_inside_output else tmp_path
    source_dir.mkdir(parents=True, exist_ok=True)
    source = source_dir / "input.csv"
    source.write_text("Time,Signal\ns,kW\n0,1\n1,2\n", encoding="utf-8")

    config_files = {}
    for name in (
        "channel_selection",
        "math_channels",
        "statistics",
        "plotting",
        "excel_statistics",
        "excel_report",
    ):
        path = tmp_path / f"{name}.yaml"
        path.write_text("version: 1\n", encoding="utf-8")
        config_files[name] = path.name

    config = tmp_path / "pipeline.yaml"
    source_value = source.relative_to(tmp_path).as_posix()
    config.write_text(
        f"""version: 1
input:
  file: {source_value}
  strict: true
configs:
  channel_selection: {config_files['channel_selection']}
  math_channels: {config_files['math_channels']}
  statistics: {config_files['statistics']}
  plotting: {config_files['plotting']}
  excel_statistics: {config_files['excel_statistics']}
  excel_report: {config_files['excel_report']}
output:
  root_dir: output
  clean_before_run: true
""",
        encoding="utf-8",
    )
    return config


def test_atomic_write_text_replaces_existing_file(tmp_path: Path) -> None:
    path = tmp_path / "metadata.json"
    path.write_text("old", encoding="utf-8")
    atomic_write_text(path, "new\n")
    assert path.read_text(encoding="utf-8") == "new\n"
    assert not list(tmp_path.glob("*.tmp"))


def test_safe_output_clean_preserves_unknown_user_files(tmp_path: Path) -> None:
    output = tmp_path / "output"
    (output / "01_inspection").mkdir(parents=True)
    (output / "01_inspection" / "generated.txt").write_text("generated", encoding="utf-8")
    (output / "user_notes.txt").write_text("keep me", encoding="utf-8")
    (output / "pipeline_manifest.json").write_text("{}", encoding="utf-8")

    _prepare_output_root(output, clean=True)

    assert not (output / "01_inspection").exists()
    assert not (output / "pipeline_manifest.json").exists()
    assert (output / "user_notes.txt").read_text(encoding="utf-8") == "keep me"
    assert (output / ".vsm_postprocessing_output").exists()


def test_pipeline_config_rejects_input_inside_output_root(tmp_path: Path) -> None:
    config = _write_pipeline_config(tmp_path, input_inside_output=True)
    with pytest.raises(ConfigurationError, match="dedicated output directory|must not be located inside"):
        load_pipeline_config(config)


def test_unexpected_pipeline_error_is_logged_and_manifested(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    config = _write_pipeline_config(tmp_path)

    def explode(*args: object, **kwargs: object) -> object:
        raise ValueError("synthetic unexpected failure")

    monkeypatch.setattr("vsm_postprocessing.pipeline_engine.inspect_data_file", explode)

    with pytest.raises(PipelineError, match="Unexpected ValueError"):
        run_pipeline(config)

    output = tmp_path / "output"
    manifest = json.loads((output / "pipeline_manifest.json").read_text(encoding="utf-8"))
    assert manifest["status"] == "FAIL"
    assert manifest["software"]["name"] == "vsm-postprocessing"
    assert manifest["stages"][0]["name"] == "inspection"
    assert manifest["stages"][0]["status"] == "FAIL"
    assert "Unexpected ValueError" in manifest["stages"][0]["error"]
    assert manifest["started_at_utc"]
    assert manifest["completed_at_utc"]
    assert manifest["duration_seconds"] >= 0
    assert (output / "pipeline.log").exists()
    log = (output / "pipeline.log").read_text(encoding="utf-8")
    assert "synthetic unexpected failure" in log


def test_doctor_reports_no_blocking_failure_for_project(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(
        "vsm_postprocessing.doctor._check_packages",
        lambda: [DoctorCheck("VSM package", "PASS", "test")],
    )
    monkeypatch.setattr(
        "vsm_postprocessing.doctor._check_uv",
        lambda: DoctorCheck("uv launcher", "PASS", "test"),
    )
    report = run_doctor(PROJECT_ROOT)
    assert report.status == "PASS"
    assert report.error_count == 0
    assert any(check.name == "Python version" and check.status == "PASS" for check in report.checks)
    assert any(check.name == "VSM package" and check.status == "PASS" for check in report.checks)


@pytest.mark.skipif(not SOURCE_WORKBOOK.exists(), reason="Client source workbook is not present")
def test_doctor_validates_default_pipeline_bundle_when_reference_present(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(
        "vsm_postprocessing.doctor._check_packages",
        lambda: [DoctorCheck("VSM package", "PASS", "test")],
    )
    monkeypatch.setattr(
        "vsm_postprocessing.doctor._check_uv",
        lambda: DoctorCheck("uv launcher", "PASS", "test"),
    )
    report = run_doctor(PROJECT_ROOT)
    config_checks = [check for check in report.checks if "configuration" in check.name.lower()]
    assert config_checks
    assert all(check.status == "PASS" for check in config_checks)
