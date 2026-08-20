from __future__ import annotations

import json
import shutil
from pathlib import Path

import pytest

from vsm_postprocessing.errors import ConfigurationError, PipelineError
from vsm_postprocessing.pipeline_engine import load_pipeline_config, run_pipeline

PROJECT_ROOT = Path(__file__).resolve().parents[1]
PIPELINE_CONFIG = PROJECT_ROOT / "config" / "end_to_end_example.yaml"
TINY_INPUT = PROJECT_ROOT / "tests" / "fixtures" / "tiny_vsm_input.csv"


def _write_minimal_pipeline_config(tmp_path: Path, *, extra_root: str = "") -> Path:
    source = tmp_path / "input.csv"
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
    config.write_text(
        f"""version: 1
{extra_root}input:
  file: {source.name}
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


def test_pipeline_config_resolves_relative_paths(tmp_path: Path) -> None:
    path = _write_minimal_pipeline_config(tmp_path)
    config = load_pipeline_config(path)
    assert config.input_file == (tmp_path / "input.csv").resolve()
    assert config.output_root == (tmp_path / "output").resolve()
    assert config.import_options.strict is True
    assert config.math_config == (tmp_path / "math_channels.yaml").resolve()


def test_pipeline_config_rejects_unknown_root_keys(tmp_path: Path) -> None:
    path = _write_minimal_pipeline_config(tmp_path, extra_root="unexpected: true\n")
    with pytest.raises(ConfigurationError, match="Unknown key"):
        load_pipeline_config(path)


def test_pipeline_config_rejects_removed_scenario_block(tmp_path: Path) -> None:
    path = _write_minimal_pipeline_config(tmp_path, extra_root="duty" + "_cycle: {}\n")
    with pytest.raises(ConfigurationError, match="Unknown key"):
        load_pipeline_config(path)


def test_pipeline_failure_writes_diagnostic_manifest(tmp_path: Path) -> None:
    path = _write_minimal_pipeline_config(tmp_path)
    # Selection config is intentionally incomplete, so inspection passes and stage 2 fails.
    with pytest.raises(PipelineError, match="channel_selection"):
        run_pipeline(path)

    manifest_path = tmp_path / "output" / "pipeline_manifest.json"
    summary_path = tmp_path / "output" / "pipeline_summary.txt"
    assert manifest_path.exists()
    assert summary_path.exists()
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    assert manifest["status"] == "FAIL"
    assert manifest["stages"][0]["name"] == "inspection"
    assert manifest["stages"][0]["status"] == "PASS"
    assert manifest["stages"][1]["name"] == "channel_selection"
    assert manifest["stages"][1]["status"] == "FAIL"


def test_example_pipeline_config_loads_without_removed_scenario_block(tmp_path: Path) -> None:
    config_dir = tmp_path / "config"
    fixture_dir = tmp_path / "tests" / "fixtures"
    config_dir.mkdir()
    fixture_dir.mkdir(parents=True)
    shutil.copy2(TINY_INPUT, fixture_dir / TINY_INPUT.name)
    for name in (
        "channel_selection_example.yaml",
        "math_channels_example.yaml",
        "statistics_example.yaml",
        "plotting_example.yaml",
        "statistics_excel_report.yaml",
        "excel_report_example.yaml",
        "powerpoint_report_example.yaml",
    ):
        shutil.copy2(PROJECT_ROOT / "config" / name, config_dir / name)

    portable_config = config_dir / "end_to_end_example.yaml"
    source = PIPELINE_CONFIG.read_text(encoding="utf-8")
    source = source.replace(
        "../reference_files/RoboSprayer_3500Kg_Electric_12kph_Batt_50kW_Mot_63RPM_Susp_Cool_Rough_Grad_Discharge.csv",
        "../tests/fixtures/tiny_vsm_input.csv",
    )
    portable_config.write_text(source, encoding="utf-8")

    config = load_pipeline_config(portable_config)

    assert config.input_file == (fixture_dir / TINY_INPUT.name).resolve()
    assert config.channel_selection_config == (config_dir / "channel_selection_example.yaml").resolve()
