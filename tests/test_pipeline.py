from __future__ import annotations

import json
from pathlib import Path

import pytest

from vsm_postprocessing.errors import ConfigurationError, PipelineError
from vsm_postprocessing.pipeline_engine import load_pipeline_config, run_pipeline

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SOURCE_WORKBOOK = PROJECT_ROOT / "reference_files" / "Sprayer_Caiman_SP_9300Kg_Hybrid_Gen80kW_30kph_74Ht_4000KgAQ_57-4pcSOC_5-80_1C2G_02.xlsx"
PIPELINE_CONFIG = PROJECT_ROOT / "config" / "end_to_end_example.yaml"


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


@pytest.mark.skipif(not SOURCE_WORKBOOK.exists(), reason="Client source workbook is not present")
def test_supplied_source_workbook_end_to_end_acceptance(tmp_path: Path) -> None:
    raw = PIPELINE_CONFIG.read_text(encoding="utf-8")
    raw = raw.replace("../outputs/end_to_end", str((tmp_path / "end_to_end").resolve()).replace("\\", "/"))
    config_path = tmp_path / "end_to_end_acceptance.yaml"
    # Config references are normally relative to config/. This temp config lives elsewhere,
    # so replace them with absolute paths for the acceptance test.
    raw = raw.replace(
        "../reference_files/Sprayer_Caiman_SP_9300Kg_Hybrid_Gen80kW_30kph_74Ht_4000KgAQ_57-4pcSOC_5-80_1C2G_02.xlsx",
        str(SOURCE_WORKBOOK.resolve()).replace("\\", "/"),
    )
    for filename in (
        "channel_selection_example.yaml",
        "math_channels_example.yaml",
        "statistics_example.yaml",
        "plotting_example.yaml",
        "statistics_excel_report.yaml",
        "excel_report_example.yaml",
        "powerpoint_report_example.yaml",
    ):
        raw = raw.replace(filename, str((PROJECT_ROOT / "config" / filename).resolve()).replace("\\", "/"))
    config_path.write_text(raw, encoding="utf-8")

    result = run_pipeline(config_path)
    assert result.status == "PASS"
    assert result.completed_stage_count == 7
    assert result.report_path is not None and result.report_path.exists()
    assert result.powerpoint_path is not None and result.powerpoint_path.exists()
    assert result.manifest_path.exists()
    assert result.summary_path.exists()
    assert [stage.name for stage in result.stages] == [
        "inspection",
        "channel_selection",
        "math_channels",
        "statistics",
        "plotting",
        "excel_report",
        "powerpoint_report",
    ]
    assert all(stage.status == "PASS" for stage in result.stages)
    assert result.stages[0].metrics["samples"] == 1866
    assert result.stages[1].metrics["selected_channels"] == 12
    assert result.stages[2].metrics["math_channels"] == 13
    assert result.stages[4].metrics["plots"] == 24
    assert result.stages[5].metrics["report_channels"] == 21
    assert result.stages[6].metrics["slides"] == 4
