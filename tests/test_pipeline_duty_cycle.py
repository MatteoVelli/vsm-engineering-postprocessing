from __future__ import annotations

import json
from pathlib import Path

import numpy as np
import pytest

from vsm_postprocessing.duty_cycle import (
    WorkbookRowProfileProvider,
    compose_duty_cycle,
    export_pipeline_dataset,
    load_duty_cycle_config,
    load_profile_provider_config,
)
from vsm_postprocessing.importer import ImportOptions, load_data_file
from vsm_postprocessing.pipeline_engine import load_pipeline_config, run_pipeline

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SOURCE_WORKBOOK = PROJECT_ROOT / "reference_files" / "Sprayer_Caiman_SP_9300Kg_Hybrid_Gen80kW_30kph_74Ht_4000KgAQ_57-4pcSOC_5-80_1C2G_02.xlsx"
REFERENCE_WORKBOOK = PROJECT_ROOT / "reference_files" / "Sprayer_Caiman_SP_9300Kg_Electrification_03.xlsx"
SCENARIO_CONFIG = PROJECT_ROOT / "config" / "duty_cycle_sergio_reference.yaml"
PROVIDER_CONFIG = PROJECT_ROOT / "config" / "duty_cycle_profiles_sergio_reference.yaml"
PIPELINE_CONFIG = PROJECT_ROOT / "config" / "end_to_end_sergio_duty_cycle.yaml"


@pytest.mark.skipif(
    not SOURCE_WORKBOOK.exists() or not REFERENCE_WORKBOOK.exists(),
    reason="Sergio reference workbooks are not present",
)
def test_pipeline_dataset_export_roundtrips_full_composition(tmp_path: Path) -> None:
    source = load_data_file(SOURCE_WORKBOOK)
    scenario = load_duty_cycle_config(SCENARIO_CONFIG)
    provider = WorkbookRowProfileProvider(load_profile_provider_config(PROVIDER_CONFIG), REFERENCE_WORKBOOK)
    composition = compose_duty_cycle(scenario, source, provider)

    export_path = export_pipeline_dataset(composition, tmp_path / "duty_cycle_dataset.csv")
    reloaded = load_data_file(
        export_path,
        ImportOptions(
            header_row=1,
            unit_row=2,
            data_start_row=3,
            last_channel_column=source.quality.channel_count,
            time_channel=source.channels[source.channel_index(source.quality.time_channel_id)].source_name,
            strict=True,
        ),
    )

    assert reloaded.quality.sample_count == 17418
    assert reloaded.quality.channel_count == 70
    assert [channel.channel_id for channel in reloaded.channels] == [channel.channel_id for channel in source.channels]
    np.testing.assert_allclose(reloaded.values, composition.values, rtol=0.0, atol=0.0)


def test_full_duty_cycle_pipeline_config_loads_optional_composer_block(tmp_path: Path) -> None:
    source = tmp_path / "input.csv"
    source.write_text("Time,Signal\ns,kW\n0,1\n1,2\n", encoding="utf-8")
    scenario = tmp_path / "scenario.yaml"
    scenario.write_text("scenario: {}\n", encoding="utf-8")
    provider = tmp_path / "provider.yaml"
    provider.write_text("provider: {}\n", encoding="utf-8")
    profile = tmp_path / "profile.xlsx"
    profile.write_bytes(b"placeholder")
    config_names = (
        "channel_selection",
        "math_channels",
        "statistics",
        "plotting",
        "excel_statistics",
        "excel_report",
    )
    for name in config_names:
        (tmp_path / f"{name}.yaml").write_text("version: 1\n", encoding="utf-8")
    pipeline = tmp_path / "pipeline.yaml"
    pipeline.write_text(
        """version: 1
input:
  file: input.csv
duty_cycle:
  scenario: scenario.yaml
  profile_provider: provider.yaml
  profile_workbook: profile.xlsx
configs:
  channel_selection: channel_selection.yaml
  math_channels: math_channels.yaml
  statistics: statistics.yaml
  plotting: plotting.yaml
  excel_statistics: excel_statistics.yaml
  excel_report: excel_report.yaml
output:
  root_dir: output
""",
        encoding="utf-8",
    )
    config = load_pipeline_config(pipeline)
    assert config.duty_cycle is not None
    assert config.duty_cycle.scenario_config == scenario.resolve()
    assert config.duty_cycle.profile_provider_config == provider.resolve()
    assert config.duty_cycle.profile_workbook == profile.resolve()


@pytest.mark.skipif(
    not SOURCE_WORKBOOK.exists() or not REFERENCE_WORKBOOK.exists(),
    reason="Sergio reference workbooks are not present",
)
def test_full_duty_cycle_runs_through_normal_reporting_pipeline(tmp_path: Path) -> None:
    raw = PIPELINE_CONFIG.read_text(encoding="utf-8")
    raw = raw.replace(
        "../outputs/end_to_end_sergio_duty_cycle",
        str((tmp_path / "full_mission").resolve()).replace("\\", "/"),
    )
    raw = raw.replace(
        "../reference_files/Sprayer_Caiman_SP_9300Kg_Hybrid_Gen80kW_30kph_74Ht_4000KgAQ_57-4pcSOC_5-80_1C2G_02.xlsx",
        str(SOURCE_WORKBOOK.resolve()).replace("\\", "/"),
    )
    raw = raw.replace(
        "../reference_files/Sprayer_Caiman_SP_9300Kg_Electrification_03.xlsx",
        str(REFERENCE_WORKBOOK.resolve()).replace("\\", "/"),
    )
    for filename in (
        "duty_cycle_sergio_reference.yaml",
        "duty_cycle_profiles_sergio_reference.yaml",
        "channel_selection_example.yaml",
        "math_channels_example.yaml",
        "statistics_excel_report.yaml",
        "plotting_example.yaml",
        "excel_report_duty_cycle.yaml",
        "powerpoint_report_duty_cycle.yaml",
    ):
        raw = raw.replace(filename, str((PROJECT_ROOT / "config" / filename).resolve()).replace("\\", "/"))
    config_path = tmp_path / "full_duty_cycle_pipeline.yaml"
    config_path.write_text(raw, encoding="utf-8")

    result = run_pipeline(config_path)

    assert result.status == "PASS"
    assert result.completed_stage_count == 8
    assert [stage.name for stage in result.stages] == [
        "inspection",
        "duty_cycle",
        "channel_selection",
        "math_channels",
        "statistics",
        "plotting",
        "excel_report",
        "powerpoint_report",
    ]
    assert result.stages[0].metrics["samples"] == 1866
    assert result.stages[1].metrics["samples"] == 17418
    assert result.stages[1].metrics["phases"] == 12
    assert result.stages[1].metrics["external_profile_phases"] == 4
    assert result.stages[1].metrics["final_time_min"] == pytest.approx(290.28333333333336)
    assert result.stages[1].metrics["final_distance_km"] == pytest.approx(114.0011, abs=1e-4)
    assert result.stages[1].metrics["final_soc_pct"] == pytest.approx(23.9383, abs=1e-4)
    assert result.stages[1].metrics["max_speed_kph"] == pytest.approx(62.6233, abs=1e-4)
    assert result.stages[1].metrics["max_generator_kw"] == pytest.approx(80.02669, abs=1e-4)

    for stage in result.stages[2:]:
        assert stage.metrics["samples"] == 17418
    assert result.report_path is not None and result.report_path.exists()
    assert result.powerpoint_path is not None and result.powerpoint_path.exists()
    assert result.processing_input_file.name == "duty_cycle_dataset.csv"
    assert result.processing_input_file.exists()

    manifest = json.loads(result.manifest_path.read_text(encoding="utf-8"))
    assert manifest["duty_cycle"] is not None
    assert manifest["processing_input_file"] == str(result.processing_input_file)
    assert manifest["stage_count"] == 8
