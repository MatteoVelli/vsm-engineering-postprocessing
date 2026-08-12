from __future__ import annotations

import csv
from pathlib import Path

import pytest
import yaml

from vsm_postprocessing.errors import ConfigurationError
from vsm_postprocessing.importer import inspect_data_file, load_data_file
from vsm_postprocessing.pipeline_engine import load_pipeline_config, run_pipeline
from vsm_postprocessing.ui_config import (
    build_engineering_report_runtime_bundle,
    build_full_duty_cycle_runtime_bundle,
    build_runtime_bundle,
    default_full_duty_cycle_scenario,
    default_ui_profile,
    load_ui_profile,
    load_ui_templates,
    save_ui_profile,
)

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SOURCE_WORKBOOK = PROJECT_ROOT / "reference_files" / "Sprayer_Caiman_SP_9300Kg_Hybrid_Gen80kW_30kph_74Ht_4000KgAQ_57-4pcSOC_5-80_1C2G_02.xlsx"
PROFILE_WORKBOOK = PROJECT_ROOT / "reference_files" / "Sprayer_Caiman_SP_9300Kg_Electrification_03.xlsx"
CANONICAL_DUTY_CYCLE_PIPELINE = PROJECT_ROOT / "config" / "end_to_end_sergio_duty_cycle.yaml"


def test_default_ui_profile_matches_templates() -> None:
    templates = load_ui_templates(PROJECT_ROOT)
    profile = default_ui_profile(templates)
    assert profile["version"] == 1
    assert "chassis_speed__col_005" in profile["export_channels"]
    assert "calc_total_generator_power" in profile["math_channels"]
    assert "report_battery_power_rms" in profile["statistics"]
    assert "speed_vs_time" in profile["plots"]
    assert profile["generate_powerpoint"] is True


def test_ui_profile_round_trip(tmp_path: Path) -> None:
    templates = load_ui_templates(PROJECT_ROOT)
    profile = default_ui_profile(templates)
    path = save_ui_profile(tmp_path / "profile.yaml", profile)
    assert load_ui_profile(path) == profile


def test_ui_profile_rejects_duplicate_values(tmp_path: Path) -> None:
    path = tmp_path / "bad.yaml"
    path.write_text(
        """version: 1
export_channels: [a, a]
math_channels: []
report_channels: [a]
statistics: [s]
kpis: []
plots: [p]
""",
        encoding="utf-8",
    )
    with pytest.raises(ConfigurationError, match="duplicates"):
        load_ui_profile(path)


def test_runtime_bundle_adds_required_math_dependencies(tmp_path: Path) -> None:
    source = tmp_path / "source.csv"
    source.write_text("Time,Signal\ns,kW\n0,1\n1,2\n", encoding="utf-8")
    templates = load_ui_templates(PROJECT_ROOT)
    bundle = build_runtime_bundle(
        source_file=source,
        runtime_dir=tmp_path / "run",
        templates=templates,
        time_channel_id="track_time__col_001",
        export_channel_ids=["chassis_speed__col_005"],
        selected_math_channel_ids=[],
        report_channel_ids=["calc_total_generator_power"],
        selected_statistic_ids=["report_total_generator_power_max"],
        kpi_statistic_ids=["report_total_generator_power_max"],
        selected_plot_ids=["speed_vs_time"],
    )
    assert "calc_total_generator_power" in bundle.effective_math_channel_ids
    assert "calc_generator_power_1" in bundle.effective_math_channel_ids
    assert "calc_generator_power_2" in bundle.effective_math_channel_ids
    raw = yaml.safe_load(bundle.math_config.read_text(encoding="utf-8"))
    ids = [item["channel_id"] for item in raw["math_channels"]]
    assert ids == list(bundle.effective_math_channel_ids)


def test_runtime_bundle_writes_loadable_pipeline_config(tmp_path: Path) -> None:
    source = tmp_path / "source.csv"
    source.write_text("Time,Signal\ns,kW\n0,1\n1,2\n", encoding="utf-8")
    templates = load_ui_templates(PROJECT_ROOT)
    profile = default_ui_profile(templates)
    bundle = build_runtime_bundle(
        source_file=source,
        runtime_dir=tmp_path / "run",
        templates=templates,
        time_channel_id="track_time__col_001",
        export_channel_ids=profile["export_channels"],
        selected_math_channel_ids=profile["math_channels"],
        report_channel_ids=profile["report_channels"],
        selected_statistic_ids=profile["statistics"],
        kpi_statistic_ids=profile["kpis"],
        selected_plot_ids=profile["plots"],
    )
    config = load_pipeline_config(bundle.pipeline_config)
    assert config.input_file == source.resolve()
    assert config.duty_cycle is None
    assert config.output_root == bundle.output_root


@pytest.mark.skipif(not SOURCE_WORKBOOK.exists(), reason="Client source workbook is not present")
def test_engineering_report_runtime_bundle_is_single_file_xlsx(tmp_path: Path) -> None:
    bundle = build_engineering_report_runtime_bundle(
        source_file=SOURCE_WORKBOOK,
        runtime_dir=tmp_path / "engineering_report",
        project_root=PROJECT_ROOT,
    )
    config = load_pipeline_config(bundle.pipeline_config)
    raw = yaml.safe_load(bundle.pipeline_config.read_text(encoding="utf-8"))

    assert config.input_file == SOURCE_WORKBOOK.resolve()
    assert config.duty_cycle is not None
    assert "duty_cycle" in raw
    assert raw["duty_cycle"]["profile_workbook"].endswith("assets\\scenarios\\caiman_sp_hybrid\\missing_phase_profiles.csv") or raw[
        "duty_cycle"
    ]["profile_workbook"].endswith("assets/scenarios/caiman_sp_hybrid/missing_phase_profiles.csv")
    assert "reference_files" not in raw["duty_cycle"]["profile_workbook"]
    assert bundle.excel_report_config.name == "excel_report_duty_cycle.yaml"
    assert bundle.powerpoint_report_config is not None
    assert bundle.powerpoint_report_config.name == "powerpoint_report_duty_cycle.yaml"


@pytest.mark.skipif(not SOURCE_WORKBOOK.exists(), reason="Client source workbook is not present")
def test_engineering_report_runtime_bundle_accepts_csv_without_filename_or_sample_count_assumption(tmp_path: Path) -> None:
    inspection = inspect_data_file(SOURCE_WORKBOOK)
    dataset = load_data_file(SOURCE_WORKBOOK)
    csv_path = tmp_path / "new_vsm_simulation.csv"
    with csv_path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow([""] * len(inspection.channels))
        writer.writerow([""] * len(inspection.channels))
        writer.writerow([channel.display_name for channel in inspection.channels])
        writer.writerow([channel.unit or "" for channel in inspection.channels])
        writer.writerows(dataset.values)

    bundle = build_engineering_report_runtime_bundle(
        source_file=csv_path,
        runtime_dir=tmp_path / "engineering_csv",
        project_root=PROJECT_ROOT,
    )
    config = load_pipeline_config(bundle.pipeline_config)

    assert config.input_file == csv_path.resolve()
    assert config.duty_cycle is not None
    assert inspect_data_file(config.input_file, config.import_options).quality.sample_count == 1866


def test_engineering_report_runtime_bundle_reports_missing_required_channels(tmp_path: Path) -> None:
    csv_path = tmp_path / "missing_required.csv"
    csv_path.write_text("Time,Speed\ns,kph\n0,0\n1,1\n", encoding="utf-8")

    with pytest.raises(ConfigurationError, match="requires channel"):
        build_engineering_report_runtime_bundle(
            source_file=csv_path,
            runtime_dir=tmp_path / "bad_engineering",
            project_root=PROJECT_ROOT,
        )


def test_full_duty_cycle_runtime_bundle_uses_client_compatible_profile_validation(tmp_path: Path) -> None:
    source = tmp_path / "source.xlsx"
    profile = tmp_path / "profile.xlsx"
    source.write_bytes(b"placeholder source")
    profile.write_bytes(b"placeholder profile")
    scenario = default_full_duty_cycle_scenario(PROJECT_ROOT)

    bundle = build_full_duty_cycle_runtime_bundle(
        source_file=source,
        profile_workbook=profile,
        profile_original_filename="Sprayer_Caiman_SP_9300Kg_Electrification_03.xlsx",
        runtime_dir=tmp_path / "run",
        scenario=scenario,
    )
    config = load_pipeline_config(bundle.pipeline_config)
    raw = yaml.safe_load(bundle.pipeline_config.read_text(encoding="utf-8"))

    assert config.duty_cycle is not None
    assert config.duty_cycle.profile_validation_mode == "compatible"
    assert config.duty_cycle.profile_original_filename == "Sprayer_Caiman_SP_9300Kg_Electrification_03.xlsx"
    assert raw["duty_cycle"]["profile_validation_mode"] == "compatible"
    assert raw["duty_cycle"]["profile_original_filename"] == "Sprayer_Caiman_SP_9300Kg_Electrification_03.xlsx"
    assert config.input_file == source.resolve()
    assert config.duty_cycle.profile_workbook == profile.resolve()


def test_canonical_duty_cycle_pipeline_keeps_strict_profile_validation() -> None:
    config = load_pipeline_config(CANONICAL_DUTY_CYCLE_PIPELINE)

    assert config.duty_cycle is not None
    assert config.duty_cycle.profile_validation_mode == "strict"
    assert config.duty_cycle.profile_original_filename is None


@pytest.mark.skipif(not SOURCE_WORKBOOK.exists(), reason="Client source workbook is not present")
def test_ui_default_profile_end_to_end_acceptance(tmp_path: Path) -> None:
    templates = load_ui_templates(PROJECT_ROOT)
    profile = default_ui_profile(templates)
    bundle = build_runtime_bundle(
        source_file=SOURCE_WORKBOOK,
        runtime_dir=tmp_path / "ui_run",
        templates=templates,
        time_channel_id="track_time__col_001",
        export_channel_ids=profile["export_channels"],
        selected_math_channel_ids=profile["math_channels"],
        report_channel_ids=profile["report_channels"],
        selected_statistic_ids=profile["statistics"],
        kpi_statistic_ids=profile["kpis"],
        selected_plot_ids=profile["plots"],
    )
    result = run_pipeline(bundle.pipeline_config)
    assert result.status == "PASS"
    assert result.completed_stage_count == 7
    assert result.config.duty_cycle is None
    assert next(stage for stage in result.stages if stage.name == "statistics").metrics["samples"] == 1866
    assert result.report_path is not None and result.report_path.exists()
    assert result.powerpoint_path is not None and result.powerpoint_path.exists()


@pytest.mark.skipif(not SOURCE_WORKBOOK.exists(), reason="Client source workbook is not present")
def test_engineering_report_single_file_pipeline_generates_final_reports(tmp_path: Path) -> None:
    bundle = build_engineering_report_runtime_bundle(
        source_file=SOURCE_WORKBOOK,
        runtime_dir=tmp_path / "engineering_report_run",
        project_root=PROJECT_ROOT,
    )

    result = run_pipeline(bundle.pipeline_config)

    assert result.status == "PASS"
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
    assert result.config.duty_cycle is not None
    assert result.report_path is not None and result.report_path.exists()
    assert result.powerpoint_path is not None and result.powerpoint_path.exists()
    duty_stage = next(stage for stage in result.stages if stage.name == "duty_cycle")
    assert duty_stage.metrics["samples"] == 17418
    assert duty_stage.metrics["final_distance_km"] == pytest.approx(114.0011)
    assert duty_stage.metrics["final_soc_pct"] == pytest.approx(23.9383)
    assert duty_stage.metrics["final_fuel_kg"] == pytest.approx(39.84212)
    assert duty_stage.metrics["max_generator_kw"] == pytest.approx(80.02669061662199)
    assert next(stage for stage in result.stages if stage.name == "statistics").metrics["samples"] == 17418


def test_math_availability_respects_source_dependencies() -> None:
    from vsm_postprocessing.ui_config import available_math_channel_ids

    templates = load_ui_templates(PROJECT_ROOT)
    minimal_source = {
        "track_time__col_001",
        "vehicle_distance__col_006",
    }
    available = available_math_channel_ids(templates.math_channels, minimal_source)
    assert "calc_time_minutes" in available
    assert "calc_distance_km" in available
    assert "calc_engine_power_required" not in available
    assert "calc_total_generator_power" not in available
