from __future__ import annotations

from pathlib import Path

import pytest
import yaml

from vsm_postprocessing.errors import ConfigurationError
from vsm_postprocessing.pipeline_engine import load_pipeline_config, run_pipeline
from vsm_postprocessing.ui_config import (
    build_runtime_bundle,
    default_ui_profile,
    load_ui_profile,
    load_ui_templates,
    save_ui_profile,
)

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SOURCE_WORKBOOK = PROJECT_ROOT / "reference_files" / "Sprayer_Caiman_SP_9300Kg_Hybrid_Gen80kW_30kph_74Ht_4000KgAQ_57-4pcSOC_5-80_1C2G_02.xlsx"


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
    assert config.output_root == bundle.output_root


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
    assert result.report_path is not None and result.report_path.exists()
    assert result.powerpoint_path is not None and result.powerpoint_path.exists()


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
