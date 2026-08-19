from __future__ import annotations

from pathlib import Path

import pytest
import yaml

from vsm_postprocessing.errors import ConfigurationError
from vsm_postprocessing.pipeline_engine import load_pipeline_config
from vsm_postprocessing.ui_config import (
    build_runtime_bundle,
    default_ui_profile,
    discover_reporting_profiles,
    generate_reporting_profile_excel_report,
    generate_reporting_profile_engineering_report,
    load_ui_profile,
    load_ui_templates,
    save_ui_profile,
    supported_profile_upload_extensions,
    validate_profile_upload_extension,
    validate_reporting_profile_source,
)
from conftest import (
    ROBOSPRAYER_REFERENCE_CSV,
    ROBOSPRAYER_REFERENCE_DESCRIPTION,
    require_private_reference_file,
)

PROJECT_ROOT = Path(__file__).resolve().parents[1]
ELECTRIC_PROFILE = PROJECT_ROOT / "config" / "report_profiles" / "robosprayer_electric.yaml"
HYBRID_PROFILE = PROJECT_ROOT / "config" / "report_profiles" / "robosprayer_hybrid.yaml"


def _robosprayer_csv() -> Path:
    return require_private_reference_file(ROBOSPRAYER_REFERENCE_CSV, ROBOSPRAYER_REFERENCE_DESCRIPTION)


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


def test_profile_driven_upload_formats_are_csv_and_xlsx_only(tmp_path: Path) -> None:
    assert supported_profile_upload_extensions() == (".csv", ".xlsx")
    validate_profile_upload_extension(tmp_path / "source.csv")
    validate_profile_upload_extension(tmp_path / "source.xlsx")
    with pytest.raises(ConfigurationError, match="Unsupported upload type"):
        validate_profile_upload_extension(tmp_path / "source.xlsm")


def test_reporting_profile_discovery_lists_electric_and_hybrid() -> None:
    profiles = discover_reporting_profiles(PROJECT_ROOT)
    by_id = {profile.profile_id: profile for profile in profiles}

    assert by_id["robosprayer_electric"].display_name == "RoboSprayer Electric"
    assert by_id["robosprayer_hybrid"].display_name == "RoboSprayer Hybrid"


def test_electric_reporting_profile_validation_summary_matches_reference_csv() -> None:
    summary = validate_reporting_profile_source(_robosprayer_csv(), ELECTRIC_PROFILE)

    assert summary.profile_name == "RoboSprayer Electric"
    assert summary.is_valid
    assert summary.sample_count == 3853
    assert summary.source_raw_channel_count == 607
    assert summary.required_raw_count == 287
    assert summary.resolved_raw_count == 288
    assert summary.missing_required_count == 0
    assert summary.math_count == 29
    assert summary.statistic_count == 27
    assert summary.kpi_count == 9
    assert summary.plot_count == 12
    assert summary.duration_minutes == pytest.approx(64.2)


def test_hybrid_reporting_profile_validation_treats_inactive_channels_as_resolved() -> None:
    summary = validate_reporting_profile_source(_robosprayer_csv(), HYBRID_PROFILE)

    assert summary.profile_name == "RoboSprayer Hybrid"
    assert summary.is_valid
    assert summary.required_raw_count == 293
    assert summary.resolved_raw_count == 294
    assert summary.missing_required_count == 0
    assert summary.math_count == 32
    assert summary.statistic_count == 36
    assert summary.kpi_count == 9
    assert summary.plot_count == 18
    assert summary.all_zero_resolved_count > 0
    assert any("Engine" in name or "Generator" in name for name in summary.all_zero_names)


def test_reporting_profile_validation_reports_missing_required_channel(tmp_path: Path) -> None:
    csv_path = tmp_path / "missing_required.csv"
    csv_path.write_text("Track_Time,Chassis_Speed\ns,kph\n0,0\n1,1\n", encoding="utf-8")

    summary = validate_reporting_profile_source(csv_path, ELECTRIC_PROFILE)

    assert not summary.is_valid
    assert summary.missing_required_count > 0
    assert "Battery Power" in summary.missing_required_names


def test_profile_report_generation_wrapper_calls_profile_excel_engine(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    calls: dict[str, object] = {}
    expected_result = object()

    def fake_generate(source_file, profile_file, output_dir, import_options):
        calls["source_file"] = source_file
        calls["profile_file"] = profile_file
        calls["output_dir"] = output_dir
        calls["import_options"] = import_options
        return expected_result

    monkeypatch.setattr("vsm_postprocessing.ui_config.generate_profile_excel_report", fake_generate)

    result = generate_reporting_profile_excel_report(
        tmp_path / "source.csv",
        ELECTRIC_PROFILE,
        tmp_path / "out",
    )

    assert result is expected_result
    assert calls["profile_file"] == ELECTRIC_PROFILE
    assert calls["output_dir"] == tmp_path / "out"


def test_profile_engineering_report_wrapper_reuses_excel_result_for_powerpoint(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    calls: dict[str, object] = {}
    expected_excel = object()
    expected_powerpoint = object()

    class FakeProfile:
        profile_id = "fake_profile"

        class metadata:
            name = "Fake Profile"

    def fake_load_profile(profile_file):
        calls["loaded_profile_file"] = profile_file
        return FakeProfile()

    def fake_generate(source_file, profile_file, output_dir, import_options, *, output_filename=None):
        calls["source_file"] = source_file
        calls["profile_file"] = profile_file
        calls["excel_output_dir"] = output_dir
        calls["excel_output_filename"] = output_filename
        calls["import_options"] = import_options
        return expected_excel

    def fake_build(excel_result, output_dir, *, output_filename=None):
        calls["ppt_excel_result"] = excel_result
        calls["ppt_output_dir"] = output_dir
        calls["ppt_output_filename"] = output_filename
        return expected_powerpoint

    monkeypatch.setattr("vsm_postprocessing.ui_config.load_reporting_profile", fake_load_profile)
    monkeypatch.setattr("vsm_postprocessing.ui_config.generate_profile_excel_report", fake_generate)
    monkeypatch.setattr("vsm_postprocessing.ui_config.build_profile_powerpoint_report", fake_build)

    result = generate_reporting_profile_engineering_report(
        tmp_path / "source.csv",
        ELECTRIC_PROFILE,
        tmp_path / "out",
    )

    assert result.excel_result is expected_excel
    assert result.powerpoint_result is expected_powerpoint
    assert calls["ppt_excel_result"] is expected_excel
    assert calls["excel_output_dir"] == tmp_path / "out" / "profile_excel_report"
    assert calls["ppt_output_dir"] == tmp_path / "out" / "profile_powerpoint_report"
    assert calls["excel_output_filename"] == "Fake_Profile_Engineering_Report.xlsx"
    assert calls["ppt_output_filename"] == "Fake_Profile_Engineering_Report.pptx"


def test_ui_app_preserves_profile_flow_and_removes_legacy_flow() -> None:
    source = (PROJECT_ROOT / "src" / "vsm_postprocessing" / "ui_app.py").read_text(encoding="utf-8")

    assert "Custom Analysis" in source
    assert "Engineering Report" in source
    assert "Select Report Profile" in source
    assert "Legacy " + "Cai" + "man" not in source
    assert "duty" + "-cycle" not in source.casefold()
    assert 'type=["xlsx", "xlsm", "csv"]' not in source


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
    raw = yaml.safe_load(bundle.pipeline_config.read_text(encoding="utf-8"))
    assert config.input_file == source.resolve()
    assert "duty" + "_cycle" not in raw
    assert config.output_root == bundle.output_root


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
