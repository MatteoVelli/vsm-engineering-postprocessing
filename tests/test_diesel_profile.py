from __future__ import annotations

import csv
import math
from dataclasses import replace
from pathlib import Path

import numpy as np
import pytest

from conftest import DIESEL_REFERENCE_CSV, DIESEL_REFERENCE_DESCRIPTION, PROJECT_ROOT, require_private_reference_file
from vsm_postprocessing.errors import ConfigurationError, DataValidationError, FileImportError, MathChannelError
from vsm_postprocessing.importer import load_data_file
from vsm_postprocessing.math_engine import _cumulative_trapezoid, _time_average
from vsm_postprocessing.profile_math import calculate_profile_math_channels
from vsm_postprocessing.profile_statistics import calculate_profile_statistics
from vsm_postprocessing.report_metadata import resolve_report_metadata
from vsm_postprocessing.report_profile import load_reporting_profile, resolve_profile
from vsm_postprocessing.ui_config import discover_reporting_profiles, generate_reporting_profile_engineering_report, validate_reporting_profile_source


PROFILE_PATH = PROJECT_ROOT / "config/report_profiles/full_size_sprayer_diesel.yaml"


@pytest.fixture(scope="module")
def profile():
    return load_reporting_profile(PROFILE_PATH)


@pytest.fixture(scope="module")
def reference():
    return load_data_file(require_private_reference_file(DIESEL_REFERENCE_CSV, DIESEL_REFERENCE_DESCRIPTION))


def _source(tmp_path, profile, *, omitted=(), extra_electric=False):
    definitions = [c for c in profile.raw_channels if c.semantic_name not in omitted]
    signals = {
        "track_time": [10, 11, 13],
        "track_distance": [100, 101, 105],
        "chassis_speed": [0, 10, 20],
        "track_gradient": [-2, 0, 5],
        "engine_speed": [0, 600, 1200],
        "engine_torque": [-30, 100, -100],
        "engine_fuel_consumption": [10, 10.1, 10.3],
        "fuel_flow": [2, 4, 8],
    }
    names = [c.source_name for c in definitions]
    units = [c.unit for c in definitions]
    columns = [signals.get(c.semantic_name, [1, 2, 3]) for c in definitions]
    if extra_electric:
        names.append("ElectricSystem_Battery_Power")
        units.append("kW")
        columns.append([0, 0, 0])
    path = tmp_path / "synthetic.csv"
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(names)
        writer.writerow(units)
        writer.writerows(zip(*columns))
    return path


def _values(dataset, name):
    index = next(i for i, c in enumerate(dataset.channels) if c.source_name == name)
    return dataset.values[:, index]


def test_diesel_profile_is_independent_and_discovered_last(profile):
    assert profile.profile_id == "full_size_sprayer_diesel"
    assert profile.metadata.powertrain == "diesel"
    assert profile.metadata.extends is None
    assert len(profile.plots) == 18 and len(profile.presentation.slides) == 10
    assert sum(c.required for c in profile.raw_channels) == 7
    assert not any("battery" in c.source_name.lower() or "electricsystem_em" in c.source_name.lower() for c in profile.raw_channels)
    assert not profile.raw_by_semantic_name()["track_height"].required
    assert profile.raw_by_semantic_name()["engine_load"].unit == "Nm"
    discovered = discover_reporting_profiles(PROJECT_ROOT)
    assert [p.profile_id for p in discovered[:2]] == ["robosprayer_electric", "robosprayer_hybrid"]
    assert discovered[-1].display_name == "Diesel"
    assert profile.metadata.name == "Full Size Sprayer Diesel"
    metadata = resolve_report_metadata(DIESEL_REFERENCE_CSV, discovered[-1])
    assert metadata.machine_name == "Full Size Sprayer"
    assert metadata.report_title == "Full Size Sprayer Diesel"


@pytest.mark.parametrize("missing", ["track_time", "track_distance", "chassis_speed", "track_gradient", "engine_speed", "engine_torque", "engine_fuel_consumption"])
def test_missing_required_channel_fails(tmp_path, profile, missing):
    source = _source(tmp_path, profile, omitted=[missing])
    if missing == "track_time":
        # The strict importer itself requires a time channel.
        with pytest.raises(FileImportError, match="unique time channel"):
            load_data_file(source)
        return
    dataset = load_data_file(source)
    resolution = resolve_profile(dataset, profile)
    assert not resolution.is_valid
    assert [c.definition.semantic_name for c in resolution.missing_required] == [missing]
    summary = validate_reporting_profile_source(source, PROFILE_PATH)
    assert not summary.is_valid and summary.missing_required_count == 1
    with pytest.raises(DataValidationError, match=profile.raw_by_semantic_name()[missing].source_name):
        calculate_profile_statistics(dataset, profile)


def test_all_optional_channels_may_be_absent(tmp_path, profile):
    optional = [c.semantic_name for c in profile.raw_channels if not c.required]
    source = _source(tmp_path, profile, omitted=optional, extra_electric=True)
    dataset = load_data_file(source)
    result = calculate_profile_statistics(dataset, profile)
    assert result.resolution.is_valid and result.is_complete and result.math_result.is_complete
    assert len(result.resolution.missing_optional) == 41
    assert "fuel_volume" in [c.definition.semantic_name for c in result.math_result.unavailable_optional]
    assert {"time_minutes", "distance_km", "engine_energy_delivered"} <= result.math_result.values_by_semantic_name.keys()
    assert [c.definition.kpi_id for c in result.unavailable_optional_kpis] == ["fuel_flow_average"]
    assert all("battery" not in m.metric_id for m in result.canonical_metrics)
    assert validate_reporting_profile_source(source, PROFILE_PATH).is_valid


def test_signed_power_statistics_offsets_and_irregular_elapsed_time(tmp_path, profile):
    dataset = load_data_file(_source(tmp_path, profile))
    result = calculate_profile_statistics(dataset, profile)
    power = result.math_result.values_by_semantic_name["engine_power"]
    np.testing.assert_allclose(power, [0, 2 * math.pi, -4 * math.pi], rtol=1e-14)
    # 1-second then 2-second trapezoids; no invented interval before t=10.
    np.testing.assert_allclose(result.math_result.values_by_semantic_name["engine_mechanical_energy"], [0, math.pi / 3600, -math.pi / 3600])
    stats = {s.definition.statistic_id: s.value for s in result.statistics}
    assert stats["engine_power_rms"] == pytest.approx(math.sqrt(20 * math.pi**2 / 3))
    assert stats["engine_power_min"] == pytest.approx(-4 * math.pi)
    assert stats["engine_power_max"] == pytest.approx(2 * math.pi)
    assert stats["fuel_volume_last"] == pytest.approx(15 / 3600)
    kpis = {k.definition.kpi_id: k.value for k in result.kpis}
    assert kpis == pytest.approx(dict(drive_cycle_minutes=0.05, distance_km=0.005, fuel_consumed_kg=0.3, engine_speed_average=700, fuel_flow_average=5))
    assert result.is_complete and not result.diagnostics
    assert all(s.used_sample_count == 3 and s.omitted_sample_count == 0 for s in result.statistics)
    # Existing Hybrid conversion constant must not affect the exact Diesel formula.
    other = calculate_profile_math_channels(dataset, profile, constants={"rpm_nm_to_kw_divisor": 1})
    np.testing.assert_array_equal(other.values_by_semantic_name["engine_power"], power)


@pytest.mark.parametrize("name,unit", [("Engine_Speed", "rad/s"), ("Engine_Torque", "kNm"), ("Engine_Load", "%"), ("Engine_FuelConsumption_absolut", "l")])
def test_wrong_units_are_rejected(tmp_path, profile, name, unit):
    dataset = load_data_file(_source(tmp_path, profile))
    dataset.channels = [replace(c, unit=unit) if c.source_name == name else c for c in dataset.channels]
    assert not resolve_profile(dataset, profile).is_valid
    with pytest.raises(DataValidationError, match="unit_mismatches"):
        calculate_profile_math_channels(dataset, profile)


@pytest.mark.parametrize("name,values", [
    ("Engine_FuelConsumption_absolut", [10, 11, 1]),
    ("Engine_FuelConsumption_absolut", [-1, 0, 1]),
    ("Track_Distance", [100, 101, 99]),
    ("Engine_FuelConsumption_volumeflow", [0, -1, 2]),
    ("Engine_Speed", [0, -1, 2]),
    ("Engine_Torque", [0, np.nan, 2]),
    ("Engine_Oil_Temperature", [0, np.inf, 2]),
    ("Track_Time", [0, 0, 1]),
])
def test_unsafe_diesel_values_are_rejected(tmp_path, profile, name, values):
    dataset = load_data_file(_source(tmp_path, profile))
    _values(dataset, name)[:] = values
    with pytest.raises(DataValidationError, match="Diesel"):
        calculate_profile_statistics(dataset, profile)


def test_normalized_engine_names_resolve_without_alias_guessing(tmp_path, profile):
    dataset = load_data_file(_source(tmp_path, profile))
    dataset.channels = [replace(c, source_name="Engine Speed") if c.source_name == "Engine_Speed" else c for c in dataset.channels]
    assert resolve_profile(dataset, profile).resolved["engine_speed"].match_type == "normalized"
    assert calculate_profile_statistics(dataset, profile).is_complete


@pytest.mark.parametrize("values,time", [([1], [0]), ([1, 2], [0, 0]), ([1, 2], [2, 1]), ([1, np.nan], [0, 1]), ([1, 2], [0, np.inf]), ([1, 2], [0, 1, 2]), ([[1, 2]], [[0, 1]])])
def test_elapsed_integral_rejects_invalid_input(values, time):
    with pytest.raises(MathChannelError):
        _cumulative_trapezoid(values, time)


def test_elapsed_integral_and_average_are_invariant_to_time_origin():
    for time in ([0, 1, 4], [100, 101, 104]):
        np.testing.assert_allclose(_cumulative_trapezoid([2, -2, 6], time), [0, 0, 6])
        assert _time_average([2, -2, 6], time) == pytest.approx(1.5)


def test_diesel_combined_reports_reuse_the_same_excel_result(tmp_path, profile):
    result = generate_reporting_profile_engineering_report(_source(tmp_path, profile), PROFILE_PATH, tmp_path / "report")
    assert result.powerpoint_result.excel_result is result.excel_result
    assert result.report_path.exists() and result.presentation_path.exists()


def test_diesel_ui_offers_combined_reports_and_does_not_show_stale_downloads():
    from vsm_postprocessing.ui_app import _render_profile_engineering_report_workflow

    class UI:
        session_state = {"profile_report_result": object()}

        def __init__(self):
            self.buttons = []
            self.messages = []

        def write(self, message):
            pass

        def info(self, message):
            self.messages.append(message)

        def file_uploader(self, *args, **kwargs):
            return None

        def selectbox(self, label, options, **kwargs):
            assert "full_size_sprayer_diesel" in options
            assert kwargs["format_func"]("full_size_sprayer_diesel") == "Diesel"
            return "full_size_sprayer_diesel"

        def button(self, label, **kwargs):
            self.buttons.append(label)
            return False

    ui = UI()
    _render_profile_engineering_report_workflow(ui)
    assert ui.buttons == ["STEP 3 - Validate", "STEP 4 - Generate Engineering Report"]
    assert any("Excel and PowerPoint" in message for message in ui.messages)


@pytest.mark.private_reference
def test_reference_import_resolution_and_profile_mismatch(reference, profile):
    assert reference.source_path.name == DIESEL_REFERENCE_CSV.name
    quality = reference.quality
    assert (quality.channel_count, quality.sample_count) == (586, 223)
    assert (quality.header_row, quality.unit_row, quality.data_start_row) == (1, 2, 3)
    assert quality.is_valid and quality.nominal_time_step == 1
    assert quality.source_sha256 == "e089fd94a9d6c6657982198c2a7106ac5cd36e0431cf5fb870c9565bb764d7b1"
    resolution = resolve_profile(reference, profile)
    assert resolution.is_valid and len(resolution.resolved) == 47
    assert [x.definition.semantic_name for x in resolution.missing_optional] == ["track_height"]
    electric = [i for i, c in enumerate(reference.channels) if c.source_name.startswith("ElectricSystem_")]
    assert len(electric) == 38 and np.all(reference.values[:, electric] == 0)
    for name in ["robosprayer_electric", "robosprayer_hybrid"]:
        other = load_reporting_profile(PROFILE_PATH.with_name(name + ".yaml"))
        result = resolve_profile(reference, other)
        assert not result.is_valid and len(result.missing_required) == 34
    summary = validate_reporting_profile_source(reference.source_path, PROFILE_PATH)
    assert summary.is_valid
    assert summary.profile_name == "Diesel"
    assert summary.missing_optional_names == ("Road Height",)


@pytest.mark.private_reference
def test_reference_engine_power_independent_csv_calculation(reference, profile):
    # Read original rows independently of the importer and vectorized evaluator.
    with reference.source_path.open(encoding="cp1252", newline="") as handle:
        reader = csv.reader(handle)
        names = next(reader)
        next(reader)
        rows = list(reader)
    independent = [float(row[names.index("Engine_Torque")]) * float(row[names.index("Engine_Speed")]) * math.tau / 60000 for row in rows]
    result = calculate_profile_math_channels(reference, profile)
    np.testing.assert_allclose(result.values_by_semantic_name["engine_power"], independent, rtol=2e-15)
    assert max(independent) == pytest.approx(165.93862633964403)
    assert independent.index(max(independent)) == 145
    assert sum(p < 0 for p in independent) == 12
    for index, expected in [(0, -0.0447520857147), (1, 2.06148871147884), (50, 65.5488921183535), (151, -8.23009219414374), (222, 11.2353108575304)]:
        assert independent[index] == pytest.approx(expected, rel=1e-12)


@pytest.mark.private_reference
def test_reference_fuel_semantics_statistics_and_kpis(reference, profile):
    result = calculate_profile_statistics(reference, profile)
    assert result.is_complete and result.math_result.is_complete
    assert len(result.statistics) == 51 and len(result.kpis) == 5
    stats = {s.definition.statistic_id: s.value for s in result.statistics}
    for key, value in dict(chassis_speed_max=22.4686, road_gradient_min=-15, road_gradient_max=15, engine_speed_max=1996.81, engine_torque_max=826.219, engine_power_max=165.93862633964403, engine_mechanical_energy_last=3.7647817811605133, fuel_volume_last=0.9421207194583333, fuel_flow_max=41.4741).items():
        assert stats[key] == pytest.approx(value, rel=1e-11)
    kpis = {k.definition.kpi_id: k.value for k in result.kpis}
    assert kpis == pytest.approx(dict(drive_cycle_minutes=3.7, distance_km=0.999242, fuel_consumed_kg=0.772836, engine_speed_average=1537.2141644144144, fuel_flow_average=15.277633288513513))
    time = _values(reference, "Track_Time")
    fuel = _values(reference, "Engine_FuelConsumption_absolut")
    flow = _values(reference, "Engine_FuelConsumption_volumeflow")
    assert fuel[0] == 0 and np.all(np.diff(fuel) >= 0)
    assert np.all(flow > 0) and np.isfinite(fuel).all()
    assert stats["fuel_volume_last"] == pytest.approx(np.trapezoid(flow, time) / 3600)
    assert stats["engine_mechanical_energy_last"] == pytest.approx(np.trapezoid(result.math_result.values_by_semantic_name["engine_power"], time) / 3600)
    assert {s.definition.target for s in result.unavailable_optional_statistics} == {"track_height"}
