"""Keep active gearbox measurements separate from inactive legacy shaft signals."""
import csv
from dataclasses import replace
from pathlib import Path

import numpy as np
from openpyxl import load_workbook
from pptx import Presentation
import pytest

from conftest import REFERENCE_FILES_DIR, require_private_reference_file
from test_profile_math import _channel
from test_sergio_october_15 import csv_for, dataset_for, profile_for, profile_path, remove_channel
from vsm_postprocessing.errors import DataValidationError
from vsm_postprocessing.excel_report_engine import generate_profile_excel_report, _template_channel_order_entries
from vsm_postprocessing.importer import ImportOptions, load_data_file
from vsm_postprocessing.profile_math import calculate_profile_math_channels
from vsm_postprocessing.profile_statistics import calculate_profile_statistics
from vsm_postprocessing.profile_powerpoint_report_engine import (
    build_profile_powerpoint_report, _profile_statistics_as_powerpoint_statistics,
)
from vsm_postprocessing.report_profile import StatisticDefinition, resolve_profile

WHEELS = ("fl", "fr", "rl", "rr")


def competing_sources(kind):
    data = dataset_for(kind)
    for i, wheel in enumerate(WHEELS, 1):
        data.values[:, data.channel_index("driveshaft_torque_" + wheel)] = np.array([0, 7000, -8000, 8700]) * i
        data.values[:, data.channel_index("wheel_rotationalspeed_" + wheel)] = np.array([0, 20, 30, -10]) * i
        source = "DriveShaft_Torque_" + wheel.upper()
        if not any(c.source_name == source for c in data.channels):
            data.channels.append(_channel("legacy_" + wheel, source, "Nm", len(data.channels) + 1))
            data.values = np.column_stack((data.values, np.zeros(len(data.values))))
        else:
            c = next(c for c in data.channels if c.source_name == source)
            data.values[:, data.channel_index(c.channel_id)] = 0
    return data


@pytest.mark.parametrize("kind", ("electric", "hybrid"))
def test_four_corners_select_gearbox_even_when_zero_shafts_are_present(kind):
    data, profile = competing_sources(kind), profile_for(kind)
    before = data.values.copy()
    result = calculate_profile_math_channels(data, profile)
    total_torque, total_power = np.zeros(4), np.zeros(4)
    for i, wheel in enumerate(WHEELS, 1):
        source = result.resolution.resolved["driveshaft_torque_" + wheel]
        assert source.channel.source_name == f"KERS_MotGenUnit_GearboxOutputTorque_{i}"
        assert source.match_type == "exact" and source.definition.channel_type == "VSM"
        assert source.channel.kind == "raw"
        torque = np.array([0, 7000, -8000, 8700]) * i
        power = torque * np.array([0, 20, 30, -10]) * i * np.pi / 30000
        np.testing.assert_allclose(result.values_by_semantic_name["wheel_power_" + wheel], power, rtol=2e-15)
        total_torque += torque
        total_power += power
    np.testing.assert_allclose(result.values_by_semantic_name["wheel_total_torque"], total_torque)
    np.testing.assert_allclose(result.values_by_semantic_name["wheel_power_total"], total_power, rtol=2e-15)
    np.testing.assert_array_equal(data.values, before)


@pytest.mark.parametrize("kind", ("electric", "hybrid"))
@pytest.mark.parametrize("wheel", WHEELS)
@pytest.mark.parametrize("problem", ("missing", "nan", "unit", "duplicate"))
def test_unusable_gearbox_never_uses_existing_legacy_shaft(kind, wheel, problem):
    data, profile = competing_sources(kind), profile_for(kind)
    name = "driveshaft_torque_" + wheel
    index = data.channel_index(name)
    if problem == "missing":
        remove_channel(data, name)
    elif problem == "nan":
        data.values[1, index] = np.nan
    elif problem == "unit":
        data.channels[index] = replace(data.channels[index], unit="kNm")
    else:
        data.channels.append(replace(data.channels[index], channel_id="duplicate"))
        data.values = np.column_stack((data.values, data.values[:, index]))
    resolution = resolve_profile(data, profile)
    assert name not in resolution.resolved
    if problem == "duplicate":
        assert resolution.ambiguous and not resolution.is_valid
    else:
        assert any(m.definition.semantic_name == name for m in resolution.missing_optional)
    result = calculate_profile_math_channels(data, profile, resolution)
    assert "wheel_power_" + wheel not in result.values_by_semantic_name
    assert "wheel_total_torque" not in result.values_by_semantic_name
    assert "wheel_power_total" not in result.values_by_semantic_name


@pytest.mark.parametrize("kind", ("electric", "hybrid"))
def test_zero_gearbox_is_never_replaced_by_nonzero_shafts(kind):
    data = competing_sources(kind)
    for wheel in WHEELS:
        data.values[:, data.channel_index("driveshaft_torque_" + wheel)] = 0
        c = next(c for c in data.channels if c.source_name == "DriveShaft_Torque_" + wheel.upper())
        data.values[:, data.channel_index(c.channel_id)] = 1234
    result = calculate_profile_math_channels(data, profile_for(kind))
    np.testing.assert_array_equal(result.values_by_semantic_name["wheel_power_total"], 0)
    assert any("All four measured" in msg for msg in result.resolution.diagnostics)


@pytest.mark.parametrize("kind", ("electric", "hybrid"))
def test_signed_wheel_statistics_include_independent_rms_max_min_last(kind):
    profile, data = profile_for(kind), competing_sources(kind)
    definitions = tuple(StatisticDefinition(f"audit_{name}_{op}", name, op)
                        for wheel in WHEELS for name in ("driveshaft_torque_" + wheel, "wheel_power_" + wheel)
                        for op in ("rms", "max", "min", "last"))
    result = calculate_profile_statistics(data, replace(profile, statistics=(*profile.statistics, *definitions)))
    for statistic in result.statistics:
        d = statistic.definition
        if not d.statistic_id.startswith("audit_"):
            continue
        wheel = d.target[-2:]
        i = WHEELS.index(wheel) + 1
        expected = np.array([0, 7000, -8000, 8700]) * i
        if d.target.startswith("wheel_power"):
            expected = expected * np.array([0, 20, 30, -10]) * i * np.pi / 30000
        want = {"rms": np.sqrt(np.mean(expected ** 2)), "max": expected.max(),
                "min": expected.min(), "last": expected[-1]}[d.operation]
        assert statistic.value == pytest.approx(want, rel=2e-15, abs=1e-12)


@pytest.mark.parametrize("kind", ("electric", "hybrid"))
def test_nonnumeric_torque_is_rejected_or_diagnosed_without_zero_fill(kind, tmp_path):
    source = csv_for(competing_sources(kind), tmp_path / "bad.csv")
    with source.open(newline="") as stream:
        rows = list(csv.reader(stream))
    rows[3][rows[0].index("KERS_MotGenUnit_GearboxOutputTorque_4")] = "invalid"
    with source.open("w", newline="") as stream:
        csv.writer(stream).writerows(rows)
    with pytest.raises(DataValidationError):
        load_data_file(source)
    resolution = resolve_profile(load_data_file(source, ImportOptions(strict=False)), profile_for(kind))
    assert "driveshaft_torque_rr" not in resolution.resolved
    assert any("nonnumeric" in msg for msg in resolution.diagnostics)


@pytest.mark.parametrize("kind", ("electric", "hybrid"))
def test_corner_order_is_supported_by_client_workbook(kind):
    entries = _template_channel_order_entries(Path(profile_for(kind).metadata.excel_template))
    for i, wheel in enumerate(WHEELS, 1):
        assert any(e["source_name"] == f"KERS_MotGenUnit_GearboxOutputTorque_{i}"
                   and e["report_name"] == f"Wheel Torque {wheel.upper()}" for e in entries)


@pytest.mark.parametrize("kind", ("electric", "hybrid"))
def test_upload_defaults_to_current_mapping_and_legacy_choice_warns(kind, tmp_path, monkeypatch):
    from types import SimpleNamespace
    import streamlit as st
    from streamlit.testing.v1 import AppTest
    from vsm_postprocessing import ui_app
    source = csv_for(competing_sources(kind), tmp_path / f"measured_{kind}.csv")
    monkeypatch.setattr(st, "file_uploader", lambda *a, **kw: SimpleNamespace(name=source.name, getvalue=source.read_bytes))
    monkeypatch.setattr(ui_app, "UI_WORKSPACE", tmp_path / "workspace")
    app = AppTest.from_string("from vsm_postprocessing.ui_app import main\nmain()", default_timeout=60).run()
    assert not app.exception
    revision = next(s for s in app.selectbox if s.label == "Reporting definitions")
    assert revision.value == "Sergio - 15 October 2026"
    next(b for b in app.button if b.label.startswith("STEP 3")).click().run()
    assert not app.exception and not app.error
    assert app.session_state["profile_validation_summary"].is_valid
    next(s for s in app.selectbox if s.label == "Reporting definitions").set_value("Legacy - Electric_05 / Hybrid_06 / Diesel").run()
    assert any("Legacy wheel torque uses DriveShaft" in warning.value for warning in app.warning)
    assert "profile_validation_summary" not in app.session_state


@pytest.mark.parametrize("kind", ("electric", "hybrid"))
def test_current_rr_cards_reuse_excel_statistics_without_recalculation(kind, tmp_path, monkeypatch):
    data = competing_sources(kind)
    report = generate_profile_excel_report(csv_for(data, tmp_path / "measured.csv"), profile_path(kind), tmp_path / "excel")
    shared = {s.definition.statistic_id: s for s in report.statistics_result.statistics}
    assert {"wheel_power_rr_max", "edu_mech_power_rr_max"} <= shared.keys()
    import vsm_postprocessing.profile_powerpoint_report_engine as engine
    def unexpected(*args, **kwargs):
        pytest.fail("PowerPoint must reuse configured shared RR statistics")
    with monkeypatch.context() as patch:
        patch.setattr(engine, "compute_statistic", unexpected)
        metrics = _profile_statistics_as_powerpoint_statistics(report).statistics
        excluded = {"wheel_power_rr_max", "edu_mech_power_rr_max"}
        selected_profile = replace(report.profile, statistics=tuple(s for s in report.profile.statistics if s.statistic_id not in excluded))
        selected_statistics = replace(report.statistics_result, profile=selected_profile,
                                      statistics=[s for s in report.statistics_result.statistics if s.definition.statistic_id not in excluded])
        selected_report = replace(report, profile=selected_profile, statistics_result=selected_statistics)
        selected_metrics = _profile_statistics_as_powerpoint_statistics(selected_report).statistics
        assert not any(m.statistic_id in excluded for m in selected_metrics)
    for name in ("wheel_power_rr_max", "edu_mech_power_rr_max"):
        metric, = [m for m in metrics if m.statistic_id == name]
        assert metric.value == shared[name].value
    ppt = build_profile_powerpoint_report(report, tmp_path / "ppt")
    book = load_workbook(report.report_path, data_only=True, read_only=True)
    try:
        rows = {r[0]: r for r in book["Statistics"].iter_rows(values_only=True)}
        for name in ("wheel_power_rr_max", "edu_mech_power_rr_max"):
            assert rows[name][4] == pytest.approx(shared[name].value, rel=2e-15)
    finally:
        book.close()
    deck = Presentation(ppt.presentation_path)
    text = "\n".join(s.text for slide in deck.slides for s in slide.shapes if s.has_text_frame)
    assert "RR WHEEL MAX POWER" in text and "RR EDU MAX POWER" in text


@pytest.mark.private_reference
def test_real_inactive_shafts_do_not_hide_active_electric_gearbox_outputs():
    path = require_private_reference_file(REFERENCE_FILES_DIR / "RoboSprayer_8500Kg_Electric_02_UoL_Drive_Cycle_20kph_DriverPC_Flat_02.csv", "Electric active gearbox / inactive driveshaft regression")
    data = load_data_file(path)
    result = calculate_profile_math_channels(data, profile_for("electric"))
    for i, wheel in enumerate(WHEELS, 1):
        shaft = next(c for c in data.channels if c.source_name == "DriveShaft_Torque_" + wheel.upper())
        assert not np.any(data.values[:, data.channel_index(shaft.channel_id)])
        gearbox = result.resolution.resolved["driveshaft_torque_" + wheel].channel
        torque = data.values[:, data.channel_index(gearbox.channel_id)]
        speed = next(c for c in data.channels if c.source_name == "Wheel_RotationalSpeed_" + wheel.upper())
        expected = torque * data.values[:, data.channel_index(speed.channel_id)] * np.pi / 30000
        np.testing.assert_allclose(result.values_by_semantic_name["wheel_power_" + wheel], expected, rtol=2e-15, atol=1e-12)
