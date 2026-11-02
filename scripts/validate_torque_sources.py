"""Independently audit real CSV torque sources through Excel and PowerPoint.

Run with .venv/Scripts/python scripts/validate_torque_sources.py --output NEW_DIR.
Only disposable evidence is written; source files and previous reports are read only.
"""
from __future__ import annotations

import argparse
import csv
import hashlib
import json
from pathlib import Path
import sys
from types import SimpleNamespace
from unittest.mock import patch
from zipfile import ZipFile, ZIP_DEFLATED
from xml.etree import ElementTree as ET

import numpy as np
from openpyxl import load_workbook
from pptx import Presentation

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
from vsm_postprocessing import ui_app
from vsm_postprocessing.ui_config import generate_reporting_profile_engineering_report

WHEELS = ("fl", "fr", "rl", "rr")
REL_TOL, ABS_TOL = 2e-14, 1e-10
NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def raw_csv(source):
    # Deliberately independent of the production importer and expression engine.
    with source.open(encoding="utf-8-sig", newline="") as stream:
        reader = csv.reader(stream)
        names, units = next(reader), next(reader)
        assert len(names) == len(set(names)), "Ambiguous raw CSV headers"
        values = np.asarray([[float(x) for x in row] for row in reader], dtype=float)
    return dict(zip(names, values.T)), dict(zip(names, units))


def ui_report(source, destination):
    import streamlit as st
    from streamlit.testing.v1 import AppTest
    results = []
    original = ui_app.generate_reporting_profile_engineering_report
    def generate(*args, **kwargs):
        result = original(*args, **kwargs)
        results.append(result)
        return result
    with patch.object(st, "file_uploader", return_value=SimpleNamespace(name=source.name, getvalue=source.read_bytes)), \
         patch.object(ui_app, "UI_WORKSPACE", destination / "workspace"), \
         patch.object(ui_app, "UI_RUNS", destination / "runs"), \
         patch.object(ui_app, "generate_reporting_profile_engineering_report", generate):
        app = AppTest.from_string("from vsm_postprocessing.ui_app import main\nmain()", default_timeout=600).run()
        assert not app.exception
        revision = next(s for s in app.selectbox if s.label == "Reporting definitions")
        assert revision.value == "Sergio - 15 October 2026"
        next(b for b in app.button if b.label.startswith("STEP 3")).click().run()
        assert not app.exception and not app.error
        next(b for b in app.button if b.label.startswith("STEP 4")).click().run()
        assert not app.exception and not app.error, [e.value for e in app.error]
        assert len(results) == 1
        # Switching revisions invalidates the previous validated result and warns.
        next(s for s in app.selectbox if s.label == "Reporting definitions").set_value("Legacy - Electric_05 / Hybrid_06 / Diesel").run()
        assert any("Legacy wheel torque uses DriveShaft" in w.value for w in app.warning)
        assert "profile_report_result" not in app.session_state
    return results[0]


def clear_caches(source, destination):
    count = 0
    with ZipFile(source) as before, ZipFile(destination, "w", ZIP_DEFLATED) as after:
        for item in before.infolist():
            data = before.read(item.filename)
            if item.filename.startswith("xl/worksheets/sheet") and item.filename.endswith(".xml"):
                root = ET.fromstring(data)
                for cell in root.findall(".//m:c", NS):
                    if cell.find("m:f", NS) is not None:
                        value = cell.find("m:v", NS)
                        if value is not None:
                            cell.remove(value)
                            count += 1
                data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            after.writestr(item, data)
    assert count > 0
    return count


def verify_office(checks_path):
    records = json.loads(checks_path.read_text(encoding="utf-8-sig"))
    if isinstance(records, dict):
        records = [records]
    evidence = []
    for record in records:
        count, maximum_error = 0, 0.0
        with ZipFile(record["empty_cache_copy"]) as archive:
            for part in archive.namelist():
                if part.startswith("xl/worksheets/sheet") and part.endswith(".xml"):
                    for cell in ET.fromstring(archive.read(part)).findall(".//m:c", NS):
                        if cell.find("m:f", NS) is not None:
                            assert cell.find("m:v", NS) is None, "Formula cache was not empty"
        before = load_workbook(record["original_excel"], read_only=True, data_only=True)
        after = load_workbook(record["recalculated_excel"], read_only=True, data_only=True)
        try:
            assert before.sheetnames == after.sheetnames
            for original in before:
                recalculated = after[original.title]
                for old_row, new_row in zip(original.iter_rows(), recalculated.iter_rows(max_row=original.max_row, max_col=original.max_column)):
                    for old, new in zip(old_row, new_row):
                        assert new.data_type != "e", (original.title, new.coordinate, new.value)
                        if isinstance(old.value, (int, float)):
                            assert isinstance(new.value, (int, float)), (original.title, old.coordinate)
                            np.testing.assert_allclose(new.value, old.value, rtol=REL_TOL, atol=ABS_TOL,
                                                       err_msg=f"{original.title}!{old.coordinate}")
                            count += 1
                            maximum_error = max(maximum_error, abs(new.value - old.value))
        finally:
            before.close()
            after.close()
        evidence.append({"source": record["source"], "numeric_cells": count,
                         "max_absolute_error": maximum_error, "status": "PASS"})
        print("PASS Office recalculation", record["source"], count, maximum_error, flush=True)
    checks_path.with_name("numerical_verification.json").write_text(json.dumps(evidence, indent=2), encoding="utf-8")


def inspect(result, source, destination):
    raw, units = raw_csv(source)
    report = result.excel_result
    observed = report.plotting_result.values_by_semantic_name
    expected, checks, peaks = {}, {}, {}
    for i, wheel in enumerate(WHEELS, 1):
        torque = f"KERS_MotGenUnit_GearboxOutputTorque_{i}"
        speed = f"Wheel_RotationalSpeed_{wheel.upper()}"
        if torque in raw:
            assert units[torque] == "Nm" and units[speed] == "rpm"
            resolved = report.resolution.resolved["driveshaft_torque_" + wheel]
            assert resolved.channel.source_name == torque and resolved.match_type == "exact"
            expected["driveshaft_torque_" + wheel] = raw[torque]
            power = raw[torque] * raw[speed] * np.pi / 30000
            expected["wheel_power_" + wheel] = power
            peak = int(np.argmax(power))
            peaks[wheel] = {"sample_zero_based": peak, "timestamp_s": float(raw["Track_Time"][peak]),
                            "torque_Nm": float(raw[torque][peak]), "speed_rpm": float(raw[speed][peak]),
                            "power_kW": float(power[peak])}
        else:
            assert "driveshaft_torque_" + wheel not in observed
            assert "wheel_power_" + wheel not in observed
        motor_speed, motor_torque = f"ElectricSystem_EM{i}_Speed", f"ElectricSystem_EM{i}_Torque"
        expected["edu_mech_power_" + wheel] = raw[motor_speed] * raw[motor_torque] * np.pi / 30000
    if len(peaks) == 4:
        expected["wheel_total_torque"] = sum(expected["driveshaft_torque_" + w] for w in WHEELS)
        expected["wheel_power_total"] = sum(expected["wheel_power_" + w] for w in WHEELS)
    else:
        assert "wheel_total_torque" not in observed and "wheel_power_total" not in observed
    for name, array in expected.items():
        np.testing.assert_allclose(observed[name], array, rtol=REL_TOL, atol=ABS_TOL)
        checks[name] = {"maximum_absolute_error": float(np.max(np.abs(array - observed[name]))),
                        "max": float(array.max()), "min": float(array.min()),
                        "rms": float(np.sqrt(np.mean(array ** 2))), "last": float(array[-1])}

    # Inspect every exported numerical sample, using independent arrays for all
    # affected channels and engine arrays only for unrelated preservation checks.
    book = load_workbook(report.report_path, data_only=True, read_only=True)
    try:
        sheet = book.worksheets[0]
        n = len(raw["Track_Time"])
        matrix = np.asarray(list(sheet.iter_rows(min_row=5, max_row=n+4, max_col=report.report_channel_count, values_only=True)), dtype=float)
        for index, channel in enumerate(report.report_channels):
            want = expected.get(channel.channel_id, observed[channel.channel_id])
            np.testing.assert_allclose(matrix[:, index], want, rtol=REL_TOL, atol=ABS_TOL)
        bottom = list(sheet.iter_rows(min_row=n+5, max_row=n+8, max_col=report.report_channel_count, values_only=True))
        for index, channel in enumerate(report.report_channels):
            if channel.channel_id in checks:
                for row, operation in zip(bottom[:3], ("max", "min", "last")):
                    np.testing.assert_allclose(row[index], checks[channel.channel_id][operation], rtol=REL_TOL, atol=ABS_TOL)
        statistics = {r[0]: r for r in book["Statistics"].iter_rows(values_only=True) if r[0]}
        # All configured affected stats use the same full-array independent oracle.
        for statistic in report.statistics_result.statistics:
            d = statistic.definition
            if d.target not in checks:
                continue
            want = checks[d.target][d.operation]
            np.testing.assert_allclose(statistic.value, want, rtol=REL_TOL, atol=ABS_TOL)
            np.testing.assert_allclose(statistics[d.statistic_id][4], want, rtol=REL_TOL, atol=ABS_TOL)
    finally:
        book.close()
    formulas = load_workbook(report.report_path)
    try:
        sheet = formulas.worksheets[0]
        mapping = {r[5]: r for r in formulas["Rename From VSM to Astauto"].iter_rows(min_row=3, values_only=True)}
        chart_checks = []
        for chart in sheet._charts:
            text = ET.tostring(chart.to_tree(), encoding="unicode")
            if not any(t in text for t in ("Wheel Power", "Wheel Torque", "Power at Wheels")):
                continue
            references = [s.yVal.numRef.f for c in chart._charts for s in c.series]
            assert references and all("!" in ref for ref in references)
            chart_checks.append({"xml_title_and_units": text[:600], "y_references": references})
        assert chart_checks or not peaks
        for col, channel in enumerate(report.report_channels, 1):
            if channel.channel_id in expected:
                assert mapping[channel.channel_id][0] == channel.source_name
                assert mapping[channel.channel_id][2] == channel.kind.upper()
                assert sheet.cell(4, col).value == channel.unit
    finally:
        formulas.close()

    deck = Presentation(result.presentation_path)
    metrics = {s.statistic_id: s for s in result.powerpoint_result.powerpoint_result.statistics_result.statistics}
    shared = {s.definition.statistic_id: s for s in report.statistics_result.statistics}
    for name in ("wheel_power_rr_max", "edu_mech_power_rr_max", "edu_speed_rl_max", "edu_torque_rl_max"):
        if name == "wheel_power_rr_max" and not peaks:
            assert name not in metrics and name not in shared
        else:
            assert metrics[name].value == shared[name].value
    slide_index = next(i for i,s in enumerate(result.powerpoint_result.powerpoint_result.config.slides) if s.slide_id == "traction_auxiliaries")
    slide = deck.slides[slide_index]
    notes = json.loads(slide.notes_slide.notes_text_frame.text)
    texts = [s.text for s in slide.shapes if s.has_text_frame]
    for name in ("wheel_power_rr_max", "edu_mech_power_rr_max"):
        if name in metrics:
            assert notes["statistics"][name]["value"] == shared[name].value
            assert f"{shared[name].value:,.2f} kW" in texts
    plots = result.powerpoint_result.powerpoint_result.plotting_result
    if peaks:
        plot = next(p for p in plots.rendered_plots if p.plot_id == "power_at_wheels_and_edu")
        assert plot.primary_series_ids == ("wheel_power_rr", "edu_mech_power_rr")
        assert Path(plot.png_file).read_bytes() in [s.image.blob for s in slide.shapes if s.shape_type == 13]

    cache_free = destination / "cache_free.xlsx"
    caches = clear_caches(report.report_path, cache_free)
    return {"source": str(source.resolve()), "source_sha256": digest(source), "kind": report.profile.metadata.powertrain,
            "excel": str(report.report_path), "powerpoint": str(result.presentation_path),
            "cache_cleared_workbook": str(cache_free.resolve()), "profile": report.profile.profile_id,
            "samples": n, "numeric_time_series_cells_checked": int(matrix.size),
            "checks": checks, "peaks": peaks, "native_wheel_charts": chart_checks,
            "formula_caches_removed": caches, "slides": len(deck.slides),
            "rr_wheel_max_power_kW": peaks.get("rr", {}).get("power_kW"),
            "diagnostics": list(report.statistics_result.diagnostics),
            "status": "PASS" if peaks else "PASS_WITH_MISSING_GEARBOX_SOURCES"}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--verify-office", type=Path, help="Independently compare Office checks.json against original numeric results")
    parser.add_argument("--source", type=Path, action="append")
    parser.add_argument("--ui", action="store_true", help="Exercise actual UI upload/validate/generate/rerun")
    args = parser.parse_args()
    if args.verify_office:
        verify_office(args.verify_office)
        return
    if args.output is None:
        parser.error("--output is required for report generation")
    destination = args.output.resolve()
    destination.mkdir(parents=True, exist_ok=False)
    sources = args.source or [p for p in Path("reference_files").glob("*.csv") if "Diesel" not in p.name]
    assert sources, "No real CSV sources found"
    hashes = {str(p): digest(p) for p in Path("reference_files").iterdir() if p.is_file()}
    records = []
    for index, source in enumerate(sources):
        print("Generating", source.name, flush=True)
        output = destination / f"case_{index}"
        output.mkdir()
        kind = "hybrid" if "Hybrid" in source.name else "electric"
        if args.ui:
            result = ui_report(source, output)
        else:
            result = generate_reporting_profile_engineering_report(source, f"config/report_profiles/sergio_2026_10/robosprayer_{kind}.yaml", output)
        record = inspect(result, source, output)
        record["ui_exercised"] = args.ui
        records.append(record)
        (destination / "audit.json").write_text(json.dumps(records, indent=2), encoding="utf-8")
        print(record["status"], record["samples"], "samples; RR wheel max:", record["rr_wheel_max_power_kW"], flush=True)
    assert all(digest(Path(p)) == value for p,value in hashes.items()), "Reference changed"
    (destination / "reference_hashes.json").write_text(json.dumps(hashes, indent=2), encoding="utf-8")


if __name__ == "__main__":
    main()
