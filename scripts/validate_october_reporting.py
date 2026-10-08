"""Generate real-data reports and compare them with pre-change workbooks.

Run --baseline before editing the reporting code, then run without that flag.
Artifacts stay in ignored outputs/october_2026_validation; sources are preserved.
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path
import re
from zipfile import ZipFile

import numpy as np
from openpyxl import load_workbook
from openpyxl.utils.cell import range_to_tuple
from pptx import Presentation

from vsm_postprocessing.excel_report_engine import generate_profile_excel_report
from vsm_postprocessing.profile_powerpoint_report_engine import build_profile_powerpoint_report

ROOT = Path(__file__).resolve().parents[1]
CASES = {
    "full_size_sprayer_diesel": "Full_Size_Sprayer_8500Kg_Diesel_Test_03_Crop_Field_MatLab_20kph_Gradient.csv",
    "robosprayer_hybrid": "Sprayer_Caiman_SP_9300Kg_Hybrid_12x1Km_4000Kg_Chem.csv",
    "robosprayer_electric": "RoboSprayer_3500Kg_Electric_12kph_Batt_50kW_Mot_63RPM_Susp_Cool_Rough_Grad_Discharge.csv",
}


def title(value):
    return "".join(run.t for p in value.tx.rich.p for run in p.r)


def prepare_recalculation(audit_path):
    """Remove every formula cache before independent desktop Excel calculation."""
    audits = json.loads(audit_path.read_text(encoding="utf-8"))
    destination = audit_path.parent / "cache_cleared"
    destination.mkdir(exist_ok=True)
    for audit in audits:
        path = destination / Path(audit["workbook"]).name
        if path.exists():
            raise FileExistsError(path)
        count = 0
        with ZipFile(audit["workbook"]) as source, ZipFile(path, "w") as target:
            for part in source.infolist():
                data = source.read(part.filename)
                if part.filename.startswith("xl/worksheets/") and part.filename.endswith(".xml"):
                    def clear(match):
                        nonlocal count
                        cell = match.group(0)
                        if b"<f" not in cell:
                            return cell
                        count += 1
                        return re.sub(rb"<v(?:\s[^>]*)?>.*?</v>", b"<v></v>", cell, flags=re.S)
                    data = re.sub(rb"<c\b[^>]*>.*?</c>", clear, data, flags=re.S)
                target.writestr(part, data)
        assert count > 0
        formulas = load_workbook(path)
        numeric = load_workbook(path, data_only=True)
        try:
            for sheet in formulas:
                for row in sheet:
                    for cell in row:
                        if cell.data_type == "f":
                            assert numeric[sheet.title][cell.coordinate].value is None
        finally:
            formulas.close()
            numeric.close()
        audit.update(cache_cleared_workbook=str(path.resolve()), caches_removed=count)
        print(f"{audit['profile']}: removed {count} formula caches", flush=True)
    audit_path.write_text(json.dumps(audits, indent=2), encoding="utf-8")


def audit_recalculation(audit_path):
    audits = json.loads(audit_path.read_text(encoding="utf-8"))
    for audit in audits:
        path = audit_path.parent / "recalculated" / Path(audit["workbook"]).name
        original = load_workbook(audit["workbook"], data_only=True)
        recalculated = load_workbook(path, data_only=True)
        formulas = load_workbook(path)
        maximum = 0.0
        checked = 0
        rms = []
        try:
            for name in (original.worksheets[0].title, "Statistics"):
                for row in original[name]:
                    for cell in row:
                        if not isinstance(cell.value, (float, int)):
                            continue
                        actual = recalculated[name][cell.coordinate].value
                        assert isinstance(actual, (int, float)), (audit["profile"], name, cell.coordinate, actual)
                        assert np.isclose(actual, cell.value, rtol=2e-14, atol=1e-12), (audit["profile"], cell.coordinate, actual, cell.value)
                        maximum = max(maximum, abs(actual - cell.value))
                        checked += 1
            sheet = formulas.worksheets[0]
            original_formulas = load_workbook(audit["workbook"])
            try:
                for col in range(1, sheet.max_column + 1):
                    cell = original_formulas.worksheets[0].cell(2, col)
                    if cell.data_type == "f" and "SUMSQ(" in cell.value:
                        rms.append({"cell": cell.coordinate, "formula": cell.value,
                                    "python": original.worksheets[0][cell.coordinate].value,
                                    "excel_recalculated": recalculated.worksheets[0][cell.coordinate].value})
                assert len(rms) == 2
                for channel in audit["live_math"]:
                    index = audit["equation_inventory"].index(channel)  # inventory membership also checked below
                    mapping = original_formulas["Rename From VSM to Astauto"]
                    col = next(row[6].value for row in mapping.iter_rows(min_row=3) if row[5].value == channel)
                    assert all(sheet.cell(row, col).data_type == "f" for row in range(5, audit["samples"] + 5))
            finally:
                original_formulas.close()
        finally:
            original.close()
            recalculated.close()
            formulas.close()
        audit.update(live_recalculation="PASS: desktop Excel CalculateFullRebuild on cache-cleared copy",
                     recalculated_workbook=str(path.resolve()), recalculated_numeric_cells=checked,
                     maximum_recalculation_difference=maximum, recalculated_rms=rms)
        print(f"{audit['profile']}: {checked} recalculated numeric cells, max difference {maximum:.12g}", flush=True)
    audit_path.write_text(json.dumps(audits, indent=2), encoding="utf-8")


def audit_report(report, baseline_path, *, refinements=False):
    book = load_workbook(report.report_path)
    cached = load_workbook(report.report_path, data_only=True)
    baseline = load_workbook(baseline_path, data_only=True)
    try:
        sheet, numeric, old = book.worksheets[0], cached.worksheets[0], baseline.worksheets[0]
        mapping = book["Rename From VSM to Astauto"]
        assert mapping.sheet_state == "visible"
        assert book.sheetnames == baseline.sheetnames
        assert "Plot Templates" not in book.sheetnames and not sheet._images
        assert sheet.freeze_panes == old.freeze_panes
        assert list(sheet.merged_cells) == list(old.merged_cells)
        assert {k: dict(v) for k, v in sheet.column_dimensions.items()} == {k: dict(v) for k, v in old.column_dimensions.items()}
        assert {k: dict(v) for k, v in sheet.row_dimensions.items()} == {k: dict(v) for k, v in old.row_dimensions.items()}
        # All numbers, including top RMS, bottom statistics and summary references,
        # are compared across the entire primary worksheet and Statistics sheet.
        numeric_cells = 0
        max_error = 0.0
        for current, previous in ((numeric, old), (cached["Statistics"], baseline["Statistics"])):
            assert (current.max_row, current.max_column) == (previous.max_row, previous.max_column)
            for row in current:
                for cell in row:
                    before = previous[cell.coordinate]
                    if isinstance(before.value, (int, float)):
                        assert isinstance(cell.value, (int, float)), cell.coordinate
                        assert np.isclose(cell.value, before.value, rtol=2e-14, atol=1e-12), cell.coordinate
                        max_error = max(max_error, abs(cell.value - before.value))
                        numeric_cells += 1
        columns = {c.channel_id: i for i, c in enumerate(report.report_channels, 1)}
        math = {c.semantic_name: c for c in report.profile.math_channels}
        exported_math = {c.channel_id for c in report.report_channels if c.kind == "math"}
        expected_math = set(report.math_result.calculation_order)
        assert exported_math == expected_math
        coverage = []
        live_math = []
        retained_math = []
        for channel in report.report_channels:
            col = columns[channel.channel_id]
            row = col + 2
            assert mapping.cell(row, 6).value == channel.channel_id
            assert mapping.cell(row, 3).value == channel.kind.upper()
            if channel.kind != "math":
                continue
            definition = math[channel.channel_id]
            equation = mapping.cell(row, 8)
            assert equation.data_type == "s"
            assert equation.value.splitlines()[0] == f"Formula: {channel.channel_id} = {definition.expression}"
            assert mapping.cell(row, 2).value == channel.display_name
            assert mapping.cell(row, 5).value == (channel.unit or "-")
            assert equation.alignment.wrap_text and mapping.row_dimensions[row].height >= 30
            assert sheet.cell(3, col).fill.fgColor.rgb.endswith("C65911")
            if refinements:
                note = sheet.cell(3, col).comment
                assert note and definition.expression in note.text
                assert f"Formula: {channel.channel_id} = {definition.expression}" in note.text
                if sheet.cell(5, col).data_type == "f":
                    assert all(sheet.cell(row, col).data_type == "f" for row in range(5, report.sample_count + 5))
                    live_math.append(channel.channel_id)
                else:
                    assert all(sheet.cell(row, col).data_type != "f" for row in range(5, report.sample_count + 5))
                    retained_math.append(channel.channel_id)
            coverage.append(channel.channel_id)
        assert set(coverage) == exported_math
        plots = []
        old_charts = baseline.worksheets[0]._charts
        assert len(sheet._charts) == len(old_charts)
        for chart, before in zip(sheet._charts, old_charts):
            expected_anchor = before.anchor
            if refinements and report.profile.metadata.powertrain == "diesel":
                expected_anchor._from.row += 5
            assert chart.anchor == expected_anchor
            assert chart.layout == before.layout
            assert chart.x_axis.__class__.__name__ == "NumericAxis"
            series = [s for part in chart._charts for s in part.series]
            for item in series:
                for ref in (item.xVal.numRef.f, item.yVal.numRef.f):
                    name, (_, first, _, last) = range_to_tuple(ref)
                    assert name == sheet.title and (first, last) == (5, report.sample_count + 4)
            assert [s.xVal.numRef.f for s in series] == [s.xVal.numRef.f for p in before._charts for s in p.series]
            assert [s.yVal.numRef.f for s in series] == [s.yVal.numRef.f for p in before._charts for s in p.series]
            plots.append({"title": title(chart.title), "x": title(chart.x_axis.title),
                          "y": [title(p.y_axis.title) for p in chart._charts],
                          "series": [s.yVal.numRef.f for s in series],
                          "anchor": [chart.anchor._from.col, chart.anchor._from.row]})
        auxiliary = None
        if report.profile.metadata.powertrain == "diesel":
            for i in (1, 2):
                col = columns[f"engine_auxiliary_torque_{i}"]
                assert sheet.cell(3, col).value == f"Auxiliary Torque_{i}"
                assert sheet.cell(4, col).value == "Nm"
            power = [c for c in sheet._charts if title(c.title) == "Auxiliary Power"]
            assert len(power) == 1
            chart = power[0]
            assert title(chart.x_axis.title) == "Distance [km]"
            assert title(chart.y_axis.title) == "Power [kW]"
            assert [s.tx.v for s in chart.series] == ["Total Auxiliary Power", "Auxiliary Power_1", "Auxiliary Power_2"]
            values = report.plotting_result.values_by_semantic_name
            rpm = values["engine_speed"]
            for i in (1, 2):
                np.testing.assert_allclose(values[f"auxiliary_power_{i}"],
                                           rpm * values[f"engine_auxiliary_torque_{i}"] * 2 * np.pi / 60000,
                                           rtol=2e-14, atol=1e-12)
            total = values["total_auxiliary_power"]
            np.testing.assert_allclose(total, values["auxiliary_power_1"] + values["auxiliary_power_2"], rtol=2e-14, atol=1e-12)
            selected = sorted(set([0, len(total) // 2, len(total) - 1, int(np.argmax(total)), int(np.argmin(total))]))
            auxiliary = {"zero_samples": int(np.isclose(total, 0, atol=1e-12).sum()),
                         "nonzero_samples": int((~np.isclose(total, 0, atol=1e-12)).sum()),
                         "samples": [{"index": j, "distance_km": float(values["distance_km"][j]),
                                      "total_kw": float(total[j]), "power_1_kw": float(values["auxiliary_power_1"][j]),
                                      "power_2_kw": float(values["auxiliary_power_2"][j])} for j in selected]}
        return {"profile": report.profile.profile_id, "source": str(report.dataset.source_path),
                "source_sha256": report.dataset.quality.source_sha256, "workbook": str(report.report_path),
                "baseline": str(baseline_path), "samples": report.sample_count,
                "math_exported": len(exported_math), "equations_displayed": len(coverage), "unresolved": 0,
                "numeric_cells_compared": numeric_cells, "maximum_numeric_difference": max_error,
                "statistics": report.statistic_count, "kpis": report.kpi_count,
                "equation_inventory": coverage, "charts": plots, "auxiliary": auxiliary,
                "live_math": live_math, "retained_python_math": retained_math,
                "live_recalculation": "not yet performed",
                "programmatic_inspection": "PASS", "visual_inspection": "not performed"}
    finally:
        book.close()
        cached.close()
        baseline.close()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--baseline", action="store_true")
    parser.add_argument("--run-name", default="updated")
    parser.add_argument("--refinements", action="store_true")
    parser.add_argument("--prepare-recalculation", type=Path)
    parser.add_argument("--audit-recalculation", type=Path)
    args = parser.parse_args()
    if args.prepare_recalculation:
        prepare_recalculation(args.prepare_recalculation.resolve())
        return
    if args.audit_recalculation:
        audit_recalculation(args.audit_recalculation.resolve())
        return
    root = ROOT / "outputs/october_2026_validation"
    phase = "baseline" if args.baseline else args.run_name
    audits = []
    for profile, filename in CASES.items():
        destination = root / phase / profile
        if destination.exists():
            raise FileExistsError(f"Preserve existing artifacts: {destination}")
        result = generate_profile_excel_report(ROOT / "reference_files" / filename,
                                              ROOT / "config/report_profiles" / f"{profile}.yaml", destination)
        if not args.baseline:
            baseline_phase = "final" if args.refinements else "baseline"
            audit = audit_report(result, root / baseline_phase / profile / result.report_path.name,
                                 refinements=args.refinements)
            powerpoint = build_profile_powerpoint_report(result, destination / "powerpoint")
            deck = Presentation(powerpoint.presentation_path)
            audit.update({"powerpoint": str(powerpoint.presentation_path), "slides": len(deck.slides)})
            audits.append(audit)
        print(f"{phase}: {profile}: {result.sample_count} samples, {result.math_count} math channels", flush=True)
    if not args.baseline:
        audit_path = root / phase / "audit.json" if args.refinements else root / "audit.json"
        audit_path.write_text(json.dumps(audits, indent=2), encoding="utf-8")


if __name__ == "__main__":
    main()
