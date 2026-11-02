"""Regenerate Excel only; audit native maps and preserve approved artifacts."""
from __future__ import annotations

import argparse
import csv
import hashlib
import json
from pathlib import Path
import sys
from zipfile import ZipFile

import numpy as np
from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
from vsm_postprocessing.excel_report_engine import generate_profile_excel_report
from validate_torque_maps import inspect_excel, numeric_record


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def audit_maps(report):
    with report.dataset.source_path.open(encoding="utf-8-sig", newline="") as stream:
        reader = csv.reader(stream)
        names, units = next(reader), next(reader)
        raw = dict(zip(names, np.array([[float(v) for v in row] for row in reader]).T))
    book = load_workbook(report.report_path)
    try:
        sheet = book["Torque Maps"]
        records = report.plotting_result.series_diagnostics
        result = []
        for chart, key in zip(sheet._charts, ("torque_map_powertrain", "torque_map_wheels")):
            included = [r for r in records[key] if r["included"]]
            assert len(chart.series) == len(included) == 4
            for axis in (chart.x_axis, chart.y_axis):
                assert axis.scaling.min is None and axis.scaling.max is None and axis.majorUnit is None
            for index, (series, record) in enumerate(zip(chart.series, included)):
                assert series.marker.size == 3 and series.graphicalProperties.line.noFill
                for reference, field in ((series.xVal.numRef.f, "x_source"), (series.yVal.numRef.f, "y_source")):
                    # Source identities come from explicit diagnostic records.
                    semantic = record["x" if field == "x_source" else "y"]
                    source = report.plotting_result.channels_by_semantic_name[semantic].source_name
                    name, extent = reference.rsplit("!", 1)
                    source_sheet = book[name.strip("'").replace("''", "'")]
                    assert source_sheet is sheet
                    values = np.asarray([r[0].value for r in source_sheet[extent]])
                    np.testing.assert_allclose(values, raw[source], rtol=1e-14, atol=0)
                    assert len(values) == len(raw["Track_Time"])
                if key == "torque_map_wheels":
                    assert series.tx.v == ("FL", "FR", "RL", "RR")[index] + " Wheel"
            result.append({"plot": key, "samples_per_series": len(raw["Track_Time"]),
                           "labels": [s.tx.v for s in chart.series]})
        assert "Motor Curves" not in book
        assert book.calculation.calcMode == "auto" and book.calculation.fullCalcOnLoad
        assert not book.calculation.forceFullCalc
        assert sheet.max_row == 39 + report.sample_count
        return result
    finally:
        book.close()


def compare_previous_reports(before_root, report):
    old_excel = next(before_root.rglob("profile_excel_report/" + report.report_path.name))
    baseline = json.loads((before_root/"validation.json").read_text(encoding="utf-8"))[report.dataset.source_path.name]
    numerics = numeric_record(report.dataset, report.statistics_result)
    assert numerics == baseline["numerics"]
    old = load_workbook(old_excel, read_only=True, data_only=True)
    new = load_workbook(report.report_path, read_only=True, data_only=True)
    try:
        assert new.sheetnames == [name for name in old.sheetnames if name != "Motor Curves"]
        n, count = report.sample_count, report.report_channel_count
        a = np.asarray(list(old.worksheets[0].iter_rows(min_row=5,max_row=n+4,max_col=count,values_only=True)))
        b = np.asarray(list(new.worksheets[0].iter_rows(min_row=5,max_row=n+4,max_col=count,values_only=True)))
        np.testing.assert_array_equal(a,b)
        for name in new.sheetnames:
            if name not in {"Torque Maps", "Metadata"}:
                assert list(old[name].values) == list(new[name].values), name
        # The new run timestamp and removed sheet inventory are expected metadata changes.
        def metadata(book):
            return [row for row in book["Metadata"].values if row[0] not in {"Visible sheets", "Generated UTC"}]
        assert metadata(old) == metadata(new)
        assert list(old["Torque Maps"].iter_rows(min_row=34,max_row=38,values_only=True)) == list(
            new["Torque Maps"].iter_rows(min_row=34,max_row=38,values_only=True))
    finally:
        old.close()
        new.close()
    unchanged_parts = []
    with ZipFile(old_excel) as a, ZipFile(report.report_path) as b:
        for name in a.namelist():
            if (name.startswith("xl/worksheets/sheet") and name.endswith(".xml")
                    and name not in {"xl/worksheets/sheet3.xml", "xl/worksheets/sheet6.xml", "xl/worksheets/sheet7.xml"}):
                assert a.read(name) == b.read(name), name
                unchanged_parts.append(name)
        chart_parts = sorted(name for name in a.namelist() if name.startswith("xl/charts/chart") and name.endswith(".xml"))
        for name in chart_parts:
            if "Torque Maps" not in a.read(name).decode():
                assert a.read(name) == b.read(name), name
                unchanged_parts.append(name)
    return {"status":"PASS", "time_series_cells_equal":n*count,
            "unrelated_worksheet_values_equal":True, "provenance_equal":True,
            "expected_metadata_changes":["Generated UTC", "Visible sheets"],
            "exact_math_hashes_equal":len(numerics["math"]), "exact_statistics_equal":len(numerics["statistics"]),
            "exact_kpis_equal":len(numerics["kpis"]),
            "unchanged_xml_parts":unchanged_parts}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--source", type=Path, action="append")
    parser.add_argument("--baseline-report", type=Path, required=True)
    args = parser.parse_args()
    root = args.output.resolve()
    root.mkdir(parents=True, exist_ok=False)
    sources = args.source or [p for p in Path("reference_files").glob("*.csv") if "Electric" in p.name and "UoL" in p.name]
    assert sources
    hashes = {str(p):sha(p) for p in Path("reference_files").iterdir() if p.is_file()}
    records = {}
    for i,source in enumerate(sources):
        print("Generating",source.name,flush=True)
        approved_ppt = next(args.baseline_report.rglob("profile_powerpoint_report/" + source.stem + ".pptx"))
        hashes[str(approved_ppt)] = sha(approved_ppt)
        report = generate_profile_excel_report(source, "config/report_profiles/sergio_2026_10/robosprayer_electric.yaml",
                                               root/f"case_{i}"/"profile_excel_report")
        checks = {"kind":"electric", "source":str(source.resolve()), "excel":str(report.report_path),
                  "powerpoint":str(approved_ppt.resolve()), "approved_powerpoint_sha256":sha(approved_ppt),
                  "powerpoint_regenerated":False, "excel_checks":inspect_excel(report), "maps":audit_maps(report),
                  "series":report.plotting_result.series_diagnostics,
                  "numerics":numeric_record(report.dataset,report.statistics_result)}
        if args.baseline_report:
            checks["preservation"] = compare_previous_reports(args.baseline_report,report)
        records[source.name]=checks
        (root/"validation.json").write_text(json.dumps(records,indent=2),encoding="utf-8")
        print("PASS",source.name,checks["maps"],flush=True)
    assert all(sha(Path(p))==v for p,v in hashes.items())
    (root/"reference_hashes.json").write_text(json.dumps(hashes,indent=2),encoding="utf-8")


if __name__ == "__main__":
    main()
