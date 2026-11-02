"""Regenerate Excel 03 and compare actual user-provided baseline workbooks."""
from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import sys
from zipfile import ZipFile

from lxml import etree
import numpy as np
from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
from vsm_postprocessing.excel_report_engine import generate_profile_excel_report
from validate_final_torque_refinements import audit_maps
from validate_torque_maps import inspect_excel, numeric_record


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def compare_workbooks(report, layout_path, current_path):
    """Require baseline maps and unchanged current engineering worksheet XML."""
    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with ZipFile(layout_path) as layout, ZipFile(current_path) as current, ZipFile(report.report_path) as new:
        def sheets(archive):
            root = etree.fromstring(archive.read("xl/workbook.xml"))
            return {s.get("name"): f"xl/worksheets/sheet{i}.xml"
                    for i, s in enumerate(root.find("m:sheets", ns), 1)}

        old_sheets, current_sheets, new_sheets = sheets(layout), sheets(current), sheets(new)
        assert list(new_sheets) == [name for name in old_sheets if name != "Motor Curves"]
        assert list(new_sheets) == list(current_sheets)
        map_part = new_sheets["Torque Maps"]
        assert layout.read(old_sheets["Torque Maps"]) == new.read(map_part), "Baseline map layout/data changed"
        unchanged_sheets = []
        for name in new_sheets:
            if name not in {"Torque Maps", "Metadata"}:
                assert current.read(current_sheets[name]) == new.read(new_sheets[name]), name
                unchanged_sheets.append(name)
        unchanged_charts, restored_charts = [], []
        for name in new.namelist():
            if name.startswith("xl/charts/chart") and name.endswith(".xml"):
                expected = layout if b"'Torque Maps'!" in new.read(name) else current
                assert expected.read(name) == new.read(name), name
                (restored_charts if expected is layout else unchanged_charts).append(name)
            if name.startswith("xl/drawings/") and name.endswith(".xml"):
                assert current.read(name) == new.read(name), name
        assert len(restored_charts) == 2

    baseline = load_workbook(layout_path, read_only=True, data_only=True)
    current = load_workbook(current_path, read_only=True, data_only=True)
    new = load_workbook(report.report_path, read_only=True, data_only=True)
    try:
        assert list(baseline["Torque Maps"].values) == list(new["Torque Maps"].values)
        assert list(current["Statistics"].values) == list(new["Statistics"].values)
        n, count = report.sample_count, report.report_channel_count
        for name, values in report.plotting_result.values_by_semantic_name.items():
            # Every exported processed channel is also checked through the
            # main worksheet's exact XML comparison above, including formulas.
            assert len(values) == n, name
        a = np.asarray(list(current.worksheets[0].iter_rows(min_row=5, max_row=n+4, max_col=count, values_only=True)))
        b = np.asarray(list(new.worksheets[0].iter_rows(min_row=5, max_row=n+4, max_col=count, values_only=True)))
        np.testing.assert_array_equal(a, b)
        def metadata(book):
            return {row[0]: row[1] for row in book["Metadata"].values if row[0] != "Generated UTC"}
        assert metadata(current) == metadata(new)
    finally:
        baseline.close()
        current.close()
        new.close()
    return {"status": "PASS", "baseline_map_worksheet_xml_identical": True,
            "baseline_map_values_identical": True, "baseline_chart_xml_identical": restored_charts,
            "current_main_worksheet_xml_identical": True, "unchanged_worksheets": unchanged_sheets,
            "unchanged_chart_xml": unchanged_charts, "unchanged_statistics": True,
            "time_series_cells_equal": n*count, "metadata_only_generation_time_changed": True}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--layout-baseline", type=Path, required=True)
    parser.add_argument("--current", type=Path, required=True)
    parser.add_argument("--approved-powerpoint", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    root = args.output.resolve()
    root.mkdir(parents=True, exist_ok=False)
    protected = [args.source, args.layout_baseline, args.current, args.approved_powerpoint]
    hashes = {str(p.resolve()): sha(p) for p in protected}
    current = load_workbook(args.current, read_only=True, data_only=True)
    try:
        metadata = {row[0]: row[1] for row in current["Metadata"].values}
    finally:
        current.close()
    assert sha(args.source) == metadata["Source SHA-256"]
    print("Generating Excel only:", args.source.name, flush=True)
    report = generate_profile_excel_report(args.source, "config/report_profiles/sergio_2026_10/robosprayer_electric.yaml",
        root/"case_0"/"profile_excel_report", machine_name_override=metadata["Machine / Vehicle name"],
        reporting_inputs=json.loads(metadata["Reporting inputs"]))
    record = {"kind": "electric", "source": str(args.source.resolve()), "excel": str(report.report_path),
              "powerpoint": str(args.approved_powerpoint.resolve()), "powerpoint_regenerated": False,
              "approved_powerpoint_sha256": sha(args.approved_powerpoint), "excel_checks": inspect_excel(report),
              "maps": audit_maps(report), "series": report.plotting_result.series_diagnostics,
              "numerics": numeric_record(report.dataset, report.statistics_result),
              "preservation": compare_workbooks(report, args.layout_baseline, args.current)}
    (root/"validation.json").write_text(json.dumps({args.source.name: record}, indent=2), encoding="utf-8")
    assert all(sha(Path(p)) == digest for p, digest in hashes.items())
    (root/"protected_hashes.json").write_text(json.dumps(hashes, indent=2), encoding="utf-8")
    print("PASS:", record["maps"], record["preservation"], flush=True)


if __name__ == "__main__":
    main()
