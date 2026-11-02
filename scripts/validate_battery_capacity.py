"""Independent real-data battery audit and protected Torque Maps comparison."""
from __future__ import annotations

import argparse
from dataclasses import asdict
import hashlib
import io
import json
from pathlib import Path
import sys
from zipfile import ZipFile

import numpy as np
from openpyxl import load_workbook
from pptx import Presentation
from PIL import Image

sys.path.insert(0, str(Path(__file__).resolve().parents[1]/'src'))
from vsm_postprocessing.profile_statistics import calculate_profile_statistics
from vsm_postprocessing.reporting_inputs import ReportingInputs
from vsm_postprocessing.ui_config import generate_reporting_profile_engineering_report
from vsm_postprocessing.ui_app import _render_diagnostics
from vsm_postprocessing.diagnostics import diagnostic_records


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def audit(result, independent, confirmed):
    report, ppt = result.excel_result, result.powerpoint_result
    stats = report.statistics_result
    k = {item.definition.kpi_id: item.value for item in stats.kpis}
    np.testing.assert_allclose(k['battery_capacity_used'], independent['used_kwh'], rtol=1e-14)
    np.testing.assert_allclose(k['battery_energy_consumption_wh_per_km'], independent['consumption_whkm'], rtol=1e-14)
    expected = independent['reference_350_range'] if confirmed else independent['old_range']
    np.testing.assert_allclose(k['range_85_battery_km'], expected, rtol=1e-14)
    assert stats.battery_capacity['reference_status'] == ('confirmed' if confirmed else 'inferred')
    assert stats.battery_capacity['nominal_capacity_kwh'] == (350 if confirmed else None)
    estimate = stats.battery_capacity['energy_soc_equivalent_kwh']
    np.testing.assert_allclose(estimate['median'], independent['median'], rtol=1e-14)
    np.testing.assert_allclose(estimate['max_deviation'], independent['max_deviation'], rtol=1e-14)
    book = load_workbook(report.report_path, read_only=True, data_only=True)
    try:
        rows = {row[0]: row[2] for row in book['Statistics'].values if row[0]}
        for key, value in k.items():
            np.testing.assert_allclose(rows[key], value, rtol=1e-14)
        metadata = {row[0]: row[1] for row in book['Metadata'].values}
        assert json.loads(metadata['Battery capacity interpretation']) == stats.battery_capacity
        assert json.loads(metadata['Diagnostics']) == diagnostic_records(stats.diagnostics)
        assert 'Motor Curves' not in book
    finally:
        book.close()
    powerpoint = {item.statistic_id: item.value for item in ppt.powerpoint_result.statistics_result.statistics}
    for key in k:
        assert powerpoint[key] == k[key]
    prs = Presentation(ppt.presentation_path)
    assert len(prs.slides) == 16
    for slide in prs.slides:
        notes = json.loads(slide.notes_slide.notes_text_frame.text)
        assert notes['diagnostic_records'] == diagnostic_records(stats.diagnostics)
        assert notes['reporting_inputs'] == asdict(report.report_metadata.reporting_inputs)
    class UI:
        def __init__(self): self.messages = []
        def warning(self, message): self.messages.append(('warning', message))
        def info(self, message): self.messages.append(('info', message))
        def error(self, message): self.messages.append(('error', message))
    ui = UI()
    _render_diagnostics(ui, stats.diagnostics)
    assert ui.messages == [(d['severity'], d['message']) for d in diagnostic_records(stats.diagnostics) if d['user_visible']]
    if confirmed:
        unknown = calculate_profile_statistics(report.dataset, report.profile, reporting_inputs=ReportingInputs())
        unknown_kpis = {item.definition.kpi_id: item.value for item in unknown.kpis}
        np.testing.assert_allclose(unknown_kpis['range_85_battery_km'], independent['old_range'], rtol=1e-14)
        assert unknown.battery_capacity['reference_status'] == 'inferred'
    return {'kpis': k, 'capacity_interpretation': stats.battery_capacity,
            'diagnostics': diagnostic_records(stats.diagnostics), 'excel_powerpoint_ui_consistent': True,
            'excel': str(report.report_path), 'powerpoint': str(ppt.presentation_path), 'slides': 16,
            'unknown_capacity_fallback_verified': confirmed}


def preserve(result, baseline_root):
    report, ppt = result.excel_result, result.powerpoint_result
    source = report.dataset.source_path.name
    old = json.loads((baseline_root/'validation.json').read_text())[source]
    stats = report.statistics_result
    assert old['numerics']['raw_sha256'] == digest_bytes(report.dataset.values.tobytes())
    assert old['numerics']['math'] == {name: digest_bytes(a.tobytes()) for name, a in stats.math_result.values_by_semantic_name.items()}
    assert old['numerics']['statistics'] == {s.definition.statistic_id: s.value for s in stats.statistics}
    new_kpis = {k.definition.kpi_id: k.value for k in stats.kpis}
    allowed = {'battery_capacity_100', 'usable_battery_capacity_85', 'range_85_battery_km'}
    for key, value in old['numerics']['kpis'].items():
        if key not in allowed:
            assert new_kpis[key] == value, key
    old_excel, old_ppt = Path(old['excel']), Path(old['powerpoint'])
    with ZipFile(old_excel) as a, ZipFile(report.report_path) as b:
        for part in a.namelist():
            if ((part.startswith('xl/charts/') or part.startswith('xl/drawings/')) and part.endswith('.xml')
                    or part == 'xl/worksheets/sheet6.xml'):
                assert a.read(part) == b.read(part), part
    before = load_workbook(old_excel, read_only=True, data_only=True)
    after = load_workbook(report.report_path, read_only=True, data_only=True)
    try:
        n, count = report.sample_count, report.report_channel_count
        a = list(before.worksheets[0].iter_rows(min_row=5, max_row=n+4, max_col=count, values_only=True))
        b = list(after.worksheets[0].iter_rows(min_row=5, max_row=n+4, max_col=count, values_only=True))
        assert a == b
    finally:
        before.close(); after.close()
    a, b = Presentation(old_ppt), Presentation(ppt.presentation_path)
    assert len(a.slides) == len(b.slides) == 16
    changes = []
    old_range, new_range = old['numerics']['kpis']['range_85_battery_km'], new_kpis['range_85_battery_km']
    for index, (x, y) in enumerate(zip(a.slides, b.slides), 1):
        left = x.element.xml.replace(f'{old_range:.2f}', 'BATTERY_RANGE')
        right = y.element.xml.replace(f'{new_range:.2f}', 'BATTERY_RANGE')
        assert left == right, f'Unrelated slide XML changed: {index}'
        left_images = [Image.open(io.BytesIO(s.image.blob)).convert('RGBA') for s in x.shapes if s.shape_type == 13]
        right_images = [Image.open(io.BytesIO(s.image.blob)).convert('RGBA') for s in y.shapes if s.shape_type == 13]
        assert len(left_images) == len(right_images)
        assert all(p.size == q.size and p.tobytes() == q.tobytes() for p, q in zip(left_images, right_images)), f'Slide image pixels changed: {index}'
        if x.element.xml != y.element.xml: changes.append(index)
    return {'raw_arrays_equal': True, 'math_hashes_equal': len(old['numerics']['math']),
            'statistics_equal': len(stats.statistics), 'unrelated_kpis_equal': len(new_kpis)-len(allowed),
            'main_data_cells_equal': n*count, 'all_native_charts_and_torque_maps_xml_equal': True,
            'all_slide_image_pixels_equal': True, 'slide_layouts_equal': True, 'only_range_text_changed_on_slides': changes}


def digest_bytes(data):
    return hashlib.sha256(data).hexdigest()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--source', type=Path, action='append', required=True)
    parser.add_argument('--output', type=Path, required=True)
    parser.add_argument('--baseline', type=Path, required=True)
    parser.add_argument('--analysis', type=Path, required=True)
    args = parser.parse_args()
    root = args.output.resolve(); root.mkdir(exist_ok=False)
    independent = json.loads(args.analysis.read_text())
    baseline = json.loads((args.baseline/'validation.json').read_text())
    primary_record = baseline[args.source[0].name]
    book = load_workbook(primary_record['excel'], read_only=True, data_only=True)
    metadata = {row[0]: row[1] for row in book['Metadata'].values}; book.close()
    records = {}
    for index, source in enumerate(args.source):
        print('Generating', source.name, flush=True)
        options = {'reporting_inputs': json.loads(metadata['Reporting inputs']), 'machine_name_override': metadata['Machine / Vehicle name']} if index == 0 else {}
        result = generate_reporting_profile_engineering_report(source,
            'config/report_profiles/sergio_2026_10/robosprayer_electric.yaml', root/f'case_{index}', **options)
        record = audit(result, independent[source.name], index == 0)
        if index == 0: record['preservation'] = preserve(result, args.baseline)
        records[source.name] = record
        (root/'validation.json').write_text(json.dumps(records, indent=2), encoding='utf-8')
        print('PASS', source.name, record['kpis']['range_85_battery_km'], flush=True)


if __name__ == '__main__':
    main()
