"""Synthetic design curves only: these are never client/reference motor data."""
import hashlib
import json
import math
from pathlib import Path
from types import SimpleNamespace
from xml.etree import ElementTree as ET
from zipfile import ZipFile, ZIP_DEFLATED

from matplotlib.figure import Figure
from matplotlib.backends.backend_agg import FigureCanvasAgg
from openpyxl import Workbook
from openpyxl.chart import ScatterChart, Series, Reference
from openpyxl.workbook.defined_name import DefinedName
from pptx import Presentation
from pptx.util import Inches
import pytest

from vsm_postprocessing.errors import PowerPointReportError
from vsm_postprocessing.motor_capability import (read_motor_capability, render_motor_capability,
    insert_motor_slide_definition, SERIES, NS)
from vsm_postprocessing.profile_powerpoint_report_engine import (
    _mark_energy_kpi_title, _update_energy_kpi_titles, build_profile_powerpoint_report)
from vsm_postprocessing.report_profile import _parse_presentation
from test_october_soil_plots import soil_report


def source_fixture(path, *, defect=None):
    book = Workbook(); sheet = book.active; sheet.title = 'Torque Maps'
    primary = ScatterChart(); primary.title = 'Motor Torque and Power'
    primary.x_axis.title = 'Motor Speed [rpm]'; primary.y_axis.title = 'Torque [Nm]'
    secondary = ScatterChart(); secondary.y_axis.axId = 200
    secondary.x_axis.title = 'Motor Speed [rpm]'; secondary.y_axis.title = 'Power [kW]'
    secondary.y_axis.axPos = 'r'; secondary.y_axis.crosses = 'max'
    # Relocated columns, unequal independent grids, power subsets and strRef headers.
    grids = [(0, 101, 353, 702, 1300), (0, 207, 419, 835), (101, 353, 702, 1300), (207, 419, 835)]
    columns = [(3, 4), (8, 9), (13, 14), (18, 19)]
    cached = {}
    for index, (name, speeds, (xc, yc)) in enumerate(zip(SERIES, grids, columns)):
        sheet.cell(4, xc, 'Motor Speed [rpm]')
        sheet.cell(4, yc, name + (' [Nm]' if index < 2 else ' [kW]'))
        for row, speed in enumerate(speeds, 5):
            sheet.cell(row, xc, speed)
            torque = 120 if index % 2 == 0 else 75
            if index < 2:
                sheet.cell(row, yc, torque)
            else:
                c = sheet.cell(row, yc, f'={torque}*{sheet.cell(row,xc).coordinate}*2*PI()/60000')
                cached[c.coordinate] = round(torque * speed * 2 * math.pi / 60000, 2)
        series = Series(Reference(sheet, min_col=yc, min_row=4, max_row=4+len(speeds)),
                        Reference(sheet, min_col=xc, min_row=5, max_row=4+len(speeds)), title_from_data=True)
        if index == 0:
            book.defined_names.add(DefinedName('PeakSpeed', attr_text=series.xVal.numRef.f))
            series.xVal.numRef.f = 'PeakSpeed'
        (primary if index < 2 else secondary).series.append(series)
    if defect == 'length': secondary.series[0].xVal.numRef.f = "'Torque Maps'!$M$5:$M$7"
    if defect == 'missing_series': secondary.series.pop()
    if defect == 'unit': secondary.y_axis.title = 'Power [W]'
    if defect == 'shaft': primary.x_axis.title = 'Wheel Speed [rpm]'
    if defect == 'negative': sheet.cell(5, 3, -1)
    if defect == 'duplicate': primary.series.append(primary.series[0])
    if defect == 'text': sheet.cell(5, 4, 'invalid')
    if defect == 'unmatched': sheet.cell(5, 13, 102)
    if defect == 'duplicate_speed':
        sheet.cell(5, 3, 101); sheet.cell(5, 4, 121)
    if defect == 'nonmonotonic':
        sheet.cell(5, 3, 101); sheet.cell(6, 3, 0)
    if defect == 'chart_title': primary.title = 'Other chart'
    primary += secondary
    sheet.add_chart(primary, 'B15')
    if defect == 'sheet': sheet.title = 'Other'
    book.save(path)
    with ZipFile(path) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    root = ET.fromstring(parts['xl/worksheets/sheet1.xml'])
    for cell in root.findall('s:sheetData/s:row/s:c', NS):
        if cell.attrib['r'] in cached and defect != 'uncached':
            value = cached[cell.attrib['r']]
            if defect == 'discrepancy': value += 3
            cell.find('s:v', NS).text = str(value)
        if defect == 'nonfinite' and cell.attrib['r'] == 'D5': cell.find('s:v', NS).text = 'nan'
    parts['xl/worksheets/sheet1.xml'] = ET.tostring(root)
    with ZipFile(path, 'w', ZIP_DEFLATED) as z:
        for name, data in parts.items(): z.writestr(name, data)
    return path


def test_exact_independent_arrays_formulas_named_ranges_and_spot_checks(tmp_path):
    path = source_fixture(tmp_path/'synthetic.xlsx')
    before = hashlib.sha256(path.read_bytes()).hexdigest()
    data = read_motor_capability(path)
    assert [s['point_count'] for s in data['series']] == [5, 4, 4, 3]
    assert data['series'][0]['x_reference'] == 'PeakSpeed'
    assert data['series'][1]['x'] == [0, 207, 419, 835]
    assert data['series'][3]['x'] == [207, 419, 835]
    assert all(c['formula'] for s in data['series'][2:] for c in s['y_cells'])
    assert not data['warnings']
    for duty in ['Peak', 'Continuous']:
        checks = [c for c in data['power_checks'] if c['duty'] == duty and c['rpm'] > 0]
        assert len(checks) >= 3
        for check in checks[:3]:
            assert check['spreadsheet_kw'] == pytest.approx(check['torque_nm']*check['rpm']*math.pi/30000, abs=.005)
    assert hashlib.sha256(path.read_bytes()).hexdigest() == before


@pytest.mark.parametrize('defect,match', [('length','length'), ('missing_series','missing series'),
    ('unit','requires kW'), ('shaft','motor shaft'), ('negative','nonnegative'),
    ('duplicate','duplicate series'), ('text','finite evaluated'), ('nonfinite','finite evaluated'),
    ('sheet','missing "Torque Maps"'), ('uncached','recalculate'), ('discrepancy','power discrepancy'),
    ('chart_title','found 0')])
def test_invalid_source_is_controlled_error_and_never_modified(tmp_path, defect, match):
    path = source_fixture(tmp_path/'synthetic.xlsx', defect=defect)
    before = path.read_bytes()
    with pytest.raises(PowerPointReportError, match=match): read_motor_capability(path)
    assert path.read_bytes() == before


def test_missing_workbook_or_native_chart(tmp_path):
    with pytest.raises(PowerPointReportError, match='workbook'): read_motor_capability(tmp_path/'missing.xlsx')
    b = Workbook(); b.active.title = 'Torque Maps'; b.save(tmp_path/'empty.xlsx')
    with pytest.raises(PowerPointReportError, match='no native chart'): read_motor_capability(tmp_path/'empty.xlsx')


def test_unmatched_speed_is_not_interpolated_or_silently_validated(tmp_path):
    data = read_motor_capability(source_fixture(tmp_path/'synthetic.xlsx', defect='unmatched'))
    assert data['series'][2]['x'][0] == 102
    assert any('no same-speed motor torque' in w for w in data['warnings'])
    assert not any(c['rpm'] == 102 for c in data['power_checks'])


def test_external_workbook_configuration_is_rejected():
    from vsm_postprocessing.errors import ConfigurationError
    with pytest.raises(ConfigurationError):
        _parse_presentation({'motor_capability_workbook': 'other_profile.xlsx'})


def test_duplicate_speed_never_selects_an_arbitrary_torque(tmp_path):
    data = read_motor_capability(source_fixture(tmp_path/'synthetic.xlsx', defect='duplicate_speed'))
    assert data['series'][0]['x'][:2] == [101, 101]
    assert data['series'][0]['y'][:2] == [121, 120]
    assert any('ambiguous motor torque' in w for w in data['warnings'])
    assert not any(c['duty'] == 'Peak' and c['rpm'] == 101 for c in data['power_checks'])


def test_original_chart_order_is_preserved_without_sorting(tmp_path):
    data = read_motor_capability(source_fixture(tmp_path/'synthetic.xlsx', defect='nonmonotonic'))
    assert data['series'][0]['x'] == [101, 0, 353, 702, 1300]
    assert not data['warnings']


def test_export_has_four_exact_series_dual_axes_and_unclipped_labels(tmp_path, monkeypatch):
    data = read_motor_capability(source_fixture(tmp_path/'synthetic.xlsx'))
    original = Figure.savefig
    def inspect(figure, *args, **kwargs):
        canvas = FigureCanvasAgg(figure); canvas.draw(); renderer = canvas.get_renderer()
        left, right = figure.axes
        assert left.get_xlabel() == 'Motor Speed [rpm]'
        assert left.get_ylabel() == 'Torque [Nm]' and right.get_ylabel() == 'Power [kW]'
        lines = [*left.lines, *right.lines]
        assert len(lines) == 4 and len({l.get_color() for l in lines}) == 4
        for line, series in zip(lines, data['series']):
            assert list(line.get_xdata()) == series['x'] and list(line.get_ydata()) == series['y']
        for artist in [left.get_legend(), left.xaxis.label, left.yaxis.label, right.yaxis.label]:
            bounds = artist.get_window_extent(renderer)
            assert 0 <= bounds.x0 < bounds.x1 <= figure.bbox.width
            assert 0 <= bounds.y0 < bounds.y1 <= figure.bbox.height
        original(figure, *args, **kwargs)
    monkeypatch.setattr(Figure, 'savefig', inspect)
    rendered = render_motor_capability(data, tmp_path/'assets')
    assert Path(rendered['image']).is_file()


def test_insertion_is_idempotent_and_precedes_summary():
    for anchor in ['torque_maps', 'wheel_steering_angles']:
        slides = [dict(slide_id=s) for s in ['cover', anchor, 'simulation_summary']]
        once = insert_motor_slide_definition(slides)
        assert insert_motor_slide_definition(once) == once
        expected = ['cover', 'motor_torque_power', anchor, 'simulation_summary'] if anchor == 'torque_maps' else ['cover', anchor, 'motor_torque_power', 'simulation_summary']
        assert [s['slide_id'] for s in once] == expected


@pytest.mark.parametrize('shift', [False, True])
def test_only_kpi_titles_change_values_xml_retained_and_index_independent(shift):
    prs = Presentation(); definitions = []
    if shift:
        prs.slides.add_slide(prs.slide_layouts[6]); definitions.append(SimpleNamespace(slide_id='motor_torque_power'))
    slide = prs.slides.add_slide(prs.slide_layouts[6]); definitions.append(SimpleNamespace(slide_id='traction_auxiliaries'))
    untouched = []
    for index, text in enumerate(['TYRE TOTAL ENERGY','AUX ENERGY','TOTAL ENERGY','ENERGY CONSUMPTION','POWER CONSUMPTION']):
        label = slide.shapes.add_textbox(Inches(index*2), Inches(1), Inches(1.8), Inches(.4)); label.text = text
        _mark_energy_kpi_title(label, str(index))
        value = slide.shapes.add_textbox(Inches(index*2), Inches(1.6), Inches(1.8), Inches(.6)); value.text = '1,234.56 kWh'
        untouched.append((value, value.element.xml))
        if index > 2: untouched.append((label, label.element.xml))
    heading = slide.shapes.add_textbox(0, 0, Inches(3), Inches(.5)); heading.text='TOTAL ENERGY'
    untouched.append((heading, heading.element.xml))
    mapping = _update_energy_kpi_titles(prs, definitions)
    assert len(mapping) == 3
    assert all(m['slide_before'] == 1 and m['slide_after'] == 1+shift for m in mapping)
    assert all(s.element.xml == before for s, before in untouched)
    after = slide.element.xml
    assert _update_energy_kpi_titles(prs, definitions) == [] and slide.element.xml == after


def test_both_profiles_preserve_existing_slides_and_add_one_editable_motor_slide(soil_report, tmp_path):
    from vsm_postprocessing.excel_report_engine import generate_profile_excel_report
    from vsm_postprocessing.motor_performance import MotorPerformanceCurve
    from vsm_postprocessing.continuous_torque import ContinuousTorqueCurve
    from vsm_postprocessing.continuous_torque_reporting import ContinuousTorqueInputs
    from test_sergio_october_15 import csv_for, profile_path
    kind, data, baseline_excel, baseline = soil_report
    rating = 120 if kind == 'electric' else 240
    excel = generate_profile_excel_report(csv_for(data, tmp_path/'source.csv'), profile_path(kind), tmp_path/'excel',
        motor_performance=MotorPerformanceCurve((0,100,400,800),(rating,)*4,17.5,.97),
        continuous_torque_inputs=ContinuousTorqueInputs(ContinuousTorqueCurve((0,200,500,900),(rating/2,)*4)))
    result = build_profile_powerpoint_report(excel, tmp_path/'ppt')
    again = build_profile_powerpoint_report(excel, tmp_path/'ppt')
    before, after = Presentation(baseline.presentation_path), Presentation(again.presentation_path)
    assert len(after.slides) == len(before.slides)+1 == (17 if kind == 'electric' else 16)
    motor = [s for s in after.slides if s.name == 'VSM motor_torque_power']
    assert len(motor) == 1
    text = [s.text for s in motor[0].shapes if s.has_text_frame]
    assert 'MOTOR BRAND:' in text and 'MOTOR TYPE / MODEL:' in text
    assert 'Motor Torque & Power Characteristics' in text
    old = {d.slide_id:s for d,s in zip(baseline.powerpoint_result.config.slides, before.slides)}
    new_defs = result.powerpoint_result.config.slides
    ids = [d.slide_id for d in new_defs]
    assert ids.index('motor_torque_power') + 1 == ids.index('torque_maps')
    assert ids[-1] == 'simulation_summary'
    for d, slide in zip(new_defs, after.slides):
        if d.slide_id not in old or d.slide_id == 'torque_maps': continue
        for a,b in zip(old[d.slide_id].shapes, slide.shapes):
            if a.has_text_frame and '/' in a.text and a.text.strip().replace('/','').replace(' ','').isdigit(): continue
            assert str(a.element.xml) == str(b.element.xml)
    manifest = json.loads(result.manifest_path.read_text())
    capability = manifest['motor_capability']
    assert capability['status'] == 'included'
    assert capability['source_workbook'] == str(excel.report_path.resolve())
    assert capability['profile_id'] == excel.profile.profile_id
    assert capability['series'][0]['y'] == [rating]*4
    assert capability['series'][1]['x'] == [0,200,500,900]
    assert manifest['energy_title_updates']
    assert json.loads(baseline.manifest_path.read_text())['motor_capability']['status'] == 'omitted'
    assert any(getattr(d, 'code', '') == 'motor_capability_omitted' and d.user_visible for d in baseline_excel.statistics_result.diagnostics)
    assert 'WARNING: Motor Torque and Power chart and slide omitted' in baseline.summary_path.read_text()
