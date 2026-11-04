"""Isolated synthetic fixtures, never defaults or real manufacturer ratings."""
from copy import deepcopy
import math
from types import SimpleNamespace
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.chart import ScatterChart
import pytest

from vsm_postprocessing.continuous_torque import ContinuousTorqueCurve
from vsm_postprocessing.continuous_torque_reporting import ContinuousTorqueInputs
from vsm_postprocessing.continuous_torque_excel import write_continuous_torque
from vsm_postprocessing.motor_performance import MotorPerformanceCurve
from vsm_postprocessing.motor_performance_excel import write_motor_performance
from vsm_postprocessing.motor_capability_excel import write_motor_capability
from vsm_postprocessing.motor_capability import read_motor_capability, NS
from vsm_postprocessing.excel_formulas import save_report_workbook


def fixture(*, speeds=(0,100,400,800), factor=1):
    book = Workbook(); sheet = book.active; sheet.title = 'Torque Maps'
    sheet['A3'] = 'No capability curves or drivetrain conversions.'
    for title, anchor in [('Original motor map', 'A6'), ('Original wheel map', 'N6')]:
        chart = ScatterChart(); chart.title = title; sheet.add_chart(chart, anchor)
    peak = MotorPerformanceCurve(speeds, (120*factor,)*len(speeds),17.5,.97)
    assessment = SimpleNamespace(inputs=ContinuousTorqueInputs(ContinuousTorqueCurve((0,200,500,900),(75*factor,)*4)))
    write_motor_performance(book, peak)
    write_continuous_torque(book, assessment, peak)
    return book, peak, assessment


def test_live_formula_provenance_independent_grids_axes_and_preservation(tmp_path):
    book, peak, assessment = fixture()
    sheet = book['Torque Maps']
    charts = [ET.tostring(deepcopy(c)._write()) for c in sheet._charts]
    cells = {c.coordinate: (c.value,c.style_id) for row in sheet for c in row if c.value is not None}
    record = write_motor_capability(book, peak, assessment)
    assert write_motor_capability(book, peak, assessment) == record
    assert len(sheet._charts) == 3 and record['chart_anchor'] == 'AA6'
    assert charts == [ET.tostring(deepcopy(c)._write()) for c in sheet._charts[:2]]
    assert all((sheet[key].value,sheet[key].style_id)==value for key,value in cells.items())
    assert sheet['BP6'].value == '=AY6*AX6*2*PI()/60000'
    assert sheet['BQ6'].value == '=BJ6*BI6*2*PI()/60000'
    assert sheet['BP6'].font.color.rgb == '00000000'
    path = tmp_path/'test.xlsx'; save_report_workbook(book,path)
    payload = read_motor_capability(path)
    assert payload['series'][0]['x'] == [0,100,400,800]
    assert payload['series'][1]['x'] == [0,200,500,900]
    for torque, power in zip(payload['series'][:2],payload['series'][2:]):
        assert torque['x'] == power['x']
        assert power['y'] == pytest.approx([t*r*2*math.pi/60000 for t,r in zip(torque['y'],torque['x'])])
        assert power['y'][0] == 0
        assert all('BE' not in c['formula'] and 'BG' not in c['formula'] and 'AZ' not in c['formula'] for c in power['y_cells'])
    with ZipFile(path) as archive:
        chart = ET.fromstring(archive.read(payload['chart_part']))
        groups = chart.findall('.//c:scatterChart',NS)
        assert [len(c.findall('c:ser',NS)) for c in groups] == [2,2]
        ids = [[a.attrib['val'] for a in c.findall('c:axId',NS)] for c in groups]
        assert len(set(ids[0] + ids[1])) == 4  # Complete native secondary XY pair.
        axes = {a.find('c:axId',NS).attrib['val']:a for a in chart.findall('.//c:valAx',NS)}
        assert axes[ids[1][0]].find('c:delete',NS).attrib['val'] == '1'
        assert axes[ids[0][1]].find('c:axPos',NS).attrib['val'] == 'l'
        assert axes[ids[1][1]].find('c:axPos',NS).attrib['val'] == 'r'
    loaded = load_workbook(path)
    assert loaded.calculation.calcMode == 'auto' and loaded.calculation.fullCalcOnLoad
    loaded.close()


@pytest.mark.parametrize('speeds',[(0,0,100),(-1,0,100),(0,100,50)])
def test_invalid_grid_omits_without_modifying_existing_tables_or_charts(speeds):
    book, peak, assessment = fixture(speeds=speeds)
    result = write_motor_capability(book,peak,assessment)
    assert result['status'] == 'omitted' and 'strictly increasing' in result['warning']
    assert len(book['Torque Maps']._charts) == 2


@pytest.mark.parametrize('missing',['peak','continuous','sheet'])
def test_missing_inputs_never_fall_back(missing):
    book, peak, assessment = fixture()
    if missing == 'peak': peak = None
    if missing == 'continuous': assessment = None
    if missing == 'sheet': book.remove(book['Torque Maps'])
    result = write_motor_capability(book,peak,assessment)
    assert result['status'] == 'omitted' and result['warning']


def test_populated_chart_region_and_output_columns_are_preserved():
    book, peak, assessment = fixture()
    sheet = book['Torque Maps']; sheet['AA10'] = 'Reserved'; sheet['BP5'] = 'Existing output'
    record = write_motor_capability(book,peak,assessment)
    assert record['chart_anchor'] != 'AA6'
    assert sheet['AA10'].value == 'Reserved' and sheet['BP5'].value == 'Existing output'
    assert all('BP' not in s['power_reference'] for s in record['series'])


def test_rejects_stale_or_other_run_source_values_before_writing():
    book, peak, assessment = fixture()
    book['Torque Maps']['AY6'] = 999
    with pytest.raises(ValueError,match='provenance mismatch'):
        write_motor_capability(book,peak,assessment)
    assert len(book['Torque Maps']._charts) == 2
