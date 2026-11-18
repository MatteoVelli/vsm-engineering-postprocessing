"""Deterministic import fixtures; these are not manufacturer-data acceptance."""
from io import BytesIO
from zipfile import ZipFile

import numpy as np
from openpyxl import Workbook, load_workbook
from pptx import Presentation
import pytest

from vsm_postprocessing.continuous_torque import CONTINUOUS, ContinuousTorqueCurve, analyse_motor
from vsm_postprocessing.continuous_torque_reporting import ContinuousTorqueInputs
from vsm_postprocessing.errors import ConfigurationError
from vsm_postprocessing.motor_performance import SPEED, TORQUE, MotorPerformanceCurve, POSITIVE, NEGATIVE
from vsm_postprocessing.motor_torque_import import (HEADERS, motor_torque_template,
    motor_torque_worksheets, parse_motor_torque_workbook)


POINTS = [(0, 120, 60), (100, 110, 55), (400, 80, 40), (800, 40, 20)]
LEGACY_ROWS = [dict(zip((SPEED, TORQUE, CONTINUOUS), point)) for point in POINTS]


def workbook(headers=HEADERS, points=POINTS, extra_sheets=()):
    book = Workbook()
    sheet = book.active
    sheet.title = 'Motor'
    sheet.append(headers)
    for point in points:
        sheet.append(point)
    for name, rows in extra_sheets:
        sheet = book.create_sheet(name)
        for row in rows:
            sheet.append(row)
    output = BytesIO()
    book.save(output)
    book.close()
    return output.getvalue()


def test_valid_workbook_produces_exact_existing_editor_rows_and_preserves_precision():
    assert parse_motor_torque_workbook(workbook()) == LEGACY_ROWS
    value = 1068.5710547955418
    content = workbook(points=[(0, value, 50), (175, value, 40)])
    rows = parse_motor_torque_workbook(content)
    book = load_workbook(BytesIO(content))
    try:
        stored = book.active['B3'].value
    finally:
        book.close()
    # XLSX writers may round on serialization. The importer must retain the
    # value actually stored in the source, without rounding it again.
    assert rows[1][TORQUE] == stored
    assert stored == pytest.approx(value, rel=1e-15)
    assert MotorPerformanceCurve.from_rows(rows, 17.5, .97).calculated_rows()[1][POSITIVE] == stored * 17.5 * .97
    text_rows = parse_motor_torque_workbook(workbook(points=[(0, repr(value), 50), (175, repr(value), 40)]))
    assert text_rows[1][TORQUE] == value


def test_reordered_columns_and_unordered_rpm_keep_all_three_quantities_paired():
    points = [(continuous, rpm, peak) for rpm, peak, continuous in reversed(POINTS)]
    content = workbook(('Cont Torque', 'Motor RPM', 'Peak Torque'), points)
    assert parse_motor_torque_workbook(content) == LEGACY_ROWS


@pytest.mark.parametrize('headers', [
    ('  motor_rpm ', 'PEAK TORQUE', 'continuous torque'),
    (SPEED, TORQUE, CONTINUOUS),
    ('Motor Speed (RPM)', 'Peak Torque (N.m)', 'Motor Cont Torque [N·m]'),
    ('Motor Speed rpm', 'Motor Peak Torque Nm', 'Cont Torque Nm'),
])
def test_unambiguous_header_variations(headers):
    assert parse_motor_torque_workbook(workbook(headers)) == LEGACY_ROWS


@pytest.mark.parametrize('headers, message', [
    (('Motor RPM', 'Peak Torque'), 'Missing required columns: Cont Torque'),
    (('Speed', 'Peak Torque', 'Cont Torque'), 'Missing required columns: Motor RPM'),
    (('Motor Speed [rad/s]', 'Peak Torque', 'Cont Torque'), 'expected rpm'),
    (('Motor RPM', 'Peak Torque [kNm]', 'Cont Torque'), 'expected Nm'),
    (('Motor RPM', 'Peak Torque [N/m]', 'Cont Torque'), 'expected Nm'),
    (('Motor RPM', 'Peak Torque N/m', 'Cont Torque'), 'Missing required columns: Peak Torque'),
    (('Motor RPM', 'Peak Torque', 'Cont Torque [lb-ft]'), 'expected Nm'),
    (('Motor RPM', 'Peak Torque', 'Motor Peak Torque', 'Cont Torque'), 'Ambiguous columns'),
])
def test_missing_ambiguous_or_wrong_unit_columns(headers, message):
    with pytest.raises(ConfigurationError, match=message):
        parse_motor_torque_workbook(workbook(headers), 'Motor')


@pytest.mark.parametrize('content', [b'', b'not a workbook', b'Motor RPM,Peak Torque,Cont Torque\n0,1,1'])
def test_invalid_formats_are_actionable(content):
    for reader in (motor_torque_worksheets, parse_motor_torque_workbook):
        with pytest.raises(ConfigurationError, match='valid, unencrypted .xlsx'):
            reader(content)


@pytest.mark.parametrize('points', [[], [(0, 10, 5)], [(None, None, None)]])
def test_empty_or_insufficient_data(points):
    with pytest.raises(ConfigurationError, match='at least two complete'):
        parse_motor_torque_workbook(workbook(points=points))


def test_empty_sheet_is_rejected():
    with pytest.raises(ConfigurationError, match='Missing required columns'):
        parse_motor_torque_workbook(workbook(headers=(), points=[]), 'Motor')


@pytest.mark.parametrize('column, value, message', [
    (0, None, 'missing value'), (1, None, 'missing value'), (2, '', 'missing value'),
    (0, 'bad RPM', 'finite numeric'), (1, '20 Nm', 'finite numeric'), (2, 'NaN', 'finite numeric'),
    (0, 'inf', 'finite numeric'), (1, True, 'finite numeric'), (2, '#DIV/0!', 'finite numeric'),
    (1, '=10+2', 'formulas are not accepted'),
    (0, -1, 'nonnegative'), (1, -1, 'nonnegative'), (2, -1, 'positive'), (2, 0, 'positive'),
])
def test_invalid_cells_are_not_dropped_or_zero_filled(column, value, message):
    points = [list(point) for point in POINTS]
    points[1][column] = value
    with pytest.raises(ConfigurationError, match=message) as error:
        parse_motor_torque_workbook(workbook(points=points))
    assert f'{chr(65+column)}3' in str(error.value)


def test_duplicate_rpm_is_rejected_even_with_identical_torque():
    with pytest.raises(ConfigurationError, match='duplicate RPM 0'):
        parse_motor_torque_workbook(workbook(points=[POINTS[0], POINTS[0]]))


def test_blank_rows_and_irrelevant_columns_do_not_change_the_map():
    points = [(*POINTS[0], 'first'), (None, None, None, None),
              *[(*point, None) for point in POINTS[1:]], (None, None, None, None)]
    assert parse_motor_torque_workbook(workbook((*HEADERS, 'Notes'), points)) == LEGACY_ROWS


def test_trimmed_numeric_strings_are_supported_without_unit_conversion():
    points = [[f' {v} ' for v in point] for point in POINTS]
    assert parse_motor_torque_workbook(workbook(points=points)) == LEGACY_ROWS


def test_irrelevant_first_sheet_is_skipped_and_ambiguous_sheets_require_selection():
    content = workbook(('Notes',), [('Instructions',)], [('Ratings', [HEADERS, *POINTS])])
    assert motor_torque_worksheets(content) == ('Ratings',)
    assert parse_motor_torque_workbook(content) == LEGACY_ROWS
    content = workbook(extra_sheets=[('Other Motor', [HEADERS, (0, 10, 5), (100, 8, 4)])])
    assert motor_torque_worksheets(content) == ('Motor', 'Other Motor')
    with pytest.raises(ConfigurationError, match='Select the worksheet'):
        parse_motor_torque_workbook(content)
    assert parse_motor_torque_workbook(content, 'Other Motor')[1][TORQUE] == 8
    with pytest.raises(ConfigurationError, match='does not exist'):
        parse_motor_torque_workbook(content, 'Missing')


def test_no_matching_headers_gives_required_column_instructions():
    content = workbook(('Notes',), [('Instructions',)])
    assert motor_torque_worksheets(content) == ('Motor',)
    with pytest.raises(ConfigurationError, match='Motor RPM, Peak Torque and Cont Torque'):
        parse_motor_torque_workbook(content)


def test_template_has_headers_and_units_but_no_motor_data():
    content = motor_torque_template()
    book = load_workbook(BytesIO(content))
    try:
        assert list(book.active.values) == [HEADERS]
        assert [c.comment.text for c in book.active[1]] == [
            'Required numeric values in rpm.', 'Required numeric values in Nm.', 'Required numeric values in Nm.']
    finally:
        book.close()
    with pytest.raises(ConfigurationError, match='at least two complete'):
        parse_motor_torque_workbook(content)


def test_regression_against_existing_paste_row_parsers_and_interpolation():
    rows = parse_motor_torque_workbook(workbook())
    old_peak = MotorPerformanceCurve.from_rows(LEGACY_ROWS, 17.5, .97)
    new_peak = MotorPerformanceCurve.from_rows(rows, 17.5, .97)
    old_cont = ContinuousTorqueCurve.from_rows(LEGACY_ROWS)
    new_cont = ContinuousTorqueCurve.from_rows(rows)
    assert new_peak == old_peak and new_cont == old_cont
    assert new_peak.calculated_rows() == old_peak.calculated_rows()
    assert new_peak.calculated_rows()[0][POSITIVE] == 2037
    assert new_peak.calculated_rows()[0][NEGATIVE] == -120 * 17.5 / .97
    args = ([0, 1, 3, 4], [50, 250, -600, 800], [60, 45, -40, 20])
    old = analyse_motor(*args, old_cont)
    new = analyse_motor(*args, new_cont)
    assert new == old
    assert new.limits_nm == (57.5, 47.5, 30, 20)
    assert new.flags == (True, False, True, False)
    assert new.maximum_segment_s == 1 and new.total_time_s == 2
    assert analyse_motor([0, 1], [0, 801], [0, 0], new_cont).status == 'unavailable'


@pytest.mark.parametrize('kind', ['electric', 'hybrid'])
def test_excel_and_paste_inputs_preserve_full_engineering_reports(kind, tmp_path):
    from test_sergio_october_15 import csv_for, dataset_for, profile_path
    from vsm_postprocessing.ui_config import generate_reporting_profile_engineering_report
    data = dataset_for(kind)
    for index in range(1, 5):
        data.values[:, data.channel_index(f'electricsystem_em{index}_speed')] = [50, 250, 600, 800]
        data.values[:, data.channel_index(f'electricsystem_em{index}_torque')] = [60, 45, -40, 20]
    source = csv_for(data, tmp_path / f'{kind}.csv')
    results = []
    for label, rows in [('paste', LEGACY_ROWS), ('excel', parse_motor_torque_workbook(workbook()))]:
        results.append(generate_reporting_profile_engineering_report(source, profile_path(kind), tmp_path / label,
            motor_performance=MotorPerformanceCurve.from_rows(rows, 17.5, .97),
            continuous_torque_inputs=ContinuousTorqueInputs(ContinuousTorqueCurve.from_rows(rows),
                final_drive_ratio=17.5, driveline_efficiency=.97)))
    left, right = [result.excel_result for result in results]
    np.testing.assert_array_equal(left.dataset.values, right.dataset.values)
    for name, values in left.math_result.values_by_semantic_name.items():
        np.testing.assert_array_equal(values, right.math_result.values_by_semantic_name[name])
    for attribute in ('statistics', 'kpis', 'unavailable_statistics', 'unavailable_kpis',
                      'kpi_calculation_order', 'diagnostics', 'battery_capacity', 'canonical_metrics'):
        assert getattr(left.statistics_result, attribute) == getattr(right.statistics_result, attribute), attribute
    # Compare all engineering sheets, formulas, caches, and every native chart.
    books = [load_workbook(result.report_path) for result in results]
    try:
        assert books[0].sheetnames == books[1].sheetnames
        with ZipFile(results[0].report_path) as a, ZipFile(results[1].report_path) as b:
            for index, name in enumerate(books[0].sheetnames, 1):
                if name != 'Metadata':  # Run locations and timestamps differ.
                    part = f'xl/worksheets/sheet{index}.xml'
                    assert a.read(part) == b.read(part), name
            for part in a.namelist():
                if part.startswith('xl/charts/chart') and part.endswith('.xml'):
                    assert a.read(part) == b.read(part), part
    finally:
        for book in books:
            book.close()
    presentations = [Presentation(result.presentation_path) for result in results]
    assert len(presentations[0].slides) == len(presentations[1].slides)
    for a, b in zip(presentations[0].slides, presentations[1].slides):
        assert str(a.element.xml) == str(b.element.xml)
        assert [s.image.blob for s in a.shapes if s.shape_type == 13] == [s.image.blob for s in b.shapes if s.shape_type == 13]


@pytest.fixture
def upload_gui():
    from streamlit.testing.v1 import AppTest
    # Streamlit 1.49 AppTest does not drive file_uploader. Replace only that
    # widget boundary; run the real parser, selection, previews and curve UI.
    return AppTest.from_string('''
from unittest.mock import patch
import streamlit as st
from vsm_postprocessing.motor_performance_ui import render_motor_performance_inputs
st.session_state.setdefault('test_uploads', {})
def uploaded(*args, **kwargs):
    st.session_state['test_upload_key'] = kwargs['key']
    return st.session_state.test_uploads.get(kwargs['key'])
with patch.object(st, 'file_uploader', uploaded):
    st.session_state['result'] = render_motor_performance_inputs(st, st.session_state.get('scope', 'Source A'))
''', default_timeout=30).run()


def upload(app, content):
    app.session_state['test_uploads'][app.session_state['test_upload_key']] = BytesIO(content)
    app.run()
    assert not app.exception
    return app.session_state['result']


def test_upload_is_clipboard_independent_survives_reruns_and_replaces_all_values(upload_gui):
    app = upload_gui
    app.checkbox[0].check().run()
    app.number_input[0].set_value(17.5)
    app.number_input[1].set_value(.97)
    first = upload(app, workbook())
    assert first.valid and first.peak.motor_peak_torque_nm == (120, 110, 80, 40)
    assert first.continuous.curve.continuous_torque_nm == (60, 55, 40, 20)
    assert '4 valid RPM points' in app.success[0].value
    assert not any(k.endswith('_table') for k in app.session_state._state.filtered_state)
    app.run(); app.run()
    assert app.session_state['result'] == first
    app.number_input[0].set_value(35).run()
    assert app.session_state['result'].peak.motor_peak_torque_nm == first.peak.motor_peak_torque_nm
    second = upload(app, workbook(points=[(0, 20, 10), (200, 10, 5)]))
    assert second.valid and second.peak.motor_speed_rpm == (0, 200)
    assert second.peak.motor_peak_torque_nm == (20, 10)
    assert second.continuous.curve.continuous_torque_nm == (10, 5)
    invalid = upload(app, b'bad file')
    assert not invalid.valid and invalid.peak is None and invalid.continuous is None
    assert app.error and not app.success


def test_upload_scope_clear_and_explicit_reuse(upload_gui):
    app = upload_gui
    app.checkbox[0].check().run()
    original = upload(app, workbook())
    assert original.valid and original.continuous is not None
    app.session_state['scope'] = 'Source B:Hybrid'
    app.run()
    assert not app.checkbox[0].value and app.session_state['result'].continuous is None
    assert all(n.value is None for n in app.number_input)
    next(b for b in app.button if b.label.startswith('Reuse previous motor')).click().run()
    assert app.session_state['result'].continuous == original.continuous
    next(b for b in app.button if b.label == 'Clear motor curve').click().run()
    assert app.session_state['result'].valid and app.session_state['result'].continuous is None
    assert not app.checkbox[0].value


def test_sheet_selection_and_workbook_replacement_reset_selection(upload_gui):
    app = upload_gui
    app.checkbox[0].check().run()
    content = workbook(extra_sheets=[('Second', [HEADERS, (0, 10, 5), (100, 8, 4)])])
    upload(app, content)
    app.selectbox[0].select('Second').run()
    assert app.session_state['result'].continuous.curve.continuous_torque_nm == (5, 4)
    replacement = workbook(points=[(0, 22, 11), (100, 20, 10)],
        extra_sheets=[('Second', [HEADERS, (0, 8, 4), (100, 6, 3)])])
    result = upload(app, replacement)
    assert app.selectbox[0].value == 'Motor'
    assert result.continuous.curve.continuous_torque_nm == (11, 10)


def test_invalid_upload_blocks_reports_even_with_manual_duration(upload_gui):
    app = upload_gui
    app.checkbox[0].check().run()
    app.checkbox[1].check()
    app.number_input[2].set_value(12)
    app.text_input[0].set_value('External bench calculation')
    result = upload(app, b'bad file')
    assert result.continuous.manual_duration_s == 12
    assert not result.valid and result.peak is None
