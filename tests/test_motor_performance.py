"""Optional shared capability inputs and live Excel overlays; never simulation data."""
from dataclasses import FrozenInstanceError
import json
from pathlib import Path
from zipfile import ZipFile

import numpy as np
from openpyxl import load_workbook
from pptx import Presentation
import pytest

from vsm_postprocessing.errors import ConfigurationError
from vsm_postprocessing.motor_performance import (MotorPerformanceCurve, SPEED, TORQUE,
    GENERATOR, WHEEL_SPEED, POSITIVE, NEGATIVE)
from vsm_postprocessing.report_selection import ProfileReportSelection
from vsm_postprocessing.ui_config import generate_reporting_profile_engineering_report
from test_sergio_october_15 import csv_for, dataset_for, profile_path


def curve():
    return MotorPerformanceCurve((0, 175, 350), (1068.5710547955418,)*3, 17.5, .97)


def test_sergio_reference_arithmetic_units_signs_and_immutable_inputs():
    motor = curve(); row = motor.calculated_rows()[1]
    assert row[GENERATOR] == -1068.5710547955418
    assert row[WHEEL_SPEED] == 10
    assert row[POSITIVE] == pytest.approx(18138.99365515432, abs=1e-10)
    assert row[NEGATIVE] == pytest.approx(-19278.343772084518, abs=1e-10)
    assert 'not a universal' in motor.provenance()['assumption']
    with pytest.raises(FrozenInstanceError): motor.final_drive_ratio = 10


@pytest.mark.parametrize('ratio,efficiency', [(0,.97),(-1,.97),(None,.97),(True,.97),
    (float('inf'),.97),(float('nan'),.97),(17.5,0),(17.5,-1),(17.5,1.01),
    (17.5,None),(17.5,False),(17.5,float('nan')),(17.5,float('inf')),(1e-308,1e-308)])
def test_invalid_drivetrain_inputs_are_rejected(ratio, efficiency):
    with pytest.raises(ConfigurationError):
        MotorPerformanceCurve((0,175),(1000,1000),ratio,efficiency)


@pytest.mark.parametrize('rows', [[],[{SPEED:0,TORQUE:1}],
    [{SPEED:0,TORQUE:1},{SPEED:175,TORQUE:None}],
    [{SPEED:'abc',TORQUE:1},{SPEED:175,TORQUE:1}],
    [{SPEED:0,TORQUE:1},{SPEED:175,TORQUE:-1}],
    [{SPEED:float('inf'),TORQUE:1},{SPEED:175,TORQUE:1}],
    [{SPEED:0,TORQUE:True},{SPEED:175,TORQUE:1}]])
def test_invalid_or_incomplete_motor_rows_are_rejected(rows):
    with pytest.raises(ConfigurationError): MotorPerformanceCurve.from_rows(rows,17.5,.97)


def test_curve_length_is_dynamic_blank_rows_are_ignored_and_order_is_preserved():
    rows = [{SPEED:i,TORQUE:1000-i} for i in range(1000)]
    motor = MotorPerformanceCurve.from_rows([{SPEED:None,TORQUE:None},*rows,
        {SPEED:float('nan'),TORQUE:float('nan')}],17.5,.97)
    assert len(motor.motor_speed_rpm) == 1000
    assert motor.motor_speed_rpm == tuple(range(1000))
    signed = MotorPerformanceCurve((175,0,-175),(1,2,3),1,1)
    assert [r[WHEEL_SPEED] for r in signed.calculated_rows()] == [175,0,-175]


@pytest.fixture(scope='module')
def reports(tmp_path_factory):
    folder = tmp_path_factory.mktemp('motor_reports')
    source = csv_for(dataset_for('electric'),folder/'Electric.csv')
    baseline = generate_reporting_profile_engineering_report(source,profile_path('electric'),folder/'baseline')
    optional = generate_reporting_profile_engineering_report(source,profile_path('electric'),folder/'optional',
        motor_performance=curve())
    return baseline, optional


def test_motor_overlays_do_not_change_simulation_calculations_or_diagnostics(reports):
    a,b = (r.excel_result for r in reports)
    np.testing.assert_array_equal(a.dataset.values,b.dataset.values)
    for name, values in a.math_result.values_by_semantic_name.items():
        np.testing.assert_array_equal(values,b.math_result.values_by_semantic_name[name])
    assert a.statistics_result.statistics == b.statistics_result.statistics
    assert a.statistics_result.kpis == b.statistics_result.kpis
    assert a.statistics_result.diagnostics == b.statistics_result.diagnostics
    assert a.statistics_result.battery_capacity == b.statistics_result.battery_capacity
    with ZipFile(a.report_path) as left, ZipFile(b.report_path) as right:
        for part in ('xl/worksheets/sheet1.xml','xl/worksheets/sheet2.xml',
                     'xl/worksheets/sheet4.xml','xl/worksheets/sheet5.xml'):
            assert left.read(part) == right.read(part),part


def test_excel_keeps_original_native_points_and_adds_four_live_formula_series(reports):
    a,b = (load_workbook(r.report_path) for r in reports)
    cached = load_workbook(reports[1].report_path,data_only=True)
    try:
        assert a.sheetnames == b.sheetnames and 'Motor Curves' not in b
        assert len(a.worksheets[0]._charts) == len(b.worksheets[0]._charts)
        old,new = a['Torque Maps'],b['Torque Maps']
        expected_refs = [('AX','AY','AZ'),('BA','BB','BC')]
        expected_names = [('Motor Peak Torque','Generator Peak Torque'),
                          ('Positive Wheel Torque Limit','Negative Wheel Torque Limit')]
        for x,y,refs,names in zip(old._charts,new._charts,expected_refs,expected_names):
            assert len(x.series) == 4 and len(y.series) == 6
            assert y.scatterStyle == 'lineMarker'
            assert x.anchor == y.anchor
            for original, retained in zip(x.series,y.series[:4]):
                assert original.to_tree().attrib == retained.to_tree().attrib
                from openpyxl.xml.functions import tostring
                assert tostring(original.to_tree()) == tostring(retained.to_tree())
                assert retained.graphicalProperties.line.noFill
            for series,label,col in zip(y.series[4:],names,refs[1:]):
                assert series.tx.v == label and series.marker.symbol is None  # OOXML marker "none".
                assert not series.graphicalProperties.line.noFill and not series.smooth
                assert series.xVal.numRef.f == f"'Torque Maps'!${refs[0]}$5:${refs[0]}$7"
                assert series.yVal.numRef.f == f"'Torque Maps'!${col}$5:${col}$7"
            assert y.x_axis.scaling.max is None and y.y_axis.scaling.min is None
        for col in ('D','E','J','K','P','Q','V','W','AB','AC','AH','AI','AN','AO','AT','AU'):
            assert [old[f'{col}{r}'].value for r in range(39,44)] == [new[f'{col}{r}'].value for r in range(39,44)]
        assert new['BE4'].value == 17.5 and new['BG4'].value == .97
        assert [new[f'{c}6'].value for c in ('AZ','BA','BB','BC')] == ['=-AY6','=AX6/$BE$4','=AY6*$BE$4*$BG$4','=AZ6*$BE$4/$BG$4']
        for col,key in zip(('AZ','BA','BB','BC'),(GENERATOR,WHEEL_SPEED,POSITIVE,NEGATIVE)):
            assert cached['Torque Maps'][col+'6'].value == pytest.approx(curve().calculated_rows()[1][key],rel=1e-14)
            assert new[col+'6'].font.color.rgb == '00000000'
        assert new['AX6'].font.color.rgb == new['AY6'].font.color.rgb == '000000FF'
        assert b.calculation.calcMode == 'auto' and not b.calculation.forceFullCalc
        fields = {r[0]:r[1] for r in cached['Metadata'].values}
        provenance = json.loads(fields['Motor performance capability inputs'])
        assert provenance['input_table'] == 'AX4:BG7'
        assert provenance['final_drive_ratio'] == 17.5
    finally:
        a.close();b.close();cached.close()


def test_no_input_has_no_capability_artifacts_and_only_torque_slide_changes(reports):
    baseline, optional = reports
    manifest = json.loads(baseline.excel_result.manifest_path.read_text())
    assert 'motor_performance_curve' not in manifest
    a,b = (Presentation(r.presentation_path) for r in reports)
    assert len(a.slides) == len(b.slides) == 16
    for i,(x,y) in enumerate(zip(a.slides,b.slides)):
        if i==14:
            old=json.loads(x.notes_slide.notes_text_frame.text)['torque_maps_presentation']
            new=json.loads(y.notes_slide.notes_text_frame.text)['torque_maps_presentation']
            assert [len(c['series']) for c in old['charts'].values()]==[4,4]
            assert [len(c['series']) for c in new['charts'].values()]==[6,6]
            continue
        assert x.element.xml == y.element.xml
        assert x.notes_slide.notes_text_frame.text == y.notes_slide.notes_text_frame.text
        assert [s.image.blob for s in x.shapes if s.shape_type == 13] == [s.image.blob for s in y.shapes if s.shape_type == 13]


def test_deselecting_maps_cannot_silently_discard_requested_motor_curve(tmp_path):
    source = csv_for(dataset_for('electric'),tmp_path/'Electric.csv')
    with pytest.raises(ConfigurationError,match='Select both Torque Maps'):
        generate_reporting_profile_engineering_report(source,profile_path('electric'),tmp_path/'report',
            motor_performance=curve(),selection=ProfileReportSelection(plots=()))


def editor_key(app):
    return next(k for k in app.session_state._state.filtered_state if k.endswith('_table'))


def run_gui(app):
    # AppTest 1.49 has no data_editor driver. Send its real JSON widget protocol
    # alongside the supported widgets, as the browser does on every rerun.
    state = app._tree.get_widget_states()
    for editor in app.dataframe:
        if editor.proto.id:
            state.widgets.add(id=editor.proto.id,
                string_value=json.dumps(app.session_state[editor_key(app)]))
    return app._run(state)


@pytest.fixture
def motor_gui():
    from streamlit.testing.v1 import AppTest
    return AppTest.from_string('''
import streamlit as st
from vsm_postprocessing.ui_app import _render_motor_performance_inputs
motor, valid = _render_motor_performance_inputs(st, st.session_state.get('scope', 'source A:current Electric'))
st.session_state['motor_result'] = motor
st.session_state['motor_valid'] = valid
''').run()


def fill_gui(app):
    app.checkbox[0].check().run()
    app.number_input[0].set_value(17.5)
    app.number_input[1].set_value(.97)
    app.session_state[editor_key(app)] = {'edited_rows':{'0':{SPEED:0.,TORQUE:1068.5710547955418},
        '1':{SPEED:175.,TORQUE:1068.5710547955418}},'added_rows':[],'deleted_rows':[]}
    run_gui(app)
    assert not app.exception
    return app


def test_gui_is_optional_has_no_universal_defaults_and_preserves_values_across_reruns(motor_gui):
    app = motor_gui
    assert not app.exception and not app.error
    assert all(v.value is None for v in app.number_input)
    assert app.session_state['motor_result'] is None and app.session_state['motor_valid']
    fill_gui(app)
    assert app.session_state['motor_valid']
    assert app.session_state['motor_result'].calculated_rows()[1][WHEEL_SPEED] == 10
    original = app.session_state['motor_result']
    run_gui(app);run_gui(app)
    assert app.session_state['motor_result'] == original
    app.number_input[0].set_value(35);run_gui(app)
    assert app.session_state['motor_result'].calculated_rows()[1][WHEEL_SPEED] == 5
    app.number_input[1].set_value(.5);run_gui(app)
    assert app.session_state['motor_result'].calculated_rows()[1][POSITIVE] == pytest.approx(1068.5710547955418*35*.5)


def test_gui_rows_can_be_added_removed_and_cleared(motor_gui):
    app = fill_gui(motor_gui)
    app.session_state[editor_key(app)] = {'edited_rows':{'0':{SPEED:0.,TORQUE:1000.},'1':{SPEED:175.,TORQUE:900.}},
        'added_rows':[{SPEED:350.,TORQUE:800.}], 'deleted_rows':[]}
    run_gui(app)
    assert app.session_state['motor_result'].motor_speed_rpm == (0,175,350)
    app.session_state[editor_key(app)] = {'edited_rows':{'0':{SPEED:0.,TORQUE:1000.}},
        'added_rows':[{SPEED:350.,TORQUE:800.}], 'deleted_rows':[1]}
    run_gui(app)
    assert app.session_state['motor_result'].motor_speed_rpm == (0,350)
    next(b for b in app.button if b.label == 'Clear motor curve').click().run()
    assert app.session_state['motor_result'] is None and app.session_state['motor_valid']
    assert not app.checkbox[0].value and all(v.value is None for v in app.number_input)


def test_gui_invalid_inputs_and_unrelated_source_never_reuse_a_curve(motor_gui):
    app = fill_gui(motor_gui)
    app.number_input[1].set_value(0);run_gui(app)
    assert not app.session_state['motor_valid'] and app.error
    app.number_input[1].set_value(.97);run_gui(app)
    assert app.session_state['motor_valid']
    app.session_state['scope'] = 'different content SHA:current Electric'
    run_gui(app)
    assert not app.exception and not app.checkbox[0].value
    assert app.session_state['motor_result'] is None and app.session_state['motor_valid']
    assert all(v.value is None for v in app.number_input)
