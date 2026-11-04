import json
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET

import numpy as np
import pytest
from openpyxl import Workbook, load_workbook

from vsm_postprocessing.continuous_torque import ContinuousTorqueCurve
from vsm_postprocessing.continuous_torque_reporting import ContinuousTorqueInputs, assess_continuous_torque
from vsm_postprocessing.motor_duty_summary import build_motor_duty_summary
from vsm_postprocessing.motor_duty_excel import write_motor_duty_summary
from vsm_postprocessing.motor_performance import MotorPerformanceCurve
from test_sergio_october_15 import dataset_for, profile_for


def case(torque=None, times=(0, 1, 3, 6)):
    data=dataset_for('electric',times=times)
    torques=torque or ([20,20,0,20],[0,20,20,0],[-20,0,0,0],[0,0,0,0])
    for i,values in enumerate(torques,1):
        data.values[:,data.channel_index(f'electricsystem_em{i}_speed')]=[0,100,600,1000]
        data.values[:,data.channel_index(f'electricsystem_em{i}_torque')]=values
    assessment=assess_continuous_torque(data,profile_for('electric'),ContinuousTorqueInputs(ContinuousTorqueCurve((0,1000),(10,10))))
    return data,assessment


def test_independent_motor_maxima_ties_and_observed_elapsed_time():
    data,a=case();original=data.values.copy();r=build_motor_duty_summary(data,a)
    assert [m['total_time_s'] for m in r['motors']]==[3,5,1,0]
    assert [m['maximum_segment_s'] for m in r['motors']]==[3,5,1,0]
    assert [m['flagged_samples'] for m in r['motors']]==[3,2,1,0]
    assert [m['episode_count'] for m in r['motors']]==[2,1,1,0]
    assert r['kpis']['total_time_s']==dict(value=5,motors=['EM2 / FR'])
    assert r['kpis']['maximum_utilisation']['motors']==['EM1 / FL','EM2 / FR','EM3 / RL']
    assert r['kpis']['fraction_above']['value']==5/6
    np.testing.assert_array_equal(data.values,original)
    assert [(e['start_s'],e['interval_end_s'],e['duration_s']) for e in r['events'][:2]]==[(0,3,3),(6,6,0)]
    assert r['events'][0]['flagged_samples']==2 and r['events'][0]['mean_torque_nm']==20
    assert r['events'][-1]['mode']=='Regeneration'


def test_sample_rms_means_extrema_and_band_partition_preserve_sergio_definitions():
    data,a=case();r=build_motor_duty_summary(data,a)
    for i,m in enumerate(r['motors'],1):
        torque=data.values[:,data.channel_index(f'electricsystem_em{i}_torque')]
        assert m['rms_nm']==pytest.approx(np.sqrt(np.mean(torque**2)))
        assert m['mean_speed_rpm']==425
        assert m['continuous_at_mean_speed_nm']==10
        assert m['rms_continuous_ratio']==pytest.approx(m['rms_nm']/10)
        assert sum(b['sample_counts'][i-1] for b in r['bands'])==4
        energy=sum((b['rms_nm'][i-1] or 0)**2*b['sample_counts'][i-1] for b in r['bands'])
        assert energy==pytest.approx(np.sum(torque**2))
    assert r['motors'][2]['maximum_regeneration_nm']==-20
    assert r['motors'][3]['mean_driving_nm'] is None
    assert any('empty speed band' in x for x in r['diagnostics'])


@pytest.mark.parametrize('rating,above',[(20.000001,False),(21,True),(19,False)])
def test_peak_residuals_are_retained_and_display_screen_is_not_physical_tolerance(rating,above):
    data,a=case(torque=[[rating]*4]*4)
    r=build_motor_duty_summary(data,a,MotorPerformanceCurve((0,1000),(20,20),1,1))
    p=r['motors'][0]['peak']
    assert p['above_display_precision']==above
    assert p['strict_flagged_samples']==(4 if rating>20 else 0)
    assert p['maximum_positive_residual_nm']==pytest.approx(max(0,rating-20))
    assert 'not a manufacturer acceptance tolerance' in p['precision_note']
    assert not any('failure' in x.lower() or 'too small' in x.lower() or 'overheating' in x.lower() for x in r['interpretation'])


@pytest.mark.parametrize('rpm,torque,reason', [((0,900),(20,20),'outside'),((0,1000),(0,0),'positive'),((0,0),(20,20),'unique')])
def test_unsupported_peak_inputs_are_na_without_clamping_or_partial_pass(rpm,torque,reason):
    data,a=case();r=build_motor_duty_summary(data,a,MotorPerformanceCurve(rpm,torque,1,1))
    assert all(m['peak']['status']=='unavailable' and reason in m['peak']['reason'] for m in r['motors'])
    assert len(r['diagnostics'])>0


@pytest.mark.parametrize('mode',['missing','manual','unsupported','partial'])
def test_optional_sheet_is_omitted_without_valid_four_motor_automatic_assessment(mode):
    data,a=case()
    if mode=='missing':a=None
    elif mode=='manual':a=assess_continuous_torque(data,profile_for('electric'),ContinuousTorqueInputs(manual_duration_s=12,manual_provenance='Bench'))
    else:
        if mode=='partial':data.values[0,data.channel_index('electricsystem_em2_speed')]=2000
        c=ContinuousTorqueCurve((10000,20000),(10,10)) if mode=='unsupported' else a.inputs.curve
        a=assess_continuous_torque(data,profile_for('electric'),ContinuousTorqueInputs(c))
    book=Workbook();before=book.sheetnames[:]
    assert write_motor_duty_summary(book,data,profile_for('electric'),a) is None
    assert book.sheetnames==before


@pytest.mark.parametrize('zero_events',[False,True])
def test_native_charts_use_local_numeric_event_and_band_tables_and_fit_dashboard(tmp_path,zero_events):
    data,a=case(torque=[[0]*4]*4 if zero_events else None)
    book=Workbook();book.active['A1']='Untouched'
    r=write_motor_duty_summary(book,data,profile_for('electric'),a)
    sheet=book['Motor Duty Summary']
    assert book.active['A1'].value=='Untouched'
    assert len(sheet._charts)==2 and not sheet._images
    assert sheet.freeze_panes=='A4' and sheet.sheet_view.zoomScale==90
    assert [sheet.cell(6,c).value for c in (1,5,9,13)]==([0]*4 if zero_events else [5,5,5/6,2])
    assert sheet.sheet_properties.tabColor is None
    assert set(sheet.tables)=={'MotorDutySpeedBands','MotorDutyEvents'}
    assert '$A$1:$P$17' in str(sheet.print_area) and r['print_overview']=='A1:P17'
    if not zero_events:
        from openpyxl.utils.cell import range_to_tuple
        for motor, series in zip(r['evidence']['motors'], sheet._charts[0].series):
            _, (col, row, _, _) = range_to_tuple(series.yVal.numRef.f)
            formula = sheet.cell(row,col).value
            if motor['episode_count']==0:
                assert formula is None
                continue
            assert 'COUNTIFS(' in formula and 'SUMIFS(' in formula and motor['motor'] in formula
            assert 'NA()' in formula  # Missing/duplicate keys must not fabricate an event.
    for chart in sheet._charts:
        assert len(chart.series)==4
        assert chart.x_axis.title is not None and chart.y_axis.title is not None
        assert chart.width<=16.2 and chart.height<=8.4
        for series in chart.series:
            assert series.smooth is False
            assert "'Motor Duty Summary'!" in series.xVal.numRef.f
            assert "'Motor Duty Summary'!" in series.yVal.numRef.f
    path=tmp_path/'summary.xlsx';book.save(path)
    reloaded=load_workbook(path)
    assert len(reloaded['Motor Duty Summary']._charts)==2
    with ZipFile(path) as z:
        ns={'c':'http://schemas.openxmlformats.org/drawingml/2006/chart'}
        for name in ('xl/charts/chart1.xml','xl/charts/chart2.xml'):
            chart=ET.fromstring(z.read(name))
            assert len(chart.findall('.//c:scatterChart/c:ser',ns))==4
            assert chart.find('.//c:dispBlanksAs',ns).attrib['val']=='gap'


def test_episode_count_is_dynamic_not_limited_to_reference_thirteen():
    times=tuple(range(81));data=dataset_for('electric',times=times)
    for i in range(1,5):
        data.values[:,data.channel_index(f'electricsystem_em{i}_speed')]=np.linspace(0,3200,len(times))
        data.values[:,data.channel_index(f'electricsystem_em{i}_torque')]=np.tile([20,0],41)[:len(times)]
    a=assess_continuous_torque(data,profile_for('electric'),ContinuousTorqueInputs(ContinuousTorqueCurve((0,100,2000,4000),(10,10,10,10))))
    r=build_motor_duty_summary(data,a)
    assert len(r['events'])==164
    assert len(r['bands'])!=0 and r['bands'][-1]['upper_rpm']>=3200
    assert all(m['episode_count']==41 and m['total_time_s']==40 for m in r['motors'])
