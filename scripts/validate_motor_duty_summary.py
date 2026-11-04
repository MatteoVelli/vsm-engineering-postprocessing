"""Independent workbook/source comparisons and dynamic optional-sheet cases."""
import argparse
import json
from pathlib import Path
import sys

import numpy as np
from openpyxl import Workbook, load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from vsm_postprocessing.continuous_torque import ContinuousTorqueCurve
from vsm_postprocessing.continuous_torque_reporting import ContinuousTorqueInputs, assess_continuous_torque
from vsm_postprocessing.importer import ImportOptions, load_data_file
from vsm_postprocessing.motor_duty_excel import write_motor_duty_summary
from vsm_postprocessing.motor_duty_summary import build_motor_duty_summary
from vsm_postprocessing.motor_performance import MotorPerformanceCurve
from vsm_postprocessing.report_profile import load_reporting_profile, resolve_profile
from vsm_postprocessing.torque_maps_presentation import read_torque_map_series


def main():
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('--output',type=Path,required=True)
    a=p.parse_args();h=a.output
    settings=json.loads((h/'inputs.json').read_text());reference=json.loads((h/'reference.json').read_text())
    expected=json.loads((h/'evidence.json').read_text())
    book=load_workbook(settings['sergio'],read_only=True,data_only=True)
    curve_rows=list(book['Motor Curve'].iter_rows(min_row=4,values_only=True))
    continuous=ContinuousTorqueCurve(tuple(r[0] for r in curve_rows),tuple(r[3] for r in curve_rows))
    peak=MotorPerformanceCurve(tuple(r[0] for r in curve_rows),tuple(r[1] for r in curve_rows),17.5,.97)
    profile=load_reporting_profile('config/report_profiles/sergio_2026_10/robosprayer_electric.yaml')
    source=Path.home()/'Downloads'/Path(expected['source']).name
    dataset=load_data_file(source,ImportOptions(strict=False))
    sergio=build_motor_duty_summary(dataset,assess_continuous_torque(dataset,profile,ContinuousTorqueInputs(continuous)),peak)
    original_curve=continuous
    continuous=ContinuousTorqueCurve(**expected['channel_provenance']['curve'])
    peak_source=next(s for s in read_torque_map_series(Path(reference['excel']))['torque_map_powertrain']['series'] if s['name']=='Motor Peak Torque')
    peak=MotorPerformanceCurve(tuple(peak_source['x']),tuple(peak_source['y']),17.5,.97)
    rows=list(book['Drive Data'].iter_rows(min_row=4,values_only=True))
    time=np.array([r[0] for r in rows]);widths=np.r_[np.diff(time),0.]
    compared=[]
    for i,m in enumerate(expected['motors']):
        speed=np.array([r[1+i*5] for r in rows]);torque=np.array([r[2+i*5] for r in rows])
        limits=np.interp(abs(speed),continuous.motor_speed_rpm,continuous.continuous_torque_nm)
        flags=abs(torque)>limits
        # Independent grouping of flagged indices, including a flagged terminal point.
        indices=np.flatnonzero(flags);groups=np.split(indices,np.flatnonzero(np.diff(indices)>1)+1)
        groups=[g for g in groups if len(g)]
        assert flags.sum()==m['flagged_samples']==829
        assert widths[flags].sum()==m['total_time_s']==828
        assert len(groups)==m['episode_count']==13
        assert max(sum(widths[g]) for g in groups)==m['maximum_segment_s']==82
        assert m['fraction_above']==828/2160
        np.testing.assert_allclose(m['rms_nm'],np.sqrt(np.mean(torque**2)),rtol=1e-14)
        np.testing.assert_allclose(m['rms_nm'],book['Duty Summary'].cell(12,i+2).value,rtol=1e-13)
        for key,row in [('maximum_speed_rpm',7),('mean_speed_rpm',8),('maximum_driving_nm',9),
                        ('maximum_regeneration_nm',10),('mean_driving_nm',11),
                        ('continuous_at_mean_speed_nm',13),('rms_continuous_ratio',14),('maximum_utilisation',18)]:
            np.testing.assert_allclose(sergio['motors'][i][key],book['Duty Summary'].cell(row,i+2).value,rtol=1e-13)
            if key not in ('continuous_at_mean_speed_nm','rms_continuous_ratio','maximum_utilisation'):
                np.testing.assert_allclose(m[key],sergio['motors'][i][key],rtol=1e-13)
        np.testing.assert_allclose(m['continuous_at_mean_speed_nm'],np.interp(speed.mean(),continuous.motor_speed_rpm,continuous.continuous_torque_nm),rtol=1e-14)
        np.testing.assert_allclose(m['maximum_utilisation'],max(abs(torque)/limits),rtol=1e-14)
        np.testing.assert_allclose(m['rms_continuous_ratio'],m['rms_nm']/m['continuous_at_mean_speed_nm'],rtol=1e-14)
        events=[e for e in expected['events'] if e['motor']==m['motor']]
        for group,event in zip(groups,events):
            assert event['first_sample']==group[0] and event['last_sample']==group[-1]
            assert event['duration_s']==sum(widths[group])
            assert event['start_s']==time[group[0]] and event['interval_end_s']==time[min(group[-1]+1,len(time)-1)]
            np.testing.assert_allclose(event['mean_torque_nm'],torque[group].mean(),rtol=1e-14)
            np.testing.assert_allclose(event['mean_utilisation'],(abs(torque[group])/limits[group]).mean(),rtol=1e-14)
        peak_limits=np.interp(abs(speed),peak.motor_speed_rpm,peak.motor_peak_torque_nm)
        peak_flags=abs(torque)>peak_limits
        assert m['peak']['strict_flagged_samples']==int(peak_flags.sum())
        np.testing.assert_allclose(m['peak']['maximum_positive_residual_nm'],max(0,float((abs(torque)-peak_limits).max())),rtol=1e-12,atol=1e-12)
        assert not m['peak']['above_display_precision']
        for band_index,band in enumerate(expected['bands']):
            assert (band['lower_rpm'],band['upper_rpm'])==(band_index*500,(band_index+1)*500)
            mask=(speed>=band['lower_rpm'])&(speed<band['upper_rpm'])
            np.testing.assert_allclose(band['rms_nm'][i],np.sqrt(np.mean(torque[mask]**2)),rtol=1e-14)
            np.testing.assert_allclose(band['rms_nm'][i],book['Duty Summary'].cell(23+band_index,i+2).value,rtol=1e-13)
        compared.append(dict(motor=m['motor'],flagged_samples=829,elapsed_s=828,longest_s=82,episodes=13,
                             rms_nm=m['rms_nm'],maximum_continuous_utilisation=m['maximum_utilisation'],peak=m['peak'],
                             sergio_original_continuous_limit_nm=sergio['motors'][i]['continuous_at_mean_speed_nm'],
                             approved_continuous_limit_nm=m['continuous_at_mean_speed_nm'],
                             all_reference_duty_metrics_match_with_original_curve=True,approved_curve_results_independently_match=True,
                             all_events_independently_matched=True,all_five_speed_bands_match=True))
    book.close()
    sources={'rough':source,'flat':Path('reference_files/RoboSprayer_8500Kg_Electric_02_UoL_Drive_Cycle_20kph_DriverPC_Flat_02.csv')}
    matrix={}
    for name,file in sources.items():
        dataset=load_data_file(file,ImportOptions(strict=False))
        for label,curve in [('reference',continuous),('higher',ContinuousTorqueCurve(continuous.motor_speed_rpm,tuple(v*2 for v in continuous.continuous_torque_nm))),
                            ('zero',ContinuousTorqueCurve((0,10000),(10000,10000))),
                            ('variable',ContinuousTorqueCurve((0,3500),(500,100))),
                            ('unavailable',ContinuousTorqueCurve((10000,20000),(10,10))),('missing',None)]:
            assessment=assess_continuous_torque(dataset,profile,ContinuousTorqueInputs(curve)) if curve else None
            workbook=Workbook();workbook.active.title='Preserved';workbook.active['A1']='unchanged'
            record=write_motor_duty_summary(workbook,dataset,profile,assessment,peak)
            assert workbook['Preserved']['A1'].value=='unchanged'
            case=name+'_'+label
            if label in ('unavailable','missing'):
                assert record is None and workbook.sheetnames==['Preserved'];matrix[case]={'omitted':True};continue
            assert record is not None
            if label=='zero':
                assert not record['evidence']['events']
                assert all(record['evidence']['kpis'][key]['value']==0 for key in ('maximum_segment_s','total_time_s','fraction_above'))
            folder=h/'variants';folder.mkdir(exist_ok=True)
            target=folder/(case+'.xlsx');workbook.save(target)
            matrix[case]=dict(excel=str(target.resolve()),kpis=record['evidence']['kpis'],events=len(record['evidence']['events']),bands=record['evidence']['bands'])
    (h/'numerical_validation.json').write_text(json.dumps(dict(per_motor=compared,variants=matrix,
        continuous_curve_input_max_difference_nm=max(abs(x-y) for x,y in zip(original_curve.continuous_torque_nm,continuous.continuous_torque_nm)),
        differences_from_sergio=['Approved report curve inputs retain their existing six decimal places; original Sergio inputs have more precision. Both sets are checked independently; no input values changed.',
                                '829 flagged samples retained; accumulated elapsed duration is 828 s, not 829 s.',
                                'Cycle share uses 828/2160 = 38.3333%, not sample-count or extended-terminal conventions.',
                                'All four motors are independently analysed; close utilisation values are not tied by rounding.',
                                'Original sample RMS retained; thermal-equivalent and motor-too-small claims removed.',
                                'Strict peak residuals retained and shown as unconfirmed at report display precision.']),indent=2,allow_nan=False))
    print('PASS all four motor summaries, 52 events, sample RMS, 20 motor/band checks, peak residuals and 12 dynamic cases')


if __name__=='__main__':main()
