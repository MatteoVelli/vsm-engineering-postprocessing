"""Independently compare continuous-duty assessment with supplied real workbooks."""
import argparse
from bisect import bisect_right
from dataclasses import asdict
import json
from pathlib import Path
import sys

import numpy as np
from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from vsm_postprocessing.continuous_torque import ContinuousTorqueCurve, analyse_dataset
from vsm_postprocessing.importer import ImportOptions, load_data_file
from vsm_postprocessing.report_profile import load_reporting_profile, resolve_profile


def independent(time, speed, torque, curve):
    flags=[]; utilisation=[]
    x,y=curve.motor_speed_rpm,curve.continuous_torque_nm
    for rpm,t in zip(speed,torque):
        rpm=abs(rpm)
        assert x[0]<=rpm<=x[-1]
        i=min(bisect_right(x,rpm)-1,len(x)-2)
        limit=y[i]+(rpm-x[i])*(y[i+1]-y[i])/(x[i+1]-x[i])
        flags.append(abs(t)>limit);utilisation.append(abs(t)/limit)
    durations=[];start=None
    for i,flag in enumerate(flags):
        if flag and start is None:start=time[i]
        if start is not None and (not flag or i==len(flags)-1):
            durations.append(time[i]-start);start=None
    return flags,durations,max(utilisation)


def main():
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('--example',type=Path,required=True)
    p.add_argument('--source',type=Path,action='append',required=True)
    p.add_argument('--output',type=Path,required=True)
    args=p.parse_args();args.output.mkdir(parents=True,exist_ok=False)
    b=load_workbook(args.example,read_only=True,data_only=True)
    points=list(b['Motor Curve'].iter_rows(min_row=4,values_only=True))
    curve=ContinuousTorqueCurve(tuple(r[0] for r in points),tuple(r[3] for r in points))
    reference=list(b['Drive Data'].iter_rows(min_row=4,values_only=True))
    example={'rpm':curve.motor_speed_rpm,'continuous':curve.continuous_torque_nm,'peak':[r[1] for r in points]}
    (args.output/'curve.json').write_text(json.dumps(example,indent=2));b.close()
    profile=load_reporting_profile('config/report_profiles/sergio_2026_10/robosprayer_electric.yaml')
    records={}
    for case,source in enumerate(args.source):
        data=load_data_file(source,ImportOptions(strict=False));resolution=resolve_profile(data,profile)
        time=data.values[:,data.channel_index(resolution.resolved['track_time'].channel.channel_id)]
        results={}
        for label,factor in (('reference',1),('higher',2),('lower',.5),('no_exceedance',10000)):
            current=ContinuousTorqueCurve(curve.motor_speed_rpm,tuple(v*factor for v in curve.continuous_torque_nm))
            result=analyse_dataset(data,profile,current,resolution)
            assert result.available,result.summary()
            for i,motor in enumerate(result.motors,1):
                speed=data.values[:,data.channel_index(resolution.resolved[f'electricsystem_em{i}_speed'].channel.channel_id)]
                torque=data.values[:,data.channel_index(resolution.resolved[f'electricsystem_em{i}_torque'].channel.channel_id)]
                flags,durations,utilisation=independent(time,speed,torque,current)
                assert motor.flags==tuple(flags)
                np.testing.assert_allclose(motor.total_time_s,sum(durations),atol=1e-10)
                np.testing.assert_allclose(motor.maximum_segment_s,max(durations,default=0),atol=1e-10)
                np.testing.assert_allclose(motor.maximum_utilisation,utilisation,rtol=1e-13)
                if case==0 and label=='reference':
                    assert motor.sample_count==len(reference)==2161
                    assert motor.flagged_samples==829 and motor.episode_count==13
                    assert motor.maximum_segment_s==82
                    assert list(motor.flags)==[bool(r[5+(i-1)*5]) for r in reference]
            results[label]=result.summary()
        for i in range(4):
            assert results['higher']['motors'][i]['flagged_samples']<=results['reference']['motors'][i]['flagged_samples']<=results['lower']['motors'][i]['flagged_samples']
            assert results['no_exceedance']['motors'][i]['total_time_s']==0
        records[source.name]=results
        print('PASS',source.name,[(r['maximum_segment_s'],[m['flagged_samples'] for m in r['motors']]) for r in results.values()],flush=True)
    (args.output/'results.json').write_text(json.dumps(records,indent=2))


if __name__=='__main__':main()
