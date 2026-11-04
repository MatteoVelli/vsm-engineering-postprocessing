"""Generate real-data slide 15 cases and compare every XY coordinate with Excel."""
import argparse
import hashlib
import json
from pathlib import Path
import sys

from openpyxl import load_workbook
from openpyxl.utils.cell import range_to_tuple
from pptx import Presentation

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from vsm_postprocessing.continuous_torque import ContinuousTorqueCurve
from vsm_postprocessing.continuous_torque_reporting import ContinuousTorqueInputs
from vsm_postprocessing.motor_performance import MotorPerformanceCurve
from vsm_postprocessing.ui_config import generate_reporting_profile_engineering_report


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--source',type=Path,required=True)
    parser.add_argument('--second-source',type=Path,required=True)
    parser.add_argument('--curve',type=Path,required=True)
    parser.add_argument('--output',type=Path,required=True)
    parser.add_argument('--only-reference',action='store_true')
    parser.add_argument('--reporting-inputs',type=Path)
    args=parser.parse_args();args.output.mkdir(parents=True,exist_ok=True)
    points=json.loads(args.curve.read_text())
    cont=ContinuousTorqueCurve(tuple(points['rpm']),tuple(points['continuous']))
    peak=MotorPerformanceCurve(tuple(points['rpm']),tuple(points['peak']),17.5,.97)
    higher=ContinuousTorqueCurve(cont.motor_speed_rpm,tuple(v*2 for v in cont.continuous_torque_nm))
    alternate=MotorPerformanceCurve(tuple(points['rpm']),tuple(v*2 for v in points['peak']),10,.9)
    cases=[('reference',args.source,dict(motor_performance=peak,continuous_torque_inputs=ContinuousTorqueInputs(cont))),
           ('second',args.second_source,dict(motor_performance=peak,continuous_torque_inputs=ContinuousTorqueInputs(cont))),
           ('altered',args.source,dict(motor_performance=alternate,continuous_torque_inputs=ContinuousTorqueInputs(higher))),
           ('peak_only',args.source,dict(motor_performance=peak)),
           ('continuous_only',args.source,dict(continuous_torque_inputs=ContinuousTorqueInputs(cont))),
           ('continuous_wheel',args.source,dict(continuous_torque_inputs=ContinuousTorqueInputs(cont,final_drive_ratio=17.5,driveline_efficiency=.97))),
           ('ratio_only',args.source,dict(continuous_torque_inputs=ContinuousTorqueInputs(cont,final_drive_ratio=17.5))),
           ('efficiency_only',args.source,dict(continuous_torque_inputs=ContinuousTorqueInputs(cont,driveline_efficiency=.97))),
           ('none',args.source,{}),
           ('manual',args.source,dict(continuous_torque_inputs=ContinuousTorqueInputs(manual_duration_s=12.25,manual_provenance='QA external bench calculation'))),
           ('unavailable',args.source,dict(continuous_torque_inputs=ContinuousTorqueInputs(ContinuousTorqueCurve((10000,20000),(1,1)))))]
    profile=Path('config/report_profiles/sergio_2026_10/robosprayer_electric.yaml')
    if args.only_reference:
        cases=cases[:1]
    if args.reporting_inputs:
        inputs=json.loads(args.reporting_inputs.read_text())
        cases=[(label,source,{**kwargs,'reporting_inputs':inputs}) for label,source,kwargs in cases]
    records={}
    index=args.output/'validation.json'
    if index.exists(): records=json.loads(index.read_text())
    for label,source,kwargs in cases:
        if label in records: continue
        print('Generating',label,flush=True)
        existing=list((args.output/'reports'/label/'profile_powerpoint_report').glob('*.pptx'))
        if existing:
            presentation_path=existing[0]
        else:
            report=generate_reporting_profile_engineering_report(source,profile,args.output/'reports'/label,**kwargs)
            presentation_path=report.presentation_path
        prs=Presentation(presentation_path);assert len(prs.slides)==16
        notes=json.loads(prs.slides[14].notes_slide.notes_text_frame.text)
        metadata=notes['torque_maps_presentation']
        excel_path=Path(metadata['source_excel'])
        data=json.loads(Path(metadata['data_path']).read_text())
        book=load_workbook(excel_path,read_only=True,data_only=True)
        cell_values={s.title:{cell.coordinate:cell.value for row in s for cell in row if cell.value is not None} for s in [book['Torque Maps']]}
        compared=0
        for chart in data['charts'].values():
            for series in chart['series']:
                for axis in ('x','y'):
                    sheet,(left,top,right,bottom)=range_to_tuple(series[f'{axis}_reference'])
                    from openpyxl.utils.cell import get_column_letter
                    expected=[cell_values[sheet][f'{get_column_letter(left)}{row}'] for row in range(top,bottom+1)]
                    assert series[axis]==expected,(label,series['name'],axis)
                    compared+=len(expected)
                assert min(series['x'])>chart['x_limits'][0] and max(series['x'])<chart['x_limits'][1]
                assert min(series['y'])>chart['y_limits'][0] and max(series['y'])<chart['y_limits'][1]
        book.close()
        values=[c['display'] for c in metadata['kpi_cards']]
        if label=='reference':assert values==['82 s','828 s','13','123.6%']
        counts=[len(c['series']) for c in data['charts'].values()]
        expected_counts={'reference':[8,8],'second':[8,8],'altered':[8,8],'peak_only':[6,6],
                         'continuous_only':[6,4],'continuous_wheel':[6,6],'ratio_only':[6,4],
                         'efficiency_only':[6,4],'none':[4,4],'manual':[4,4],'unavailable':[6,4]}
        assert counts==expected_counts[label]
        if label in ('none','unavailable'):assert not values
        if label=='manual':assert values==['12.25 s']
        ppt_manifest=json.loads((presentation_path.parent/'profile_powerpoint_report_manifest.json').read_text())
        record=dict(excel=str(excel_path.resolve()),powerpoint=str(presentation_path.resolve()),
                    series_counts=counts,kpis=values,source_samples=ppt_manifest['sample_count'],compared_coordinates=compared,
                    max_absolute_xy_error=0,all_series_inside_axes=True,metadata=metadata,
                    excel_sha256=hashlib.sha256(excel_path.read_bytes()).hexdigest())
        records[label]=record;index.write_text(json.dumps(records,indent=2),encoding='utf-8')
        print('PASS',label,values,counts,'coordinates',compared,flush=True)
    if args.only_reference:
        return
    for label in ('reference','altered','peak_only','continuous_only','continuous_wheel','ratio_only','efficiency_only','manual','unavailable'):
        a,b=(Presentation(records[k]['powerpoint']) for k in ('none',label))
        for i,(old,new) in enumerate(zip(a.slides,b.slides)):
            if i==14:continue
            assert old.element.xml==new.element.xml,(label,i+1)
            assert [s.image.blob for s in old.shapes if s.shape_type==13]==[s.image.blob for s in new.shapes if s.shape_type==13],(label,i+1)
    print('PASS all other slide XML and image bytes unchanged across same-source cases',flush=True)


if __name__=='__main__':main()
