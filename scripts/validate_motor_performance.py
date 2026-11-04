"""Compare real GUI/report capability inputs with Sergio's untouched workbook."""
import argparse
import hashlib
import io
import json
from pathlib import Path
import sys
from zipfile import ZipFile

import numpy as np
from openpyxl import load_workbook
from openpyxl.utils.cell import range_to_tuple
from PIL import Image
from pptx import Presentation

sys.path.insert(0, str(Path(__file__).resolve().parents[1]/'src'))
from vsm_postprocessing.motor_performance import MotorPerformanceCurve, SPEED, TORQUE, GENERATOR, WHEEL_SPEED, POSITIVE, NEGATIVE
from vsm_postprocessing.ui_config import generate_reporting_profile_engineering_report


def numerical_record(report):
    stats = report.statistics_result
    return {'raw_sha256': hashlib.sha256(report.dataset.values.tobytes()).hexdigest(),
            'math': {k: hashlib.sha256(v.tobytes()).hexdigest() for k,v in report.math_result.values_by_semantic_name.items()},
            'statistics': {s.definition.statistic_id:s.value for s in stats.statistics},
            'kpis': {k.definition.kpi_id:k.value for k in stats.kpis}}


def compare_powerpoint(left, right):
    a,b = Presentation(left),Presentation(right)
    assert len(a.slides) == len(b.slides)
    for i,(x,y) in enumerate(zip(a.slides,b.slides),1):
        assert x.element.xml == y.element.xml, i
        for p,q in zip([s for s in x.shapes if s.shape_type == 13], [s for s in y.shapes if s.shape_type == 13]):
            p,q = Image.open(io.BytesIO(p.image.blob)).convert('RGBA'),Image.open(io.BytesIO(q.image.blob)).convert('RGBA')
            assert p.size == q.size and p.tobytes() == q.tobytes(),i


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--example', type=Path, required=True)
    parser.add_argument('--source', type=Path, action='append', required=True)
    parser.add_argument('--output', type=Path, required=True)
    parser.add_argument('--baseline', type=Path, required=True)
    args = parser.parse_args();root=args.output.resolve();root.mkdir(exist_ok=False)
    expected=json.loads(args.baseline.read_text())
    formulas=load_workbook(args.example);cached=load_workbook(args.example,data_only=True)
    sheet=formulas['Torque Maps'];numbers=cached['Torque Maps']
    chart=sheet._charts[0]
    _,(x,first,_,last)=range_to_tuple(chart.series[4].xVal.numRef.f)
    _,(y,_,_,_)=range_to_tuple(chart.series[4].yVal.numRef.f)
    rows=[{SPEED:numbers.cell(r,x).value,TORQUE:numbers.cell(r,y).value} for r in range(first,last+1)]
    ratio,efficiency=numbers['AE4'].value,numbers['AG4'].value
    motor=MotorPerformanceCurve.from_rows(rows,ratio,efficiency)
    expected_outputs=np.array([[numbers.cell(r,c).value for c in (26,27,28,29)] for r in range(first,last+1)])
    actual_outputs=np.array([[row[k] for k in (GENERATOR,WHEEL_SPEED,POSITIVE,NEGATIVE)] for row in motor.calculated_rows()])
    np.testing.assert_allclose(actual_outputs,expected_outputs,rtol=2e-14,atol=1e-11)
    (root/'example_points.json').write_text(json.dumps({'rows':rows,'ratio':ratio,'efficiency':efficiency,
        'example':str(args.example.resolve()),'example_sha256':hashlib.sha256(args.example.read_bytes()).hexdigest(),
        'expected_outputs':expected_outputs.tolist(),'max_absolute_difference':float(np.max(np.abs(actual_outputs-expected_outputs)))},indent=2))
    formulas.close();cached.close()
    baseline_report=None;records={}
    for index, source in enumerate(args.source):
        kind='hybrid' if 'hybrid' in source.name.lower() else 'electric'
        profile=Path('config/report_profiles/sergio_2026_10')/f'robosprayer_{kind}.yaml'
        if index==0:
            print('Generating real no-input baseline',flush=True)
            baseline_report=generate_reporting_profile_engineering_report(source,profile,root/'no_input')
            assert numerical_record(baseline_report.excel_result)=={k:v for k,v in expected[source.name].items() if k!='source_sha256'}
            assert 'motor_performance_curve' not in json.loads(baseline_report.excel_result.manifest_path.read_text())
        print('Generating real capability overlays:',source.name,flush=True)
        result=generate_reporting_profile_engineering_report(source,profile,root/f'case_{index}',motor_performance=motor)
        report=result.excel_result
        assert numerical_record(report)=={k:v for k,v in expected[source.name].items() if k!='source_sha256'}
        book=load_workbook(report.report_path);values=load_workbook(report.report_path,data_only=True)
        try:
            maps=book['Torque Maps'];series_counts=[]
            assert 'Motor Curves' not in book
            for chart,key in zip(maps._charts,('torque_map_powertrain','torque_map_wheels')):
                states=[s for s in report.plotting_result.series_diagnostics[key] if s['included']]
                series_counts.append(len(chart.series))
                assert len(chart.series)==len(states)+2
                for series,state in zip(chart.series,states):
                    for reference,semantic in ((series.xVal.numRef.f,state['x']),(series.yVal.numRef.f,state['y'])):
                        name,extent=reference.rsplit('!',1);assert name=="'Torque Maps'"
                        np.testing.assert_allclose([r[0].value for r in maps[extent]],
                            report.plotting_result.values_by_semantic_name[semantic],rtol=2e-14,atol=0)
                for series in chart.series[-2:]:
                    assert not series.graphicalProperties.line.noFill
            output=np.array([[values['Torque Maps'].cell(r,c).value for c in (52,53,54,55)] for r in range(5,len(rows)+5)])
            np.testing.assert_allclose(output,expected_outputs,rtol=2e-14,atol=1e-11)
            assert maps['BE4'].value==ratio and maps['BG4'].value==efficiency
            if kind=='electric':
                assert series_counts==[6,6]
                assert len(book.worksheets[0]._charts)==25
            if index==0:
                with ZipFile(baseline_report.report_path) as a,ZipFile(result.report_path) as b:
                    for part in a.namelist():
                        if part.startswith('xl/charts/') and part.endswith('.xml'):
                            # Only the Torque Maps charts gain capability series.
                            if a.read(part)==b.read(part):continue
                            assert b'Motor Peak Torque' in b.read(part) or b'Positive Wheel Torque Limit' in b.read(part),part
                        if part in ('xl/worksheets/sheet1.xml','xl/worksheets/sheet2.xml','xl/worksheets/sheet4.xml','xl/worksheets/sheet5.xml'):
                            assert a.read(part)==b.read(part),part
                compare_powerpoint(baseline_report.presentation_path,result.presentation_path)
        finally:
            book.close();values.close()
        records[source.name]={'excel':str(result.report_path),'powerpoint':str(result.presentation_path),
            'excel_checks':{'series_counts':series_counts},'points':len(rows),
            'max_absolute_difference_from_example':float(np.max(np.abs(output-expected_outputs))),
            'numerics_equal_pretask_baseline':True,'kpis':numerical_record(report)['kpis'],
            'statistics':numerical_record(report)['statistics'],'motor':motor.provenance()}
        (root/'validation.json').write_text(json.dumps(records,indent=2))
        print('PASS',source.name,series_counts,flush=True)


if __name__=='__main__':main()
