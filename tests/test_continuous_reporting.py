import json
from pathlib import Path
from zipfile import ZipFile

import numpy as np
from openpyxl import load_workbook
from pptx import Presentation
import pytest

from vsm_postprocessing.continuous_torque import CONTINUOUS, ContinuousTorqueCurve
from vsm_postprocessing.continuous_torque_reporting import ContinuousTorqueInputs, assess_continuous_torque
from vsm_postprocessing.errors import ConfigurationError
from vsm_postprocessing.motor_performance import MotorPerformanceCurve, SPEED, TORQUE
from vsm_postprocessing.ui_config import generate_reporting_profile_engineering_report
from test_sergio_october_15 import dataset_for, csv_for, profile_path, profile_for


def curve(): return ContinuousTorqueCurve((0,1000),(10,10))


@pytest.mark.parametrize('duration,source', [(-1,'external'),(float('nan'),'external'),(True,'external'),(1,''),(1,None)])
def test_manual_fallback_requires_finite_nonnegative_duration_and_provenance(duration,source):
    with pytest.raises(ConfigurationError): ContinuousTorqueInputs(manual_duration_s=duration,manual_provenance=source)


def test_automatic_zero_takes_precedence_over_manual_and_unavailable_never_fabricates_zero():
    data=dataset_for('electric')
    assessment=assess_continuous_torque(data,profile_for('electric'),ContinuousTorqueInputs(curve(),123,'External analysis'))
    assert assessment.origin=='automatic' and assessment.duration_s==0
    unsupported=ContinuousTorqueCurve((10000,20000),(1,1))
    unavailable=assess_continuous_torque(data,profile_for('electric'),ContinuousTorqueInputs(unsupported))
    assert unavailable.duration_s is None and unavailable.origin is None
    fallback=assess_continuous_torque(data,profile_for('electric'),ContinuousTorqueInputs(unsupported,12.5,'Bench analysis'))
    assert fallback.duration_s==12.5 and fallback.origin=='manual_fallback'


@pytest.fixture(scope='module')
def reports(tmp_path_factory):
    root=tmp_path_factory.mktemp('continuous_reports')
    data=dataset_for('electric',times=(0,.5,2,3.5))
    for index in range(1,5):
        data.values[:,data.channel_index(f'electricsystem_em{index}_speed')]=100
        data.values[:,data.channel_index(f'electricsystem_em{index}_torque')]=[20,20,0,20]
    source=csv_for(data,root/'New_Electric.csv')
    generate=lambda label,**kw:generate_reporting_profile_engineering_report(source,profile_path('electric'),root/label,**kw)
    return {'base':generate('base'),
        'continuous':generate('continuous',continuous_torque_inputs=ContinuousTorqueInputs(curve())),
        'both':generate('both',motor_performance=MotorPerformanceCurve((0,1000),(30,30),17.5,.97),
                        continuous_torque_inputs=ContinuousTorqueInputs(curve())),
        'manual':generate('manual',continuous_torque_inputs=ContinuousTorqueInputs(manual_duration_s=12.25,manual_provenance='External bench calculation'))}


def test_native_excel_continuous_only_and_combined_peak_curves_preserve_all_original_data(reports):
    for label,expected_counts in (('continuous',[6,4]),('both',[8,8,2])):
        baseline=reports['base'].excel_result;current=reports[label].excel_result
        np.testing.assert_array_equal(current.dataset.values,baseline.dataset.values)
        assert current.statistics_result.statistics==baseline.statistics_result.statistics
        assert current.statistics_result.kpis==baseline.statistics_result.kpis
        original_diagnostics = lambda r: [d for d in r.statistics_result.diagnostics if getattr(d, 'code', '') != 'motor_capability_omitted']
        assert original_diagnostics(current)==original_diagnostics(baseline)
        with ZipFile(baseline.report_path) as a,ZipFile(current.report_path) as b:
            for part in ('sheet1','sheet2','sheet4','sheet5'):
                assert a.read(f'xl/worksheets/{part}.xml')==b.read(f'xl/worksheets/{part}.xml')
        b=load_workbook(current.report_path)
        try:
            assert len(b.worksheets)==7 and 'Motor Duty Summary' in b and 'Motor Curves' not in b
            s=b['Torque Maps']
            assert [len(c.series) for c in s._charts]==expected_counts
            assert s['BK5'].value=='=-BJ5' and s['BJ5'].font.color.rgb=='000000FF'
            assert s['BK5'].font.color.rgb=='00000000'
            if label=='both':
                assert s['BL5'].value=='=BI5/$BE$4'
                assert s['BM5'].value=='=BJ5*$BE$4*$BG$4'
                assert s['BN5'].value=='=BK5*$BE$4/$BG$4'
            else: assert s['BL5'].value is None
        finally:b.close()


def test_only_torque_map_slide_changes_and_duration_is_calculated_from_current_samples(reports):
    a=Presentation(reports['base'].presentation_path)
    b=Presentation(reports['continuous'].presentation_path)
    assert len(a.slides)==len(b.slides)==16
    for i,(old,new) in enumerate(zip(a.slides,b.slides)):
        if i!=14:assert old.element.xml==new.element.xml
        else:
            assert next(s.text for s in new.shapes if s.name=='Continuous torque duration value')=='2 s'
            values={s.name:s.text for s in new.shapes if s.has_text_frame}
            assert values['Continuous torque accumulated value']=='2 s'
            assert values['Continuous torque episodes value']=='2'
            assert values['Continuous torque utilisation value']=='200.0%'
            notes=json.loads(new.notes_slide.notes_text_frame.text)
            assert notes['continuous_torque']['origin']=='automatic'
            assert notes['continuous_torque']['analysis']['motors'][0]['flagged_samples']==3
            charts=notes['torque_maps_presentation']['charts']
            assert [len(c['series']) for c in charts.values()]==[6,4]
            data=json.loads(Path(notes['torque_maps_presentation']['data_path']).read_text())
            baseline=json.loads(Path(json.loads(old.notes_slide.notes_text_frame.text)['torque_maps_presentation']['data_path']).read_text())
            for key in charts:
                assert data['charts'][key]['series'][:4]==baseline['charts'][key]['series']
    assert reports['continuous'].powerpoint_result.displayed_kpi_count==reports['base'].powerpoint_result.displayed_kpi_count+4
    manual=Presentation(reports['manual'].presentation_path)
    assert next(s.text for s in manual.slides[14].shapes if s.name=='Continuous torque duration value')=='12.25 s'
    assert 'External bench calculation' in manual.slides[14].notes_slide.notes_text_frame.text
    assert not any(s.name=='Continuous torque accumulated value' for s in manual.slides[14].shapes)
    assert not any(s.name=='Continuous torque duration value' for s in a.slides[14].shapes)


def test_manifest_records_separate_provenance_without_adding_engineering_kpis(reports):
    base=json.loads(reports['base'].excel_result.manifest_path.read_text())
    current=json.loads(reports['both'].excel_result.manifest_path.read_text())
    assert 'continuous_torque_assessment' not in base
    assert current['continuous_torque_assessment']['duration_s']==2
    assert current['continuous_torque_assessment']['origin']=='automatic'
    ppt=json.loads(reports['both'].manifest_path.read_text())
    assert ppt['continuous_torque_assessment']['duration_s']==2


def test_torque_cards_and_chart_containers_reuse_established_slide_geometry(reports):
    from lxml import etree
    from pptx.util import Inches
    prs=Presentation(reports['both'].presentation_path)
    reference,slide=prs.slides[13],prs.slides[15]
    standard=next(s for s in reference.shapes if s.shape_type==1 and abs(s.top-Inches(1.72))<2)
    cards=[s for s in slide.shapes if s.name.startswith('Continuous torque') and s.name.endswith(' card')]
    assert len(cards)==4
    expected_width=(6*1.89+5*.16-3*.16)/4
    for i,card in enumerate(cards):
        assert (card.top,card.height)==(standard.top,standard.height)
        assert abs(card.width-Inches(expected_width))<2
        assert abs(card.left-Inches(.6+i*(expected_width+.16)))<2
        assert card.fill.fore_color.rgb==standard.fill.fore_color.rgb
        assert card.line.color.rgb==standard.line.color.rgb
        assert card.line.width==standard.line.width
        # Native theme effect reference is the same, including inherited shadow.
        assert etree.tostring(card.element.find('{http://schemas.openxmlformats.org/presentationml/2006/main}style'))==etree.tostring(standard.element.find('{http://schemas.openxmlformats.org/presentationml/2006/main}style'))
    for key in ('duration','accumulated','episodes','utilisation'):
        label=next(s for s in slide.shapes if s.name==f'Continuous torque {key} label')
        value=next(s for s in slide.shapes if s.name==f'Continuous torque {key} value')
        dot=next(s for s in slide.shapes if s.name==f'Continuous torque {key} indicator')
        assert all(r.font.name=='Calibri' and r.font.size.pt==8.5 for p in label.text_frame.paragraphs for r in p.runs)
        assert all(r.font.name=='Calibri' and r.font.size.pt==19 for p in value.text_frame.paragraphs for r in p.runs)
        assert all(abs(a-Inches(b))<2 for a,b in zip((label.top,value.top,dot.top),(1.83,2.26,1.89)))
    pictures=lambda s:[p for p in s.shapes if p.shape_type==13 and p.top>Inches(2)]
    for actual,expected in zip(pictures(slide),pictures(reference)):
        assert (actual.left,actual.top,actual.width,actual.height)==(expected.left,expected.top,expected.width,expected.height)
        assert actual.image.size[0]/actual.image.size[1]==pytest.approx(actual.width/actual.height)
    manual=Presentation(reports['manual'].presentation_path).slides[14]
    manual_cards=[s for s in manual.shapes if s.name.startswith('Continuous torque') and s.name.endswith(' card')]
    assert len(manual_cards)==1 and manual_cards[0].top==standard.top


@pytest.fixture
def editor():
    from streamlit.testing.v1 import AppTest
    return AppTest.from_string('''
import streamlit as st
from vsm_postprocessing.motor_performance_ui import render_motor_performance_inputs
st.session_state['input_result']=render_motor_performance_inputs(st,st.session_state.get('scope','Source A'))
''',default_timeout=20).run()


def run_editor(app,rows):
    from test_motor_performance import editor_key, run_gui
    app.session_state[editor_key(app)]={'edited_rows':rows,'added_rows':[],'deleted_rows':[]}
    run_gui(app)
    assert not app.exception
    return app.session_state['input_result']


def test_continuous_only_does_not_require_peak_ratio_or_efficiency(editor):
    app=editor;app.checkbox[0].check().run()
    result=run_editor(app,{'0':{SPEED:0,CONTINUOUS:100},'1':{SPEED:1000,CONTINUOUS:80}})
    assert result.valid and result.peak is None and result.continuous.curve.continuous_torque_nm==(100,80)
    assert all(n.value is None for n in app.number_input)


def test_bad_continuous_data_does_not_reject_a_valid_peak_curve(editor):
    app=editor;app.checkbox[0].check().run();app.number_input[0].set_value(17.5);app.number_input[1].set_value(.97)
    result=run_editor(app,{'0':{SPEED:0,TORQUE:100,CONTINUOUS:50},'1':{SPEED:0,TORQUE:100,CONTINUOUS:50}})
    assert result.valid and result.peak is not None and result.continuous is None
    assert any('Continuous analysis unavailable' in e.value for e in app.error)


def test_explicit_specification_reuse_never_copies_manual_duration(editor):
    from test_motor_performance import run_gui
    app=editor;app.checkbox[0].check().run()
    run_editor(app,{'0':{SPEED:0,CONTINUOUS:100},'1':{SPEED:1000,CONTINUOUS:80}})
    app.checkbox[1].check();app.number_input[2].set_value(12.5);app.text_input[0].set_value('External test');run_gui(app)
    assert app.session_state['input_result'].continuous.manual_duration_s==12.5
    app.session_state['scope']='Source B';run_gui(app)
    assert app.session_state['input_result'].continuous is None
    next(b for b in app.button if b.label.startswith('Reuse previous motor')).click().run()
    result=app.session_state['input_result']
    assert result.continuous.curve.continuous_torque_nm==(100,80)
    assert result.continuous.manual_duration_s is None
