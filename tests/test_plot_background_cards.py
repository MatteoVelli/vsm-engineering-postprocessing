"""Plot-container styling only; synthetic engineering fixtures are not motor ratings."""
from copy import deepcopy
import hashlib

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches
import pytest

from vsm_postprocessing import profile_powerpoint_report_engine as ppt
from vsm_postprocessing.continuous_torque import ContinuousTorqueCurve
from vsm_postprocessing.continuous_torque_reporting import ContinuousTorqueInputs
from vsm_postprocessing.excel_report_engine import generate_profile_excel_report
from vsm_postprocessing.motor_performance import MotorPerformanceCurve
from test_sergio_october_15 import csv_for, dataset_for, profile_path


SHADOW = {'blurRad': '152400', 'dist': '50800', 'dir': '5400000',
          'sx': '100000', 'sy': '100000', 'kx': '0', 'ky': '0',
          'algn': 'bl', 'rotWithShape': '0'}


def assert_style(shape):
    geometry = shape.element.xpath('./p:spPr/a:prstGeom')[0]
    assert geometry.get('prst') == 'roundRect'
    assert [(a.get('name'),a.get('fmla')) for a in geometry.xpath('./a:avLst/a:gd')] == [('adj','val 1617')]
    assert shape.element.xpath('./p:spPr/a:solidFill/a:srgbClr')[0].get('val') == 'FFFFFF'
    assert len(shape.element.xpath('./p:spPr/a:ln/a:noFill')) == 1
    assert not shape.element.xpath('./p:style')
    assert not shape.element.xpath('./p:spPr/a:effectDag')
    assert len(shape.element.xpath('./p:spPr/a:effectLst')) == 1
    shadow = shape.element.xpath('./p:spPr/a:effectLst/a:outerShdw')
    assert len(shadow) == 1 and dict(shadow[0].attrib) == SHADOW
    assert shape.element.xpath('./p:spPr/a:effectLst/a:outerShdw/a:srgbClr')[0].get('val') == '0A1524'
    assert shape.element.xpath('./p:spPr/a:effectLst/a:outerShdw/a:srgbClr/a:alpha')[0].get('val') == '13000'


def bounds(shape):
    return shape.left,shape.top,shape.width,shape.height


def assert_inside(shape, deck):
    assert shape.left >= 0 and shape.top >= 0
    assert shape.left+shape.width <= deck.slide_width
    assert shape.top+shape.height <= deck.slide_height


@pytest.mark.parametrize('count',[1,2])
def test_single_and_dual_plot_cards_preserve_picture_bytes_bounds_and_z_order(tmp_path,count):
    image=tmp_path/'figure.png';Image.new('RGB',(300,160),'white').save(image)
    deck=Presentation();deck.slide_width=Inches(13.333);deck.slide_height=Inches(7.5)
    slide=deck.slides.add_slide(deck.slide_layouts[6])
    for i in range(count):
        panel=tuple(Inches(v) for v in (.6+6.1*i,2,5.9 if count==2 else 12.1,4))
        picture=(panel[0]+Inches(.2),panel[1]+Inches(.2),panel[2]-Inches(.4),panel[3]-Inches(.4))
        ppt._add_custom_plot_panel(slide,image,panel,picture)
        card,photo=list(slide.shapes)[-2:]
        assert_style(card)
        assert bounds(card)==panel and bounds(photo)==picture
        assert photo.image.blob==image.read_bytes()
        assert_inside(card,deck);assert_inside(photo,deck)
    assert len(slide.shapes)==2*count
    assert sum(s.shape_type==13 for s in slide.shapes)==count


def test_style_is_idempotent_and_does_not_touch_kpi_cards():
    deck=Presentation();slide=deck.slides.add_slide(deck.slide_layouts[6])
    cards=ppt._add_custom_kpi_cards(slide,{},(('TOTAL ENERGY','example','kWh'),))
    original=[str(s.element.xml) for s in slide.shapes]
    panel=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,0,0,Inches(5),Inches(3))
    ppt._style_plot_background_card(panel);once=str(panel.element.xml)
    ppt._style_plot_background_card(panel)
    assert str(panel.element.xml)==once
    assert [str(s.element.xml) for s in list(slide.shapes)[:-1]]==original
    assert cards[0][0].element.xpath('./p:style')


def test_unavailable_plot_uses_one_card_without_adding_a_picture():
    deck=Presentation();slide=deck.slides.add_slide(deck.slide_layouts[6])
    rect=tuple(Inches(v) for v in (.6,2,8,4))
    ppt._add_custom_plot_panel(slide,None,rect,rect)
    assert_style(slide.shapes[0])
    assert sum(s.shape_type==13 for s in slide.shapes)==0
    assert len([s for s in slide.shapes if s.element.xpath('./p:spPr/a:prstGeom[@prst="roundRect"]')])==1


@pytest.mark.parametrize('kind,expected_cards',[('electric',11),('hybrid',9)])
def test_generated_profiles_change_only_plot_card_style_preserving_notes_and_engineering(tmp_path,monkeypatch,kind,expected_cards):
    data=dataset_for(kind)
    excel=generate_profile_excel_report(csv_for(data,tmp_path/'synthetic.csv'),profile_path(kind),tmp_path/'excel',
        motor_performance=MotorPerformanceCurve((0,100,400,800),(120,100,80,40),17.5,.97),
        continuous_torque_inputs=ContinuousTorqueInputs(ContinuousTorqueCurve((0,200,500,900),(70,60,50,30))))
    workbook_before=hashlib.sha256(excel.report_path.read_bytes()).hexdigest()
    engineering=deepcopy((excel.statistics_result.statistics,excel.statistics_result.kpis,excel.workbook_structure))
    style=ppt._style_plot_background_card
    def legacy(panel):
        panel.fill.solid();panel.fill.fore_color.rgb=RGBColor.from_string('FFFFFF')
        panel.line.fill.background()
    monkeypatch.setattr(ppt,'_style_plot_background_card',legacy)
    result=ppt.build_profile_powerpoint_report(excel,tmp_path/'ppt')
    before=Presentation(result.presentation_path)
    new_cards=[]
    def approved(panel):
        style(panel);new_cards.append(panel)
    monkeypatch.setattr(ppt,'_style_plot_background_card',approved)
    result=ppt.build_profile_powerpoint_report(excel,tmp_path/'ppt')
    after=Presentation(result.presentation_path)
    assert len(before.slides)==len(after.slides)==(17 if kind=='electric' else 16)
    assert len(new_cards)==expected_cards
    changed=0;affected_picture_counts=[]
    for old_slide,new_slide in zip(before.slides,after.slides):
        assert len(old_slide.shapes)==len(new_slide.shapes)
        assert old_slide.notes_slide.notes_text_frame.text==new_slide.notes_slide.notes_text_frame.text
        slide_changes=0
        for old,new in zip(old_slide.shapes,new_slide.shapes):
            assert bounds(old)==bounds(new)
            if old.has_text_frame:assert old.text==new.text
            if old.shape_type==13:assert old.image.blob==new.image.blob
            if str(old.element.xml)!=str(new.element.xml):
                assert_style(new);assert_inside(new,after)
                # All non-style XML, including shape IDs and coordinates, is exact.
                a,b=deepcopy(old.element),deepcopy(new.element)
                for element in (a,b):
                    for child in element.xpath('./p:style|./p:spPr/a:prstGeom/a:avLst|./p:spPr/a:effectLst|./p:spPr/a:effectDag'):
                        child.getparent().remove(child)
                assert str(a.xml)==str(b.xml)
                changed+=1;slide_changes+=1
        if slide_changes:
            pictures=[s for s in new_slide.shapes if s.shape_type==13 and s.top>Inches(2)]
            assert len(pictures)==slide_changes
            for picture in pictures:assert_inside(picture,after)
            affected_picture_counts.append(len(pictures))
    assert changed==expected_cards
    assert set(affected_picture_counts)=={1,2}
    assert hashlib.sha256(excel.report_path.read_bytes()).hexdigest()==workbook_before
    assert (excel.statistics_result.statistics,excel.statistics_result.kpis,excel.workbook_structure)==engineering
