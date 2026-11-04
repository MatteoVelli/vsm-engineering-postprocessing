"""Presentation provenance, dynamic scales and per-motor aggregation."""
import hashlib
import json
from pathlib import Path
from types import SimpleNamespace

import pytest
from openpyxl import Workbook
from openpyxl.chart import ScatterChart, Series, Reference
from matplotlib.figure import Figure
from matplotlib.backends.backend_agg import FigureCanvasAgg

from vsm_postprocessing.continuous_torque_ppt import continuous_torque_cards
from vsm_postprocessing.torque_maps_presentation import read_torque_map_series, render_torque_maps, torque_maps_metadata
from vsm_postprocessing.plotting_engine import PlotDefaults, PlotDefinition, PlotSeriesDefinition, _render_one_plot


def workbook(path, factor=1, points=3):
    book = Workbook()
    sheet = book.active
    sheet.title = 'Torque Maps'
    for i in range(points):
        sheet.append([i*137, (-1)**i * 250*factor, 20+i, (-1)**i*4000*factor])
    for col, anchor, names in [(1, 'A6', ['EM1 / FL', 'Motor Peak Torque', 'Generator Peak Torque']),
                                (3, 'N6', ['FL Wheel', 'Positive Wheel Torque Limit', 'Negative Wheel Torque Limit'])]:
        chart = ScatterChart()
        for name in names:
            chart.series.append(Series(Reference(sheet, min_col=col+1, min_row=1, max_row=points),
                                       Reference(sheet, min_col=col, min_row=1, max_row=points), title=name))
        sheet.add_chart(chart, anchor)
    book.save(path)


@pytest.mark.parametrize('factor,points', [(1,3),(19,11),(.25,2)])
def test_native_sources_are_exact_signed_and_dynamic_without_mutating_excel(tmp_path, factor, points):
    path=tmp_path/'source.xlsx'; workbook(path, factor, points)
    digest=hashlib.sha256(path.read_bytes()).hexdigest()
    result=render_torque_maps(path,tmp_path/'assets')
    motor,wheel=result['charts'].values()
    for chart,expected in [(motor,250*factor),(wheel,4000*factor)]:
        assert chart['series'][0]['y']==[(-1)**i*expected for i in range(points)]
        assert [s['kind'] for s in chart['series']]==['simulation','peak','peak']
        assert chart['y_limits'][0]<-expected and chart['y_limits'][1]>expected
        assert Path(chart['image']).is_file()
        for s in chart['series']:
            assert len(s['x'])==points  # Only original points, no extrapolated or interpolated samples.
    assert hashlib.sha256(path.read_bytes()).hexdigest()==digest
    artifact=json.loads(Path(result['data_path']).read_text())
    assert artifact['charts']==result['charts']
    compact=torque_maps_metadata(result)
    assert compact['charts']['torque_map_powertrain']['series'][0]['point_count']==points
    assert 'x' not in compact['charts']['torque_map_powertrain']['series'][0]


def test_no_torque_sheet_does_not_invent_a_chart(tmp_path):
    book=Workbook();path=tmp_path/'empty.xlsx';book.save(path)
    assert read_torque_map_series(path)=={}


def test_complete_curve_styles_and_legends_fit_the_rendered_canvas(tmp_path, monkeypatch):
    book=Workbook();sheet=book.active;sheet.title='Torque Maps'
    sheet.append([0,100,-100,70,-70]);sheet.append([3500,20,-20,14,-14])
    names=['EM1 / FL','EM2 / FR','EM3 / RL','EM4 / RR','Motor Peak Torque',
           'Generator Peak Torque','Motor Continuous Torque','Generator Continuous Torque']
    for anchor in ('A6','N6'):
        chart=ScatterChart()
        for name,col in zip(names,[2,2,2,2,2,3,4,5]):
            chart.series.append(Series(Reference(sheet,min_col=col,min_row=1,max_row=2),
                                       Reference(sheet,min_col=1,min_row=1,max_row=2),title=name))
        sheet.add_chart(chart,anchor)
    source=tmp_path/'styles.xlsx';book.save(source)
    original=Figure.savefig
    inspected=[]

    def inspect(figure,*args,**kwargs):
        canvas=FigureCanvasAgg(figure);canvas.draw();renderer=canvas.get_renderer()
        ax=figure.axes[0]
        defaults=PlotDefaults()
        assert tuple(figure.get_size_inches())==(defaults.width_inches,defaults.height_inches)
        assert ax.title.get_fontweight()=='normal'
        assert ax.title.get_fontsize()==defaults.style.title_fontsize
        assert len(ax.collections)==4
        assert [(line.get_color(),line.get_linestyle()) for line in ax.lines]==[
            ('#D98A2B','-'),('#D98A2B','-'),('#008C95','--'),('#008C95','--'),('#4D5963','-')]
        assert list(ax.lines[-1].get_ydata())==[0,0]
        legend=ax.get_legend().get_window_extent(renderer)
        title=ax.title.get_window_extent(renderer)
        assert 0<=legend.x0<legend.x1<=figure.bbox.width
        assert 0<=legend.y0<legend.y1<=figure.bbox.height
        assert title.y0>legend.y1
        assert ax.bbox.x0<=legend.x0<legend.x1<=ax.bbox.x1
        assert ax.bbox.y0<=legend.y0<legend.y1<=ax.bbox.y1
        assert ax.get_position().width*ax.get_position().height>=.70
        assert ax.get_position().height>.75
        # Locators also create off-axis ticks that Matplotlib does not draw.
        x_ticks=[t for t in ax.get_xticklabels() if ax.get_xlim()[0]<=t.get_position()[0]<=ax.get_xlim()[1]]
        y_ticks=[t for t in ax.get_yticklabels() if ax.get_ylim()[0]<=t.get_position()[1]<=ax.get_ylim()[1]]
        for text in [ax.title,ax.xaxis.label,ax.yaxis.label,*x_ticks,*y_ticks]:
            if not text.get_visible():continue
            bounds=text.get_window_extent(renderer)
            assert 0<=bounds.x0<bounds.x1<=figure.bbox.width
            assert 0<=bounds.y0<bounds.y1<=figure.bbox.height
        inspected.append(True)
        original(figure,*args,**kwargs)

    monkeypatch.setattr(Figure,'savefig',inspect)
    render_torque_maps(source,tmp_path/'assets')
    assert inspected==[True,True]


@pytest.mark.parametrize('capability', ['peak','continuous','none'])
def test_standard_internal_legend_adapts_to_partial_inputs(tmp_path,monkeypatch,capability):
    book=Workbook();sheet=book.active;sheet.title='Torque Maps'
    sheet.append([0,100,-100]);sheet.append([1e6,20,-20])
    names=['EM1 / FL','EM2 / FR','EM3 / RL','EM4 / RR']
    if capability=='peak':names+=['Motor Peak Torque','Generator Peak Torque']
    if capability=='continuous':names+=['Motor Continuous Torque','Generator Continuous Torque']
    for anchor in ('A6','N6'):
        chart=ScatterChart()
        for i,name in enumerate(names):
            chart.series.append(Series(Reference(sheet,min_col=3 if i==5 else 2,min_row=1,max_row=2),
                                       Reference(sheet,min_col=1,min_row=1,max_row=2),title=name))
        sheet.add_chart(chart,anchor)
    path=tmp_path/'partial.xlsx';book.save(path)
    original=Figure.savefig;observed=[]
    def inspect(figure,*args,**kwargs):
        canvas=FigureCanvasAgg(figure);canvas.draw();renderer=canvas.get_renderer();ax=figure.axes[0]
        legend=ax.get_legend();bounds=legend.get_window_extent(renderer)
        assert ax.bbox.y0<=bounds.y0<bounds.y1<=ax.bbox.y1
        assert ax.bbox.x0<=bounds.x0<bounds.x1<=ax.bbox.x1
        assert 0<=bounds.x0<bounds.x1<=figure.bbox.width
        assert ax.title.get_window_extent(renderer).y0>bounds.y1
        rows={round(t.get_window_extent(renderer).y0,2) for t in legend.get_texts()}
        assert len(rows)==(2 if capability=='none' else 3)
        assert len(legend.get_texts())==(4 if capability=='none' else 5)
        assert ax.get_position().width*ax.get_position().height>=.70
        for text in [ax.xaxis.label,ax.yaxis.label]:
            b=text.get_window_extent(renderer)
            assert 0<=b.x0<b.x1<=figure.bbox.width and 0<=b.y0<b.y1<=figure.bbox.height
        observed.append(True);original(figure,*args,**kwargs)
    monkeypatch.setattr(Figure,'savefig',inspect)
    render_torque_maps(path,tmp_path/'assets')
    assert observed==[True,True]


def test_torque_titles_axes_and_legend_match_the_standard_report_renderer(tmp_path,monkeypatch):
    """Compare actual standard rendering, including title offset and layout engine."""
    source=tmp_path/'source.xlsx';workbook(source)
    snapshots=[]
    original=Figure.savefig
    def inspect(figure,*args,**kwargs):
        FigureCanvasAgg(figure).draw();ax=figure.axes[0];legend=ax.get_legend()
        frame=legend.get_frame()
        snapshots.append(dict(
            title_font=ax.title.get_fontproperties(),title_position=ax.title.get_position(),
            title_alignment=(ax.title.get_ha(),ax.title.get_va()),
            title_pad=ax.titleOffsetTrans.get_matrix().copy(),
            axis_fonts=(ax.xaxis.label.get_fontproperties(),ax.yaxis.label.get_fontproperties()),
            tick_font=ax.get_xticklabels()[0].get_fontproperties(),
            legend_font=legend.get_texts()[0].get_fontproperties(),legend_loc=legend._loc,
            legend_frame=(legend.get_frame_on(),frame.get_alpha(),frame.get_facecolor(),frame.get_edgecolor()),
            grid=(ax.get_xgridlines()[0].get_color(),ax.get_xgridlines()[0].get_alpha(),ax.get_xgridlines()[0].get_linewidth()),
            spines=[(s.get_edgecolor(),s.get_linewidth()) for s in ax.spines.values()],
            axes_top=ax.get_position().y1,layout=type(figure.get_layout_engine()),
        ))
        original(figure,*args,**kwargs)
    monkeypatch.setattr(Figure,'savefig',inspect)
    definition=PlotDefinition('standard','Motor Torque vs Motor Speed','x',
        (PlotSeriesDefinition('y'),),'standard.png',x_label='Motor speed [rpm]',primary_y_label='Torque [Nm]')
    channels={k:SimpleNamespace(display_name=k,unit='rpm' if k=='x' else 'Nm') for k in ('x','y')}
    _render_one_plot(definition,PlotDefaults(),tmp_path/'standard',channels,{'x':[0,137,274],'y':[250,-250,250]},3)
    render_torque_maps(source,tmp_path/'torque')
    standard,*torque=snapshots
    import numpy as np
    for actual in torque:
        assert np.array_equal(actual.pop('title_pad'),standard['title_pad'])
        assert actual=={k:v for k,v in standard.items() if k!='title_pad'}


def test_aggregation_keeps_independent_maxima_ties_and_never_sums_motors():
    motors=[SimpleNamespace(motor=name,maximum_segment_s=segment,total_time_s=total,
                            episode_count=count,maximum_utilisation=util)
            for name,segment,total,count,util in [('EM1',82,828,13,1.2),('EM2',82,800,10,1.235678),
                                                ('EM3',80,828,14,1.1),('EM4',79,700,9,1.0)]]
    assessment=SimpleNamespace(duration_s=82,origin='automatic',analysis=SimpleNamespace(motors=motors))
    cards=continuous_torque_cards(assessment)
    assert [c['display'] for c in cards]==['82 s','828 s','14','123.6%']
    assert [c['responsible_motors'] for c in cards]==[['EM1','EM2'],['EM1','EM3'],['EM3'],['EM2']]


def test_missing_and_manual_analysis_cannot_fabricate_other_indicators():
    assert continuous_torque_cards(None)==[]
    assert continuous_torque_cards(SimpleNamespace(duration_s=None))==[]
    manual=SimpleNamespace(duration_s=12.25,origin='manual_fallback',inputs=SimpleNamespace(manual_provenance='Bench run'))
    cards=continuous_torque_cards(manual)
    assert len(cards)==1 and cards[0]['display']=='12.25 s' and cards[0]['provenance']=='Bench run'
