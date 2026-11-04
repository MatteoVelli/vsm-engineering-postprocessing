"""Slide-only rendering of native Excel XY series and cached numeric cells.

No engineering conversion, interpolation or extrapolation belongs here.
"""
from __future__ import annotations

import json
import math
import posixpath
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from matplotlib.figure import Figure
from openpyxl.utils.cell import get_column_letter, range_to_tuple

from .errors import PowerPointReportError
from .plotting_engine import PlotDefaults, _apply_axis_format, _draw_zero_line

NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
      'd': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'}
PEAK = {'Motor Peak Torque', 'Generator Peak Torque',
        'Positive Wheel Torque Limit', 'Negative Wheel Torque Limit'}
CONTINUOUS = {'Motor Continuous Torque', 'Generator Continuous Torque',
              'Positive Continuous Wheel Torque Limit', 'Negative Continuous Wheel Torque Limit'}
KEYS = ('torque_map_powertrain', 'torque_map_wheels')
COLOURS = ('#2E56D0', '#76B947', '#C85778', '#7451A8')


def torque_maps_metadata(payload):
    """Keep notes compact; exact arrays remain in the referenced data artifact."""
    return {**payload, 'charts': {key: {**chart, 'series': [
        {k: v for k, v in s.items() if k not in {'x', 'y'}} | {'point_count': len(s['x'])}
        for s in chart['series']]} for key, chart in payload['charts'].items()}}


def _relationships(archive, part):
    folder, name = posixpath.split(part)
    root = ET.fromstring(archive.read(f'{folder}/_rels/{name}.rels'))
    return {r.attrib['Id']: posixpath.normpath(posixpath.join(folder, r.attrib['Target'])).lstrip('/')
            for r in root if r.attrib.get('TargetMode') != 'External'}


def read_torque_map_series(workbook: Path) -> dict:
    """Read exact XY values/references through native workbook relationships."""
    with ZipFile(workbook) as archive:
        root = ET.fromstring(archive.read('xl/workbook.xml'))
        book_rels = _relationships(archive, 'xl/workbook.xml')
        sheets = {s.attrib['name']: book_rels[s.attrib[f'{{{NS["r"]}}}id']]
                  for s in root.findall('s:sheets/s:sheet', NS)}
        if 'Torque Maps' not in sheets:
            return {}
        sheet_part = sheets['Torque Maps']
        sheet = ET.fromstring(archive.read(sheet_part))
        drawing = sheet.find('s:drawing', NS)
        if drawing is None:
            return {}
        drawing_part = _relationships(archive, sheet_part)[drawing.attrib[f'{{{NS["r"]}}}id']]
        drawing_rels = _relationships(archive, drawing_part)
        anchors = sorted(ET.fromstring(archive.read(drawing_part)),
                         key=lambda a: int(a.findtext('d:from/d:col', '0', NS)))
        cell_cache = {}

        def values(reference):
            name, (left, top, right, bottom) = range_to_tuple(reference)
            if name not in sheets or left != right:
                raise PowerPointReportError(f'Unsupported Torque Maps numeric range: {reference}')
            if name not in cell_cache:
                cells = ET.fromstring(archive.read(sheets[name])).findall('s:sheetData/s:row/s:c', NS)
                cell_cache[name] = {c.attrib['r']: c.findtext('s:v', None, NS)
                                    for c in cells if c.attrib.get('t', 'n') == 'n'}
            result = []
            for row in range(top, bottom + 1):
                value = cell_cache[name].get(f'{get_column_letter(left)}{row}')
                numeric = float(value) if value is not None else None
                if numeric is not None and not math.isfinite(numeric):
                    raise PowerPointReportError(f'Non-finite Torque Maps value: {reference}')
                result.append(numeric)
            return result

        result = {}
        for key, anchor in zip(KEYS, anchors):
            link = anchor.find('.//c:chart', NS)
            if link is None:
                continue
            chart_part = drawing_rels[link.attrib[f'{{{NS["r"]}}}id']]
            chart = ET.fromstring(archive.read(chart_part))
            series = []
            for item in chart.findall('.//c:scatterChart/c:ser', NS):
                name = item.findtext('c:tx/c:v', '', NS)
                xr = item.findtext('c:xVal/c:numRef/c:f', '', NS)
                yr = item.findtext('c:yVal/c:numRef/c:f', '', NS)
                if not xr or not yr:
                    raise PowerPointReportError('Torque Maps chart requires numeric XY cell references.')
                x, y = values(xr), values(yr)
                if len(x) != len(y):
                    raise PowerPointReportError(f'Unequal Torque Maps XY ranges: {name}')
                if not any(a is not None and b is not None for a, b in zip(x, y)):
                    continue  # Do not draw the native empty-chart placeholder.
                kind = 'peak' if name in PEAK else 'continuous' if name in CONTINUOUS else 'simulation'
                series.append(dict(name=name, kind=kind, x_reference=xr, y_reference=yr, x=x, y=y))
            result[key] = dict(chart_part=chart_part, series=series)
        return result


def _limits(series, coordinate):
    values = [v for s in series for v in s[coordinate] if v is not None]
    low, high = min(values), max(values)
    padding = (high - low) * .07 if high != low else max(abs(high) * .07, 1.)
    return [low - padding, high + padding]


def render_torque_maps(workbook: Path, output_dir: Path) -> dict:
    """Render slide assets without changing the workbook or original plots."""
    charts = read_torque_map_series(workbook)
    output_dir.mkdir(parents=True, exist_ok=True)
    for key, chart in charts.items():
        series = chart['series']
        chart['image'] = None
        if not series:
            continue
        defaults = PlotDefaults()
        style = defaults.style
        figure = Figure(figsize=(defaults.width_inches, defaults.height_inches),
                        constrained_layout=style.constrained_layout, facecolor=style.background)
        figure.patch.set_alpha(0.0 if style.transparent else 1.0)
        ax = figure.subplots()
        ax.set_facecolor(style.background)
        ax.set_title('Motor Torque vs Motor Speed' if key == KEYS[0] else 'Wheel-Side Torque vs Wheel Speed',
                     fontsize=style.title_fontsize, pad=12)
        handles, labels, kinds = [], [], set()
        simulation_index = 0
        for s in series:
            x = [float('nan') if v is None else v for v in s['x']]
            y = [float('nan') if v is None else v for v in s['y']]
            if s['kind'] == 'simulation':
                handle = ax.scatter(x, y, s=6, alpha=.55, linewidths=0,
                                    color=COLOURS[simulation_index % len(COLOURS)], rasterized=True)
                simulation_index += 1
                handles.append(handle); labels.append(s['name'])
            else:
                peak = s['kind'] == 'peak'
                handle, = ax.plot(x, y, color='#D98A2B' if peak else '#008C95',
                                  linestyle='-' if peak else '--', linewidth=1.9)
                if s['kind'] not in kinds:
                    handles.append(handle); labels.append('Peak (+ / -)' if peak else 'Continuous (+ / -)')
                    kinds.add(s['kind'])
        ax.set_xlabel('Motor speed [rpm]' if key == KEYS[0] else 'Wheel speed [rpm]', fontsize=style.axis_label_fontsize)
        ax.set_ylabel('Torque [Nm]', fontsize=style.axis_label_fontsize)
        chart['x_limits'], chart['y_limits'] = _limits(series, 'x'), _limits(series, 'y')
        ax.set_xlim(chart['x_limits']); ax.set_ylim(chart['y_limits'])
        ax.tick_params(labelsize=style.tick_fontsize)
        _apply_axis_format(ax, ax.get_xlabel(), axis_name='x')
        _apply_axis_format(ax, ax.get_ylabel(), axis_name='y')
        ax.grid(True, color='#B8C2CC', alpha=style.grid_alpha, linewidth=style.grid_linewidth)
        ax.set_axisbelow(True)
        if style.zero_line:
            _draw_zero_line(ax)
        for spine in ax.spines.values():
            spine.set_color('#7A8691'); spine.set_linewidth(.8)
        # Use the standard report legend and layout; two columns keep the
        # additional capability entries compact inside the numeric axes.
        ax.legend(handles, labels, loc=style.legend_location, ncols=2,
                  fontsize=style.legend_fontsize, frameon=True, framealpha=.92,
                  facecolor='white', edgecolor='#B8C2CC')
        if style.tight_layout and not style.constrained_layout:
            figure.tight_layout()
        target = output_dir / f'{key}.png'
        figure.savefig(target, dpi=defaults.dpi)
        figure.clear()
        chart['image'] = str(target.resolve())
    payload = dict(source_excel=str(workbook.resolve()), source='Native Excel XY series and cached numeric source cells',
                   calculations='None; no conversion, interpolation or extrapolation in the presentation layer',
                   styles={'peak': 'orange solid', 'continuous': 'teal dashed', 'simulation': 'small scatter'}, charts=charts)
    data_path = output_dir / 'torque_maps_presentation_data.json'
    payload['data_path'] = str(data_path.resolve())
    data_path.write_text(json.dumps(payload, indent=2, allow_nan=False), encoding='utf-8')
    return payload
