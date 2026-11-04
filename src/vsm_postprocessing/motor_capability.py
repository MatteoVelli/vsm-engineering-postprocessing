"""Read-only motor design curves; never imports or calculates VSM channels.

Only native XY references from the named chart are accepted. Excel's evaluated
cell values are authoritative; uncached formulas are a controlled error.
"""
from __future__ import annotations

import json
import math
import re
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import BadZipFile, ZipFile

from matplotlib.figure import Figure
from openpyxl.utils.cell import get_column_letter, range_to_tuple

from .errors import PowerPointReportError
from .torque_maps_presentation import NS as BASE_NS, _relationships
from .utils import sha256_file

NS = {**BASE_NS, 'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
TITLE = 'MOTOR TORQUE AND POWER'
SLIDE_TITLE = 'Motor Torque & Power Characteristics'
SLIDE_ID = 'motor_torque_power'
SERIES = ('Motor Peak Torque', 'Motor Continuous Torque', 'Motor Peak Power', 'Motor Continuous Power')
UNITS = ('Nm', 'Nm', 'kW', 'kW')
ABS_TOL_KW, REL_TOL = .01, .001


def _text(element):
    return ' '.join(t.text or '' for t in element.findall('.//a:t', NS)) if element is not None else ''


def read_motor_capability(workbook: str | Path) -> dict:
    """Validate all four independent XY arrays and record cell/formula provenance."""
    path = Path(workbook).expanduser().resolve()
    try:
        with ZipFile(path) as archive:
            return _read(archive, path)
    except PowerPointReportError:
        raise
    except (OSError, BadZipFile, KeyError, ValueError, ET.ParseError) as exc:
        raise PowerPointReportError(f'Motor capability workbook {path}: {exc}') from exc


def _read(archive, path):
    book = ET.fromstring(archive.read('xl/workbook.xml'))
    rels = _relationships(archive, 'xl/workbook.xml')
    sheets = {s.attrib['name']: rels[s.attrib[f'{{{NS["r"]}}}id']]
              for s in book.findall('s:sheets/s:sheet', NS)}
    if 'Torque Maps' not in sheets:
        raise PowerPointReportError(f'Motor capability: missing "Torque Maps" sheet in {path}')
    names = {n.attrib['name']: n.text for n in book.findall('s:definedNames/s:definedName', NS)
             if 'localSheetId' not in n.attrib}
    shared = []
    if 'xl/sharedStrings.xml' in archive.namelist():
        shared = [''.join(t.text or '' for t in s.findall('.//s:t', NS))
                  for s in ET.fromstring(archive.read('xl/sharedStrings.xml'))]
    cache = {}

    def cells(reference):
        resolved = names.get(reference, reference)
        sheet, (left, top, right, bottom) = range_to_tuple(resolved)
        if sheet not in sheets or (left != right and top != bottom):
            raise PowerPointReportError(f'Motor capability: unsupported cell range {reference}')
        if sheet not in cache:
            cache[sheet] = {c.attrib['r']: c for c in ET.fromstring(archive.read(sheets[sheet])).findall('s:sheetData/s:row/s:c', NS)}
        result = []
        for row in range(top, bottom + 1):
            for col in range(left, right + 1):
                address = f'{get_column_letter(col)}{row}'
                cell = cache[sheet].get(address)
                value = cell.findtext('s:v', None, NS) if cell is not None else None
                kind = cell.attrib.get('t', 'n') if cell is not None else 'n'
                if kind == 's' and value is not None:
                    value = shared[int(value)]
                elif kind == 'inlineStr':
                    value = ''.join(t.text or '' for t in cell.findall('.//s:t', NS))
                elif kind == 'n':
                    value = float(value) if value else None
                result.append(dict(cell=f"'{sheet}'!{address}", value=value,
                                   formula=cell.findtext('s:f', None, NS) if cell is not None else None,
                                   cell_type=kind))
        return result

    sheet = ET.fromstring(archive.read(sheets['Torque Maps']))
    drawing = sheet.find('s:drawing', NS)
    if drawing is None:
        raise PowerPointReportError('Motor capability: Torque Maps has no native chart.')
    drawing_part = _relationships(archive, sheets['Torque Maps'])[drawing.attrib[f'{{{NS["r"]}}}id']]
    drawing_rels = _relationships(archive, drawing_part)
    matches = []
    for link in ET.fromstring(archive.read(drawing_part)).findall('.//c:chart', NS):
        part = drawing_rels[link.attrib[f'{{{NS["r"]}}}id']]
        chart = ET.fromstring(archive.read(part))
        title = chart.find('c:chart/c:title', NS)
        title_text = _text(title)
        title_ref = title.findtext('c:tx/c:strRef/c:f', '', NS) if title is not None else ''
        if title_ref:
            title_text = str(cells(title_ref)[0]['value'])
        if ' '.join(title_text.upper().split()) == TITLE:
            matches.append((part, chart))
    if len(matches) != 1:
        raise PowerPointReportError(f'Motor capability: expected one "Motor Torque and Power" chart, found {len(matches)}.')
    part, chart = matches[0]
    axes = {axis.find('c:axId', NS).attrib['val']: _text(axis.find('c:title', NS))
            for axis in chart.findall('.//c:plotArea/c:valAx', NS)}
    mapped = {}
    for group in chart.findall('.//c:plotArea/c:scatterChart', NS):
        ids = [a.attrib['val'] for a in group.findall('c:axId', NS)]
        if len(ids) != 2:
            raise PowerPointReportError('Motor capability requires explicit XY axes.')
        x_title, y_title = [axes.get(i, '') for i in ids]
        if not re.search(r'\bmotor\b.*\bspeed\b.*\brpm\b', x_title, re.I):
            raise PowerPointReportError(f'Motor capability speed must identify the motor shaft and rpm: {x_title!r}')
        for item in group.findall('c:ser', NS):
            name_ref = item.findtext('c:tx/c:strRef/c:f', '', NS)
            name = str(cells(name_ref)[0]['value']) if name_ref else item.findtext('c:tx/c:v', '', NS)
            normalized = re.sub(r'\s*[\[(](?:Nm|kW)[\])]\s*$', '', name, flags=re.I).strip()
            canonical = next((s for s in SERIES if s.casefold() == normalized.casefold()), None)
            if canonical is None or canonical in mapped:
                raise PowerPointReportError(f'Motor capability: unexpected or duplicate series {name!r}')
            unit = UNITS[SERIES.index(canonical)]
            quantity = 'Torque' if unit == 'Nm' else 'Power'
            if not re.search(rf'\b{quantity}\b.*\b{unit}\b', y_title, re.I):
                raise PowerPointReportError(f'Motor capability: {canonical} requires {unit} axis, found {y_title!r}')
            explicit_unit = re.search(r'[\[(]([^\])]+)[\])]', name)
            if explicit_unit and explicit_unit[1].casefold() != unit.casefold():
                raise PowerPointReportError(f'Motor capability: invalid series unit {name!r}')
            xr = item.findtext('c:xVal/c:numRef/c:f', '', NS)
            yr = item.findtext('c:yVal/c:numRef/c:f', '', NS)
            if not xr or not yr:
                raise PowerPointReportError(f'Motor capability: missing XY references for {name}')
            xc, yc = cells(xr), cells(yr)
            if len(xc) != len(yc) or len(xc) < 2:
                raise PowerPointReportError(f'Motor capability: invalid XY lengths for {name}')
            for cell in xc + yc:
                if cell['cell_type'] != 'n' or not isinstance(cell['value'], (int, float)) or not math.isfinite(cell['value']):
                    raise PowerPointReportError(f'Motor capability: finite evaluated numeric value required at {cell["cell"]}; recalculate and save Excel if this is a formula.')
            x, y = [c['value'] for c in xc], [c['value'] for c in yc]
            if min(x) < 0:
                raise PowerPointReportError(f'Motor capability: speeds must be nonnegative for {name}')
            mapped[canonical] = dict(name=canonical, source_header=name, header_reference=name_ref,
                                     unit=unit, x_unit='rpm', x_reference=xr, y_reference=yr,
                                     x=x, y=y, x_cells=xc, y_cells=yc, point_count=len(x))
    if set(mapped) != set(SERIES):
        raise PowerPointReportError(f'Motor capability: missing series {sorted(set(SERIES) - set(mapped))}')
    checks, warnings = [], []
    for duty in ('Peak', 'Continuous'):
        torque, power = mapped[f'Motor {duty} Torque'], mapped[f'Motor {duty} Power']
        by_speed = {}
        for speed, value in zip(torque['x'], torque['y']):
            by_speed.setdefault(speed, set()).add(value)
        checked = 0
        for speed, value, cell in zip(power['x'], power['y'], power['y_cells']):
            if speed not in by_speed:
                warnings.append(f'{duty}: no same-speed motor torque for {cell["cell"]} at {speed:g} rpm; no interpolation used.')
                continue
            if len(by_speed[speed]) != 1:
                warnings.append(f'{duty}: ambiguous motor torque at {speed:g} rpm for {cell["cell"]}; no torque chosen from duplicate speeds.')
                continue
            torque_value = next(iter(by_speed[speed]))
            expected = torque_value * speed * 2 * math.pi / 60000
            tolerance = ABS_TOL_KW + REL_TOL * abs(expected)
            checks.append(dict(duty=duty, cell=cell['cell'], rpm=speed, torque_nm=torque_value,
                               spreadsheet_kw=value, expected_kw=expected, tolerance_kw=tolerance))
            if abs(value - expected) > tolerance:
                raise PowerPointReportError(f'Motor capability power discrepancy at {cell["cell"]}: spreadsheet {value:g} kW, T*RPM*2*pi/60000 {expected:g} kW (tolerance {tolerance:g}); source unchanged.')
            checked += speed > 0
        if checked < 3:
            warnings.append(f'{duty}: only {checked} nonzero same-speed checks available; three are required for release acceptance.')
    return dict(source_workbook=str(path), source_sha256=sha256_file(path), chart_part=part,
                source='Torque Maps native chart references and evaluated workbook cells',
                series=[mapped[s] for s in SERIES], axes=axes, warnings=warnings, power_checks=checks,
                power_formula='P[kW] = T[Nm] * Motor Speed[rpm] * 2*pi / 60000',
                tolerance='0.01 kW + 0.1% of calculated power; source values never changed',
                shaft_evidence='Motor series names and Motor Speed [rpm] axes; no wheel/gearbox conversions')


def render_motor_capability(payload: dict, output_dir: Path) -> dict:
    """Export four exact arrays with independent speed grids, dual axes and legend."""
    output_dir.mkdir(parents=True, exist_ok=True)
    fig = Figure(figsize=(11.6, 4.5), layout='constrained', facecolor='white')
    left = fig.subplots()
    right = left.twinx()
    handles = []
    for series, colour, style in zip(payload['series'], ('#2E56D0', '#379047', '#C43D42', '#7451A8'), ('-', '--', '-', '--')):
        axis = left if series['unit'] == 'Nm' else right
        line, = axis.plot(series['x'], series['y'], color=colour, linestyle=style,
                          linewidth=2.2, label=f'{series["name"]} [{series["unit"]}]')
        handles.append(line)
    left.set_xlabel('Motor Speed [rpm]')
    left.set_ylabel('Torque [Nm]')
    right.set_ylabel('Power [kW]')
    left.grid(True, color='#B8C2CC', alpha=.4)
    left.set_axisbelow(True)
    left.legend(handles=handles, loc='upper center', bbox_to_anchor=(.5, 1.18), ncols=2, fontsize=10)
    target = output_dir / 'motor_torque_power.png'
    fig.savefig(target, dpi=240)
    fig.clear()
    result = {**payload, 'image': str(target.resolve())}
    (output_dir / 'motor_capability_data.json').write_text(json.dumps(result, indent=2, allow_nan=False), encoding='utf-8')
    return result


def insert_motor_slide_definition(slides: list[dict]) -> list[dict]:
    """Insert once before operating maps (after steering in older profiles)."""
    result = [s for s in slides if s['slide_id'] != SLIDE_ID]
    anchor = 'torque_maps' if any(s['slide_id'] == 'torque_maps' for s in result) else 'wheel_steering_angles'
    index = next(i for i, s in enumerate(result) if s['slide_id'] == anchor) + (anchor != 'torque_maps')
    result.insert(index, dict(slide_id=SLIDE_ID, type='overview', title=SLIDE_TITLE,
                             subtitle='Motor capability reference / design data', statistics=[], plots=[]))
    return result
