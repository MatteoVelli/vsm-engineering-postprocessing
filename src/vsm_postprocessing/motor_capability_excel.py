"""Live shaft-power outputs from this run's editable motor reference inputs."""
import math

from openpyxl.chart import Reference, Series
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils.cell import range_to_tuple

from .excel_formulas import write_formula


def write_motor_capability(workbook, peak, assessment):
    """Add one four-curve chart, preserving the existing operating-map charts."""
    previous = getattr(workbook, '_motor_capability_record', None)
    if previous is not None:
        return previous
    continuous = assessment.inputs.curve if assessment is not None else None
    reason = None
    if peak is None or continuous is None:
        reason = 'Supply both peak and continuous motor torque curves for this run.'
    elif 'Torque Maps' not in workbook:
        reason = 'Include the Torque Maps plots in this report.'
    if reason:
        return {'status': 'omitted', 'warning': 'Motor Torque and Power chart and slide omitted: ' + reason}

    sheet = workbook['Torque Maps']
    inputs = ((peak.motor_speed_rpm, peak.motor_peak_torque_nm, 'Motor Peak Torque'),
              (continuous.motor_speed_rpm, continuous.continuous_torque_nm, 'Motor Continuous Torque'))
    sources = []
    for speeds, torques, name in inputs:
        if (len(speeds) < 2 or any(x < 0 or not math.isfinite(x) for x in speeds)
                or any(b <= a for a, b in zip(speeds, speeds[1:]))
                or any(not math.isfinite(t * r * 2 * math.pi / 60000) for t, r in zip(torques, speeds))):
            return {'status': 'omitted', 'warning': 'Motor Torque and Power chart and slide omitted: '
                    'each motor speed grid must be nonnegative and strictly increasing, with finite shaft power.'}
        matches = [s for s in sheet._charts[0].series if s.tx and s.tx.v == name]
        if len(matches) != 1:
            raise ValueError('Missing current-run motor input series: ' + name)
        series = matches[0]
        refs = []
        for reference, expected in ((series.xVal.numRef.f, speeds), (series.yVal.numRef.f, torques)):
            tab, (left, top, right, bottom) = range_to_tuple(reference)
            if tab != sheet.title or left != right or tuple(sheet.cell(r, left).value for r in range(top, bottom + 1)) != tuple(expected):
                raise ValueError('Motor capability input provenance mismatch: ' + reference)
            refs.append((left, top, bottom))
        sources.append(refs)

    # A6/N6 contain the existing maps. AA6:AL32 is a reserved vacant chart
    # region above the original data tables (row 34) and left of AX inputs.
    # If a future layout uses it, append beyond all occupied columns instead.
    anchor = 'AA6'
    if ([c.anchor for c in sheet._charts] != ['A6', 'N6'] or any(sheet.cell(r, c).value is not None
            for r in range(6, 33) for c in range(27, 39))):
        anchor = sheet.cell(6, sheet.max_column + 2).coordinate
    output_col = sheet.max_column + 2
    if anchor != 'AA6':
        output_col += 16

    from .excel_report_engine import _make_scatter_shell, _apply_profile_chart_layout
    primary = _make_scatter_shell('Motor Torque and Power', x_title='Motor Speed [rpm]',
        y_title='Torque [Nm]', x_axis_id=91001, y_axis_id=91002, y_position='l')
    secondary = _make_scatter_shell('', x_title='Motor Speed [rpm]',
        y_title='Power [kW]', x_axis_id=91004, y_axis_id=91003, y_position='r')
    # Excel requires a complete axis pair for a native secondary XY group.
    # Both pairs reference identical RPM arrays, so their automatic bounds stay
    # aligned after edits. Only the primary speed axis is visible.
    secondary.x_axis.delete = True
    secondary.x_axis.majorGridlines = None
    secondary.y_axis.crosses = 'max'
    secondary.y_axis.majorGridlines = None
    primary.x_axis.numFmt = secondary.x_axis.numFmt = '0'
    primary.y_axis.numFmt = '0'
    secondary.y_axis.numFmt = '0.0'
    primary.width, primary.height = 23, 10.5
    records = []
    for i, ((speeds, torques, name), ((xc, start, end), (yc, ystart, yend))) in enumerate(zip(inputs, sources)):
        col = output_col + i
        header = sheet.cell(4, col, name.replace('Torque', 'Power') + ' [kW]')
        header.font = Font(name='Calibri', bold=True, color='000000')
        header.fill = PatternFill('solid', fgColor='E2EFDA')
        header.alignment = Alignment(wrap_text=True, vertical='top')
        sheet.column_dimensions[header.column_letter].width = 27
        sheet.cell(2, col, 'Calculated motor-shaft power').font = Font(name='Calibri', bold=True, color='1F4E78')
        sheet.cell(3, col, 'Black: calculated; torque/RPM inputs in blue').alignment = Alignment(wrap_text=True)
        for offset, (rpm, torque) in enumerate(zip(speeds, torques)):
            target = sheet.cell(5 + offset, col)
            formula = f'={sheet.cell(ystart + offset, yc).coordinate}*{sheet.cell(start + offset, xc).coordinate}*2*PI()/60000'
            write_formula(target, formula, torque * rpm * 2 * math.pi / 60000)
            target.font = Font(name='Calibri', color='000000')
            target.number_format = '0.000000'
        xref = Reference(sheet, min_col=xc, min_row=start, max_row=end)
        torque_ref = Reference(sheet, min_col=yc, min_row=ystart, max_row=yend)
        power_ref = Reference(sheet, min_col=col, min_row=5, max_row=4 + len(speeds))
        for chart, ref, title, color in ((primary, torque_ref, name, ('2E56D0', '379047')[i]),
                (secondary, power_ref, name.replace('Torque', 'Power'), ('C43D42', '7451A8')[i])):
            series = Series(ref, xref, title=title)
            series.marker.symbol = 'none'
            series.smooth = False
            series.graphicalProperties.line.solidFill = color
            series.graphicalProperties.line.width = 25400
            if i:
                series.graphicalProperties.line.prstDash = 'dash'
            chart.series.append(series)
        records.append({'name': name, 'speed_reference': str(xref), 'torque_reference': str(torque_ref),
                        'power_reference': str(power_ref), 'point_count': len(speeds)})
    primary += secondary
    _apply_profile_chart_layout(primary)
    sheet.add_chart(primary, anchor)
    record = {'status': 'included', 'sheet': sheet.title, 'chart_title': 'Motor Torque and Power',
              'chart_anchor': anchor, 'source': 'Current-run user-provided motor-shaft reference curves',
              'formula': 'torque_Nm * speed_rpm * 2*PI()/60000', 'series': records,
              'drivetrain_conversion_applied': False}
    workbook._motor_capability_record = record
    return record
