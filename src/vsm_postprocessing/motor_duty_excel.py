"""One optional executive-first worksheet; existing report sheets stay intact."""
from openpyxl.chart import Reference, Series
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.pagebreak import Break
from openpyxl.worksheet.table import Table, TableStyleInfo

from .errors import ExcelReportError
from .excel_formulas import write_formula
from .motor_duty_summary import build_motor_duty_summary

SHEET = 'Motor Duty Summary'
NAVY, BLUE = '17365D', '2E56D0'
COLOURS = (BLUE, '76B947', 'C85778', '7451A8')


def write_motor_duty_summary(workbook, dataset, profile, assessment, peak=None):
    evidence = build_motor_duty_summary(dataset, assessment, peak)
    if evidence is None:
        return None
    sheet = workbook.create_sheet(SHEET)
    sheet.sheet_view.showGridLines = False
    sheet.sheet_view.zoomScale = 90
    sheet.freeze_panes = 'A4'
    sheet.sheet_format.defaultRowHeight = 17
    for i in range(1, 17):
        sheet.column_dimensions[get_column_letter(i)].width = 9.5

    def text(row, first, last, value, *, size=11, bold=False, fill='FFFFFF', colour='30343B', height=None, fmt=None):
        for col in range(first, last + 1):
            cell = sheet.cell(row, col)
            cell.fill = PatternFill('solid', fgColor=fill)
        if last > first:
            sheet.merge_cells(start_row=row, start_column=first, end_row=row, end_column=last)
        cell = sheet.cell(row, first, value if value is not None else 'N/A')
        cell.font = Font(name='Calibri', size=size, bold=bold, color=colour)
        cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='left')
        if fmt:
            # Optional decimal placeholders otherwise leave a trailing decimal
            # separator on whole-second durations in native Excel.
            if fmt.startswith('0.##') and isinstance(value, (int, float)) and float(value).is_integer():
                fmt = fmt.replace('0.##', '0')
            cell.number_format = fmt
        if height:
            sheet.row_dimensions[row].height = height
        return cell

    def section(row, value):
        text(row, 1, 16, value, bold=True, colour=NAVY, height=23)

    def table(row, headers, rows, name=None):
        for col, header in enumerate(headers, 1):
            text(row, col, col, header, size=10, bold=True, fill='F2F2F2', height=42)
        for offset, values in enumerate(rows, 1):
            for col, value in enumerate(values, 1):
                text(row + offset, col, col, value, size=10, height=30)
        if name:
            ref = f'A{row}:{get_column_letter(len(headers))}{row + max(1, len(rows))}'
            native = Table(displayName=name, ref=ref)
            native.tableStyleInfo = TableStyleInfo(name='TableStyleLight1', showRowStripes=False)
            sheet.add_table(native)

    text(1, 1, 16, 'ASTAUTO | Motor Duty Summary', size=20, bold=True, colour=NAVY, height=32)
    # Bound the visible identifier; retain the full source in its comment and provenance.
    identifier = dataset.source_path.stem.replace('_', ' ')
    short_identifier = identifier if len(identifier) <= 100 else identifier[:97] + '...'
    header = text(2, 1, 16, short_identifier, size=10, height=20)
    header.comment = Comment(f'Simulation: {dataset.source_path.stem}\nPowertrain: {profile.metadata.powertrain}\nProfile: {profile.profile_id}', 'ASTAUTO')
    count = len(assessment.inputs.curve.motor_speed_rpm)
    text(3, 1, 16, f"Cycle: {evidence['duration_s']:g} s  |  {evidence['sample_count']:,} samples  |  Supplied continuous curve: {count} points, full RPM coverage", size=10, height=20)

    indicators = [('maximum_segment_s', 'Longest exceedance', '0.##" s"'),
                  ('total_time_s', 'Total above / motor (max)', '0.##" s"'),
                  ('fraction_above', 'Cycle above rating (max)', '0.0%'),
                  ('maximum_utilisation', 'Continuous utilisation (max)', '0.0%')]
    for i, (key, label, fmt) in enumerate(indicators):
        first = i * 4 + 1
        value = evidence['kpis'][key]
        text(5, first, first + 3, label, size=11, height=24)
        text(6, first, first + 3, value['value'], size=22, bold=True, height=32, fmt=fmt)
        responsible = 'All four motors (tie)' if len(value['motors']) == 4 else ', '.join(value['motors'])
        text(7, first, first + 3, responsible, size=10, height=28)
    section(9, 'Above supplied continuous rating')
    columns = [('motor', 'Motor', None, 1, 3),
               ('maximum_segment_s', 'Longest [s]', '0.##', 4, 6),
               ('total_time_s', 'Total [s]', '0.##', 7, 9),
               ('fraction_above', 'Cycle [%]', '0.0%', 10, 11),
               ('maximum_utilisation', 'Max utilisation [%]', '0.0%', 12, 14),
               ('episode_count', 'Episodes', '0', 15, 16)]
    for key, label, fmt, first, last_col in columns:
        text(10, first, last_col, label, bold=True, size=10, fill='F2F2F2', height=25)
        maximum = max(m[key] for m in evidence['motors']) if key != 'motor' else None
        for row, motor in enumerate(evidence['motors'], 11):
            demanding = key != 'motor' and motor[key] == maximum
            c = text(row, first, last_col, motor[key], fmt=fmt, size=11,
                     bold=demanding, height=23)
            if key != 'motor':
                c.alignment = Alignment(horizontal='right', vertical='center', indent=1)
    # Use the engine's existing deterministic observation, including zero-duration flags.
    text(16, 1, 16, evidence['interpretation'][0], size=11, height=23)
    text(17, 1, 16, 'Thermal acceptability requires motor manufacturer confirmation.', size=10, height=22)
    # Keep all technical source ranges stable while bringing the charts directly
    # below the compact overview. Row heights reserve exactly their vertical space.
    for row in range(19, 44):
        sheet.row_dimensions[row].height = 10

    section(45, 'Detailed motor statistics - numerical evidence and definitions')
    metrics = [('maximum_speed_rpm', 'Maximum |speed|', 'rpm', 'Largest absolute motor RPM.'),
               ('mean_speed_rpm', 'Mean |speed|', 'rpm', 'Arithmetic mean of all absolute-RPM samples.'),
               ('maximum_driving_nm', 'Maximum positive driving torque', 'Nm', 'Largest positive source torque; zero when none.'),
               ('maximum_regeneration_nm', 'Maximum regenerative torque', 'Nm', 'Most negative source torque; zero when none.'),
               ('mean_driving_nm', 'Mean positive driving torque', 'Nm', 'Mean over strictly positive torque samples only; N/A when none.'),
               ('rms_nm', 'Full-cycle sample RMS torque', 'Nm', evidence['definitions']['rms']),
               ('continuous_at_mean_speed_nm', 'Continuous limit at mean |speed|', 'Nm', 'Linear interpolation of the supplied continuous curve; no extrapolation.'),
               ('rms_continuous_ratio', 'Indicative RMS / continuous limit', 'ratio', evidence['definitions']['ratio']),
               ('flagged_samples', 'Flagged source samples', 'count', 'Strict abs(torque) > continuous limit, including any flagged terminal sample.'),
               ('maximum_segment_s', 'Maximum consecutive exceedance', 's', 'Maximum duration from the existing validated segmentation.'),
               ('total_time_s', 'Accumulated elapsed exceedance', 's', evidence['definitions']['time']),
               ('fraction_above', 'Fraction of cycle above continuous', '%', 'Accumulated elapsed exceedance divided by observed cycle duration.'),
               ('episode_count', 'Exceedance episodes', 'count', 'Existing contiguous flagged runs; includes a terminal zero-duration run if present.'),
               ('episodes_per_hour', 'Exceedance episode frequency', '1/h', 'Episode count / observed duration in hours.'),
               ('maximum_utilisation', 'Maximum instantaneous continuous use', '%', 'max(abs(torque) / interpolated continuous limit).'),
               ('peak.maximum_utilisation', 'Maximum peak-curve utilisation', '%', evidence['definitions']['peak']),
               ('peak.strict_flagged_samples', 'Strict peak residual samples', 'count', 'All strict positive numerical residuals retained; not confirmed physical overloads.'),
               ('peak.strict_elapsed_s', 'Elapsed strict peak residual time', 's', 'Observed forward intervals; terminal sample has zero duration.'),
               ('peak.maximum_positive_residual_nm', 'Maximum positive peak residual', 'Nm', 'max(0, abs(torque) - interpolated peak); full precision retained.'),
               ('peak.precision_note', 'Peak-check interpretation', '', 'Display resolution is not a manufacturer acceptance tolerance.')]
    text(46, 1, 4, 'Metric / unit', bold=True, fill='F2F2F2', height=24)
    for i, motor in enumerate(evidence['motors'], 5):
        text(46, i, i, motor['motor'], size=10, bold=True, fill='F2F2F2')
    text(46, 9, 16, 'Definition / limitation', bold=True, fill='F2F2F2')
    for row, (key, label, unit, definition) in enumerate(metrics, 47):
        text(row, 1, 4, label + (f' [{unit}]' if unit else ''), size=10, height=42)
        for i, motor in enumerate(evidence['motors'], 5):
            value = motor
            for part in key.split('.'):
                value = value.get(part) if isinstance(value, dict) else None
            if key == 'peak.precision_note':
                p = motor['peak']
                value = 'N/A' if p['status'] != 'available' else 'Above display precision' if p['above_display_precision'] else 'Small residual only' if p['strict_flagged_samples'] else 'No residual'
            text(row, i, i, value, size=10, fmt='0.0%' if unit == '%' else '0.000000' if 'residual_nm' in key else '0.000' if unit != 'count' else '0')
        text(row, 9, 16, definition, size=10)
    band_header = 69
    section(band_header, 'RMS torque by absolute motor-speed band - local chart data')
    headers = ['Lower RPM', 'Centre RPM', 'EM1 RMS [Nm]', 'EM2 RMS [Nm]', 'EM3 RMS [Nm]', 'EM4 RMS [Nm]',
               'Upper RPM', 'EM1 samples', 'EM2 samples', 'EM3 samples', 'EM4 samples']
    band_rows = [[b['lower_rpm'], b['centre_rpm'], *b['rms_nm'], b['upper_rpm'], *b['sample_counts']] for b in evidence['bands']]
    table(band_header + 1, headers, band_rows, 'MotorDutySpeedBands')
    # Empty RMS cells are real blanks, so Excel charts show gaps rather than a
    # false zero. The adjacent sample counts and diagnostics explain each gap.
    for row, band in enumerate(evidence['bands'], band_header + 2):
        for col, value in enumerate(band['rms_nm'], 3):
            if value is None:
                sheet.cell(row, col).value = None
                sheet.cell(row, col).comment = Comment('N/A: no motor samples in this speed band. Chart displays a gap.', 'ASTAUTO')
    end_bands = band_header + 1 + len(band_rows)
    text(end_bands + 1, 1, 16, evidence['definitions']['bands'], size=10, height=30)
    event_section = end_bands + 3
    section(event_section, 'Detailed exceedance event history - all motors, validated segmentation')
    headers = ['Motor', 'Episode', 'Start [s]', 'End [s]', 'Duration [s]', 'Flagged samples', 'Mean RPM',
               'Mean T [Nm]', 'Max T [Nm]', 'Max |T| [Nm]', 'Mean use [%]', 'Max use [%]',
               'Mode (T sign)', 'Last flag [s]', 'First index (1)', 'Last index (1)']
    event_rows = [[e['motor'], e['episode'], e['start_s'], e['interval_end_s'], e['duration_s'], e['flagged_samples'],
                   e['mean_speed_rpm'], e['mean_torque_nm'], e['maximum_torque_nm'], e['maximum_absolute_torque_nm'],
                   e['mean_utilisation'], e['maximum_utilisation'], e['mode'], e['last_flagged_sample_s'],
                   e['first_sample'] + 1, e['last_sample'] + 1] for e in evidence['events']]
    if event_section + len(event_rows) + max(m['episode_count'] for m in evidence['motors']) + len(evidence['diagnostics']) + 50 > 1048576:
        raise ExcelReportError('Motor Duty Summary exceeds Excel row capacity; no event history was truncated.')
    table(event_section + 1, headers, event_rows or [['No episodes'] + [None] * 15], 'MotorDutyEvents')
    start_events = event_section + 2
    for row in range(start_events, start_events + len(event_rows)):
        for col in (11, 12):
            sheet.cell(row, col).number_format = '0.0%'
    empty_row = start_events + max(1, len(event_rows)) + 1
    section(empty_row + 2, 'Episode chart data - keyed by motor and episode; event history may be sorted')
    chart_header = empty_row + 3
    episode_groups = [[e for e in evidence['events'] if e['motor'] == m['motor']] for m in evidence['motors']]
    longest_group = max(1, max(map(len, episode_groups)))
    end_events = start_events + max(1, len(event_rows)) - 1
    for index, events in enumerate(episode_groups):
        first_col = index * 4 + 1
        for offset, label in enumerate(('Episode', 'Start [s]', 'Duration [s]', evidence['motors'][index]['motor'])):
            text(chart_header, first_col + offset, first_col + offset, label, size=10, bold=True, fill='F2F2F2', height=32)
        for offset, event in enumerate(events, 1):
            row = chart_header + offset
            text(row, first_col, first_col, event['episode'], size=10, height=24)
            key = f'${get_column_letter(first_col)}{row}'
            motor = event['motor']
            criteria = f'$A${start_events}:$A${end_events},"{motor}",$B${start_events}:$B${end_events},{key}'
            for column, source_column, value in ((first_col+1, 'C', event['start_s']), (first_col+2, 'E', event['duration_s'])):
                cell = text(row, column, column, value, size=10, fmt='General')
                write_formula(cell, f'=IF(COUNTIFS({criteria})=1,SUMIFS(${source_column}${start_events}:${source_column}${end_events},{criteria}),NA())', value)
    diagnostics_start = chart_header + longest_group + 2
    section(diagnostics_start, 'Definitions, diagnostics and provenance')
    notes = [f"Source: {evidence['source']}", f"SHA-256: {evidence['source_sha256']}",
             f'Profile: {profile.profile_id} | Powertrain: {profile.metadata.powertrain}',
             'Per-motor maxima include exact ties; simultaneous durations are never summed. Bold comparison values identify column maxima.',
             'Capability curve is user-provided, not manufacturer-verified.',
             *evidence['interpretation'], evidence['continuous_policy'], *evidence['definitions'].values(),
             'Assessment values are generated snapshots. Re-run the application after changing simulation or capability inputs. Charts update from their local numeric tables.',
             *evidence['diagnostics']]
    for row, note in enumerate(notes, diagnostics_start + 1):
        text(row, 1, 16, note, size=10, height=45 if len(note) > 230 else 30)
    last = diagnostics_start + len(notes)

    def chart(title, xlabel, ylabel, anchor):
        from .excel_report_engine import _make_scatter_shell
        axis_id = 700001 if anchor == 'A19' else 700011
        result = _make_scatter_shell(title, x_title=xlabel, y_title=ylabel,
                                     x_axis_id=axis_id, y_axis_id=axis_id + 1, y_position='l')
        result.scatterStyle = 'lineMarker'
        result.legend.position = 'b'
        result.width, result.height = 14.2, 8.4
        result.display_blanks = 'gap'
        result.x_axis.scaling.min = 0
        result.y_axis.scaling.min = 0
        from .excel_report_engine import _apply_profile_chart_layout
        sheet.add_chart(result, anchor)
        return result, _apply_profile_chart_layout

    episodes, apply_layout = chart('Continuous-Torque Exceedance Episodes', 'Episode start time [s]', 'Episode duration [s]', 'A19')
    bands, _ = chart('RMS Torque by Motor-Speed Band', 'Motor-speed band centre [rpm]', 'RMS torque [Nm]', 'I19')
    for index, motor in enumerate(evidence['motors']):
        first = chart_header + 1
        last_event = chart_header + max(1, len(episode_groups[index]))
        series = Series(Reference(sheet, min_col=index*4+3, min_row=first, max_row=last_event),
                        Reference(sheet, min_col=index*4+2, min_row=first, max_row=last_event), title=motor['motor'])
        series.graphicalProperties.line.noFill = True
        series.smooth = False
        series.marker.symbol = ('circle', 'diamond', 'triangle', 'square')[index]
        series.marker.size = 10 - index * 2
        series.marker.graphicalProperties.solidFill = 'FFFFFF'
        series.marker.graphicalProperties.line.solidFill = COLOURS[index]
        episodes.series.append(series)
        series = Series(Reference(sheet, min_col=index + 3, min_row=band_header + 2, max_row=end_bands), Reference(sheet, min_col=2, min_row=band_header + 2, max_row=end_bands), title=motor['motor'])
        series.graphicalProperties.line.solidFill = COLOURS[index]
        series.smooth = False
        series.graphicalProperties.line.width = 19050
        series.marker.symbol = 'circle'; series.marker.size = 4
        series.marker.graphicalProperties.solidFill = COLOURS[index]
        series.marker.graphicalProperties.line.solidFill = COLOURS[index]
        bands.series.append(series)
    for item in (episodes, bands):
        apply_layout(item)
    sheet.print_options.horizontalCentered = True
    sheet.page_setup.orientation = 'landscape'
    sheet.page_setup.paperSize = sheet.PAPERSIZE_A3
    sheet.page_setup.fitToWidth, sheet.page_setup.fitToHeight = 1, 0
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.print_area = ['A1:P17', f'A19:P{last}']
    sheet.row_breaks.append(Break(id=17))
    return dict(sheet=SHEET, overview='A1:P43', print_overview='A1:P17',
                event_table=f'A{event_section + 1}:P{start_events + max(1, len(event_rows)) - 1}', evidence=evidence)
