"""Add native capability envelopes without touching the eight simulation tables."""
from openpyxl.chart import Reference, Series
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

from .excel_formulas import write_formula
from .motor_performance import SPEED, TORQUE, GENERATOR, WHEEL_SPEED, POSITIVE, NEGATIVE, ASSUMPTION


def write_motor_performance(workbook, curve):
    sheet = workbook['Torque Maps']
    # AX:BG is beyond the complete original eight data blocks (ending at AU).
    # Keep charts at A6/N6 and every original local numeric cell in place.
    headers = (SPEED, TORQUE, GENERATOR, WHEEL_SPEED, POSITIVE, NEGATIVE)
    for column, label in enumerate(headers, 50):
        cell = sheet.cell(4, column, label)
        cell.font = Font(name='Calibri', bold=True, color='0000FF' if column < 52 else '000000')
        cell.fill = PatternFill('solid', fgColor='D9EAF7')
        cell.alignment = Alignment(wrap_text=True, vertical='top')
        sheet.column_dimensions[cell.column_letter].width = 24
    for col in (56, 58):
        sheet.column_dimensions[sheet.cell(4, col).column_letter].width = 25
    for col in (57, 59):
        sheet.column_dimensions[sheet.cell(4, col).column_letter].width = 16
    sheet['BD4'] = 'Final-drive ratio [dimensionless]'
    sheet['BE4'] = curve.final_drive_ratio
    sheet['BF4'] = 'Driveline efficiency [fraction]'
    sheet['BG4'] = curve.driveline_efficiency
    for key in ('BE4', 'BG4'):
        sheet[key].font = Font(name='Calibri', color='0000FF')
        sheet[key].number_format = '0.000000'
    for key in ('BD4', 'BF4'):
        sheet[key].alignment = Alignment(wrap_text=True, vertical='top')
    sheet.merge_cells('AX2:BG2'); sheet['AX2'] = 'Motor Performance Curves — user inputs in blue; calculated values in black'
    sheet['AX2'].font = Font(name='Calibri', bold=True, color='1F4E78', size=14)
    sheet.merge_cells('AX3:BG3'); sheet['AX3'] = ASSUMPTION
    sheet['AX3'].alignment = Alignment(wrap_text=True, vertical='top')
    sheet['A4'] = 'Motor capability inputs and formulas: AX4:BG' + str(len(curve.motor_speed_rpm)+4)
    sheet['A4'].hyperlink = "#'Torque Maps'!AX4"
    sheet['A4'].style = 'Hyperlink'
    sheet['A3'] = sheet['A3'].value.replace('No capability curves or drivetrain conversions.',
        'VSM samples unchanged; optional user capability inputs and their explicit drivetrain conversion are at AX4.')
    for row, rpm, torque, outputs in zip(range(5, 5+len(curve.motor_speed_rpm)),
                                       curve.motor_speed_rpm, curve.motor_peak_torque_nm, curve.calculated_rows()):
        sheet.cell(row, 50, rpm); sheet.cell(row, 51, torque)
        formulas = (f'=-AY{row}', f'=AX{row}/$BE$4', f'=AY{row}*$BE$4*$BG$4', f'=AZ{row}*$BE$4/$BG$4')
        for col, formula, name in zip(range(52, 56), formulas, (GENERATOR, WHEEL_SPEED, POSITIVE, NEGATIVE)):
            write_formula(sheet.cell(row, col), formula, outputs[name])
        for col in range(50, 56):
            sheet.cell(row, col).font = Font(name='Calibri', color='0000FF' if col < 52 else '000000')
            sheet.cell(row, col).number_format = '0.000000'
    ratio_validation = DataValidation(type='decimal', operator='greaterThan', formula1=0, allow_blank=False)
    efficiency_validation = DataValidation(type='custom', formula1='AND(ISNUMBER(BG4),BG4>0,BG4<=1)', allow_blank=False)
    for validation, key in ((ratio_validation, 'BE4'), (efficiency_validation, 'BG4')):
        validation.showErrorMessage = True; validation.errorStyle = 'stop'
        validation.error = 'Enter a positive ratio, or efficiency greater than zero and no greater than one.'
        sheet.add_data_validation(validation); validation.add(sheet[key])
    end = 4+len(curve.motor_speed_rpm)
    for chart, x_col, ys in ((sheet._charts[0], 50, ((51, 'Motor Peak Torque'), (52, 'Generator Peak Torque'))),
                              (sheet._charts[1], 53, ((54, 'Positive Wheel Torque Limit'), (55, 'Negative Wheel Torque Limit')))):
        # The existing VSM series explicitly have no lines; only the two new
        # series have visible lines. Keep straight manufacturer segments.
        chart.scatterStyle = 'lineMarker'
        for index, (y_col, label) in enumerate(ys):
            series = Series(Reference(sheet, min_col=y_col, min_row=5, max_row=end),
                            Reference(sheet, min_col=x_col, min_row=5, max_row=end), title=label)
            series.marker.symbol = 'none'
            series.graphicalProperties.line.solidFill = ('17365D', 'C0504D')[index]
            series.graphicalProperties.line.width = 25400
            series.graphicalProperties.line.noFill = False
            series.smooth = False
            chart.series.append(series)
    return {'sheet': sheet.title, 'input_table': f'AX4:BG{end}',
            'ratio_cell': 'BE4', 'efficiency_cell': 'BG4', **curve.provenance()}
