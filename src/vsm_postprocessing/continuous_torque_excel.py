"""Native editable continuous-limit overlays on the existing Torque Maps sheet."""
from openpyxl.chart import Reference, Series
from openpyxl.styles import Alignment, Font, PatternFill

from .excel_formulas import write_formula
from .continuous_torque import CONTINUOUS, POLICY
from .motor_performance import SPEED
from .errors import ConfigurationError


def write_continuous_torque(workbook, assessment, peak=None, selected_ids=None):
    curve = assessment.inputs.curve
    if curve is None or 'Torque Maps' not in workbook:
        return None
    sheet = workbook['Torque Maps']
    ratio = assessment.inputs.final_drive_ratio
    efficiency = assessment.inputs.driveline_efficiency
    if peak is not None and ((ratio is not None and ratio != peak.final_drive_ratio)
                             or (efficiency is not None and efficiency != peak.driveline_efficiency)):
        raise ConfigurationError('Peak and continuous wheel overlays must use the same explicit drivetrain parameters.')
    if ratio is None and peak is not None:
        ratio = peak.final_drive_ratio
    if efficiency is None and peak is not None:
        efficiency = peak.driveline_efficiency
    wheel = ratio is not None and efficiency is not None
    if wheel and peak is None:
        for address, value in (('BE4', ratio), ('BG4', efficiency)):
            sheet[address] = value
            sheet[address].font = Font(name='Calibri', color='0000FF')
            sheet[address].number_format = '0.000000'
        sheet['BD4'] = 'Final-drive ratio [dimensionless]'
        sheet['BF4'] = 'Driveline efficiency [fraction]'
    headers = (SPEED, CONTINUOUS, 'Generator continuous torque [Nm]', 'Continuous wheel speed [rpm]',
               'Positive continuous wheel torque [Nm]', 'Negative continuous wheel torque [Nm]')
    for col, text in enumerate(headers, 61):
        cell = sheet.cell(4, col, text)
        cell.font = Font(name='Calibri', bold=True, color='0000FF' if col < 63 else '000000')
        cell.fill = PatternFill('solid', fgColor='D9EAF7')
        cell.alignment = Alignment(wrap_text=True, vertical='top')
        sheet.column_dimensions[cell.column_letter].width = 28
    sheet.merge_cells('BI2:BN2'); sheet['BI2'] = 'Continuous capability — inputs in blue; formulas in black'
    sheet['BI2'].font = Font(name='Calibri', bold=True, color='1F4E78', size=14)
    sheet.merge_cells('BI3:BN3')
    sheet['BI3'] = ('One shared curve; symmetric regeneration assumed. No extrapolation. '
                    'Wheel limits require explicit ratio and efficiency. Re-run the app after Excel edits to update assessment metrics.')
    sheet['BI3'].alignment = Alignment(wrap_text=True, vertical='top')
    link = 'A5' if peak else 'A4'
    sheet[link] = 'Continuous capability inputs and formulas: BI4:BN' + str(len(curve.motor_speed_rpm)+4)
    sheet[link].hyperlink = "#'Torque Maps'!BI4"; sheet[link].style = 'Hyperlink'
    sheet['A3'] = sheet['A3'].value.replace('No capability curves or drivetrain conversions.',
        'VSM samples unchanged; user-provided continuous capability inputs are at BI4.')
    for row, rpm, torque in zip(range(5, 5+len(curve.motor_speed_rpm)), curve.motor_speed_rpm, curve.continuous_torque_nm):
        sheet.cell(row,61,rpm); sheet.cell(row,62,torque)
        write_formula(sheet.cell(row,63), f'=-BJ{row}', -torque)
        if wheel:
            for col, formula, cached in ((64,f'=BI{row}/$BE$4',rpm/ratio),
                                        (65,f'=BJ{row}*$BE$4*$BG$4',torque*ratio*efficiency),
                                        (66,f'=BK{row}*$BE$4/$BG$4',-torque*ratio/efficiency)):
                write_formula(sheet.cell(row,col),formula,cached)
        for col in range(61,67):
            sheet.cell(row,col).font = Font(name='Calibri',color='0000FF' if col < 63 else '000000')
            sheet.cell(row,col).number_format = '0.000000'
    end = len(curve.motor_speed_rpm)+4
    definitions = ((0,'torque_map_powertrain',61,((62,'Motor Continuous Torque'),(63,'Generator Continuous Torque'))),
                   (1,'torque_map_wheels',64,((65,'Positive Continuous Wheel Torque Limit'),(66,'Negative Continuous Wheel Torque Limit'))))
    for index, plot_id, xcol, ys in definitions:
        if index == 1 and not wheel or selected_ids is not None and plot_id not in selected_ids:
            continue
        chart = sheet._charts[index]; chart.scatterStyle = 'lineMarker'
        for (ycol, label), color in zip(ys, ('00806A','9E480E')):
            series = Series(Reference(sheet,min_col=ycol,min_row=5,max_row=end),
                            Reference(sheet,min_col=xcol,min_row=5,max_row=end),title=label)
            series.marker.symbol = 'none'; series.smooth = False
            series.graphicalProperties.line.solidFill = color
            series.graphicalProperties.line.width = 25400
            series.graphicalProperties.line.noFill = False
            chart.series.append(series)
    return {'sheet':'Torque Maps','table':f'BI4:BN{end}','wheel_overlays':wheel,
            'ratio':ratio,'efficiency':efficiency,'policy':POLICY}
