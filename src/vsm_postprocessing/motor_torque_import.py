"""Strict XLSX input adapter for the existing shared motor curve row format."""
from io import BytesIO
import math
from numbers import Real
import re

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter

from .continuous_torque import CONTINUOUS, ContinuousTorqueCurve
from .errors import ConfigurationError
from .motor_performance import SPEED, TORQUE


HEADERS = ('Motor RPM', 'Peak Torque', 'Cont Torque')
FIELDS = (SPEED, TORQUE, CONTINUOUS)
ALIASES = ({'motorrpm', 'motorspeed', 'rpm'},
           {'peaktorque', 'motorpeaktorque'},
           {'conttorque', 'continuoustorque', 'motorconttorque', 'motorcontinuoustorque'})


def _empty(value):
    return value is None or isinstance(value, str) and not value.strip()


def _compact(value):
    return re.sub(r'[^a-z0-9]', '', value.casefold())


def _columns(headers):
    columns = {}
    for column, value in enumerate(headers):
        if not isinstance(value, str):
            continue
        name, unit = value.strip(), None
        match = re.fullmatch(r'(.*?)\s*[\[(]([^\[\]()]+)[\])]\s*', name)
        if match is None:
            suffix = re.fullmatch(r'(.*?)\s+(rpm|n\s*[.·⋅* -]?\s*m)\s*', name, re.IGNORECASE)
            if suffix and any(_compact(suffix[1]) in aliases for aliases in ALIASES):
                match = suffix
        if match:
            name, unit = match.groups()
        normalized = _compact(name)
        for field, aliases, expected in zip(FIELDS, ALIASES, ('rpm', 'nm', 'nm')):
            if normalized not in aliases:
                continue
            valid_unit = unit is None or (unit.strip().casefold() == 'rpm' if expected == 'rpm'
                else re.fullmatch(r'n\s*[.·⋅* -]?\s*m', unit.strip(), re.IGNORECASE) is not None)
            if not valid_unit:
                raise ConfigurationError(f'Column {value!r}: expected {"rpm" if field == SPEED else "Nm"}; '
                                         'convert the workbook values and label the units before uploading.')
            if field in columns:
                raise ConfigurationError(f'Ambiguous columns for {field}; keep exactly one column for each required quantity.')
            columns[field] = column
    missing = [header for field, header in zip(FIELDS, HEADERS) if field not in columns]
    if missing:
        raise ConfigurationError('Missing required columns: ' + ', '.join(missing) +
                                 '. Put headers in row 1; use Motor RPM [rpm], Peak Torque [Nm], Cont Torque [Nm].')
    return columns


def _open_workbook(content):
    try:
        # Read formulas as formulas: neither missing caches nor stale caches may
        # silently become engineering inputs. No external links or Office calls.
        return load_workbook(BytesIO(content), read_only=True, data_only=False, keep_links=False)
    except Exception as exc:
        raise ConfigurationError('Cannot read motor workbook. Upload a valid, unencrypted .xlsx file '
                                 '(not CSV, .xls or a renamed file).') from exc


def _matching_sheets(book):
    names = []
    for sheet in book.worksheets:
        try:
            _columns(next(sheet.iter_rows(min_row=1, max_row=1, values_only=True), ()))
        except ConfigurationError:
            continue
        names.append(sheet.title)
    return tuple(names)


def motor_torque_worksheets(content):
    """Prefer sheets with all required headers; expose other sheets for diagnostics."""
    book = _open_workbook(content)
    try:
        names = _matching_sheets(book) or tuple(s.title for s in book.worksheets)
        if not names:
            raise ConfigurationError('Motor workbook has no worksheets. Add a sheet with the three required columns.')
        return names
    except ConfigurationError:
        raise
    except Exception as exc:
        raise ConfigurationError('Cannot read motor worksheets. Save a valid, unencrypted .xlsx workbook and retry.') from exc
    finally:
        book.close()


def parse_motor_torque_workbook(content, sheet_name=None):
    """Return paired, RPM-sorted rows accepted by both existing from_rows methods.

    Header-only templates, partial rows, nonfinite data and formulas are errors.
    Only completely empty rows are ignored; no engineering samples are dropped.
    """
    book = _open_workbook(content)
    try:
        if sheet_name is None:
            names = _matching_sheets(book)
            if len(names) > 1:
                raise ConfigurationError('Multiple motor torque worksheets found. Select the worksheet to import.')
            if not names:
                raise ConfigurationError('No motor torque worksheet found. Put Motor RPM, Peak Torque and Cont Torque '
                                         'in row 1 (units rpm and Nm); select a sheet to see its column errors.')
            sheet_name = names[0]
        if sheet_name not in book.sheetnames:
            raise ConfigurationError(f'Worksheet {sheet_name!r} does not exist. Select a worksheet from this workbook.')
        sheet = book[sheet_name]
        iterator = sheet.iter_rows(values_only=True)
        columns = _columns(next(iterator, ()))
        rows, speeds = [], set()
        for row_number, values in enumerate(iterator, 2):
            if all(_empty(value) for value in values):
                continue
            row = {}
            for field in FIELDS:
                column = columns[field]
                value = values[column] if column < len(values) else None
                location = f'{sheet_name}!{get_column_letter(column + 1)}{row_number} ({field})'
                if _empty(value):
                    raise ConfigurationError(f'{location}: missing value. Complete all three cells or remove the row.')
                if isinstance(value, str) and value.lstrip().startswith('='):
                    raise ConfigurationError(f'{location}: formulas are not accepted. Save numeric values in the motor map.')
                try:
                    if isinstance(value, bool) or not isinstance(value, (Real, str)):
                        raise ValueError
                    number = float(value)
                    if not math.isfinite(number):
                        raise ValueError
                except (ValueError, TypeError, OverflowError) as exc:
                    raise ConfigurationError(f'{location}: enter a finite numeric value in rpm or Nm, without unit text.') from exc
                if number < 0 or field == CONTINUOUS and number == 0:
                    rule = 'positive' if field == CONTINUOUS else 'nonnegative'
                    raise ConfigurationError(f'{location}: value must be {rule}. Torque ratings are magnitudes.')
                row[field] = number
            if row[SPEED] in speeds:
                raise ConfigurationError(f'{sheet_name}!row {row_number}: duplicate RPM {row[SPEED]:g}. '
                                         'Keep one complete torque pair per RPM.')
            speeds.add(row[SPEED])
            rows.append(row)
        if len(rows) < 2:
            raise ConfigurationError('Motor worksheet needs at least two complete RPM / peak / continuous torque rows.')
        rows.sort(key=lambda row: row[SPEED])
        # Reuse the engine's continuous curve constraints, including row capacity.
        ContinuousTorqueCurve.from_rows(rows)
        return rows
    except ConfigurationError:
        raise
    except Exception as exc:
        raise ConfigurationError('Cannot read motor worksheet data. Save a valid .xlsx workbook with numeric values and retry.') from exc
    finally:
        book.close()


def motor_torque_template():
    """Create a headers-only template; never supply invented motor ratings."""
    book = Workbook()
    try:
        sheet = book.active
        sheet.title = 'Motor Torque Map'
        sheet.append(HEADERS)
        for column, unit in enumerate(('rpm', 'Nm', 'Nm'), 1):
            sheet.cell(1, column).comment = Comment(f'Required numeric values in {unit}.', 'VSM')
            sheet.column_dimensions[get_column_letter(column)].width = 22
        sheet.freeze_panes = 'A2'
        output = BytesIO()
        book.save(output)
        return output.getvalue()
    finally:
        book.close()
