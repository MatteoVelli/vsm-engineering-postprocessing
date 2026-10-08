"""Live report formulas with deterministic initial values for non-calculating readers."""

import ast
from io import BytesIO
import math
import re
from zipfile import ZipFile

import numpy as np
from openpyxl.utils import get_column_letter


FORMULA_OPERATIONS = {"max", "min", "last", "first"}


def profile_arithmetic_formula(expression, channel_columns, constants):
    """Translate only pointwise arithmetic, preserving the engine's AST grouping.

    Return a row template; functions/integration or unexported dependencies retain
    Python values. COUNT guards prevent Excel treating missing inputs as zero.
    """
    references = []
    operators = {ast.Add: "+", ast.Sub: "-", ast.Mult: "*", ast.Div: "/", ast.Pow: "^"}

    def translate(node):
        if isinstance(node, ast.Constant) and type(node.value) in {int, float}:
            return repr(float(node.value))
        if isinstance(node, ast.Name):
            if node.id in constants:
                return repr(float(constants[node.id]))
            if node.id not in channel_columns:
                raise ValueError("dependency not exported")
            reference = "$" + get_column_letter(channel_columns[node.id]) + "{row}"
            if reference not in references:
                references.append(reference)
            return reference
        if isinstance(node, ast.BinOp) and type(node.op) in operators:
            return f"({translate(node.left)}{operators[type(node.op)]}{translate(node.right)})"
        if isinstance(node, ast.UnaryOp) and isinstance(node.op, (ast.UAdd, ast.USub)):
            return f"({'+' if isinstance(node.op, ast.UAdd) else '-'}{translate(node.operand)})"
        raise ValueError("not pointwise arithmetic")

    try:
        body = translate(ast.parse(expression, mode="eval").body)
    except (ValueError, SyntaxError):
        return None
    if references:
        return f"=IF(COUNT({','.join(references)})={len(references)},{body},NA())"
    return "=" + body


def write_statistic(cell, operation, start_row, end_row, values, value, nan_policy="error"):
    """Keep unsupported engineering statistics numeric and expose simple statistics."""
    cell.value = value
    if operation not in FORMULA_OPERATIONS:
        return
    col = get_column_letter(cell.column)
    extent = f"{col}{start_row}:{col}{end_row}"
    if operation in {"max", "min"}:
        formula = f"{operation.upper()}({extent})"
    else:
        formula = f"{col}{start_row if operation == 'first' else end_row}"
    if not np.isfinite(values).all():
        if nan_policy == "propagate":
            formula = f"IF(COUNT({extent})=ROWS({extent}),{formula},NA())"
        elif nan_policy == "omit" and operation == "first":
            formula = f"INDEX({extent},MATCH(TRUE,INDEX(ISNUMBER({extent}),0),0))"
        elif nan_policy == "omit" and operation == "last":
            formula = f"LOOKUP(2,1/ISNUMBER({extent}),{extent})"
    write_formula(cell, "=" + formula, value)


def write_formula(cell, formula, value):
    cell.value = formula
    workbook = cell.parent.parent
    if not hasattr(workbook, "_report_formula_values"):
        workbook._report_formula_values = {}
    workbook._report_formula_values.setdefault(cell.parent.title, {})[cell.coordinate] = value


def save_report_workbook(workbook, output_path):
    """openpyxl writes formulas but cannot write their cached numeric values.

    Fill only the registered formula caches in the serialized worksheet XML;
    preserve all other package parts. Excel recalculates the live dependencies
    on opening and after edits, rather than treating caches as authoritative.
    """
    workbook.calculation.calcMode = "auto"
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    buffer = BytesIO()
    workbook.save(buffer)
    caches = getattr(workbook, "_report_formula_values", {})
    by_part = {f"xl/worksheets/sheet{i}.xml": caches[sheet.title]
               for i, sheet in enumerate(workbook.worksheets, 1) if sheet.title in caches}
    with ZipFile(buffer) as source, ZipFile(output_path, "w") as target:
        for entry in source.infolist():
            data = source.read(entry.filename)
            if entry.filename in by_part:
                values = by_part[entry.filename]

                def fill_cache(match):
                    coordinate = match[1].decode("ascii")
                    if coordinate not in values or not math.isfinite(values[coordinate]):
                        return match[0]
                    cached = repr(float(values[coordinate])).encode("ascii")
                    return match[0].replace(b"<v></v>", b"<v>" + cached + b"</v>")

                data = re.sub(rb'<c\b[^>]*\br="([A-Z]+[0-9]+)"[^>]*>.*?</c>', fill_cache, data)
            target.writestr(entry, data)
