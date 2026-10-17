"""Compile configured KPI dependency chains into auditable workbook formulas.

Python remains the numerical reference. Helpers carry formulas, never copies of
final metrics; raw inputs omitted from the visible selection retain provenance.
"""
import ast
import math

import numpy as np
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter, quote_sheetname

from .battery import SOC_FRACTION_EPSILON
from .excel_formulas import statistic_formula, write_formula
from .math_engine import _evaluate_node
from .profile_math import DEFAULT_PROFILE_MATH_CONSTANTS


def write_profile_kpi_formulas(workbook, report, result, report_channels):
    """Write live statistics and every available configured KPI, preserving styles."""
    compiler = _Compiler(workbook, report, result, report_channels)
    compiler.write()


class _Compiler:
    def __init__(self, workbook, report, result, report_channels):
        self.report, self.result = report, result
        self.sheet = workbook.create_sheet("Formula Support")
        self.sheet.sheet_state = "hidden"
        self.sheet["A1"] = "Formula dependency helpers; unhide to audit raw inputs and engineering parameters."
        self.sheet["A2"] = "Rows 5 onward are aligned with the main report data; parameters follow the data."
        self.end = result.dataset.quality.sample_count + 4
        self.parameter_row = self.end + 4
        self.column = 1
        self.parameters = {}
        self.vector_expressions = {}
        self.vectors = {}
        self.context = {}
        self.scalars = {}
        self.live_main_channels = {}
        self.columns = {c.channel_id: i for i, c in enumerate(report_channels, 1)}
        self.math = result.profile.math_by_semantic_name()
        self.stats = {s.definition.statistic_id: (row, s) for row, s in enumerate(result.statistics, 2)}
        start = len(result.statistics) + 5
        self.kpis = {k.definition.kpi_id: (row, k) for row, k in enumerate(result.kpis, start)}
        for name, resolved in result.resolution.resolved.items():
            self.context[name] = result.dataset.values[:, result.dataset.channel_index(resolved.channel.channel_id)]
        self.context.update(result.math_result.values_by_semantic_name)
        self.context.update(DEFAULT_PROFILE_MATH_CONSTANTS)
        self.context.update({name: item.value for name, (_, item) in self.stats.items()})
        self.context.update({name: item.value for name, (_, item) in self.kpis.items()})

    def ref(self, sheet, coordinate):
        return f"{quote_sheetname(sheet.title)}!{coordinate}"

    def parameter(self, value, description):
        key = (float(value), description)
        if key not in self.parameters:
            row = self.parameter_row
            self.parameter_row += 1
            self.sheet.cell(row, 1, description)
            cell = self.sheet.cell(row, 2, float(value))
            cell.comment = Comment(description, "Astauto")
            self.parameters[key] = self.ref(self.sheet, f"$B${row}")
        return self.parameters[key]

    def helper(self, label, values, formulas=None):
        col = self.column
        self.column += 1
        letter = get_column_letter(col)
        self.sheet.cell(3, col, label)
        for row, value in enumerate(values, 5):
            cell = self.sheet.cell(row, col)
            if formulas is None:
                if math.isfinite(float(value)):
                    cell.value = float(value)
            else:
                write_formula(cell, "=" + formulas(row), float(value))
        return self.ref(self.sheet, f"${letter}$5:${letter}${self.end}")

    def at(self, extent, row):
        return f"INDEX({extent},{row-4})"

    def live_main(self, name):
        """A formula over static MATH copies is not a live dependency chain."""
        if name in self.live_main_channels:
            return self.live_main_channels[name]
        resolved = self.result.resolution.resolved.get(name)
        column = self.columns.get(resolved.channel.channel_id, self.columns.get(name)) if resolved else self.columns.get(name)
        if resolved:
            live = bool(column)
        else:
            definition = self.math.get(name)
            live = bool(column and definition and self.report.cell(5, column).data_type == "f"
                        and all(dep in DEFAULT_PROFILE_MATH_CONSTANTS or self.live_main(dep)
                                for dep in definition.dependencies))
        self.live_main_channels[name] = live
        return live

    def vector(self, name):
        if name in self.vectors:
            return self.vectors[name]
        resolved = self.result.resolution.resolved.get(name)
        column = self.columns.get(name)
        if resolved:
            column = self.columns.get(resolved.channel.channel_id, column)
        if column and self.live_main(name):
            letter = get_column_letter(column)
            extent = self.ref(self.report, f"${letter}$5:${letter}${self.end}")
        elif resolved:
            extent = self.helper(f"{name}: raw {resolved.channel.source_name} [{resolved.effective_unit}]",
                                 self.context[name])
        elif name in self.math and self.math[name].expression:
            node = ast.parse(self.math[name].expression, mode="eval").body
            extent = self.expression_vector(node, self.math[name].expression)
        else:
            raise ValueError(f"No live Excel dependency for {name}")
        self.vectors[name] = extent
        return extent

    def expression_vector(self, node, label):
        key = ast.dump(node)
        if key in self.vector_expressions:
            return self.vector_expressions[key]
        values = np.asarray(_evaluate_node(node, self.context), dtype=float)
        if isinstance(node, ast.Call) and node.func.id in {
            "cumulative_sum", "cumulative_trapezoid", "interval_energy_kwh", "sample_energy_kwh"
        }:
            name = node.func.id
            inputs = [self.vector(arg.id) for arg in node.args]
            own_column = get_column_letter(self.column)

            def formula(row):
                current = self.at(inputs[0], row)
                previous = self.at(inputs[0], row-1)
                own_previous = f"${own_column}{row-1}"
                if name == "cumulative_sum":
                    return current if row == 5 else f"{own_previous}+{current}"
                time = self.at(inputs[1], row)
                prior_time = self.at(inputs[1], row-1)
                if name == "sample_energy_kwh":
                    duration = f"({self.at(inputs[1], 6)}-{time})" if row == 5 else f"({time}-{prior_time})"
                    return f"{current}*{duration}/3600"
                if row == 5:
                    return "0*" + current
                increment = f"({previous}/2+{current}/2)*({time}-{prior_time})"
                return f"{increment}/3600" if name == "interval_energy_kwh" else f"{own_previous}+{increment}"

            extent = self.helper(label, values, formula)
            self.vector_expressions[key] = extent
            return extent
        # Arithmetic can wrap a temporal expression (e.g. integral / 3600).
        extent = self.helper(label, values, lambda row: self.translate(node, label, row))
        self.vector_expressions[key] = extent
        return extent

    def translate(self, node, label, row=None):
        if isinstance(node, ast.Constant):
            return self.parameter(node.value, f"Profile expression {label}: literal {node.value}")
        if isinstance(node, ast.Name):
            name = node.id
            if name in self.stats or name in self.kpis:
                return self.scalar(name)
            if name in DEFAULT_PROFILE_MATH_CONSTANTS:
                return self.parameter(self.context[name], f"Profile math constant: {name}")
            extent = self.vector(name)
            if row is None:
                raise ValueError(f"Vector {name} needs a scalar reduction")
            return self.at(extent, row)
        if isinstance(node, ast.BinOp):
            op = {ast.Add: "+", ast.Sub: "-", ast.Mult: "*", ast.Div: "/", ast.Pow: "^"}[type(node.op)]
            return f"({self.translate(node.left,label,row)}{op}{self.translate(node.right,label,row)})"
        if isinstance(node, ast.UnaryOp):
            return f"({'-' if isinstance(node.op, ast.USub) else '+'}{self.translate(node.operand,label,row)})"
        if isinstance(node, ast.Call):
            name = node.func.id
            if name in {"maximum", "minimum", "abs", "sqrt", "square"}:
                args = [self.translate(a, label, row) for a in node.args]
                if name == "square":
                    return f"({args[0]}^2)"
                return f"{dict(maximum='MAX',minimum='MIN',abs='ABS',sqrt='SQRT')[name]}({','.join(args)})"
            if name == "nominal_battery_capacity_kwh":
                energy, soc = [self.vector(a.id) for a in node.args]
                epsilon = self.parameter(SOC_FRACTION_EPSILON, "Battery inference: minimum valid SOC fraction (battery.py)")
                percent = self.parameter(100, "Battery SOC conversion: percent to fraction")
                e, s = [self.context[a.id] for a in node.args]
                with np.errstate(all="ignore"):
                    estimates = np.where(np.isfinite(e) & np.isfinite(s) & (s/100 > SOC_FRACTION_EPSILON), e/(s/100), np.nan)
                extent = self.helper(label + ": valid energy / SOC fraction", estimates,
                    lambda r: f'IFERROR(IF(AND(ISNUMBER({self.at(energy,r)}),ISNUMBER({self.at(soc,r)}),{self.at(soc,r)}/{percent}>{epsilon}),{self.at(energy,r)}/({self.at(soc,r)}/{percent}),""),"")')
                return f"MEDIAN({extent})"
            if name == "time_average":
                integral_node = ast.Call(func=ast.Name(id="cumulative_trapezoid"), args=node.args, keywords=[])
                integral = self.expression_vector(integral_node, label + ": signed trapezoidal integral")
                time = self.vector(node.args[1].id)
                return f"INDEX({integral},ROWS({integral}))/(INDEX({time},ROWS({time}))-INDEX({time},1))"
            if row is not None:
                extent = self.expression_vector(node, ast.unparse(node))
                return self.at(extent, row)
        raise ValueError(f"Unsupported Excel KPI expression: {ast.dump(node)}")

    def scalar(self, name):
        if name in self.scalars:
            return self.scalars[name]
        sheet = self.report.parent["Statistics"]
        if name in self.stats:
            row, item = self.stats[name]
            formula = statistic_formula(item.definition.operation, self.vector(item.target_channel), item.definition.nan_policy)
            if formula is None:
                raise ValueError(f"Unsupported live KPI statistic: {item.definition.operation}")
            cell = sheet.cell(row, 5)
        else:
            row, item = self.kpis[name]
            expression = item.definition.expression
            formula = "=" + self.translate(ast.parse(expression, mode="eval").body, expression)
            cell = sheet.cell(row, 3)
            cell.comment = Comment("Live profile KPI: " + expression + "\nParameters and helper calculations: Formula Support (hidden sheet).", "Astauto")
        write_formula(cell, formula, item.value)
        reference = self.ref(sheet, cell.coordinate)
        self.scalars[name] = reference
        return reference

    def write(self):
        for name in self.kpis:
            self.scalar(name)
