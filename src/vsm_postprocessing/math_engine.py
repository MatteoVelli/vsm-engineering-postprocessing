from __future__ import annotations

import ast
import csv
import difflib
import json
import math
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Mapping, Sequence

import numpy as np
import yaml

from .errors import ConfigurationError, MathChannelError
from .importer import ImportOptions, load_data_file
from .models import ChannelInfo, ImportedDataset

_IDENTIFIER = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")


@dataclass(frozen=True)
class MathOutputOptions:
    """Output controls for a math-channel calculation run."""

    data_filename: str = "dataset_with_math_channels.csv"
    include_units_row: bool = True
    float_precision: int = 12


@dataclass(frozen=True)
class MathComparisonDefinition:
    """Optional deterministic comparison against a channel already present in the source."""

    channel_id: str
    absolute_tolerance: float = 1e-9
    relative_tolerance: float = 1e-9
    required: bool = True


@dataclass(frozen=True)
class MathChannelDefinition:
    """One user-configured calculated channel."""

    channel_id: str
    display_name: str
    unit: str | None
    expression: str
    description: str | None = None
    comparison: MathComparisonDefinition | None = None


@dataclass(frozen=True)
class MathChannelsConfig:
    """Strict, versioned configuration for source export and math-channel definitions."""

    version: int
    include_time: bool
    export_source_channels: tuple[str, ...]
    constants: Mapping[str, float]
    math_channels: tuple[MathChannelDefinition, ...]
    output: MathOutputOptions


@dataclass(frozen=True)
class MathComparisonResult:
    math_channel_id: str
    reference_channel_id: str
    absolute_tolerance: float
    relative_tolerance: float
    required: bool
    passed: bool
    max_absolute_error: float
    rms_error: float
    mismatch_count: int

    def to_dict(self) -> dict[str, Any]:
        return {
            "math_channel_id": self.math_channel_id,
            "reference_channel_id": self.reference_channel_id,
            "absolute_tolerance": self.absolute_tolerance,
            "relative_tolerance": self.relative_tolerance,
            "required": self.required,
            "passed": self.passed,
            "max_absolute_error": self.max_absolute_error,
            "rms_error": self.rms_error,
            "mismatch_count": self.mismatch_count,
        }


@dataclass
class MathChannelsResult:
    """Calculated data, metadata and verification evidence."""

    dataset: ImportedDataset
    source_channels: list[ChannelInfo]
    calculated_channels: list[ChannelInfo]
    source_values: np.ndarray
    calculated_values: np.ndarray
    output_channels: list[ChannelInfo]
    output_values: np.ndarray
    config_path: Path
    config: MathChannelsConfig
    calculation_order: list[str]
    comparisons: list[MathComparisonResult]

    @property
    def sample_count(self) -> int:
        return int(self.output_values.shape[0])

    @property
    def source_channel_count(self) -> int:
        return len(self.source_channels)

    @property
    def math_channel_count(self) -> int:
        return len(self.calculated_channels)

    @property
    def output_channel_count(self) -> int:
        return len(self.output_channels)


@dataclass(frozen=True)
class _CompiledExpression:
    text: str
    tree: ast.Expression
    dependencies: tuple[str, ...]


@dataclass(frozen=True)
class _FunctionSpec:
    function: Callable[..., Any]
    min_args: int
    max_args: int


def _sample_energy_kwh(power_kw: np.ndarray, time_s: np.ndarray) -> np.ndarray:
    power = np.asarray(power_kw, dtype=np.float64)
    time = np.asarray(time_s, dtype=np.float64)
    if power.shape != time.shape:
        raise MathChannelError(
            f"sample_energy_kwh requires power and time with identical shapes; got {power.shape} and {time.shape}"
        )
    if time.ndim != 1:
        raise MathChannelError("sample_energy_kwh expects one-dimensional channels")
    if time.size < 2:
        raise MathChannelError("sample_energy_kwh requires at least two time samples")
    intervals = np.empty_like(time)
    intervals[0] = time[1] - time[0]
    intervals[1:] = np.diff(time)
    if not np.all(np.isfinite(intervals)) or np.any(intervals <= 0.0):
        raise MathChannelError("sample_energy_kwh requires a finite, strictly increasing time channel")
    return power * intervals / 3600.0


_ALLOWED_FUNCTIONS: dict[str, _FunctionSpec] = {
    "abs": _FunctionSpec(np.abs, 1, 1),
    "sqrt": _FunctionSpec(np.sqrt, 1, 1),
    "square": _FunctionSpec(np.square, 1, 1),
    "minimum": _FunctionSpec(np.minimum, 2, 2),
    "maximum": _FunctionSpec(np.maximum, 2, 2),
    "clip": _FunctionSpec(np.clip, 3, 3),
    "cumulative_sum": _FunctionSpec(np.cumsum, 1, 1),
    "sample_energy_kwh": _FunctionSpec(_sample_energy_kwh, 2, 2),
}

_ALLOWED_BINARY_OPERATORS: dict[type[ast.operator], Callable[[Any, Any], Any]] = {
    ast.Add: lambda left, right: left + right,
    ast.Sub: lambda left, right: left - right,
    ast.Mult: lambda left, right: left * right,
    ast.Div: lambda left, right: left / right,
    ast.Pow: lambda left, right: left**right,
}
_ALLOWED_UNARY_OPERATORS: dict[type[ast.unaryop], Callable[[Any], Any]] = {
    ast.UAdd: lambda value: +value,
    ast.USub: lambda value: -value,
}


def load_math_config(path: str | Path) -> MathChannelsConfig:
    """Load and strictly validate a versioned YAML math-channel file."""

    config_path = Path(path).expanduser().resolve()
    if not config_path.exists():
        raise ConfigurationError(f"Math-channel configuration does not exist: {config_path}")
    if not config_path.is_file():
        raise ConfigurationError(f"Math-channel configuration is not a file: {config_path}")

    try:
        raw = yaml.safe_load(config_path.read_text(encoding="utf-8"))
    except UnicodeDecodeError as exc:
        raise ConfigurationError(f"Math-channel configuration must be UTF-8: {config_path}") from exc
    except yaml.YAMLError as exc:
        raise ConfigurationError(f"Invalid YAML in math-channel configuration '{config_path}': {exc}") from exc

    if not isinstance(raw, dict):
        raise ConfigurationError("Math-channel configuration root must be a YAML mapping")
    _reject_unknown_keys(raw, {"version", "selection", "constants", "math_channels", "output"}, "root")

    version = raw.get("version")
    if version != 1:
        raise ConfigurationError("Math-channel configuration 'version' must be 1")

    selection = raw.get("selection", {})
    if not isinstance(selection, dict):
        raise ConfigurationError("selection must be a YAML mapping")
    _reject_unknown_keys(selection, {"include_time", "export_source_channels"}, "selection")
    include_time = selection.get("include_time", True)
    if not isinstance(include_time, bool):
        raise ConfigurationError("selection.include_time must be true or false")
    source_raw = selection.get("export_source_channels", [])
    if not isinstance(source_raw, list):
        raise ConfigurationError("selection.export_source_channels must be a YAML list")
    if any(not isinstance(value, str) or not value.strip() for value in source_raw):
        raise ConfigurationError(
            "Every selection.export_source_channels entry must be a non-empty channel_id string"
        )
    export_source_channels = tuple(value.strip() for value in source_raw)
    duplicates = _duplicates(export_source_channels)
    if duplicates:
        raise ConfigurationError(
            "selection.export_source_channels contains duplicate channel IDs: " + ", ".join(duplicates)
        )

    constants_raw = raw.get("constants", {})
    if constants_raw is None:
        constants_raw = {}
    if not isinstance(constants_raw, dict):
        raise ConfigurationError("constants must be a YAML mapping")
    constants: dict[str, float] = {}
    for name, value in constants_raw.items():
        if not isinstance(name, str) or not _IDENTIFIER.fullmatch(name):
            raise ConfigurationError(f"Invalid constant name '{name}'; use a valid Python-style identifier")
        if name in _ALLOWED_FUNCTIONS:
            raise ConfigurationError(f"Constant name '{name}' conflicts with a supported math function")
        if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(float(value)):
            raise ConfigurationError(f"Constant '{name}' must be a finite number")
        constants[name] = float(value)

    definitions_raw = raw.get("math_channels")
    if not isinstance(definitions_raw, list) or not definitions_raw:
        raise ConfigurationError("math_channels must be a non-empty YAML list")
    definitions = tuple(_load_math_definition(item, index) for index, item in enumerate(definitions_raw, start=1))
    definition_ids = [definition.channel_id for definition in definitions]
    duplicate_math_ids = _duplicates(definition_ids)
    if duplicate_math_ids:
        raise ConfigurationError("math_channels contains duplicate channel IDs: " + ", ".join(duplicate_math_ids))
    collisions = sorted(set(definition_ids) & set(constants))
    if collisions:
        raise ConfigurationError("Math channel IDs conflict with constants: " + ", ".join(collisions))

    output_raw = raw.get("output", {})
    if output_raw is None:
        output_raw = {}
    if not isinstance(output_raw, dict):
        raise ConfigurationError("output must be a YAML mapping")
    _reject_unknown_keys(output_raw, {"data_filename", "include_units_row", "float_precision"}, "output")

    data_filename = output_raw.get("data_filename", "dataset_with_math_channels.csv")
    if not isinstance(data_filename, str) or not data_filename.strip():
        raise ConfigurationError("output.data_filename must be a non-empty string")
    data_filename = data_filename.strip()
    filename_path = Path(data_filename)
    if filename_path.name != data_filename or filename_path.suffix.lower() != ".csv":
        raise ConfigurationError("output.data_filename must be a plain .csv filename without directories")

    include_units_row = output_raw.get("include_units_row", True)
    if not isinstance(include_units_row, bool):
        raise ConfigurationError("output.include_units_row must be true or false")
    float_precision = output_raw.get("float_precision", 12)
    if isinstance(float_precision, bool) or not isinstance(float_precision, int) or not 1 <= float_precision <= 17:
        raise ConfigurationError("output.float_precision must be an integer from 1 to 17")

    return MathChannelsConfig(
        version=version,
        include_time=include_time,
        export_source_channels=export_source_channels,
        constants=constants,
        math_channels=definitions,
        output=MathOutputOptions(
            data_filename=data_filename,
            include_units_row=include_units_row,
            float_precision=float_precision,
        ),
    )


def _load_math_definition(raw: object, index: int) -> MathChannelDefinition:
    if not isinstance(raw, dict):
        raise ConfigurationError(f"math_channels entry {index} must be a YAML mapping")
    _reject_unknown_keys(
        raw,
        {"channel_id", "display_name", "unit", "description", "expression", "compare_to"},
        f"math_channels[{index}]",
    )

    channel_id = raw.get("channel_id")
    if not isinstance(channel_id, str) or not _IDENTIFIER.fullmatch(channel_id):
        raise ConfigurationError(
            f"math_channels[{index}].channel_id must be a valid Python-style identifier"
        )
    if channel_id in _ALLOWED_FUNCTIONS:
        raise ConfigurationError(
            f"math_channels[{index}].channel_id '{channel_id}' conflicts with a supported math function"
        )

    display_name = raw.get("display_name")
    if not isinstance(display_name, str) or not display_name.strip():
        raise ConfigurationError(f"math_channels[{index}].display_name must be a non-empty string")

    unit = raw.get("unit")
    if unit is not None and (not isinstance(unit, str) or not unit.strip()):
        raise ConfigurationError(f"math_channels[{index}].unit must be null or a non-empty string")

    description = raw.get("description")
    if description is not None and (not isinstance(description, str) or not description.strip()):
        raise ConfigurationError(f"math_channels[{index}].description must be null or a non-empty string")

    expression = raw.get("expression")
    if not isinstance(expression, str) or not expression.strip():
        raise ConfigurationError(f"math_channels[{index}].expression must be a non-empty string")
    _compile_expression(expression.strip(), context=f"math channel '{channel_id}'")

    comparison_raw = raw.get("compare_to")
    comparison = None
    if comparison_raw is not None:
        if not isinstance(comparison_raw, dict):
            raise ConfigurationError(f"math_channels[{index}].compare_to must be a YAML mapping")
        _reject_unknown_keys(
            comparison_raw,
            {"channel_id", "absolute_tolerance", "relative_tolerance", "required"},
            f"math_channels[{index}].compare_to",
        )
        reference_id = comparison_raw.get("channel_id")
        if not isinstance(reference_id, str) or not reference_id.strip():
            raise ConfigurationError(
                f"math_channels[{index}].compare_to.channel_id must be a non-empty string"
            )
        absolute_tolerance = _finite_nonnegative_number(
            comparison_raw.get("absolute_tolerance", 1e-9),
            f"math_channels[{index}].compare_to.absolute_tolerance",
        )
        relative_tolerance = _finite_nonnegative_number(
            comparison_raw.get("relative_tolerance", 1e-9),
            f"math_channels[{index}].compare_to.relative_tolerance",
        )
        required = comparison_raw.get("required", True)
        if not isinstance(required, bool):
            raise ConfigurationError(f"math_channels[{index}].compare_to.required must be true or false")
        comparison = MathComparisonDefinition(
            channel_id=reference_id.strip(),
            absolute_tolerance=absolute_tolerance,
            relative_tolerance=relative_tolerance,
            required=required,
        )

    return MathChannelDefinition(
        channel_id=channel_id,
        display_name=display_name.strip(),
        unit=unit.strip() if isinstance(unit, str) else None,
        expression=expression.strip(),
        description=description.strip() if isinstance(description, str) else None,
        comparison=comparison,
    )


def calculate_math_channels(
    input_file: str | Path,
    config_file: str | Path,
    import_options: ImportOptions | None = None,
) -> MathChannelsResult:
    """Calculate configured math channels and validate optional source comparisons."""

    config_path = Path(config_file).expanduser().resolve()
    config = load_math_config(config_path)
    dataset = load_data_file(input_file, import_options)
    channels_by_id = {channel.channel_id: channel for channel in dataset.channels}
    source_values_by_id = {
        channel.channel_id: dataset.values[:, index]
        for index, channel in enumerate(dataset.channels)
    }

    math_ids = [definition.channel_id for definition in config.math_channels]
    source_collisions = sorted(set(math_ids) & set(channels_by_id))
    if source_collisions:
        raise MathChannelError(
            "Configured math channel IDs collide with source channel IDs: " + ", ".join(source_collisions)
        )
    constant_collisions = sorted(set(config.constants) & set(channels_by_id))
    if constant_collisions:
        raise MathChannelError(
            "Configured constant names collide with source channel IDs: " + ", ".join(constant_collisions)
        )

    compiled = {
        definition.channel_id: _compile_expression(
            definition.expression,
            context=f"math channel '{definition.channel_id}'",
        )
        for definition in config.math_channels
    }
    definitions_by_id = {definition.channel_id: definition for definition in config.math_channels}
    allowed_names = set(channels_by_id) | set(math_ids) | set(config.constants)
    missing_by_math: dict[str, list[str]] = {}
    for channel_id, expression in compiled.items():
        missing = sorted(set(expression.dependencies) - allowed_names)
        if missing:
            missing_by_math[channel_id] = missing
    if missing_by_math:
        raise MathChannelError(_format_missing_dependencies(missing_by_math, allowed_names))

    calculation_order = _topological_order(math_ids, compiled)
    context: dict[str, Any] = {**source_values_by_id, **config.constants}
    calculated_by_id: dict[str, np.ndarray] = {}

    for channel_id in calculation_order:
        expression = compiled[channel_id]
        with np.errstate(all="ignore"):
            raw_value = _evaluate_node(expression.tree.body, context)
        values = _normalise_result(raw_value, dataset.quality.sample_count, channel_id)
        non_finite = ~np.isfinite(values)
        if non_finite.any():
            first_index = int(np.flatnonzero(non_finite)[0])
            raise MathChannelError(
                f"Math channel '{channel_id}' produced {int(non_finite.sum())} non-finite values; "
                f"first failure at sample index {first_index}"
            )
        calculated_by_id[channel_id] = values
        context[channel_id] = values

    comparisons = _compare_calculated_channels(
        config.math_channels,
        calculated_by_id,
        source_values_by_id,
    )
    failed_required = [comparison for comparison in comparisons if comparison.required and not comparison.passed]
    if failed_required:
        details = "; ".join(
            f"{item.math_channel_id} vs {item.reference_channel_id}: "
            f"{item.mismatch_count} mismatches, max abs error {item.max_absolute_error:.12g}"
            for item in failed_required
        )
        raise MathChannelError("Required math-channel comparison failed: " + details)

    requested_source_ids = list(config.export_source_channels)
    if config.include_time:
        time_id = dataset.quality.time_channel_id
        if time_id is None:
            raise MathChannelError("The imported dataset does not define a time channel")
        requested_source_ids = [time_id, *[item for item in requested_source_ids if item != time_id]]
    missing_source = [channel_id for channel_id in requested_source_ids if channel_id not in channels_by_id]
    if missing_source:
        raise MathChannelError(_format_missing_source_channels(missing_source, dataset.channels))

    source_channels = [channels_by_id[channel_id] for channel_id in requested_source_ids]
    if requested_source_ids:
        source_indices = [dataset.channel_index(channel_id) for channel_id in requested_source_ids]
        source_values = dataset.values[:, source_indices].copy()
    else:
        source_values = np.empty((dataset.quality.sample_count, 0), dtype=np.float64)

    calculated_channels: list[ChannelInfo] = []
    for index, definition in enumerate(config.math_channels, start=1):
        expression = compiled[definition.channel_id]
        calculated_channels.append(
            ChannelInfo(
                channel_id=definition.channel_id,
                source_name=definition.display_name,
                display_name=definition.display_name,
                unit=definition.unit,
                source_column_index=dataset.quality.channel_count + index,
                source_column_label=f"MATH{index:03d}",
                kind="math",
                dtype="float64",
                provenance=f"{config_path.name}:math_channels[{index}]",
                dependencies=expression.dependencies,
                formula_example=definition.expression,
            )
        )

    calculated_values = np.column_stack(
        [calculated_by_id[definition.channel_id] for definition in config.math_channels]
    )
    output_channels = [*source_channels, *calculated_channels]
    output_values = np.column_stack([source_values, calculated_values])

    return MathChannelsResult(
        dataset=dataset,
        source_channels=source_channels,
        calculated_channels=calculated_channels,
        source_values=source_values,
        calculated_values=calculated_values,
        output_channels=output_channels,
        output_values=output_values,
        config_path=config_path,
        config=config,
        calculation_order=calculation_order,
        comparisons=comparisons,
    )


def export_math_channels(result: MathChannelsResult, output_dir: str | Path) -> dict[str, Path]:
    """Export calculated data, catalogue, validation evidence, manifest and summary."""

    output_path = Path(output_dir).expanduser().resolve()
    output_path.mkdir(parents=True, exist_ok=True)

    data_path = output_path / result.config.output.data_filename
    catalogue_path = output_path / "math_channel_catalogue.csv"
    validation_path = output_path / "math_validation_report.json"
    manifest_path = output_path / "math_manifest.json"
    summary_path = output_path / "math_summary.txt"

    _write_output_data(result, data_path)
    _write_math_catalogue(result, catalogue_path)
    validation_path.write_text(
        json.dumps(
            {
                "status": "PASS",
                "comparison_count": len(result.comparisons),
                "comparisons": [comparison.to_dict() for comparison in result.comparisons],
            },
            indent=2,
        ),
        encoding="utf-8",
    )
    manifest_path.write_text(json.dumps(_build_manifest(result, data_path), indent=2), encoding="utf-8")
    summary_path.write_text(_format_summary(result, data_path), encoding="utf-8")

    return {
        "output_data": data_path,
        "math_channel_catalogue": catalogue_path,
        "math_validation_report": validation_path,
        "math_manifest": manifest_path,
        "math_summary": summary_path,
    }


def _compile_expression(text: str, context: str) -> _CompiledExpression:
    try:
        parsed = ast.parse(text, mode="eval")
    except SyntaxError as exc:
        raise ConfigurationError(f"Invalid expression for {context}: {exc.msg}") from exc
    dependencies: list[str] = []

    def visit(node: ast.AST) -> None:
        if isinstance(node, ast.Expression):
            visit(node.body)
        elif isinstance(node, ast.BinOp):
            if type(node.op) not in _ALLOWED_BINARY_OPERATORS:
                raise ConfigurationError(f"Unsupported operator in {context}: {type(node.op).__name__}")
            visit(node.left)
            visit(node.right)
        elif isinstance(node, ast.UnaryOp):
            if type(node.op) not in _ALLOWED_UNARY_OPERATORS:
                raise ConfigurationError(f"Unsupported unary operator in {context}: {type(node.op).__name__}")
            visit(node.operand)
        elif isinstance(node, ast.Name):
            if node.id not in _ALLOWED_FUNCTIONS and node.id not in dependencies:
                dependencies.append(node.id)
        elif isinstance(node, ast.Constant):
            if isinstance(node.value, bool) or not isinstance(node.value, (int, float)):
                raise ConfigurationError(f"Only finite numeric constants are supported in {context}")
            if not math.isfinite(float(node.value)):
                raise ConfigurationError(f"Only finite numeric constants are supported in {context}")
        elif isinstance(node, ast.Call):
            if not isinstance(node.func, ast.Name) or node.func.id not in _ALLOWED_FUNCTIONS:
                raise ConfigurationError(f"Unsupported function call in {context}")
            if node.keywords:
                raise ConfigurationError(f"Keyword arguments are not supported in {context}")
            spec = _ALLOWED_FUNCTIONS[node.func.id]
            if not spec.min_args <= len(node.args) <= spec.max_args:
                raise ConfigurationError(
                    f"Function '{node.func.id}' in {context} expects {spec.min_args} argument(s)"
                )
            for argument in node.args:
                visit(argument)
        else:
            raise ConfigurationError(f"Unsupported expression element in {context}: {type(node).__name__}")

    visit(parsed)
    return _CompiledExpression(text=text, tree=parsed, dependencies=tuple(dependencies))


def _evaluate_node(node: ast.AST, context: Mapping[str, Any]) -> Any:
    if isinstance(node, ast.BinOp):
        operator = _ALLOWED_BINARY_OPERATORS[type(node.op)]
        return operator(_evaluate_node(node.left, context), _evaluate_node(node.right, context))
    if isinstance(node, ast.UnaryOp):
        operator = _ALLOWED_UNARY_OPERATORS[type(node.op)]
        return operator(_evaluate_node(node.operand, context))
    if isinstance(node, ast.Name):
        try:
            return context[node.id]
        except KeyError as exc:  # pre-validation should prevent this
            raise MathChannelError(f"Expression dependency '{node.id}' is unavailable") from exc
    if isinstance(node, ast.Constant):
        return float(node.value)
    if isinstance(node, ast.Call):
        spec = _ALLOWED_FUNCTIONS[node.func.id]  # type: ignore[union-attr]
        arguments = [_evaluate_node(argument, context) for argument in node.args]
        return spec.function(*arguments)
    raise MathChannelError(f"Internal error: unsupported AST node {type(node).__name__}")


def _normalise_result(value: Any, sample_count: int, channel_id: str) -> np.ndarray:
    array = np.asarray(value, dtype=np.float64)
    if array.ndim == 0:
        return np.full(sample_count, float(array), dtype=np.float64)
    if array.ndim != 1 or array.shape[0] != sample_count:
        raise MathChannelError(
            f"Math channel '{channel_id}' returned shape {array.shape}; expected ({sample_count},)"
        )
    return array.copy()


def _topological_order(math_ids: Sequence[str], compiled: Mapping[str, _CompiledExpression]) -> list[str]:
    math_set = set(math_ids)
    state: dict[str, int] = {channel_id: 0 for channel_id in math_ids}
    order: list[str] = []
    path: list[str] = []

    def visit(channel_id: str) -> None:
        if state[channel_id] == 2:
            return
        if state[channel_id] == 1:
            cycle_start = path.index(channel_id)
            cycle = [*path[cycle_start:], channel_id]
            raise MathChannelError("Circular math-channel dependency detected: " + " -> ".join(cycle))
        state[channel_id] = 1
        path.append(channel_id)
        for dependency in compiled[channel_id].dependencies:
            if dependency in math_set:
                visit(dependency)
        path.pop()
        state[channel_id] = 2
        order.append(channel_id)

    for channel_id in math_ids:
        visit(channel_id)
    return order


def _compare_calculated_channels(
    definitions: Sequence[MathChannelDefinition],
    calculated: Mapping[str, np.ndarray],
    source: Mapping[str, np.ndarray],
) -> list[MathComparisonResult]:
    results: list[MathComparisonResult] = []
    for definition in definitions:
        comparison = definition.comparison
        if comparison is None:
            continue
        if comparison.channel_id not in source:
            raise MathChannelError(
                f"Comparison reference '{comparison.channel_id}' for math channel "
                f"'{definition.channel_id}' is not present in the source dataset"
            )
        actual = calculated[definition.channel_id]
        reference = source[comparison.channel_id]
        difference = actual - reference
        close = np.isclose(
            actual,
            reference,
            atol=comparison.absolute_tolerance,
            rtol=comparison.relative_tolerance,
            equal_nan=False,
        )
        results.append(
            MathComparisonResult(
                math_channel_id=definition.channel_id,
                reference_channel_id=comparison.channel_id,
                absolute_tolerance=comparison.absolute_tolerance,
                relative_tolerance=comparison.relative_tolerance,
                required=comparison.required,
                passed=bool(close.all()),
                max_absolute_error=float(np.max(np.abs(difference))),
                rms_error=float(np.sqrt(np.mean(np.square(difference)))),
                mismatch_count=int((~close).sum()),
            )
        )
    return results


def _write_output_data(result: MathChannelsResult, path: Path) -> None:
    precision = result.config.output.float_precision
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow([channel.channel_id for channel in result.output_channels])
        if result.config.output.include_units_row:
            writer.writerow([channel.unit or "" for channel in result.output_channels])
        for row in result.output_values:
            writer.writerow([_format_float(float(value), precision) for value in row])


def _write_math_catalogue(result: MathChannelsResult, path: Path) -> None:
    definitions = {definition.channel_id: definition for definition in result.config.math_channels}
    comparisons = {comparison.math_channel_id: comparison for comparison in result.comparisons}
    fieldnames = [
        "math_order",
        "channel_id",
        "display_name",
        "unit",
        "description",
        "expression",
        "dependencies",
        "calculation_order",
        "provenance",
        "comparison_reference",
        "comparison_passed",
        "max_absolute_error",
        "rms_error",
        "mismatch_count",
    ]
    order_lookup = {channel_id: index for index, channel_id in enumerate(result.calculation_order, start=1)}
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        for index, channel in enumerate(result.calculated_channels, start=1):
            definition = definitions[channel.channel_id]
            comparison = comparisons.get(channel.channel_id)
            writer.writerow(
                {
                    "math_order": index,
                    "channel_id": channel.channel_id,
                    "display_name": channel.display_name,
                    "unit": channel.unit or "",
                    "description": definition.description or "",
                    "expression": definition.expression,
                    "dependencies": ";".join(channel.dependencies),
                    "calculation_order": order_lookup[channel.channel_id],
                    "provenance": channel.provenance,
                    "comparison_reference": comparison.reference_channel_id if comparison else "",
                    "comparison_passed": comparison.passed if comparison else "",
                    "max_absolute_error": comparison.max_absolute_error if comparison else "",
                    "rms_error": comparison.rms_error if comparison else "",
                    "mismatch_count": comparison.mismatch_count if comparison else "",
                }
            )


def _build_manifest(result: MathChannelsResult, data_path: Path) -> dict[str, Any]:
    quality = result.dataset.quality
    return {
        "configuration_version": result.config.version,
        "source_file": str(result.dataset.source_path),
        "source_sha256": quality.source_sha256,
        "configuration_file": str(result.config_path),
        "sample_count": result.sample_count,
        "source_channel_count": result.source_channel_count,
        "math_channel_count": result.math_channel_count,
        "output_channel_count": result.output_channel_count,
        "time_channel_id": quality.time_channel_id,
        "time_start": quality.time_start,
        "time_end": quality.time_end,
        "nominal_time_step": quality.nominal_time_step,
        "constants": dict(result.config.constants),
        "calculation_order": result.calculation_order,
        "output_channel_ids": [channel.channel_id for channel in result.output_channels],
        "math_channels": [channel.to_dict() for channel in result.calculated_channels],
        "comparisons": [comparison.to_dict() for comparison in result.comparisons],
        "output_data_file": str(data_path),
        "include_units_row": result.config.output.include_units_row,
        "float_precision": result.config.output.float_precision,
    }


def _format_summary(result: MathChannelsResult, data_path: Path) -> str:
    lines = [
        "VSM MATH CHANNELS",
        "=================",
        "Status: PASS",
        f"Source: {result.dataset.source_path}",
        f"Source SHA-256: {result.dataset.quality.source_sha256}",
        f"Configuration: {result.config_path}",
        f"Samples: {result.sample_count}",
        f"Exported source channels: {result.source_channel_count}",
        f"Calculated math channels: {result.math_channel_count}",
        f"Output channels: {result.output_channel_count}",
        f"Output data: {data_path}",
        "",
        "Calculation order:",
    ]
    lines.extend(f"{index:02d}. {channel_id}" for index, channel_id in enumerate(result.calculation_order, start=1))
    lines.extend(["", "Math-channel definitions:"])
    definitions = {definition.channel_id: definition for definition in result.config.math_channels}
    for index, channel in enumerate(result.calculated_channels, start=1):
        definition = definitions[channel.channel_id]
        lines.append(
            f"{index:02d}. {channel.channel_id} | {channel.display_name} [{channel.unit or '-'}] "
            f"= {definition.expression}"
        )
    lines.extend(["", "Source comparisons:"])
    if result.comparisons:
        lines.extend(
            f"- {item.math_channel_id} vs {item.reference_channel_id}: "
            f"{'PASS' if item.passed else 'FAIL'}, max abs error={item.max_absolute_error:.12g}, "
            f"RMS error={item.rms_error:.12g}, mismatches={item.mismatch_count}"
            for item in result.comparisons
        )
    else:
        lines.append("- None configured")
    return "\n".join(lines) + "\n"


def _format_missing_dependencies(missing_by_math: Mapping[str, Sequence[str]], allowed_names: set[str]) -> str:
    lines = ["Math-channel expressions contain unknown dependencies:"]
    for channel_id, missing_names in missing_by_math.items():
        for missing in missing_names:
            suggestions = difflib.get_close_matches(missing, sorted(allowed_names), n=3, cutoff=0.45)
            suffix = f" (possible identifier: {', '.join(suggestions)})" if suggestions else ""
            lines.append(f"- {channel_id}: {missing}{suffix}")
    lines.append("Use stable source channel_id values, configured math channel IDs or declared constants.")
    return "\n".join(lines)


def _format_missing_source_channels(missing: Sequence[str], channels: Sequence[ChannelInfo]) -> str:
    available_ids = [channel.channel_id for channel in channels]
    lines = ["Requested source channel IDs were not found:"]
    for missing_id in missing:
        suggestions = difflib.get_close_matches(missing_id, available_ids, n=3, cutoff=0.45)
        suffix = f" (possible channel_id: {', '.join(suggestions)})" if suggestions else ""
        lines.append(f"- {missing_id}{suffix}")
    lines.append("Use channel_id values from channel_catalogue.csv.")
    return "\n".join(lines)


def _format_float(value: float, precision: int) -> str:
    if not np.isfinite(value):
        raise MathChannelError("Output data contains a non-finite value after successful calculation")
    return format(value, f".{precision}g")


def _finite_nonnegative_number(value: object, field_name: str) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ConfigurationError(f"{field_name} must be a finite non-negative number")
    number = float(value)
    if not math.isfinite(number) or number < 0.0:
        raise ConfigurationError(f"{field_name} must be a finite non-negative number")
    return number


def _duplicates(values: Sequence[str]) -> list[str]:
    seen: set[str] = set()
    duplicates: list[str] = []
    for value in values:
        if value in seen and value not in duplicates:
            duplicates.append(value)
        seen.add(value)
    return duplicates


def _reject_unknown_keys(mapping: Mapping[str, Any], allowed: set[str], context: str) -> None:
    unknown = sorted(set(mapping) - allowed)
    if unknown:
        raise ConfigurationError(f"Unknown key(s) in {context}: {', '.join(unknown)}")
