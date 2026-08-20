from __future__ import annotations

import csv
import difflib
import json
import math
import re
import zipfile
from dataclasses import dataclass
from pathlib import Path, PurePosixPath
from typing import Any, Mapping, Sequence
from xml.etree import ElementTree

import numpy as np
import yaml

from .errors import ConfigurationError, StatisticsError
from .battery import max_charging_power_kw
from .importer import ImportOptions, load_data_file
from .math_engine import MathChannelsResult, calculate_math_channels
from .models import ChannelInfo, ImportedDataset

_IDENTIFIER = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")
_CELL_REFERENCE = re.compile(r"^[A-Za-z]{1,3}[1-9][0-9]*$")
_ALLOWED_OPERATIONS = {"rms", "time_weighted_rms", "max", "min", "first", "last", "sum", "positive_max"}
_ALLOWED_NAN_POLICIES = {"error", "omit", "propagate"}
_ALLOWED_PLACEMENT_GROUPS = {"top_rms", "bottom_channel", "kpi_block"}


@dataclass(frozen=True)
class StatisticsOutputOptions:
    results_filename: str = "statistics_results.csv"
    wide_filename: str = "statistics_by_channel.csv"
    float_precision: int = 12


@dataclass(frozen=True)
class StatisticComparisonDefinition:
    cell: str
    workbook: str | None = None
    sheet: str | None = None
    absolute_tolerance: float = 1e-9
    relative_tolerance: float = 1e-9
    required: bool = True


@dataclass(frozen=True)
class StatisticDefinition:
    statistic_id: str
    channel_id: str
    operation: str
    placement_group: str
    nan_policy: str = "error"
    display_name: str | None = None
    description: str | None = None
    comparison: StatisticComparisonDefinition | None = None


@dataclass(frozen=True)
class StatisticsConfig:
    version: int
    statistics: tuple[StatisticDefinition, ...]
    output: StatisticsOutputOptions


@dataclass(frozen=True)
class StatisticComparisonResult:
    statistic_id: str
    workbook: str
    sheet: str
    cell: str
    reference_value: float
    calculated_value: float
    absolute_tolerance: float
    relative_tolerance: float
    required: bool
    passed: bool
    absolute_error: float
    relative_error: float

    def to_dict(self) -> dict[str, Any]:
        return {
            "statistic_id": self.statistic_id,
            "workbook": self.workbook,
            "sheet": self.sheet,
            "cell": self.cell,
            "reference_value": self.reference_value,
            "calculated_value": self.calculated_value,
            "absolute_tolerance": self.absolute_tolerance,
            "relative_tolerance": self.relative_tolerance,
            "required": self.required,
            "passed": self.passed,
            "absolute_error": self.absolute_error,
            "relative_error": self.relative_error,
        }


@dataclass(frozen=True)
class StatisticResult:
    statistic_id: str
    channel_id: str
    channel_display_name: str
    channel_unit: str | None
    channel_kind: str
    operation: str
    placement_group: str
    nan_policy: str
    value: float
    sample_count: int
    used_sample_count: int
    omitted_sample_count: int
    display_name: str
    description: str | None
    comparison: StatisticComparisonResult | None = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "statistic_id": self.statistic_id,
            "channel_id": self.channel_id,
            "channel_display_name": self.channel_display_name,
            "channel_unit": self.channel_unit,
            "channel_kind": self.channel_kind,
            "operation": self.operation,
            "placement_group": self.placement_group,
            "nan_policy": self.nan_policy,
            "value": self.value if math.isfinite(self.value) else None,
            "value_is_finite": math.isfinite(self.value),
            "sample_count": self.sample_count,
            "used_sample_count": self.used_sample_count,
            "omitted_sample_count": self.omitted_sample_count,
            "display_name": self.display_name,
            "description": self.description,
            "comparison": self.comparison.to_dict() if self.comparison else None,
        }


@dataclass
class StatisticsResult:
    dataset: ImportedDataset
    config_path: Path
    config: StatisticsConfig
    channels_by_id: Mapping[str, ChannelInfo]
    values_by_id: Mapping[str, np.ndarray]
    statistics: list[StatisticResult]
    math_result: MathChannelsResult | None

    @property
    def sample_count(self) -> int:
        return self.dataset.quality.sample_count

    @property
    def statistic_count(self) -> int:
        return len(self.statistics)

    @property
    def comparison_count(self) -> int:
        return sum(item.comparison is not None for item in self.statistics)

    @property
    def required_comparisons_passed(self) -> bool:
        return all(
            item.comparison is None
            or not item.comparison.required
            or item.comparison.passed
            for item in self.statistics
        )


def load_statistics_config(path: str | Path) -> StatisticsConfig:
    config_path = Path(path).expanduser().resolve()
    if not config_path.exists():
        raise ConfigurationError(f"Statistics configuration does not exist: {config_path}")
    if not config_path.is_file():
        raise ConfigurationError(f"Statistics configuration is not a file: {config_path}")

    try:
        raw = yaml.safe_load(config_path.read_text(encoding="utf-8"))
    except UnicodeDecodeError as exc:
        raise ConfigurationError(f"Statistics configuration must be UTF-8: {config_path}") from exc
    except yaml.YAMLError as exc:
        raise ConfigurationError(f"Invalid YAML in statistics configuration '{config_path}': {exc}") from exc

    if not isinstance(raw, dict):
        raise ConfigurationError("Statistics configuration root must be a YAML mapping")
    _reject_unknown_keys(raw, {"version", "statistics", "output"}, "root")

    version = raw.get("version")
    if version != 1:
        raise ConfigurationError("Statistics configuration 'version' must be 1")

    definitions_raw = raw.get("statistics")
    if not isinstance(definitions_raw, list) or not definitions_raw:
        raise ConfigurationError("statistics must be a non-empty YAML list")
    definitions = tuple(
        _load_statistic_definition(item, index)
        for index, item in enumerate(definitions_raw, start=1)
    )
    duplicate_ids = _duplicates([item.statistic_id for item in definitions])
    if duplicate_ids:
        raise ConfigurationError("statistics contains duplicate statistic IDs: " + ", ".join(duplicate_ids))

    output_raw = raw.get("output", {})
    if output_raw is None:
        output_raw = {}
    if not isinstance(output_raw, dict):
        raise ConfigurationError("output must be a YAML mapping")
    _reject_unknown_keys(
        output_raw,
        {"results_filename", "wide_filename", "float_precision"},
        "output",
    )
    results_filename = _plain_csv_filename(
        output_raw.get("results_filename", "statistics_results.csv"),
        "output.results_filename",
    )
    wide_filename = _plain_csv_filename(
        output_raw.get("wide_filename", "statistics_by_channel.csv"),
        "output.wide_filename",
    )
    if results_filename == wide_filename:
        raise ConfigurationError("output.results_filename and output.wide_filename must be different")
    float_precision = output_raw.get("float_precision", 12)
    if isinstance(float_precision, bool) or not isinstance(float_precision, int) or not 1 <= float_precision <= 17:
        raise ConfigurationError("output.float_precision must be an integer from 1 to 17")

    return StatisticsConfig(
        version=version,
        statistics=definitions,
        output=StatisticsOutputOptions(
            results_filename=results_filename,
            wide_filename=wide_filename,
            float_precision=float_precision,
        ),
    )


def _load_statistic_definition(raw: object, index: int) -> StatisticDefinition:
    context = f"statistics[{index}]"
    if not isinstance(raw, dict):
        raise ConfigurationError(f"{context} must be a YAML mapping")
    _reject_unknown_keys(
        raw,
        {
            "statistic_id",
            "channel_id",
            "operation",
            "placement_group",
            "nan_policy",
            "display_name",
            "description",
            "compare_to",
        },
        context,
    )

    statistic_id = raw.get("statistic_id")
    if not isinstance(statistic_id, str) or not _IDENTIFIER.fullmatch(statistic_id):
        raise ConfigurationError(f"{context}.statistic_id must be a valid Python-style identifier")

    channel_id = raw.get("channel_id")
    if not isinstance(channel_id, str) or not channel_id.strip():
        raise ConfigurationError(f"{context}.channel_id must be a non-empty string")

    operation = raw.get("operation")
    if operation not in _ALLOWED_OPERATIONS:
        raise ConfigurationError(
            f"{context}.operation must be one of: {', '.join(sorted(_ALLOWED_OPERATIONS))}"
        )

    placement_group = raw.get("placement_group")
    if placement_group not in _ALLOWED_PLACEMENT_GROUPS:
        raise ConfigurationError(
            f"{context}.placement_group must be one of: {', '.join(sorted(_ALLOWED_PLACEMENT_GROUPS))}"
        )

    nan_policy = raw.get("nan_policy", "error")
    if nan_policy not in _ALLOWED_NAN_POLICIES:
        raise ConfigurationError(
            f"{context}.nan_policy must be one of: {', '.join(sorted(_ALLOWED_NAN_POLICIES))}"
        )

    display_name = raw.get("display_name")
    if display_name is not None and (not isinstance(display_name, str) or not display_name.strip()):
        raise ConfigurationError(f"{context}.display_name must be null or a non-empty string")

    description = raw.get("description")
    if description is not None and (not isinstance(description, str) or not description.strip()):
        raise ConfigurationError(f"{context}.description must be null or a non-empty string")

    comparison_raw = raw.get("compare_to")
    comparison = None
    if comparison_raw is not None:
        comparison = _load_comparison(comparison_raw, context)

    return StatisticDefinition(
        statistic_id=statistic_id,
        channel_id=channel_id.strip(),
        operation=operation,
        placement_group=placement_group,
        nan_policy=nan_policy,
        display_name=display_name.strip() if isinstance(display_name, str) else None,
        description=description.strip() if isinstance(description, str) else None,
        comparison=comparison,
    )


def _load_comparison(raw: object, parent_context: str) -> StatisticComparisonDefinition:
    context = f"{parent_context}.compare_to"
    if not isinstance(raw, dict):
        raise ConfigurationError(f"{context} must be a YAML mapping")
    _reject_unknown_keys(
        raw,
        {"cell", "workbook", "sheet", "absolute_tolerance", "relative_tolerance", "required"},
        context,
    )

    cell = raw.get("cell")
    if not isinstance(cell, str) or not _CELL_REFERENCE.fullmatch(cell.strip()):
        raise ConfigurationError(f"{context}.cell must be an A1-style cell reference such as N2")

    workbook = raw.get("workbook")
    if workbook is not None and (not isinstance(workbook, str) or not workbook.strip()):
        raise ConfigurationError(f"{context}.workbook must be null or a non-empty path")

    sheet = raw.get("sheet")
    if sheet is not None and (not isinstance(sheet, str) or not sheet.strip()):
        raise ConfigurationError(f"{context}.sheet must be null or a non-empty string")

    absolute_tolerance = _finite_nonnegative_number(
        raw.get("absolute_tolerance", 1e-9),
        f"{context}.absolute_tolerance",
    )
    relative_tolerance = _finite_nonnegative_number(
        raw.get("relative_tolerance", 1e-9),
        f"{context}.relative_tolerance",
    )
    required = raw.get("required", True)
    if not isinstance(required, bool):
        raise ConfigurationError(f"{context}.required must be true or false")

    return StatisticComparisonDefinition(
        cell=cell.strip().upper(),
        workbook=workbook.strip() if isinstance(workbook, str) else None,
        sheet=sheet.strip() if isinstance(sheet, str) else None,
        absolute_tolerance=absolute_tolerance,
        relative_tolerance=relative_tolerance,
        required=required,
    )


def calculate_statistics(
    input_file: str | Path,
    config_file: str | Path,
    import_options: ImportOptions | None = None,
    math_config_file: str | Path | None = None,
) -> StatisticsResult:
    """Calculate configured statistics over raw, imported math and newly calculated math channels."""

    config_path = Path(config_file).expanduser().resolve()
    config = load_statistics_config(config_path)

    math_result: MathChannelsResult | None = None
    if math_config_file is not None:
        math_result = calculate_math_channels(input_file, math_config_file, import_options)
        dataset = math_result.dataset
    else:
        dataset = load_data_file(input_file, import_options)

    channels_by_id: dict[str, ChannelInfo] = {channel.channel_id: channel for channel in dataset.channels}
    values_by_id: dict[str, np.ndarray] = {
        channel.channel_id: dataset.values[:, index]
        for index, channel in enumerate(dataset.channels)
    }
    if math_result is not None:
        for index, channel in enumerate(math_result.calculated_channels):
            if channel.channel_id in channels_by_id:
                raise StatisticsError(
                    f"Calculated math channel '{channel.channel_id}' collides with an imported channel"
                )
            channels_by_id[channel.channel_id] = channel
            values_by_id[channel.channel_id] = math_result.calculated_values[:, index]

    missing = sorted({item.channel_id for item in config.statistics if item.channel_id not in channels_by_id})
    if missing:
        raise StatisticsError(_format_missing_channels(missing, channels_by_id))

    time_id = dataset.quality.time_channel_id
    time_values = values_by_id.get(time_id) if time_id is not None else None

    preliminary: list[tuple[StatisticDefinition, ChannelInfo, float, int, int, int]] = []
    for definition in config.statistics:
        channel = channels_by_id[definition.channel_id]
        values = values_by_id[definition.channel_id]
        statistic_value, used_count, omitted_count = compute_statistic(
            values,
            definition.operation,
            definition.nan_policy,
            time_values=time_values,
        )
        preliminary.append(
            (
                definition,
                channel,
                statistic_value,
                int(values.size),
                used_count,
                omitted_count,
            )
        )

    comparison_values = _load_all_comparison_values(
        config_path=config_path,
        source_path=dataset.source_path,
        default_sheet=dataset.quality.sheet_name,
        definitions=config.statistics,
    )

    results: list[StatisticResult] = []
    failed_required: list[StatisticComparisonResult] = []
    for definition, channel, value, sample_count, used_count, omitted_count in preliminary:
        comparison_result = None
        if definition.comparison is not None:
            comparison = definition.comparison
            reference_key = _comparison_key(
                config_path,
                dataset.source_path,
                dataset.quality.sheet_name,
                comparison,
            )
            reference_value = comparison_values[reference_key]
            absolute_error = abs(value - reference_value)
            if reference_value == 0.0:
                relative_error = 0.0 if absolute_error == 0.0 else math.inf
            else:
                relative_error = absolute_error / abs(reference_value)
            passed = bool(
                np.isclose(
                    value,
                    reference_value,
                    atol=comparison.absolute_tolerance,
                    rtol=comparison.relative_tolerance,
                    equal_nan=False,
                )
            )
            workbook_path, sheet_name, cell = reference_key
            comparison_result = StatisticComparisonResult(
                statistic_id=definition.statistic_id,
                workbook=str(workbook_path),
                sheet=sheet_name,
                cell=cell,
                reference_value=reference_value,
                calculated_value=value,
                absolute_tolerance=comparison.absolute_tolerance,
                relative_tolerance=comparison.relative_tolerance,
                required=comparison.required,
                passed=passed,
                absolute_error=absolute_error,
                relative_error=relative_error,
            )
            if comparison.required and not comparison_result.passed:
                failed_required.append(comparison_result)

        display_name = definition.display_name or f"{channel.display_name} {definition.operation}"
        results.append(
            StatisticResult(
                statistic_id=definition.statistic_id,
                channel_id=definition.channel_id,
                channel_display_name=channel.display_name,
                channel_unit=channel.unit,
                channel_kind=channel.kind,
                operation=definition.operation,
                placement_group=definition.placement_group,
                nan_policy=definition.nan_policy,
                value=value,
                sample_count=sample_count,
                used_sample_count=used_count,
                omitted_sample_count=omitted_count,
                display_name=display_name,
                description=definition.description,
                comparison=comparison_result,
            )
        )

    if failed_required:
        details = "; ".join(
            f"{item.statistic_id} vs {Path(item.workbook).name}!{item.cell}: "
            f"calculated={item.calculated_value:.12g}, reference={item.reference_value:.12g}, "
            f"abs error={item.absolute_error:.12g}"
            for item in failed_required
        )
        raise StatisticsError("Required statistic comparison failed: " + details)

    return StatisticsResult(
        dataset=dataset,
        config_path=config_path,
        config=config,
        channels_by_id=channels_by_id,
        values_by_id=values_by_id,
        statistics=results,
        math_result=math_result,
    )


def compute_statistic(
    values: Sequence[float] | np.ndarray,
    operation: str,
    nan_policy: str = "error",
    *,
    time_values: Sequence[float] | np.ndarray | None = None,
) -> tuple[float, int, int]:
    """Compute one deterministic statistic and return value, used count and omitted count."""

    if operation not in _ALLOWED_OPERATIONS:
        raise StatisticsError(f"Unsupported statistic operation '{operation}'")
    if nan_policy not in _ALLOWED_NAN_POLICIES:
        raise StatisticsError(f"Unsupported NaN policy '{nan_policy}'")

    data = np.asarray(values, dtype=np.float64)
    if data.ndim != 1 or data.size == 0:
        raise StatisticsError("Statistics require a non-empty one-dimensional channel")

    finite_mask = np.isfinite(data)
    non_finite_count = int((~finite_mask).sum())
    if nan_policy == "error" and non_finite_count:
        first = int(np.flatnonzero(~finite_mask)[0])
        raise StatisticsError(
            f"Statistic operation '{operation}' encountered {non_finite_count} non-finite values; "
            f"first at sample index {first}"
        )
    if nan_policy == "propagate" and non_finite_count:
        return math.nan, int(data.size), 0

    if nan_policy == "omit":
        working = data[finite_mask]
        omitted_count = non_finite_count
    else:
        working = data
        omitted_count = 0
    if working.size == 0:
        raise StatisticsError(f"Statistic operation '{operation}' has no finite samples to use")

    if operation == "rms":
        value = float(np.sqrt(np.mean(np.square(working))))
    elif operation == "max":
        value = float(np.max(working))
    elif operation == "min":
        value = float(np.min(working))
    elif operation == "first":
        value = float(working[0])
    elif operation == "last":
        value = float(working[-1])
    elif operation == "sum":
        value = float(np.sum(working, dtype=np.float64))
    elif operation == "positive_max":
        value = max_charging_power_kw(working)
    elif operation == "time_weighted_rms":
        if time_values is None:
            raise StatisticsError("time_weighted_rms requires a time channel")
        time = np.asarray(time_values, dtype=np.float64)
        if time.ndim != 1 or time.shape != data.shape:
            raise StatisticsError(
                f"time_weighted_rms requires time and data with identical one-dimensional shapes; "
                f"got {time.shape} and {data.shape}"
            )
        time_finite = np.isfinite(time)
        if not time_finite.all():
            raise StatisticsError("time_weighted_rms requires a fully finite time channel")
        if nan_policy == "omit":
            time = time[finite_mask]
        if time.size < 2:
            raise StatisticsError("time_weighted_rms requires at least two usable samples")
        intervals = np.diff(time)
        if np.any(intervals <= 0.0):
            raise StatisticsError("time_weighted_rms requires a strictly increasing time channel")
        duration = float(time[-1] - time[0])
        if duration <= 0.0:
            raise StatisticsError("time_weighted_rms requires a positive time duration")
        mean_square = float(np.trapezoid(np.square(working), time) / duration)
        value = math.sqrt(max(mean_square, 0.0))
    else:  # pragma: no cover
        raise StatisticsError(f"Unsupported statistic operation '{operation}'")

    if not math.isfinite(value):
        raise StatisticsError(f"Statistic operation '{operation}' produced a non-finite result")
    return value, int(working.size), omitted_count


def export_statistics(result: StatisticsResult, output_dir: str | Path) -> dict[str, Path]:
    output_path = Path(output_dir).expanduser().resolve()
    output_path.mkdir(parents=True, exist_ok=True)

    results_path = output_path / result.config.output.results_filename
    wide_path = output_path / result.config.output.wide_filename
    validation_path = output_path / "statistics_validation_report.json"
    manifest_path = output_path / "statistics_manifest.json"
    summary_path = output_path / "statistics_summary.txt"

    _write_long_results(result, results_path)
    _write_wide_results(result, wide_path)
    comparisons = [item.comparison for item in result.statistics if item.comparison is not None]
    validation_path.write_text(
        json.dumps(
            {
                "status": "PASS" if result.required_comparisons_passed else "FAIL",
                "comparison_count": len(comparisons),
                "required_comparisons_passed": result.required_comparisons_passed,
                "comparisons": [item.to_dict() for item in comparisons],
            },
            indent=2,
        ),
        encoding="utf-8",
    )
    manifest_path.write_text(json.dumps(_build_manifest(result, results_path, wide_path), indent=2), encoding="utf-8")
    summary_path.write_text(_format_summary(result, results_path, wide_path), encoding="utf-8")

    return {
        "statistics_results": results_path,
        "statistics_by_channel": wide_path,
        "statistics_validation_report": validation_path,
        "statistics_manifest": manifest_path,
        "statistics_summary": summary_path,
    }


def _write_long_results(result: StatisticsResult, path: Path) -> None:
    precision = result.config.output.float_precision
    fieldnames = [
        "statistics_order",
        "statistic_id",
        "channel_id",
        "channel_display_name",
        "channel_unit",
        "channel_kind",
        "operation",
        "value",
        "placement_group",
        "nan_policy",
        "sample_count",
        "used_sample_count",
        "omitted_sample_count",
        "display_name",
        "description",
        "comparison_reference",
        "comparison_value",
        "comparison_passed",
        "absolute_error",
        "relative_error",
        "comparison_required",
    ]
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        for index, item in enumerate(result.statistics, start=1):
            comparison = item.comparison
            writer.writerow(
                {
                    "statistics_order": index,
                    "statistic_id": item.statistic_id,
                    "channel_id": item.channel_id,
                    "channel_display_name": item.channel_display_name,
                    "channel_unit": item.channel_unit or "",
                    "channel_kind": item.channel_kind,
                    "operation": item.operation,
                    "value": _format_float(item.value, precision),
                    "placement_group": item.placement_group,
                    "nan_policy": item.nan_policy,
                    "sample_count": item.sample_count,
                    "used_sample_count": item.used_sample_count,
                    "omitted_sample_count": item.omitted_sample_count,
                    "display_name": item.display_name,
                    "description": item.description or "",
                    "comparison_reference": (
                        f"{comparison.workbook}|{comparison.sheet}!{comparison.cell}" if comparison else ""
                    ),
                    "comparison_value": _format_float(comparison.reference_value, precision) if comparison else "",
                    "comparison_passed": comparison.passed if comparison else "",
                    "absolute_error": _format_float(comparison.absolute_error, precision) if comparison else "",
                    "relative_error": (
                        _format_float(comparison.relative_error, precision)
                        if comparison and math.isfinite(comparison.relative_error)
                        else ("inf" if comparison else "")
                    ),
                    "comparison_required": comparison.required if comparison else "",
                }
            )


def _write_wide_results(result: StatisticsResult, path: Path) -> None:
    precision = result.config.output.float_precision
    ordered_channels: list[str] = []
    grouped: dict[str, dict[str, StatisticResult]] = {}
    for item in result.statistics:
        if item.channel_id not in grouped:
            grouped[item.channel_id] = {}
            ordered_channels.append(item.channel_id)
        grouped[item.channel_id][item.operation] = item

    operations = [
        operation
        for operation in ("rms", "time_weighted_rms", "max", "min", "first", "last", "sum", "positive_max")
        if any(operation in grouped[c] for c in ordered_channels)
    ]
    fieldnames = ["channel_id", "channel_display_name", "channel_unit", "channel_kind", *operations]
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        for channel_id in ordered_channels:
            channel = result.channels_by_id[channel_id]
            row: dict[str, Any] = {
                "channel_id": channel_id,
                "channel_display_name": channel.display_name,
                "channel_unit": channel.unit or "",
                "channel_kind": channel.kind,
            }
            for operation in operations:
                item = grouped[channel_id].get(operation)
                row[operation] = _format_float(item.value, precision) if item is not None else ""
            writer.writerow(row)


def _build_manifest(result: StatisticsResult, results_path: Path, wide_path: Path) -> dict[str, Any]:
    quality = result.dataset.quality
    return {
        "configuration_version": result.config.version,
        "source_file": str(result.dataset.source_path),
        "source_sha256": quality.source_sha256,
        "configuration_file": str(result.config_path),
        "math_configuration_file": str(result.math_result.config_path) if result.math_result else None,
        "sample_count": result.sample_count,
        "available_channel_count": len(result.channels_by_id),
        "statistic_count": result.statistic_count,
        "comparison_count": result.comparison_count,
        "required_comparisons_passed": result.required_comparisons_passed,
        "time_channel_id": quality.time_channel_id,
        "time_start": quality.time_start,
        "time_end": quality.time_end,
        "nominal_time_step": quality.nominal_time_step,
        "rms_definition": "sqrt(mean(x^2))",
        "time_weighted_rms_definition": "sqrt(trapezoidal_integral(x^2 dt) / elapsed_duration)",
        "statistics": [item.to_dict() for item in result.statistics],
        "statistics_results_file": str(results_path),
        "statistics_by_channel_file": str(wide_path),
        "float_precision": result.config.output.float_precision,
    }


def _format_summary(result: StatisticsResult, results_path: Path, wide_path: Path) -> str:
    lines = [
        "VSM STATISTICS",
        "==============",
        "Status: PASS",
        f"Source: {result.dataset.source_path}",
        f"Source SHA-256: {result.dataset.quality.source_sha256}",
        f"Configuration: {result.config_path}",
        f"Math configuration: {result.math_result.config_path if result.math_result else '-'}",
        f"Samples: {result.sample_count}",
        f"Available channels: {len(result.channels_by_id)}",
        f"Statistics calculated: {result.statistic_count}",
        f"Reference comparisons: {result.comparison_count}",
        f"Results: {results_path}",
        f"Wide results: {wide_path}",
        "",
        "Definitions:",
        "- RMS: sqrt(mean(x^2))",
        "- Time-weighted RMS: sqrt(trapezoidal integral of x^2 divided by elapsed duration)",
        "",
        "Calculated statistics:",
    ]
    for index, item in enumerate(result.statistics, start=1):
        comparison = ""
        if item.comparison is not None:
            comparison = (
                f" | {'PASS' if item.comparison.passed else 'FAIL'} vs "
                f"{Path(item.comparison.workbook).name}!{item.comparison.cell} "
                f"(reference={item.comparison.reference_value:.12g}, "
                f"abs error={item.comparison.absolute_error:.12g})"
            )
        lines.append(
            f"{index:02d}. {item.statistic_id} | {item.channel_id} | {item.operation} = "
            f"{item.value:.12g} [{item.channel_unit or '-'}] | {item.placement_group}{comparison}"
        )
    return "\n".join(lines) + "\n"


def _load_all_comparison_values(
    *,
    config_path: Path,
    source_path: Path,
    default_sheet: str | None,
    definitions: Sequence[StatisticDefinition],
) -> dict[tuple[Path, str, str], float]:
    grouped: dict[tuple[Path, str], set[str]] = {}
    for definition in definitions:
        if definition.comparison is None:
            continue
        workbook_path, sheet_name, cell = _comparison_key(
            config_path,
            source_path,
            default_sheet,
            definition.comparison,
        )
        grouped.setdefault((workbook_path, sheet_name), set()).add(cell)

    loaded: dict[tuple[Path, str, str], float] = {}
    for (workbook_path, sheet_name), cells in grouped.items():
        values = _read_cached_numeric_cells_xlsx(workbook_path, sheet_name, cells)
        for cell, value in values.items():
            loaded[(workbook_path, sheet_name, cell)] = value
    return loaded


def _comparison_key(
    config_path: Path,
    source_path: Path,
    default_sheet: str | None,
    comparison: StatisticComparisonDefinition,
) -> tuple[Path, str, str]:
    if comparison.workbook is None:
        workbook_path = source_path.resolve()
    else:
        candidate = Path(comparison.workbook).expanduser()
        workbook_path = (config_path.parent / candidate).resolve() if not candidate.is_absolute() else candidate.resolve()
    if workbook_path.suffix.lower() != ".xlsx":
        raise StatisticsError(
            f"Statistic comparison workbook must be an .xlsx file: {workbook_path}"
        )
    if not workbook_path.exists():
        raise StatisticsError(f"Statistic comparison workbook does not exist: {workbook_path}")
    sheet_name = comparison.sheet or default_sheet
    if not sheet_name:
        sheet_name = _first_xlsx_sheet_name(workbook_path)
    return workbook_path, sheet_name, comparison.cell


def _read_cached_numeric_cells_xlsx(
    workbook_path: Path,
    sheet_name: str,
    cells: set[str],
) -> dict[str, float]:
    requested = {cell.upper() for cell in cells}
    with zipfile.ZipFile(workbook_path) as archive:
        worksheet_path = _worksheet_xml_path(archive, sheet_name)
        namespace = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
        found: dict[str, float] = {}
        with archive.open(worksheet_path) as stream:
            for event, element in ElementTree.iterparse(stream, events=("end",)):
                if element.tag != namespace + "c":
                    continue
                address = element.attrib.get("r", "").upper()
                if address in requested:
                    value_element = element.find(namespace + "v")
                    if value_element is None or value_element.text is None:
                        raise StatisticsError(
                            f"Reference cell {sheet_name}!{address} in '{workbook_path}' has no cached value"
                        )
                    try:
                        value = float(value_element.text)
                    except ValueError as exc:
                        raise StatisticsError(
                            f"Reference cell {sheet_name}!{address} in '{workbook_path}' is not numeric"
                        ) from exc
                    if not math.isfinite(value):
                        raise StatisticsError(
                            f"Reference cell {sheet_name}!{address} in '{workbook_path}' is non-finite"
                        )
                    found[address] = value
                element.clear()
                if len(found) == len(requested):
                    break
    missing = sorted(requested - set(found))
    if missing:
        raise StatisticsError(
            f"Reference cells were not found in {workbook_path.name}!{sheet_name}: {', '.join(missing)}"
        )
    return found


def _worksheet_xml_path(archive: zipfile.ZipFile, sheet_name: str) -> str:
    main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    package_rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"

    workbook_root = ElementTree.fromstring(archive.read("xl/workbook.xml"))
    relationship_id = None
    for sheet in workbook_root.findall(f".//{{{main_ns}}}sheet"):
        if sheet.attrib.get("name") == sheet_name:
            relationship_id = sheet.attrib.get(f"{{{rel_ns}}}id")
            break
    if relationship_id is None:
        available = [sheet.attrib.get("name", "") for sheet in workbook_root.findall(f".//{{{main_ns}}}sheet")]
        raise StatisticsError(
            f"Sheet '{sheet_name}' not found in workbook. Available sheets: {', '.join(available)}"
        )

    relationships_root = ElementTree.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    target = None
    for relationship in relationships_root.findall(f"{{{package_rel_ns}}}Relationship"):
        if relationship.attrib.get("Id") == relationship_id:
            target = relationship.attrib.get("Target")
            break
    if not target:
        raise StatisticsError(f"Could not resolve worksheet '{sheet_name}' in workbook")

    target_path = PurePosixPath(target)
    if target_path.is_absolute():
        target_path = PurePosixPath(str(target_path).lstrip("/"))
    elif not str(target_path).startswith("xl/"):
        target_path = PurePosixPath("xl") / target_path
    worksheet_path = str(target_path)
    if worksheet_path not in archive.namelist():
        raise StatisticsError(f"Worksheet XML '{worksheet_path}' is missing from workbook")
    return worksheet_path


def _first_xlsx_sheet_name(workbook_path: Path) -> str:
    with zipfile.ZipFile(workbook_path) as archive:
        main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        root = ElementTree.fromstring(archive.read("xl/workbook.xml"))
        first = root.find(f".//{{{main_ns}}}sheet")
        if first is None or not first.attrib.get("name"):
            raise StatisticsError(f"Workbook '{workbook_path}' contains no worksheets")
        return first.attrib["name"]


def _format_missing_channels(missing: Sequence[str], available: Mapping[str, ChannelInfo]) -> str:
    available_ids = sorted(available)
    lines = ["Configured statistic channel IDs were not found:"]
    for channel_id in missing:
        suggestions = difflib.get_close_matches(channel_id, available_ids, n=3, cutoff=0.45)
        suffix = f" (possible channel_id: {', '.join(suggestions)})" if suggestions else ""
        lines.append(f"- {channel_id}{suffix}")
    lines.append("Use stable channel_id values from the channel catalogue or configured math channels.")
    return "\n".join(lines)


def _plain_csv_filename(value: object, field_name: str) -> str:
    if not isinstance(value, str) or not value.strip():
        raise ConfigurationError(f"{field_name} must be a non-empty string")
    filename = value.strip()
    path = Path(filename)
    if path.name != filename or path.suffix.lower() != ".csv":
        raise ConfigurationError(f"{field_name} must be a plain .csv filename without directories")
    return filename


def _format_float(value: float, precision: int) -> str:
    if math.isnan(value):
        return "nan"
    if not math.isfinite(value):
        raise StatisticsError("Statistics output contains an infinite value")
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
