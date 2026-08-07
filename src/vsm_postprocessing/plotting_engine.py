from __future__ import annotations

import csv
import difflib
import json
import math
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Mapping, Sequence

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import yaml

from .errors import ConfigurationError, PlottingError
from .importer import ImportOptions, load_data_file
from .math_engine import MathChannelsResult, calculate_math_channels
from .models import ChannelInfo, ImportedDataset
from .utils import sha256_file

_IDENTIFIER = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")
_ALLOWED_AXES = {"primary", "secondary"}


@dataclass(frozen=True)
class PlotDefaults:
    width_inches: float = 12.0
    height_inches: float = 6.0
    dpi: int = 150
    grid: bool = True
    legend: bool = True
    line_width: float = 1.2


@dataclass(frozen=True)
class PlotSeriesDefinition:
    channel_id: str
    axis: str = "primary"
    label: str | None = None


@dataclass(frozen=True)
class PlotDefinition:
    plot_id: str
    title: str
    x_channel_id: str
    series: tuple[PlotSeriesDefinition, ...]
    output_filename: str
    x_label: str | None = None
    primary_y_label: str | None = None
    secondary_y_label: str | None = None
    reference_chart_number: int | None = None


@dataclass(frozen=True)
class PlottingConfig:
    version: int
    defaults: PlotDefaults
    plots: tuple[PlotDefinition, ...]


@dataclass(frozen=True)
class RenderedPlot:
    plot_id: str
    title: str
    output_file: str
    x_channel_id: str
    x_display_name: str
    x_unit: str | None
    primary_series_ids: tuple[str, ...]
    secondary_series_ids: tuple[str, ...]
    reference_chart_number: int | None
    sample_count: int

    def to_dict(self) -> dict[str, Any]:
        return {
            "plot_id": self.plot_id,
            "title": self.title,
            "output_file": self.output_file,
            "x_channel_id": self.x_channel_id,
            "x_display_name": self.x_display_name,
            "x_unit": self.x_unit,
            "primary_series_ids": list(self.primary_series_ids),
            "secondary_series_ids": list(self.secondary_series_ids),
            "reference_chart_number": self.reference_chart_number,
            "sample_count": self.sample_count,
        }


@dataclass
class PlottingResult:
    dataset: ImportedDataset
    config_path: Path
    config: PlottingConfig
    channels_by_id: dict[str, ChannelInfo]
    values_by_id: dict[str, np.ndarray]
    rendered_plots: list[RenderedPlot]
    math_result: MathChannelsResult | None = None

    @property
    def sample_count(self) -> int:
        return self.dataset.quality.sample_count

    @property
    def plot_count(self) -> int:
        return len(self.rendered_plots)

    @property
    def series_count(self) -> int:
        return sum(
            len(item.primary_series_ids) + len(item.secondary_series_ids)
            for item in self.rendered_plots
        )


def load_plotting_config(path: str | Path) -> PlottingConfig:
    config_path = Path(path).expanduser().resolve()
    if not config_path.exists():
        raise ConfigurationError(f"Plotting configuration does not exist: {config_path}")
    if not config_path.is_file():
        raise ConfigurationError(f"Plotting configuration is not a file: {config_path}")

    try:
        raw = yaml.safe_load(config_path.read_text(encoding="utf-8"))
    except UnicodeDecodeError as exc:
        raise ConfigurationError(f"Plotting configuration must be UTF-8: {config_path}") from exc
    except yaml.YAMLError as exc:
        raise ConfigurationError(f"Invalid YAML in plotting configuration '{config_path}': {exc}") from exc

    if not isinstance(raw, dict):
        raise ConfigurationError("Plotting configuration root must be a YAML mapping")
    _reject_unknown_keys(raw, {"version", "defaults", "plots"}, "root")

    version = raw.get("version")
    if version != 1:
        raise ConfigurationError("Plotting configuration 'version' must be 1")

    defaults = _load_defaults(raw.get("defaults", {}))
    plots_raw = raw.get("plots")
    if not isinstance(plots_raw, list) or not plots_raw:
        raise ConfigurationError("plots must be a non-empty YAML list")
    plots = tuple(_load_plot(item, index) for index, item in enumerate(plots_raw, start=1))

    duplicate_ids = _duplicates([item.plot_id for item in plots])
    if duplicate_ids:
        raise ConfigurationError("plots contains duplicate plot IDs: " + ", ".join(duplicate_ids))
    duplicate_files = _duplicates([item.output_filename.casefold() for item in plots])
    if duplicate_files:
        raise ConfigurationError("plots contains duplicate output filenames")
    reference_numbers = [item.reference_chart_number for item in plots if item.reference_chart_number is not None]
    duplicate_reference_numbers = _duplicates(reference_numbers)
    if duplicate_reference_numbers:
        raise ConfigurationError(
            "plots contains duplicate reference_chart_number values: "
            + ", ".join(str(value) for value in duplicate_reference_numbers)
        )

    return PlottingConfig(version=version, defaults=defaults, plots=plots)


def _load_defaults(raw: object) -> PlotDefaults:
    if raw is None:
        raw = {}
    if not isinstance(raw, dict):
        raise ConfigurationError("defaults must be a YAML mapping")
    _reject_unknown_keys(
        raw,
        {"width_inches", "height_inches", "dpi", "grid", "legend", "line_width"},
        "defaults",
    )

    width = _positive_number(raw.get("width_inches", 12.0), "defaults.width_inches")
    height = _positive_number(raw.get("height_inches", 6.0), "defaults.height_inches")
    dpi = raw.get("dpi", 150)
    if isinstance(dpi, bool) or not isinstance(dpi, int) or not 50 <= dpi <= 600:
        raise ConfigurationError("defaults.dpi must be an integer from 50 to 600")
    grid = raw.get("grid", True)
    legend = raw.get("legend", True)
    if not isinstance(grid, bool):
        raise ConfigurationError("defaults.grid must be true or false")
    if not isinstance(legend, bool):
        raise ConfigurationError("defaults.legend must be true or false")
    line_width = _positive_number(raw.get("line_width", 1.2), "defaults.line_width")
    return PlotDefaults(width, height, dpi, grid, legend, line_width)


def _load_plot(raw: object, index: int) -> PlotDefinition:
    context = f"plots[{index}]"
    if not isinstance(raw, dict):
        raise ConfigurationError(f"{context} must be a YAML mapping")
    _reject_unknown_keys(
        raw,
        {
            "plot_id",
            "title",
            "x_channel_id",
            "x_label",
            "primary_y_label",
            "secondary_y_label",
            "output_filename",
            "reference_chart_number",
            "series",
        },
        context,
    )

    plot_id = raw.get("plot_id")
    if not isinstance(plot_id, str) or not _IDENTIFIER.fullmatch(plot_id):
        raise ConfigurationError(f"{context}.plot_id must be a valid Python-style identifier")

    title = _nonempty_string(raw.get("title"), f"{context}.title")
    x_channel_id = _nonempty_string(raw.get("x_channel_id"), f"{context}.x_channel_id")
    output_filename = _png_filename(raw.get("output_filename", f"{plot_id}.png"), f"{context}.output_filename")

    x_label = _optional_string(raw.get("x_label"), f"{context}.x_label")
    primary_y_label = _optional_string(raw.get("primary_y_label"), f"{context}.primary_y_label")
    secondary_y_label = _optional_string(raw.get("secondary_y_label"), f"{context}.secondary_y_label")

    reference_chart_number = raw.get("reference_chart_number")
    if reference_chart_number is not None and (
        isinstance(reference_chart_number, bool)
        or not isinstance(reference_chart_number, int)
        or reference_chart_number < 1
    ):
        raise ConfigurationError(f"{context}.reference_chart_number must be null or a positive integer")

    series_raw = raw.get("series")
    if not isinstance(series_raw, list) or not series_raw:
        raise ConfigurationError(f"{context}.series must be a non-empty YAML list")
    series = tuple(_load_series(item, context, idx) for idx, item in enumerate(series_raw, start=1))

    return PlotDefinition(
        plot_id=plot_id,
        title=title,
        x_channel_id=x_channel_id,
        series=series,
        output_filename=output_filename,
        x_label=x_label,
        primary_y_label=primary_y_label,
        secondary_y_label=secondary_y_label,
        reference_chart_number=reference_chart_number,
    )


def _load_series(raw: object, plot_context: str, index: int) -> PlotSeriesDefinition:
    context = f"{plot_context}.series[{index}]"
    if not isinstance(raw, dict):
        raise ConfigurationError(f"{context} must be a YAML mapping")
    _reject_unknown_keys(raw, {"channel_id", "axis", "label"}, context)
    channel_id = _nonempty_string(raw.get("channel_id"), f"{context}.channel_id")
    axis = raw.get("axis", "primary")
    if axis not in _ALLOWED_AXES:
        raise ConfigurationError(f"{context}.axis must be 'primary' or 'secondary'")
    label = _optional_string(raw.get("label"), f"{context}.label")
    return PlotSeriesDefinition(channel_id=channel_id, axis=axis, label=label)


def render_plots(
    input_file: str | Path,
    config_file: str | Path,
    output_dir: str | Path,
    import_options: ImportOptions | None = None,
    math_config_file: str | Path | None = None,
) -> PlottingResult:
    config_path = Path(config_file).expanduser().resolve()
    config = load_plotting_config(config_path)

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
                raise PlottingError(f"Calculated math channel '{channel.channel_id}' collides with an imported channel")
            channels_by_id[channel.channel_id] = channel
            values_by_id[channel.channel_id] = math_result.calculated_values[:, index]

    requested = {definition.x_channel_id for definition in config.plots}
    requested.update(series.channel_id for definition in config.plots for series in definition.series)
    missing = sorted(channel_id for channel_id in requested if channel_id not in channels_by_id)
    if missing:
        raise PlottingError(_format_missing_channels(missing, channels_by_id))

    destination = Path(output_dir).expanduser().resolve()
    destination.mkdir(parents=True, exist_ok=True)

    rendered: list[RenderedPlot] = []
    for definition in config.plots:
        rendered.append(
            _render_one_plot(
                definition,
                config.defaults,
                destination,
                channels_by_id,
                values_by_id,
                dataset.quality.sample_count,
            )
        )

    result = PlottingResult(
        dataset=dataset,
        config_path=config_path,
        config=config,
        channels_by_id=channels_by_id,
        values_by_id=values_by_id,
        rendered_plots=rendered,
        math_result=math_result,
    )
    _write_metadata(result, destination)
    return result


def _render_one_plot(
    definition: PlotDefinition,
    defaults: PlotDefaults,
    output_dir: Path,
    channels_by_id: Mapping[str, ChannelInfo],
    values_by_id: Mapping[str, np.ndarray],
    sample_count: int,
) -> RenderedPlot:
    x_channel = channels_by_id[definition.x_channel_id]
    x_values = np.asarray(values_by_id[definition.x_channel_id], dtype=np.float64)
    _validate_series_values(x_values, definition.x_channel_id, "x-axis")

    figure, primary_axis = plt.subplots(figsize=(defaults.width_inches, defaults.height_inches))
    secondary_axis = None
    primary_ids: list[str] = []
    secondary_ids: list[str] = []

    try:
        for item in definition.series:
            channel = channels_by_id[item.channel_id]
            values = np.asarray(values_by_id[item.channel_id], dtype=np.float64)
            _validate_series_values(values, item.channel_id, "series")
            axis = primary_axis
            if item.axis == "secondary":
                if secondary_axis is None:
                    secondary_axis = primary_axis.twinx()
                axis = secondary_axis
                secondary_ids.append(item.channel_id)
            else:
                primary_ids.append(item.channel_id)
            axis.plot(
                x_values,
                values,
                label=item.label or channel.display_name,
                linewidth=defaults.line_width,
                linestyle="--" if item.axis == "secondary" else "-",
            )

        primary_axis.set_title(definition.title)
        primary_axis.set_xlabel(definition.x_label or _axis_label(x_channel))
        primary_axis.set_ylabel(
            definition.primary_y_label
            or _automatic_y_label(primary_ids, channels_by_id, fallback="Value")
        )
        if secondary_axis is not None:
            secondary_axis.set_ylabel(
                definition.secondary_y_label
                or _automatic_y_label(secondary_ids, channels_by_id, fallback="Value")
            )
        if defaults.grid:
            primary_axis.grid(True)
        if defaults.legend:
            handles, labels = primary_axis.get_legend_handles_labels()
            if secondary_axis is not None:
                second_handles, second_labels = secondary_axis.get_legend_handles_labels()
                handles += second_handles
                labels += second_labels
            if handles:
                primary_axis.legend(handles, labels, loc="best")

        figure.tight_layout()
        output_path = output_dir / definition.output_filename
        figure.savefig(output_path, dpi=defaults.dpi, format="png")
    finally:
        plt.close(figure)

    return RenderedPlot(
        plot_id=definition.plot_id,
        title=definition.title,
        output_file=str((output_dir / definition.output_filename).resolve()),
        x_channel_id=definition.x_channel_id,
        x_display_name=x_channel.display_name,
        x_unit=x_channel.unit,
        primary_series_ids=tuple(primary_ids),
        secondary_series_ids=tuple(secondary_ids),
        reference_chart_number=definition.reference_chart_number,
        sample_count=sample_count,
    )


def _write_metadata(result: PlottingResult, output_dir: Path) -> None:
    catalogue_path = output_dir / "plot_catalogue.csv"
    manifest_path = output_dir / "plot_manifest.json"
    summary_path = output_dir / "plotting_summary.txt"

    with catalogue_path.open("w", newline="", encoding="utf-8") as handle:
        fieldnames = [
            "plot_order",
            "plot_id",
            "title",
            "reference_chart_number",
            "output_file",
            "x_channel_id",
            "x_display_name",
            "x_unit",
            "primary_series_ids",
            "secondary_series_ids",
            "sample_count",
        ]
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        for index, item in enumerate(result.rendered_plots, start=1):
            writer.writerow(
                {
                    "plot_order": index,
                    "plot_id": item.plot_id,
                    "title": item.title,
                    "reference_chart_number": item.reference_chart_number or "",
                    "output_file": item.output_file,
                    "x_channel_id": item.x_channel_id,
                    "x_display_name": item.x_display_name,
                    "x_unit": item.x_unit or "",
                    "primary_series_ids": ";".join(item.primary_series_ids),
                    "secondary_series_ids": ";".join(item.secondary_series_ids),
                    "sample_count": item.sample_count,
                }
            )

    manifest = {
        "configuration_version": result.config.version,
        "source_file": str(result.dataset.source_path),
        "source_sha256": result.dataset.quality.source_sha256,
        "configuration_file": str(result.config_path),
        "configuration_sha256": sha256_file(result.config_path),
        "sample_count": result.sample_count,
        "available_channel_count": len(result.channels_by_id),
        "plot_count": result.plot_count,
        "series_count": result.series_count,
        "defaults": {
            "width_inches": result.config.defaults.width_inches,
            "height_inches": result.config.defaults.height_inches,
            "dpi": result.config.defaults.dpi,
            "grid": result.config.defaults.grid,
            "legend": result.config.defaults.legend,
            "line_width": result.config.defaults.line_width,
        },
        "plots": [item.to_dict() for item in result.rendered_plots],
    }
    manifest_path.write_text(json.dumps(manifest, indent=2), encoding="utf-8")

    lines = [
        "VSM PLOTTING",
        "============",
        "Status: PASS",
        f"Source: {result.dataset.source_path}",
        f"Source SHA-256: {result.dataset.quality.source_sha256}",
        f"Configuration: {result.config_path}",
        f"Samples: {result.sample_count}",
        f"Available channels: {len(result.channels_by_id)}",
        f"Plots rendered: {result.plot_count}",
        f"Series rendered: {result.series_count}",
        "",
        "Plots:",
    ]
    for index, item in enumerate(result.rendered_plots, start=1):
        reference = f" | reference chart {item.reference_chart_number}" if item.reference_chart_number else ""
        lines.append(
            f"{index:02d}. {item.plot_id} | {item.title} | x={item.x_channel_id} | "
            f"primary={','.join(item.primary_series_ids) or '-'} | "
            f"secondary={','.join(item.secondary_series_ids) or '-'}{reference}"
        )
    summary_path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def _validate_series_values(values: np.ndarray, channel_id: str, role: str) -> None:
    if values.ndim != 1 or values.size == 0:
        raise PlottingError(f"Plot {role} channel '{channel_id}' must be a non-empty one-dimensional array")
    non_finite = int((~np.isfinite(values)).sum())
    if non_finite:
        first = int(np.flatnonzero(~np.isfinite(values))[0])
        raise PlottingError(
            f"Plot {role} channel '{channel_id}' contains {non_finite} non-finite values; first at sample index {first}"
        )


def _axis_label(channel: ChannelInfo) -> str:
    return f"{channel.display_name} [{channel.unit}]" if channel.unit else channel.display_name


def _automatic_y_label(
    channel_ids: Sequence[str], channels_by_id: Mapping[str, ChannelInfo], *, fallback: str
) -> str:
    if not channel_ids:
        return fallback
    units = {channels_by_id[channel_id].unit for channel_id in channel_ids}
    names = [channels_by_id[channel_id].display_name for channel_id in channel_ids]
    if len(channel_ids) == 1:
        return _axis_label(channels_by_id[channel_ids[0]])
    if len(units) == 1:
        unit = next(iter(units))
        return f"Value [{unit}]" if unit else "Value"
    return " / ".join(names)


def _format_missing_channels(missing: Sequence[str], channels_by_id: Mapping[str, ChannelInfo]) -> str:
    available = sorted(channels_by_id)
    lines = ["Configured plotting channel IDs were not found:"]
    for channel_id in missing:
        suggestions = difflib.get_close_matches(channel_id, available, n=3, cutoff=0.45)
        suffix = f" (possible channel_id: {', '.join(suggestions)})" if suggestions else ""
        lines.append(f"- {channel_id}{suffix}")
    lines.append("Use stable channel_id values from the channel catalogue or configured math-channel IDs.")
    return "\n".join(lines)


def _positive_number(value: object, context: str) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ConfigurationError(f"{context} must be a positive finite number")
    result = float(value)
    if not math.isfinite(result) or result <= 0.0:
        raise ConfigurationError(f"{context} must be a positive finite number")
    return result


def _nonempty_string(value: object, context: str) -> str:
    if not isinstance(value, str) or not value.strip():
        raise ConfigurationError(f"{context} must be a non-empty string")
    return value.strip()


def _optional_string(value: object, context: str) -> str | None:
    if value is None:
        return None
    if not isinstance(value, str) or not value.strip():
        raise ConfigurationError(f"{context} must be null or a non-empty string")
    return value.strip()


def _png_filename(value: object, context: str) -> str:
    filename = _nonempty_string(value, context)
    path = Path(filename)
    if path.name != filename or path.suffix.lower() != ".png":
        raise ConfigurationError(f"{context} must be a plain .png filename without directories")
    return filename


def _duplicates(values: Sequence[Any]) -> list[Any]:
    seen: set[Any] = set()
    duplicates: list[Any] = []
    for value in values:
        if value in seen and value not in duplicates:
            duplicates.append(value)
        seen.add(value)
    return duplicates


def _reject_unknown_keys(raw: Mapping[str, Any], allowed: set[str], context: str) -> None:
    unknown = sorted(set(raw) - allowed)
    if unknown:
        raise ConfigurationError(f"Unknown key(s) in {context}: {', '.join(unknown)}")
