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
from matplotlib.ticker import MaxNLocator, StrMethodFormatter
import numpy as np
import yaml

from .errors import ConfigurationError, PlottingError
from .importer import ImportOptions, load_data_file
from .math_engine import MathChannelsResult, calculate_math_channels
from .models import ChannelInfo, ImportedDataset
from .utils import sha256_file

_IDENTIFIER = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")
_ALLOWED_AXES = {"primary", "secondary"}
_ALLOWED_OUTPUT_FORMATS = {"png", "svg"}
_RAW_ID_PATTERN = re.compile(r"__[A-Za-z]+_\d{3}\b")


@dataclass(frozen=True)
class PlotStyle:
    title_fontsize: float = 15.0
    axis_label_fontsize: float = 11.0
    tick_fontsize: float = 9.0
    legend_fontsize: float = 9.0
    legend_location: str = "best"
    legend_columns: int = 1
    grid_alpha: float = 0.28
    grid_linewidth: float = 0.6
    primary_line_style: str = "-"
    secondary_line_style: str = "--"
    background: str = "white"
    transparent: bool = False
    constrained_layout: bool = False
    tight_layout: bool = True
    zero_line: bool = True
    output_formats: tuple[str, ...] = ("png",)


@dataclass(frozen=True)
class PlotDefaults:
    width_inches: float = 12.0
    height_inches: float = 6.0
    dpi: int = 150
    grid: bool = True
    legend: bool = True
    line_width: float = 1.2
    style: PlotStyle = PlotStyle()


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
    legend_labels: tuple[str, ...]
    reference_chart_number: int | None
    sample_count: int
    figure_width_inches: float
    figure_height_inches: float
    dpi: int
    axes_count: int
    png_file: str
    svg_file: str | None = None

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
            "legend_labels": list(self.legend_labels),
            "reference_chart_number": self.reference_chart_number,
            "sample_count": self.sample_count,
            "figure_width_inches": self.figure_width_inches,
            "figure_height_inches": self.figure_height_inches,
            "dpi": self.dpi,
            "axes_count": self.axes_count,
            "png_file": self.png_file,
            "svg_file": self.svg_file,
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

    @property
    def secondary_axis_plot_count(self) -> int:
        return sum(1 for item in self.rendered_plots if item.secondary_series_ids)

    @property
    def svg_count(self) -> int:
        return sum(1 for item in self.rendered_plots if item.svg_file is not None)


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
    _reject_unknown_keys(raw, {"version", "defaults", "style", "plots"}, "root")

    version = raw.get("version")
    if version != 1:
        raise ConfigurationError("Plotting configuration 'version' must be 1")

    defaults = _load_defaults(raw.get("defaults", {}), raw.get("style", {}))
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


def _load_defaults(raw: object, style_raw: object | None = None) -> PlotDefaults:
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
    return PlotDefaults(width, height, dpi, grid, legend, line_width, _load_style(style_raw))


def _load_style(raw: object | None) -> PlotStyle:
    if raw is None:
        raw = {}
    if not isinstance(raw, dict):
        raise ConfigurationError("style must be a YAML mapping")
    _reject_unknown_keys(
        raw,
        {
            "title_fontsize",
            "axis_label_fontsize",
            "tick_fontsize",
            "legend_fontsize",
            "legend_location",
            "legend_columns",
            "grid_alpha",
            "grid_linewidth",
            "primary_line_style",
            "secondary_line_style",
            "background",
            "transparent",
            "constrained_layout",
            "tight_layout",
            "zero_line",
            "output_formats",
        },
        "style",
    )
    legend_columns = raw.get("legend_columns", 1)
    if isinstance(legend_columns, bool) or not isinstance(legend_columns, int) or not 1 <= legend_columns <= 6:
        raise ConfigurationError("style.legend_columns must be an integer from 1 to 6")
    output_formats_raw = raw.get("output_formats", ["png"])
    if not isinstance(output_formats_raw, list) or not output_formats_raw:
        raise ConfigurationError("style.output_formats must be a non-empty list")
    output_formats = tuple(str(item).strip().lower() for item in output_formats_raw)
    invalid_formats = sorted(set(output_formats) - _ALLOWED_OUTPUT_FORMATS)
    if invalid_formats:
        raise ConfigurationError("style.output_formats contains unsupported format(s): " + ", ".join(invalid_formats))
    if "png" not in output_formats:
        raise ConfigurationError("style.output_formats must include png")
    for key in ("transparent", "constrained_layout", "tight_layout", "zero_line"):
        if not isinstance(raw.get(key, getattr(PlotStyle(), key)), bool):
            raise ConfigurationError(f"style.{key} must be true or false")
    return PlotStyle(
        title_fontsize=_positive_number(raw.get("title_fontsize", 15.0), "style.title_fontsize"),
        axis_label_fontsize=_positive_number(raw.get("axis_label_fontsize", 11.0), "style.axis_label_fontsize"),
        tick_fontsize=_positive_number(raw.get("tick_fontsize", 9.0), "style.tick_fontsize"),
        legend_fontsize=_positive_number(raw.get("legend_fontsize", 9.0), "style.legend_fontsize"),
        legend_location=_nonempty_string(raw.get("legend_location", "best"), "style.legend_location"),
        legend_columns=legend_columns,
        grid_alpha=_ratio(raw.get("grid_alpha", 0.28), "style.grid_alpha"),
        grid_linewidth=_positive_number(raw.get("grid_linewidth", 0.6), "style.grid_linewidth"),
        primary_line_style=_nonempty_string(raw.get("primary_line_style", "-"), "style.primary_line_style"),
        secondary_line_style=_nonempty_string(raw.get("secondary_line_style", "--"), "style.secondary_line_style"),
        background=_nonempty_string(raw.get("background", "white"), "style.background"),
        transparent=raw.get("transparent", False),
        constrained_layout=raw.get("constrained_layout", False),
        tight_layout=raw.get("tight_layout", True),
        zero_line=raw.get("zero_line", True),
        output_formats=output_formats,
    )


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

    style = defaults.style
    figure, primary_axis = plt.subplots(
        figsize=(defaults.width_inches, defaults.height_inches),
        constrained_layout=style.constrained_layout,
        facecolor=style.background,
    )
    figure.patch.set_alpha(0.0 if style.transparent else 1.0)
    primary_axis.set_facecolor(style.background)
    secondary_axis = None
    primary_ids: list[str] = []
    secondary_ids: list[str] = []
    legend_labels: list[str] = []

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
            label = _clean_visible_text(item.label or channel.display_name)
            legend_labels.append(label)
            axis.plot(
                x_values,
                values,
                label=label,
                linewidth=defaults.line_width,
                linestyle=style.secondary_line_style if item.axis == "secondary" else style.primary_line_style,
                antialiased=True,
            )

        primary_axis.set_title(
            _clean_visible_text(definition.title),
            fontsize=style.title_fontsize,
            pad=12,
        )
        primary_axis.set_xlabel(
            _clean_visible_text(definition.x_label or _axis_label(x_channel)),
            fontsize=style.axis_label_fontsize,
        )
        primary_axis.set_ylabel(
            _clean_visible_text(
                definition.primary_y_label
                or _automatic_y_label(primary_ids, channels_by_id, fallback="Value")
            ),
            fontsize=style.axis_label_fontsize,
        )
        if secondary_axis is not None:
            secondary_axis.set_facecolor("none")
            secondary_axis.set_ylabel(
                _clean_visible_text(
                    definition.secondary_y_label
                    or _automatic_y_label(secondary_ids, channels_by_id, fallback="Value")
                ),
                fontsize=style.axis_label_fontsize,
            )
        _apply_axis_format(primary_axis, definition.x_label or _axis_label(x_channel), axis_name="x")
        _apply_axis_format(
            primary_axis,
            definition.primary_y_label or _automatic_y_label(primary_ids, channels_by_id, fallback="Value"),
            axis_name="y",
        )
        if secondary_axis is not None:
            _apply_axis_format(
                secondary_axis,
                definition.secondary_y_label or _automatic_y_label(secondary_ids, channels_by_id, fallback="Value"),
                axis_name="y",
            )
        if defaults.grid:
            primary_axis.grid(True, color="#B8C2CC", alpha=style.grid_alpha, linewidth=style.grid_linewidth)
            primary_axis.set_axisbelow(True)
        if style.zero_line:
            _draw_zero_line(primary_axis)
            if secondary_axis is not None:
                _draw_zero_line(secondary_axis)
        if defaults.legend:
            handles, labels = primary_axis.get_legend_handles_labels()
            if secondary_axis is not None:
                second_handles, second_labels = secondary_axis.get_legend_handles_labels()
                handles += second_handles
                labels += second_labels
            if handles:
                primary_axis.legend(
                    handles,
                    labels,
                    loc=style.legend_location,
                    ncols=style.legend_columns,
                    fontsize=style.legend_fontsize,
                    frameon=True,
                    framealpha=0.92,
                    facecolor="white",
                    edgecolor="#B8C2CC",
                )

        for axis in (primary_axis, secondary_axis):
            if axis is not None:
                axis.tick_params(labelsize=style.tick_fontsize)
                for spine in axis.spines.values():
                    spine.set_color("#7A8691")
                    spine.set_linewidth(0.8)
        if style.tight_layout and not style.constrained_layout:
            figure.tight_layout()
        png_dir = output_dir / "png"
        png_dir.mkdir(parents=True, exist_ok=True)
        output_path = png_dir / definition.output_filename
        figure.savefig(
            output_path,
            dpi=defaults.dpi,
            format="png",
            transparent=style.transparent,
            facecolor=figure.get_facecolor(),
        )
        svg_path: Path | None = None
        if "svg" in style.output_formats:
            svg_dir = output_dir / "svg"
            svg_dir.mkdir(parents=True, exist_ok=True)
            svg_path = svg_dir / f"{Path(definition.output_filename).stem}.svg"
            figure.savefig(svg_path, format="svg", transparent=style.transparent, facecolor=figure.get_facecolor())
    finally:
        plt.close(figure)

    return RenderedPlot(
        plot_id=definition.plot_id,
        title=_clean_visible_text(definition.title),
        output_file=str(output_path.resolve()),
        x_channel_id=definition.x_channel_id,
        x_display_name=x_channel.display_name,
        x_unit=x_channel.unit,
        primary_series_ids=tuple(primary_ids),
        secondary_series_ids=tuple(secondary_ids),
        legend_labels=tuple(legend_labels),
        reference_chart_number=definition.reference_chart_number,
        sample_count=sample_count,
        figure_width_inches=defaults.width_inches,
        figure_height_inches=defaults.height_inches,
        dpi=defaults.dpi,
        axes_count=2 if secondary_axis is not None else 1,
        png_file=str(output_path.resolve()),
        svg_file=str(svg_path.resolve()) if svg_path is not None else None,
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
            "legend_labels",
            "sample_count",
            "axes_count",
            "png_file",
            "svg_file",
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
                    "legend_labels": ";".join(item.legend_labels),
                    "sample_count": item.sample_count,
                    "axes_count": item.axes_count,
                    "png_file": item.png_file,
                    "svg_file": item.svg_file or "",
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
            "style": {
                "title_fontsize": result.config.defaults.style.title_fontsize,
                "axis_label_fontsize": result.config.defaults.style.axis_label_fontsize,
                "tick_fontsize": result.config.defaults.style.tick_fontsize,
                "legend_fontsize": result.config.defaults.style.legend_fontsize,
                "legend_location": result.config.defaults.style.legend_location,
                "legend_columns": result.config.defaults.style.legend_columns,
                "grid_alpha": result.config.defaults.style.grid_alpha,
                "grid_linewidth": result.config.defaults.style.grid_linewidth,
                "transparent": result.config.defaults.style.transparent,
                "constrained_layout": result.config.defaults.style.constrained_layout,
                "tight_layout": result.config.defaults.style.tight_layout,
                "zero_line": result.config.defaults.style.zero_line,
                "output_formats": list(result.config.defaults.style.output_formats),
            },
        },
        "secondary_axis_plot_count": result.secondary_axis_plot_count,
        "svg_count": result.svg_count,
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
        f"Secondary-axis plots: {result.secondary_axis_plot_count}",
        f"SVG plots: {result.svg_count}",
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


def _apply_axis_format(plot_axis, label: str, *, axis_name: str) -> None:
    label_lower = label.casefold()
    axis_obj = getattr(plot_axis, f"{axis_name}axis")
    axis_obj.set_major_locator(MaxNLocator(nbins=8, min_n_ticks=4))
    if "%" in label or "soc" in label_lower:
        axis_obj.set_major_formatter(StrMethodFormatter("{x:.0f}"))
    elif any(unit in label_lower for unit in ("km", "kph", "kw", "kwh", "kg", "l/h", "litre", "rpm", "nm")):
        axis_obj.set_major_formatter(StrMethodFormatter("{x:.1f}"))
    elif "time" in label_lower or "min" in label_lower or "s]" in label_lower:
        axis_obj.set_major_formatter(StrMethodFormatter("{x:.1f}"))
    else:
        axis_obj.set_major_formatter(StrMethodFormatter("{x:g}"))


def _draw_zero_line(axis) -> None:
    ymin, ymax = axis.get_ylim()
    if ymin < 0.0 < ymax:
        axis.axhline(0.0, color="#4D5963", linewidth=0.75, alpha=0.55, zorder=0)


def _clean_visible_text(value: str) -> str:
    text = value.strip()
    if _RAW_ID_PATTERN.search(text):
        text = _RAW_ID_PATTERN.sub("", text)
    text = text.replace("_", " ")
    return " ".join(text.split())


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


def _ratio(value: object, context: str) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ConfigurationError(f"{context} must be from 0.0 to 1.0")
    result = float(value)
    if not math.isfinite(result) or not 0.0 <= result <= 1.0:
        raise ConfigurationError(f"{context} must be from 0.0 to 1.0")
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
