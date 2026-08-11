from __future__ import annotations

import json
import math
import shutil
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Mapping

import yaml
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .errors import ConfigurationError, ExcelReportError
from .importer import ImportOptions
from .models import ChannelInfo
from .plotting_engine import PlottingResult, load_plotting_config, render_plots
from .statistics_engine import StatisticResult, StatisticsResult, calculate_statistics
from .utils import client_display_filename, sha256_file
from .version import __version__

_ALLOWED_BOTTOM_OPERATIONS = ("max", "min", "last", "sum", "rms", "time_weighted_rms")


@dataclass(frozen=True)
class ExcelReportConfig:
    version: int
    title: str
    subtitle: str | None
    report_sheet: str
    metadata_sheet: str
    channel_ids: tuple[str, ...]
    top_rms_statistic_ids: tuple[str, ...]
    kpi_statistic_ids: tuple[str, ...]
    bottom_operations: tuple[str, ...]
    bottom_summary_statistic_ids: tuple[str, ...]
    layout_profile: str
    plot_placement: str
    blank_separator_columns: int
    channel_width: int
    header_row_height: int
    unit_row_height: int
    plot_ids: tuple[str, ...]
    plot_columns: int
    plot_width_px: int
    plot_height_px: int
    output_filename: str
    keep_plot_assets: bool


@dataclass
class ExcelReportResult:
    report_path: Path
    manifest_path: Path
    summary_path: Path
    plot_assets_dir: Path
    config_path: Path
    config: ExcelReportConfig
    statistics_result: StatisticsResult
    plotting_result: PlottingResult
    report_channels: list[ChannelInfo]

    @property
    def sample_count(self) -> int:
        return self.statistics_result.sample_count

    @property
    def channel_count(self) -> int:
        return len(self.report_channels)

    @property
    def statistic_count(self) -> int:
        return self.statistics_result.statistic_count

    @property
    def plot_count(self) -> int:
        return self.plotting_result.plot_count


def load_excel_report_config(path: str | Path) -> ExcelReportConfig:
    config_path = Path(path).expanduser().resolve()
    if not config_path.exists():
        raise ConfigurationError(f"Excel-report configuration does not exist: {config_path}")
    if not config_path.is_file():
        raise ConfigurationError(f"Excel-report configuration is not a file: {config_path}")

    try:
        raw = yaml.safe_load(config_path.read_text(encoding="utf-8"))
    except UnicodeDecodeError as exc:
        raise ConfigurationError(f"Excel-report configuration must be UTF-8: {config_path}") from exc
    except yaml.YAMLError as exc:
        raise ConfigurationError(f"Invalid YAML in Excel-report configuration '{config_path}': {exc}") from exc

    if not isinstance(raw, dict):
        raise ConfigurationError("Excel-report configuration root must be a YAML mapping")
    _reject_unknown_keys(raw, {"version", "report", "channels", "statistics", "layout", "plots", "output"}, "root")

    version = raw.get("version")
    if version != 1:
        raise ConfigurationError("Excel-report configuration 'version' must be 1")

    report_raw = raw.get("report", {})
    if not isinstance(report_raw, dict):
        raise ConfigurationError("report must be a YAML mapping")
    _reject_unknown_keys(report_raw, {"title", "subtitle", "report_sheet", "metadata_sheet"}, "report")
    title = _nonempty_string(report_raw.get("title", "VSM Engineering Report"), "report.title")
    subtitle = _optional_string(report_raw.get("subtitle"), "report.subtitle")
    report_sheet = _sheet_name(report_raw.get("report_sheet", "Report"), "report.report_sheet")
    metadata_sheet = _sheet_name(report_raw.get("metadata_sheet", "Metadata"), "report.metadata_sheet")
    if report_sheet == metadata_sheet:
        raise ConfigurationError("report.report_sheet and report.metadata_sheet must be different")

    channels_raw = raw.get("channels")
    if not isinstance(channels_raw, list) or not channels_raw:
        raise ConfigurationError("channels must be a non-empty YAML list of channel_id values")
    channel_ids = tuple(_nonempty_string(value, "channels[]") for value in channels_raw)
    duplicates = _duplicates(channel_ids)
    if duplicates:
        raise ConfigurationError("channels contains duplicate channel IDs: " + ", ".join(duplicates))

    statistics_raw = raw.get("statistics", {})
    if not isinstance(statistics_raw, dict):
        raise ConfigurationError("statistics must be a YAML mapping")
    _reject_unknown_keys(
        statistics_raw,
        {"top_rms", "kpis", "bottom_operations", "bottom_summary"},
        "statistics",
    )
    top_rms = _string_list(statistics_raw.get("top_rms", []), "statistics.top_rms")
    kpis = _string_list(statistics_raw.get("kpis", []), "statistics.kpis")
    bottom_raw = statistics_raw.get("bottom_operations", list(_ALLOWED_BOTTOM_OPERATIONS))
    bottom_operations = _string_list(bottom_raw, "statistics.bottom_operations")
    invalid_ops = sorted(set(bottom_operations) - set(_ALLOWED_BOTTOM_OPERATIONS))
    if invalid_ops:
        raise ConfigurationError(
            "statistics.bottom_operations contains unsupported operations: " + ", ".join(invalid_ops)
        )
    if len(set(bottom_operations)) != len(bottom_operations):
        raise ConfigurationError("statistics.bottom_operations contains duplicates")
    bottom_summary = _string_list(statistics_raw.get("bottom_summary", []), "statistics.bottom_summary")

    layout_raw = raw.get("layout", {})
    if not isinstance(layout_raw, dict):
        raise ConfigurationError("layout must be a YAML mapping")
    _reject_unknown_keys(
        layout_raw,
        {
            "profile",
            "plot_placement",
            "blank_separator_columns",
            "channel_width",
            "header_row_height",
            "unit_row_height",
        },
        "layout",
    )
    layout_profile = _choice(
        layout_raw.get("profile", "engineering"),
        "layout.profile",
        {"engineering", "sergio_reference"},
    )
    default_plot_placement = "kpi_panel" if layout_profile == "sergio_reference" else "below_data"
    plot_placement = _choice(
        layout_raw.get("plot_placement", default_plot_placement),
        "layout.plot_placement",
        {"below_data", "kpi_panel"},
    )
    blank_separator_columns = _positive_int(
        layout_raw.get("blank_separator_columns", 1),
        "layout.blank_separator_columns",
        minimum=1,
        maximum=4,
    )
    channel_width = _positive_int(layout_raw.get("channel_width", 12), "layout.channel_width", minimum=7, maximum=30)
    header_row_height = _positive_int(
        layout_raw.get("header_row_height", 106 if layout_profile == "sergio_reference" else 44),
        "layout.header_row_height",
        minimum=24,
        maximum=180,
    )
    unit_row_height = _positive_int(
        layout_raw.get("unit_row_height", 16 if layout_profile == "sergio_reference" else 26),
        "layout.unit_row_height",
        minimum=12,
        maximum=60,
    )

    plots_raw = raw.get("plots", {})
    if not isinstance(plots_raw, dict):
        raise ConfigurationError("plots must be a YAML mapping")
    _reject_unknown_keys(plots_raw, {"include", "columns", "width_px", "height_px"}, "plots")
    plot_ids = _string_list(plots_raw.get("include", []), "plots.include")
    plot_columns = _positive_int(plots_raw.get("columns", 2), "plots.columns", maximum=4)
    plot_width_px = _positive_int(plots_raw.get("width_px", 720), "plots.width_px", minimum=240, maximum=1600)
    plot_height_px = _positive_int(plots_raw.get("height_px", 360), "plots.height_px", minimum=160, maximum=1000)

    output_raw = raw.get("output", {})
    if not isinstance(output_raw, dict):
        raise ConfigurationError("output must be a YAML mapping")
    _reject_unknown_keys(output_raw, {"filename", "keep_plot_assets"}, "output")
    output_filename = _plain_xlsx_filename(output_raw.get("filename", "vsm_engineering_report.xlsx"), "output.filename")
    keep_plot_assets = output_raw.get("keep_plot_assets", True)
    if not isinstance(keep_plot_assets, bool):
        raise ConfigurationError("output.keep_plot_assets must be true or false")

    return ExcelReportConfig(
        version=version,
        title=title,
        subtitle=subtitle,
        report_sheet=report_sheet,
        metadata_sheet=metadata_sheet,
        channel_ids=channel_ids,
        top_rms_statistic_ids=top_rms,
        kpi_statistic_ids=kpis,
        bottom_operations=bottom_operations,
        bottom_summary_statistic_ids=bottom_summary,
        layout_profile=layout_profile,
        plot_placement=plot_placement,
        blank_separator_columns=blank_separator_columns,
        channel_width=channel_width,
        header_row_height=header_row_height,
        unit_row_height=unit_row_height,
        plot_ids=plot_ids,
        plot_columns=plot_columns,
        plot_width_px=plot_width_px,
        plot_height_px=plot_height_px,
        output_filename=output_filename,
        keep_plot_assets=keep_plot_assets,
    )


def generate_excel_report(
    input_file: str | Path,
    report_config_file: str | Path,
    statistics_config_file: str | Path,
    plotting_config_file: str | Path,
    output_dir: str | Path,
    import_options: ImportOptions | None = None,
    math_config_file: str | Path | None = None,
    *,
    precomputed_statistics_result: StatisticsResult | None = None,
    precomputed_plotting_result: PlottingResult | None = None,
) -> ExcelReportResult:
    """Generate a deterministic Excel engineering report from the validated processing layers."""

    config_path = Path(report_config_file).expanduser().resolve()
    config = load_excel_report_config(config_path)
    destination = Path(output_dir).expanduser().resolve()
    destination.mkdir(parents=True, exist_ok=True)
    generated_plot_assets = precomputed_plotting_result is None
    plot_assets_dir = destination / "plot_assets"

    if precomputed_statistics_result is None:
        statistics_result = calculate_statistics(
            input_file,
            statistics_config_file,
            import_options,
            math_config_file=math_config_file,
        )
    else:
        statistics_result = precomputed_statistics_result

    if precomputed_plotting_result is None:
        plot_assets_dir.mkdir(parents=True, exist_ok=True)
        plotting_result = render_plots(
            input_file,
            plotting_config_file,
            plot_assets_dir,
            import_options,
            math_config_file=math_config_file,
        )
    else:
        plotting_result = precomputed_plotting_result
        if plotting_result.rendered_plots:
            plot_assets_dir = Path(plotting_result.rendered_plots[0].output_file).expanduser().resolve().parent
        else:
            plot_assets_dir.mkdir(parents=True, exist_ok=True)

    channels_by_id = statistics_result.channels_by_id
    values_by_id = statistics_result.values_by_id
    missing_channels = [channel_id for channel_id in config.channel_ids if channel_id not in channels_by_id]
    if missing_channels:
        raise ExcelReportError("Configured report channel IDs were not found: " + ", ".join(missing_channels))
    report_channels = [channels_by_id[channel_id] for channel_id in config.channel_ids]

    statistics_by_id = {item.statistic_id: item for item in statistics_result.statistics}
    missing_statistics = sorted(
        set(config.top_rms_statistic_ids + config.kpi_statistic_ids + config.bottom_summary_statistic_ids) - set(statistics_by_id)
    )
    if missing_statistics:
        raise ExcelReportError(
            "Configured report statistic IDs were not found in the statistics result: "
            + ", ".join(missing_statistics)
        )

    if config.layout_profile == "sergio_reference":
        invisible_rms = [
            statistic_id
            for statistic_id in config.top_rms_statistic_ids
            if statistics_by_id[statistic_id].channel_id not in config.channel_ids
        ]
        invisible_bottom = [
            statistic_id
            for statistic_id in config.bottom_summary_statistic_ids
            if statistics_by_id[statistic_id].channel_id not in config.channel_ids
        ]
        if invisible_rms:
            raise ExcelReportError(
                "Sergio-reference RMS statistics must target exported report channels: "
                + ", ".join(invisible_rms)
            )
        if invisible_bottom:
            raise ExcelReportError(
                "Sergio-reference bottom statistics must target exported report channels: "
                + ", ".join(invisible_bottom)
            )

    plots_by_id = {item.plot_id: item for item in plotting_result.rendered_plots}
    missing_plots = sorted(set(config.plot_ids) - set(plots_by_id))
    if missing_plots:
        raise ExcelReportError(
            "Configured report plot IDs were not found in the plotting result: " + ", ".join(missing_plots)
        )

    report_path = destination / config.output_filename
    manifest_path = destination / "excel_report_manifest.json"
    summary_path = destination / "excel_report_summary.txt"

    _write_workbook(
        report_path,
        config,
        input_file=Path(input_file).expanduser().resolve(),
        report_config_file=config_path,
        math_config_file=Path(math_config_file).expanduser().resolve() if math_config_file else None,
        statistics_config_file=Path(statistics_config_file).expanduser().resolve(),
        plotting_config_file=Path(plotting_config_file).expanduser().resolve(),
        statistics_result=statistics_result,
        report_channels=report_channels,
        values_by_id=values_by_id,
        statistics_by_id=statistics_by_id,
        plots_by_id=plots_by_id,
    )

    result = ExcelReportResult(
        report_path=report_path,
        manifest_path=manifest_path,
        summary_path=summary_path,
        plot_assets_dir=plot_assets_dir,
        config_path=config_path,
        config=config,
        statistics_result=statistics_result,
        plotting_result=plotting_result,
        report_channels=report_channels,
    )
    manifest_path.write_text(json.dumps(_manifest(result, statistics_config_file, plotting_config_file, math_config_file), indent=2), encoding="utf-8")
    summary_path.write_text(_summary(result), encoding="utf-8")

    if not config.keep_plot_assets and generated_plot_assets:
        shutil.rmtree(plot_assets_dir, ignore_errors=True)

    return result


def _write_workbook(
    output_path: Path,
    config: ExcelReportConfig,
    *,
    input_file: Path,
    report_config_file: Path,
    math_config_file: Path | None,
    statistics_config_file: Path,
    plotting_config_file: Path,
    statistics_result: StatisticsResult,
    report_channels: list[ChannelInfo],
    values_by_id: Mapping[str, Any],
    statistics_by_id: Mapping[str, StatisticResult],
    plots_by_id: Mapping[str, Any],
) -> None:
    if config.layout_profile == "sergio_reference":
        _write_sergio_reference_workbook(
            output_path,
            config,
            input_file=input_file,
            report_config_file=report_config_file,
            math_config_file=math_config_file,
            statistics_config_file=statistics_config_file,
            plotting_config_file=plotting_config_file,
            statistics_result=statistics_result,
            report_channels=report_channels,
            values_by_id=values_by_id,
            statistics_by_id=statistics_by_id,
            plots_by_id=plots_by_id,
        )
        return

    workbook = Workbook()
    report = workbook.active
    report.title = config.report_sheet
    metadata = workbook.create_sheet(config.metadata_sheet)

    thin = Side(style="thin", color="D9E2F3")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    raw_fill = PatternFill("solid", fgColor="1F4E78")
    math_fill = PatternFill("solid", fgColor="C65911")
    title_fill = PatternFill("solid", fgColor="17365D")
    rms_fill = PatternFill("solid", fgColor="7030A0")
    kpi_fill = PatternFill("solid", fgColor="548235")
    stat_fill = PatternFill("solid", fgColor="D9EAF7")
    metadata_fill = PatternFill("solid", fgColor="5B9BD5")
    white_bold = Font(color="FFFFFF", bold=True)

    channel_count = len(report_channels)
    last_channel_col = get_column_letter(channel_count)
    report.merge_cells(start_row=1, start_column=1, end_row=1, end_column=min(4, channel_count))
    report.cell(1, 1, config.title)
    report.cell(1, 1).fill = title_fill
    report.cell(1, 1).font = Font(color="FFFFFF", bold=True, size=14)
    report.cell(1, 1).alignment = Alignment(horizontal="left", vertical="center")
    report.merge_cells(start_row=2, start_column=1, end_row=2, end_column=min(4, channel_count))
    report.cell(2, 1, config.subtitle or input_file.name)
    report.cell(2, 1).font = Font(italic=True, color="666666")

    # Top RMS strip starts after the title block and may continue beyond data columns if needed.
    rms_start_col = 5
    for index, statistic_id in enumerate(config.top_rms_statistic_ids):
        item = statistics_by_id[statistic_id]
        col = rms_start_col + index * 2
        report.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + 1)
        report.merge_cells(start_row=2, start_column=col, end_row=2, end_column=col + 1)
        report.cell(1, col, item.display_name)
        report.cell(2, col, item.value)
        for row in (1, 2):
            cell = report.cell(row, col)
            cell.fill = rms_fill
            cell.font = white_bold
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border
        report.cell(2, col).number_format = "0.000000"

    # Two-row channel header: name, then explicit type + unit so math/raw is not color-only.
    for col_index, channel in enumerate(report_channels, start=1):
        name_cell = report.cell(3, col_index, channel.display_name)
        meta_cell = report.cell(4, col_index, f"{channel.kind.upper()} | {channel.unit or '-'}")
        fill = math_fill if channel.kind == "math" else raw_fill
        for cell in (name_cell, meta_cell):
            cell.fill = fill
            cell.font = white_bold
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    data_start_row = 5
    sample_count = statistics_result.sample_count
    for col_index, channel in enumerate(report_channels, start=1):
        values = values_by_id[channel.channel_id]
        for row_offset, value in enumerate(values, start=0):
            cell = report.cell(data_start_row + row_offset, col_index, float(value))
            cell.number_format = _number_format(channel.unit)

    data_end_row = data_start_row + sample_count - 1
    bottom_start_row = data_end_row + 2
    bottom_statistics = [
        item
        for item in statistics_result.statistics
        if item.operation in config.bottom_operations and item.channel_id in config.channel_ids
    ]
    stats_by_operation_channel = {(item.operation, item.channel_id): item for item in bottom_statistics}
    for offset, operation in enumerate(config.bottom_operations):
        row = bottom_start_row + offset
        label = report.cell(row, 1, operation.replace("_", " ").upper())
        label.fill = stat_fill
        label.font = Font(bold=True)
        label.border = border
        for col_index, channel in enumerate(report_channels, start=1):
            item = stats_by_operation_channel.get((operation, channel.channel_id))
            if item is None:
                continue
            cell = report.cell(row, col_index, item.value)
            cell.fill = stat_fill
            cell.font = Font(bold=True)
            cell.border = border
            cell.number_format = _number_format(item.channel_unit)

    # KPI strip is placed to the right of the selected data channels, like the reference workbook.
    kpi_start_col = channel_count + 2
    for index, statistic_id in enumerate(config.kpi_statistic_ids):
        item = statistics_by_id[statistic_id]
        col = kpi_start_col + index
        header = report.cell(3, col, item.display_name)
        value = report.cell(4, col, item.value)
        for cell in (header, value):
            cell.fill = kpi_fill
            cell.font = white_bold
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        value.number_format = _number_format(item.channel_unit)
        report.column_dimensions[get_column_letter(col)].width = 20

    # Plot images are embedded below statistics in a deterministic grid.
    chart_start_row = bottom_start_row + len(config.bottom_operations) + 3
    horizontal_span = max(8, math.ceil(channel_count / config.plot_columns))
    vertical_span = 20
    for plot_index, plot_id in enumerate(config.plot_ids):
        item = plots_by_id[plot_id]
        grid_row = plot_index // config.plot_columns
        grid_col = plot_index % config.plot_columns
        anchor_row = chart_start_row + grid_row * vertical_span
        anchor_col = 1 + grid_col * horizontal_span
        image = XLImage(item.output_file)
        image.width = config.plot_width_px
        image.height = config.plot_height_px
        report.add_image(image, f"{get_column_letter(anchor_col)}{anchor_row}")

    report.freeze_panes = "B5"
    report.row_dimensions[1].height = 28
    report.row_dimensions[2].height = 24
    report.row_dimensions[3].height = 44
    report.row_dimensions[4].height = 26
    for col_index in range(1, channel_count + 1):
        report.column_dimensions[get_column_letter(col_index)].width = 17

    _write_metadata_sheet(
        metadata,
        input_file=input_file,
        report_config_file=report_config_file,
        config=config,
        math_config_file=math_config_file,
        statistics_config_file=statistics_config_file,
        plotting_config_file=plotting_config_file,
        statistics_result=statistics_result,
        report_channels=report_channels,
        statistics_by_id=statistics_by_id,
        plots_by_id=plots_by_id,
        header_fill=metadata_fill,
        white_bold=white_bold,
        border=border,
    )

    workbook.save(output_path)


def _write_sergio_reference_workbook(
    output_path: Path,
    config: ExcelReportConfig,
    *,
    input_file: Path,
    report_config_file: Path,
    math_config_file: Path | None,
    statistics_config_file: Path,
    plotting_config_file: Path,
    statistics_result: StatisticsResult,
    report_channels: list[ChannelInfo],
    values_by_id: Mapping[str, Any],
    statistics_by_id: Mapping[str, StatisticResult],
    plots_by_id: Mapping[str, Any],
) -> None:
    """Write the compact client-style layout reverse-engineered from Sergio's workbook.

    The visible report intentionally reserves rows 1-2 for RMS, rows 3-4 for
    channel/KPI headers, starts samples on row 5, writes configured per-channel
    summary values immediately after the final sample, and places plot assets
    under the KPI strip. Engineering calculations remain external to Excel.
    """

    workbook = Workbook()
    report = workbook.active
    report.title = config.report_sheet
    metadata = workbook.create_sheet(config.metadata_sheet)

    # Reference-inspired palette: raw channels stay white; calculated channels
    # use the same pale-yellow visual cue found in the client workbook.
    border_side = Side(style="thin", color="A6A6A6")
    border = Border(left=border_side, right=border_side, top=border_side, bottom=border_side)
    raw_fill = PatternFill("solid", fgColor="FFFFFF")
    math_fill = PatternFill("solid", fgColor="FFF2CC")
    rms_fill = PatternFill("solid", fgColor="5B9BD5")
    metadata_fill = PatternFill("solid", fgColor="5B9BD5")
    white_bold = Font(name="Calibri", size=11, color="FFFFFF", bold=True)
    header_font = Font(name="Calibri", size=11, color="000000", bold=False)

    channel_count = len(report_channels)
    channel_column_by_id = {
        channel.channel_id: index for index, channel in enumerate(report_channels, start=1)
    }

    # Rows 1 and 2: RMS labels and values directly above the relevant channel.
    for statistic_id in config.top_rms_statistic_ids:
        item = statistics_by_id[statistic_id]
        col = channel_column_by_id[item.channel_id]
        label = report.cell(1, col, item.display_name)
        value = report.cell(2, col, item.value)
        for cell in (label, value):
            cell.fill = rms_fill
            cell.font = white_bold
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        value.number_format = _number_format(item.channel_unit)

        # Mirror the client helper-column convention when a squared helper is
        # immediately adjacent to the RMS source channel.
        if col < channel_count:
            helper = report_channels[col]
            if "squared" in helper.channel_id and item.channel_id in helper.dependencies:
                report.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + 1)

    # Rows 3 and 4: selected channel name and unit, with math channels highlighted.
    for col_index, channel in enumerate(report_channels, start=1):
        fill = math_fill if channel.kind == "math" else raw_fill
        name_cell = report.cell(3, col_index, channel.display_name)
        unit_cell = report.cell(4, col_index, channel.unit or "")
        for cell in (name_cell, unit_cell):
            cell.fill = fill
            cell.font = header_font
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Data begins on row 5, exactly like the reference workbook.
    data_start_row = 5
    sample_count = statistics_result.sample_count
    for col_index, channel in enumerate(report_channels, start=1):
        values = values_by_id[channel.channel_id]
        for row_offset, value in enumerate(values):
            cell = report.cell(data_start_row + row_offset, col_index, float(value))
            cell.number_format = _number_format(channel.unit)

    data_end_row = data_start_row + sample_count - 1

    # Per-channel bottom summary. First configured statistic for a channel goes
    # directly below the data; a second statistic (e.g. battery MIN) uses the
    # following row. This reproduces the MAX/MIN/last intent without a mixed
    # operation label row.
    bottom_offsets: dict[str, int] = {}
    for statistic_id in config.bottom_summary_statistic_ids:
        item = statistics_by_id[statistic_id]
        col = channel_column_by_id[item.channel_id]
        offset = bottom_offsets.get(item.channel_id, 0)
        row = data_end_row + 1 + offset
        bottom_offsets[item.channel_id] = offset + 1
        cell = report.cell(row, col, item.value)
        cell.number_format = _number_format(item.channel_unit)
        cell.font = Font(name="Calibri", size=11, bold=False)
        cell.border = Border(top=border_side)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # One blank separator column followed by the secondary KPI selection in rows 3-4.
    kpi_start_col = channel_count + config.blank_separator_columns + 1
    for index, statistic_id in enumerate(config.kpi_statistic_ids):
        item = statistics_by_id[statistic_id]
        col = kpi_start_col + index
        header = report.cell(3, col, item.display_name)
        value = report.cell(4, col, item.value)
        for cell in (header, value):
            cell.fill = raw_fill
            cell.font = header_font
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        value.number_format = _number_format(item.channel_unit)
        report.column_dimensions[get_column_letter(col)].width = max(12, config.channel_width)

    # Plots sit directly below the KPI/secondary-selection strip, as in the
    # supplied report. The assets remain deterministic PNGs from the plotting engine.
    if config.plot_placement == "kpi_panel":
        chart_start_row = 6
        chart_start_col = kpi_start_col
    else:
        max_bottom_depth = max(bottom_offsets.values(), default=0)
        chart_start_row = data_end_row + max_bottom_depth + 3
        chart_start_col = 1

    approx_col_px = 64
    approx_row_px = 20
    horizontal_span = max(8, math.ceil(config.plot_width_px / approx_col_px) + 1)
    vertical_span = max(16, math.ceil(config.plot_height_px / approx_row_px) + 2)
    for plot_index, plot_id in enumerate(config.plot_ids):
        item = plots_by_id[plot_id]
        grid_row = plot_index // config.plot_columns
        grid_col = plot_index % config.plot_columns
        anchor_row = chart_start_row + grid_row * vertical_span
        anchor_col = chart_start_col + grid_col * horizontal_span
        image = XLImage(item.output_file)
        image.width = config.plot_width_px
        image.height = config.plot_height_px
        report.add_image(image, f"{get_column_letter(anchor_col)}{anchor_row}")

    # Match the visible mechanics of the supplied workbook.
    report.freeze_panes = "B5"
    report.row_dimensions[2].height = 16
    report.row_dimensions[3].height = config.header_row_height
    report.row_dimensions[4].height = config.unit_row_height
    for col_index in range(1, channel_count + 1):
        report.column_dimensions[get_column_letter(col_index)].width = config.channel_width
    for gap in range(channel_count + 1, kpi_start_col):
        report.column_dimensions[get_column_letter(gap)].width = 2.5

    # Give the plot/KPI panel stable widths so embedded images do not crowd the data area.
    if config.plot_ids:
        panel_end = chart_start_col + config.plot_columns * horizontal_span
        for col_index in range(chart_start_col, panel_end + 1):
            letter = get_column_letter(col_index)
            if report.column_dimensions[letter].width is None:
                report.column_dimensions[letter].width = 9.0

    _write_metadata_sheet(
        metadata,
        input_file=input_file,
        report_config_file=report_config_file,
        config=config,
        math_config_file=math_config_file,
        statistics_config_file=statistics_config_file,
        plotting_config_file=plotting_config_file,
        statistics_result=statistics_result,
        report_channels=report_channels,
        statistics_by_id=statistics_by_id,
        plots_by_id=plots_by_id,
        header_fill=metadata_fill,
        white_bold=white_bold,
        border=border,
    )

    workbook.save(output_path)


def _write_metadata_sheet(
    sheet: Any,
    *,
    input_file: Path,
    report_config_file: Path,
    config: ExcelReportConfig,
    math_config_file: Path | None,
    statistics_config_file: Path,
    plotting_config_file: Path,
    statistics_result: StatisticsResult,
    report_channels: list[ChannelInfo],
    statistics_by_id: Mapping[str, StatisticResult],
    plots_by_id: Mapping[str, Any],
    header_fill: PatternFill,
    white_bold: Font,
    border: Border,
) -> None:
    sheet.freeze_panes = "A2"
    sheet.append(["VSM REPORT METADATA", "Value"])
    source_hash = statistics_result.dataset.quality.source_sha256
    metadata_rows = [
        ("Source file", client_display_filename(input_file)),
        ("Source SHA-256", source_hash),
        ("Samples", statistics_result.sample_count),
        ("Time channel", statistics_result.dataset.quality.time_channel_id or ""),
        ("Time start", statistics_result.dataset.quality.time_start),
        ("Time end", statistics_result.dataset.quality.time_end),
        ("Nominal time step", statistics_result.dataset.quality.nominal_time_step),
        ("Report configuration", Path(report_config_file).name),
        ("Math configuration", Path(math_config_file).name if math_config_file else ""),
        ("Statistics configuration", Path(statistics_config_file).name),
        ("Plotting configuration", Path(plotting_config_file).name),
    ]
    for row in metadata_rows:
        sheet.append(list(row))

    row = sheet.max_row + 2
    sheet.cell(row, 1, "REPORT CHANNELS")
    headers = ["Order", "channel_id", "display_name", "unit", "kind", "provenance", "dependencies"]
    for col, value in enumerate(headers, start=1):
        sheet.cell(row + 1, col, value)
    for index, channel in enumerate(report_channels, start=1):
        values = [
            index,
            channel.channel_id,
            channel.display_name,
            channel.unit or "",
            channel.kind,
            channel.provenance.replace(input_file.name, client_display_filename(input_file)),
            ";".join(channel.dependencies),
        ]
        for col, value in enumerate(values, start=1):
            sheet.cell(row + 1 + index, col, value)

    row = sheet.max_row + 2
    sheet.cell(row, 1, "STATISTICS")
    stat_headers = ["statistic_id", "channel_id", "operation", "placement_group", "value", "unit", "display_name"]
    for col, value in enumerate(stat_headers, start=1):
        sheet.cell(row + 1, col, value)
    for index, item in enumerate(statistics_by_id.values(), start=1):
        values = [item.statistic_id, item.channel_id, item.operation, item.placement_group, item.value, item.channel_unit or "", item.display_name]
        for col, value in enumerate(values, start=1):
            sheet.cell(row + 1 + index, col, value)

    row = sheet.max_row + 2
    sheet.cell(row, 1, "PLOTS")
    plot_headers = ["plot_id", "title", "x_channel_id", "primary_series", "secondary_series", "output_file"]
    for col, value in enumerate(plot_headers, start=1):
        sheet.cell(row + 1, col, value)
    for index, item in enumerate(plots_by_id.values(), start=1):
        values = [
            item.plot_id,
            item.title,
            item.x_channel_id,
            ";".join(item.primary_series_ids),
            ";".join(item.secondary_series_ids),
            Path(item.output_file).name,
        ]
        for col, value in enumerate(values, start=1):
            sheet.cell(row + 1 + index, col, value)

    # Style section labels and table headers.
    for row_index in range(1, sheet.max_row + 1):
        first = sheet.cell(row_index, 1).value
        if first in {"VSM REPORT METADATA", "REPORT CHANNELS", "STATISTICS", "PLOTS"}:
            for col in range(1, 8):
                cell = sheet.cell(row_index, col)
                cell.fill = header_fill
                cell.font = white_bold
                cell.border = border
        if first in {"Order", "statistic_id", "plot_id"}:
            for col in range(1, 8):
                cell = sheet.cell(row_index, col)
                cell.fill = header_fill
                cell.font = white_bold
                cell.border = border

    for col, width in {"A": 28, "B": 42, "C": 30, "D": 18, "E": 18, "F": 45, "G": 45}.items():
        sheet.column_dimensions[col].width = width
    for row in sheet.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)


def _manifest(
    result: ExcelReportResult,
    statistics_config_file: str | Path,
    plotting_config_file: str | Path,
    math_config_file: str | Path | None,
) -> dict[str, Any]:
    payload = {
        "configuration_version": result.config.version,
        "software_version": __version__,
        "source_file": str(result.statistics_result.dataset.source_path),
        "source_sha256": result.statistics_result.dataset.quality.source_sha256,
        "report_configuration_file": str(result.config_path),
        "report_configuration_sha256": sha256_file(result.config_path),
        "statistics_configuration_file": str(Path(statistics_config_file).expanduser().resolve()),
        "statistics_configuration_sha256": sha256_file(Path(statistics_config_file).expanduser().resolve()),
        "plotting_configuration_file": str(Path(plotting_config_file).expanduser().resolve()),
        "plotting_configuration_sha256": sha256_file(Path(plotting_config_file).expanduser().resolve()),
        "sample_count": result.sample_count,
        "report_channel_count": result.channel_count,
        "statistic_count": result.statistic_count,
        "plot_count": result.plot_count,
        "report_channel_ids": [channel.channel_id for channel in result.report_channels],
        "top_rms_statistic_ids": list(result.config.top_rms_statistic_ids),
        "kpi_statistic_ids": list(result.config.kpi_statistic_ids),
        "bottom_operations": list(result.config.bottom_operations),
        "bottom_summary_statistic_ids": list(result.config.bottom_summary_statistic_ids),
        "layout_profile": result.config.layout_profile,
        "plot_placement": result.config.plot_placement,
        "plot_ids": list(result.config.plot_ids),
        "report_file": str(result.report_path),
    }
    if math_config_file is not None:
        path = Path(math_config_file).expanduser().resolve()
        payload["math_configuration_file"] = str(path)
        payload["math_configuration_sha256"] = sha256_file(path)
    else:
        payload["math_configuration_file"] = None
        payload["math_configuration_sha256"] = None
    return payload


def _summary(result: ExcelReportResult) -> str:
    lines = [
        "VSM EXCEL REPORT SUMMARY",
        "========================",
        f"Status: PASS",
        f"Source: {result.statistics_result.dataset.source_path}",
        f"Samples: {result.sample_count}",
        f"Report channels: {result.channel_count}",
        f"Statistics available: {result.statistic_count}",
        f"Plots embedded: {result.plot_count}",
        f"Layout profile: {result.config.layout_profile}",
        f"Plot placement: {result.config.plot_placement}",
        f"Workbook: {result.report_path}",
        "",
        "Top RMS:",
    ]
    stats_by_id = {item.statistic_id: item for item in result.statistics_result.statistics}
    for statistic_id in result.config.top_rms_statistic_ids:
        item = stats_by_id[statistic_id]
        lines.append(f"  - {item.display_name}: {item.value:.12g} [{item.channel_unit or '-'}]")
    lines.append("")
    lines.append("KPI strip:")
    for statistic_id in result.config.kpi_statistic_ids:
        item = stats_by_id[statistic_id]
        lines.append(f"  - {item.display_name}: {item.value:.12g} [{item.channel_unit or '-'}]")
    return "\n".join(lines) + "\n"


def _number_format(unit: str | None) -> str:
    if unit is None:
        return "0.000000"
    normalized = unit.strip().lower()
    if normalized in {"%", "percent"}:
        return "0.0000"
    if normalized in {"s", "min", "m", "km", "kph", "rpm", "nm", "kw", "kwh", "kg"}:
        return "0.000000"
    return "0.000000"


def _plain_xlsx_filename(value: object, context: str) -> str:
    filename = _nonempty_string(value, context)
    path = Path(filename)
    if path.name != filename or path.suffix.lower() != ".xlsx":
        raise ConfigurationError(f"{context} must be a plain .xlsx filename without directories")
    return filename


def _sheet_name(value: object, context: str) -> str:
    name = _nonempty_string(value, context)
    if len(name) > 31 or any(character in name for character in "[]:*?/\\"):
        raise ConfigurationError(f"{context} is not a valid Excel sheet name")
    return name


def _string_list(value: object, context: str) -> tuple[str, ...]:
    if not isinstance(value, list):
        raise ConfigurationError(f"{context} must be a YAML list")
    items = tuple(_nonempty_string(item, f"{context}[]") for item in value)
    duplicates = _duplicates(items)
    if duplicates:
        raise ConfigurationError(f"{context} contains duplicates: " + ", ".join(duplicates))
    return items


def _nonempty_string(value: object, context: str) -> str:
    if not isinstance(value, str) or not value.strip():
        raise ConfigurationError(f"{context} must be a non-empty string")
    return value.strip()


def _optional_string(value: object, context: str) -> str | None:
    if value is None:
        return None
    return _nonempty_string(value, context)


def _choice(value: object, context: str, allowed: set[str]) -> str:
    choice = _nonempty_string(value, context)
    if choice not in allowed:
        raise ConfigurationError(f"{context} must be one of: " + ", ".join(sorted(allowed)))
    return choice


def _positive_int(value: object, context: str, minimum: int = 1, maximum: int | None = None) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < minimum:
        raise ConfigurationError(f"{context} must be an integer >= {minimum}")
    if maximum is not None and value > maximum:
        raise ConfigurationError(f"{context} must be <= {maximum}")
    return value


def _reject_unknown_keys(mapping: Mapping[str, object], allowed: set[str], context: str) -> None:
    unknown = sorted(set(mapping) - allowed)
    if unknown:
        raise ConfigurationError(f"Unknown key(s) in {context}: " + ", ".join(unknown))


def _duplicates(values: tuple[str, ...]) -> list[str]:
    seen: set[str] = set()
    duplicates: set[str] = set()
    for value in values:
        if value in seen:
            duplicates.add(value)
        seen.add(value)
    return sorted(duplicates)
