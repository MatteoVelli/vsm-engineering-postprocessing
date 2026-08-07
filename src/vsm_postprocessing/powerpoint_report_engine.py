from __future__ import annotations

import json
import math
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable

import yaml
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt

from .errors import ConfigurationError, PowerPointReportError
from .importer import ImportOptions
from .plotting_engine import PlottingResult, render_plots
from .statistics_engine import StatisticResult, StatisticsResult, calculate_statistics
from .utils import sha256_file


@dataclass(frozen=True)
class PowerPointSlideDefinition:
    slide_id: str
    slide_type: str
    title: str
    subtitle: str | None
    statistic_ids: tuple[str, ...]
    plot_ids: tuple[str, ...]
    footer: str | None


@dataclass(frozen=True)
class PowerPointTheme:
    title_font_size: float = 28.0
    subtitle_font_size: float = 14.0
    kpi_label_font_size: float = 10.0
    kpi_value_font_size: float = 19.0
    body_font_size: float = 11.0
    accent_fill: str = "FFF2CC"
    accent_border: str = "D6B656"
    text_color: str = "111111"
    muted_text_color: str = "666666"


@dataclass(frozen=True)
class PowerPointReportConfig:
    version: int
    title: str
    subtitle: str | None
    output_filename: str
    footer: str | None
    theme: PowerPointTheme
    slides: tuple[PowerPointSlideDefinition, ...]


@dataclass
class PowerPointReportResult:
    presentation_path: Path
    manifest_path: Path
    summary_path: Path
    plot_assets_dir: Path
    config_path: Path
    config: PowerPointReportConfig
    statistics_result: StatisticsResult
    plotting_result: PlottingResult

    @property
    def sample_count(self) -> int:
        return self.statistics_result.sample_count

    @property
    def slide_count(self) -> int:
        return len(self.config.slides)

    @property
    def plot_count(self) -> int:
        return len({pid for slide in self.config.slides for pid in slide.plot_ids})

    @property
    def statistic_count(self) -> int:
        return len({sid for slide in self.config.slides for sid in slide.statistic_ids})


_ALLOWED_SLIDE_TYPES = {"summary", "plot_pair"}


def load_powerpoint_report_config(path: str | Path) -> PowerPointReportConfig:
    config_path = Path(path).expanduser().resolve()
    if not config_path.exists():
        raise ConfigurationError(f"PowerPoint-report configuration does not exist: {config_path}")
    if not config_path.is_file():
        raise ConfigurationError(f"PowerPoint-report configuration is not a file: {config_path}")
    try:
        raw = yaml.safe_load(config_path.read_text(encoding="utf-8"))
    except UnicodeDecodeError as exc:
        raise ConfigurationError(f"PowerPoint-report configuration must be UTF-8: {config_path}") from exc
    except yaml.YAMLError as exc:
        raise ConfigurationError(f"Invalid YAML in PowerPoint-report configuration '{config_path}': {exc}") from exc
    if not isinstance(raw, dict):
        raise ConfigurationError("PowerPoint-report configuration root must be a YAML mapping")
    _reject_unknown_keys(raw, {"version", "presentation", "theme", "slides"}, "root")
    if raw.get("version") != 1:
        raise ConfigurationError("PowerPoint-report configuration 'version' must be 1")

    presentation = raw.get("presentation", {})
    if not isinstance(presentation, dict):
        raise ConfigurationError("presentation must be a YAML mapping")
    _reject_unknown_keys(presentation, {"title", "subtitle", "output_filename", "footer"}, "presentation")
    title = _nonempty_string(presentation.get("title", "VSM Engineering Report"), "presentation.title")
    subtitle = _optional_string(presentation.get("subtitle"), "presentation.subtitle")
    footer = _optional_string(presentation.get("footer"), "presentation.footer")
    output_filename = _pptx_filename(presentation.get("output_filename", "vsm_engineering_report.pptx"))

    theme = _load_theme(raw.get("theme", {}))
    slides_raw = raw.get("slides")
    if not isinstance(slides_raw, list) or not slides_raw:
        raise ConfigurationError("slides must be a non-empty YAML list")
    slides = tuple(_load_slide(item, index) for index, item in enumerate(slides_raw, start=1))
    duplicate_ids = _duplicates(slide.slide_id for slide in slides)
    if duplicate_ids:
        raise ConfigurationError("slides contains duplicate slide IDs: " + ", ".join(duplicate_ids))
    return PowerPointReportConfig(
        version=1,
        title=title,
        subtitle=subtitle,
        output_filename=output_filename,
        footer=footer,
        theme=theme,
        slides=slides,
    )


def generate_powerpoint_report(
    input_file: str | Path,
    report_config_file: str | Path,
    statistics_config_file: str | Path,
    plotting_config_file: str | Path,
    output_dir: str | Path,
    import_options: ImportOptions | None = None,
    math_config_file: str | Path | None = None,
) -> PowerPointReportResult:
    destination = Path(output_dir).expanduser().resolve()
    destination.mkdir(parents=True, exist_ok=True)
    plot_assets_dir = destination / "plot_assets"
    plot_assets_dir.mkdir(parents=True, exist_ok=True)
    statistics_result = calculate_statistics(
        input_file,
        statistics_config_file,
        import_options,
        math_config_file=math_config_file,
    )
    plotting_result = render_plots(
        input_file,
        plotting_config_file,
        plot_assets_dir,
        import_options,
        math_config_file=math_config_file,
    )
    return build_powerpoint_report(
        statistics_result,
        plotting_result,
        report_config_file,
        destination,
        plot_assets_dir=plot_assets_dir,
    )


def build_powerpoint_report(
    statistics_result: StatisticsResult,
    plotting_result: PlottingResult,
    report_config_file: str | Path,
    output_dir: str | Path,
    *,
    plot_assets_dir: str | Path,
) -> PowerPointReportResult:
    config_path = Path(report_config_file).expanduser().resolve()
    config = load_powerpoint_report_config(config_path)
    destination = Path(output_dir).expanduser().resolve()
    destination.mkdir(parents=True, exist_ok=True)
    assets = Path(plot_assets_dir).expanduser().resolve()

    statistics_by_id = {item.statistic_id: item for item in statistics_result.statistics}
    plots_by_id = {item.plot_id: item for item in plotting_result.rendered_plots}
    required_statistics = {sid for slide in config.slides for sid in slide.statistic_ids}
    required_plots = {pid for slide in config.slides for pid in slide.plot_ids}
    missing_statistics = sorted(required_statistics - set(statistics_by_id))
    missing_plots = sorted(required_plots - set(plots_by_id))
    if missing_statistics:
        raise PowerPointReportError("Configured PowerPoint statistic IDs were not found: " + ", ".join(missing_statistics))
    if missing_plots:
        raise PowerPointReportError("Configured PowerPoint plot IDs were not found: " + ", ".join(missing_plots))

    for plot_id in required_plots:
        path = _plot_path(plots_by_id[plot_id].output_file, assets)
        if not path.exists():
            raise PowerPointReportError(f"Plot asset for '{plot_id}' does not exist: {path}")

    prs = Presentation()
    prs.slide_width = Inches(13.333333)
    prs.slide_height = Inches(7.5)
    prs.core_properties.title = config.title
    prs.core_properties.subject = config.subtitle or "Deterministic VSM engineering report"
    prs.core_properties.author = "VSM Post-Processing Tool"

    for slide_number, slide_def in enumerate(config.slides, start=1):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        _add_slide_title(slide, slide_def.title, slide_def.subtitle, config.theme)
        if slide_def.slide_type == "summary":
            _render_summary_slide(slide, slide_def, statistics_by_id, statistics_result, config.theme)
        elif slide_def.slide_type == "plot_pair":
            _render_plot_pair_slide(slide, slide_def, statistics_by_id, plots_by_id, assets, config.theme)
        else:  # defensive: loader already validates
            raise PowerPointReportError(f"Unsupported PowerPoint slide type: {slide_def.slide_type}")
        _add_footer(slide, slide_def.footer or config.footer, slide_number, config.theme)

    presentation_path = destination / config.output_filename
    try:
        prs.save(presentation_path)
    except Exception as exc:  # python-pptx exposes several low-level exceptions
        raise PowerPointReportError(f"Could not save PowerPoint report '{presentation_path}': {exc}") from exc

    manifest_path = destination / "powerpoint_report_manifest.json"
    summary_path = destination / "powerpoint_report_summary.txt"
    result = PowerPointReportResult(
        presentation_path=presentation_path,
        manifest_path=manifest_path,
        summary_path=summary_path,
        plot_assets_dir=assets,
        config_path=config_path,
        config=config,
        statistics_result=statistics_result,
        plotting_result=plotting_result,
    )
    _write_metadata(result)
    return result


def _render_summary_slide(
    slide: Any,
    definition: PowerPointSlideDefinition,
    statistics_by_id: dict[str, StatisticResult],
    statistics_result: StatisticsResult,
    theme: PowerPointTheme,
) -> None:
    source_name = statistics_result.dataset.source_path.name
    source_box = slide.shapes.add_textbox(Inches(0.7), Inches(1.55), Inches(11.9), Inches(0.42))
    p = source_box.text_frame.paragraphs[0]
    p.text = f"Source: {source_name}  |  Samples: {statistics_result.sample_count}"
    p.font.size = Pt(theme.body_font_size)
    p.font.color.rgb = _rgb(theme.muted_text_color)
    p.alignment = PP_ALIGN.CENTER
    stats = [statistics_by_id[sid] for sid in definition.statistic_ids]
    _add_kpi_grid(slide, stats, y=2.15, theme=theme, max_columns=3, card_height=1.45)


def _render_plot_pair_slide(
    slide: Any,
    definition: PowerPointSlideDefinition,
    statistics_by_id: dict[str, StatisticResult],
    plots_by_id: dict[str, Any],
    assets: Path,
    theme: PowerPointTheme,
) -> None:
    stats = [statistics_by_id[sid] for sid in definition.statistic_ids]
    if stats:
        _add_kpi_strip(slide, stats, y=1.45, theme=theme)
    plot_y = 2.45 if stats else 1.7
    plot_h = 3.95 if stats else 4.7
    plot_ids = list(definition.plot_ids)
    if len(plot_ids) == 1:
        path = _plot_path(plots_by_id[plot_ids[0]].output_file, assets)
        _add_contained_picture(slide, path, 1.35, plot_y, 10.65, plot_h)
    elif len(plot_ids) == 2:
        positions = [(0.45, plot_y, 6.1, plot_h), (6.78, plot_y, 6.1, plot_h)]
        for plot_id, (x, y, w, h) in zip(plot_ids, positions):
            path = _plot_path(plots_by_id[plot_id].output_file, assets)
            _add_contained_picture(slide, path, x, y, w, h)


def _add_slide_title(slide: Any, title: str, subtitle: str | None, theme: PowerPointTheme) -> None:
    title_box = slide.shapes.add_textbox(Inches(0.35), Inches(0.26), Inches(12.63), Inches(0.72))
    tf = title_box.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.text = title
    p.alignment = PP_ALIGN.CENTER
    p.font.size = Pt(theme.title_font_size)
    p.font.bold = False
    p.font.color.rgb = _rgb(theme.text_color)
    if subtitle:
        subtitle_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.93), Inches(12.33), Inches(0.34))
        sp = subtitle_box.text_frame.paragraphs[0]
        sp.text = subtitle
        sp.alignment = PP_ALIGN.CENTER
        sp.font.size = Pt(theme.subtitle_font_size)
        sp.font.color.rgb = _rgb(theme.muted_text_color)
    accent = slide.shapes.add_shape(1, Inches(4.55), Inches(1.29), Inches(4.23), Inches(0.045))
    accent.fill.solid()
    accent.fill.fore_color.rgb = _rgb(theme.accent_border)
    accent.line.fill.background()


def _add_kpi_strip(slide: Any, statistics: list[StatisticResult], *, y: float, theme: PowerPointTheme) -> None:
    if not statistics:
        return
    max_items = min(len(statistics), 6)
    stats = statistics[:max_items]
    total_w = 12.1
    gap = 0.08
    card_w = (total_w - gap * (len(stats) - 1)) / len(stats)
    x = 0.62
    for stat in stats:
        _add_kpi_card(slide, stat, x, y, card_w, 0.78, theme, compact=True)
        x += card_w + gap


def _add_kpi_grid(
    slide: Any,
    statistics: list[StatisticResult],
    *,
    y: float,
    theme: PowerPointTheme,
    max_columns: int,
    card_height: float,
) -> None:
    if not statistics:
        return
    stats = statistics[:6]
    columns = min(max_columns, len(stats))
    gap_x = 0.25
    gap_y = 0.24
    total_w = 11.5
    card_w = (total_w - gap_x * (columns - 1)) / columns
    rows = math.ceil(len(stats) / columns)
    start_x = (13.333333 - total_w) / 2
    for index, stat in enumerate(stats):
        row, col = divmod(index, columns)
        x = start_x + col * (card_w + gap_x)
        card_y = y + row * (card_height + gap_y)
        _add_kpi_card(slide, stat, x, card_y, card_w, card_height, theme, compact=False)


def _add_kpi_card(
    slide: Any,
    stat: StatisticResult,
    x: float,
    y: float,
    w: float,
    h: float,
    theme: PowerPointTheme,
    *,
    compact: bool,
) -> None:
    shape = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = _rgb(theme.accent_fill)
    shape.line.color.rgb = _rgb(theme.accent_border)
    shape.line.width = Pt(0.8)
    tf = shape.text_frame
    tf.clear()
    tf.margin_left = Inches(0.06)
    tf.margin_right = Inches(0.06)
    tf.margin_top = Inches(0.03)
    tf.margin_bottom = Inches(0.02)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p1 = tf.paragraphs[0]
    p1.text = stat.display_name
    p1.alignment = PP_ALIGN.CENTER
    p1.font.size = Pt(theme.kpi_label_font_size if compact else theme.kpi_label_font_size + 1)
    p1.font.color.rgb = _rgb(theme.text_color)
    p2 = tf.add_paragraph()
    p2.text = _format_statistic(stat)
    p2.alignment = PP_ALIGN.CENTER
    p2.font.size = Pt(theme.kpi_value_font_size if compact else theme.kpi_value_font_size + 3)
    p2.font.bold = True
    p2.font.color.rgb = _rgb(theme.text_color)


def _add_footer(slide: Any, footer: str | None, slide_number: int, theme: PowerPointTheme) -> None:
    if footer:
        box = slide.shapes.add_textbox(Inches(0.38), Inches(7.12), Inches(11.8), Inches(0.22))
        p = box.text_frame.paragraphs[0]
        p.text = footer
        p.font.size = Pt(8)
        p.font.color.rgb = _rgb(theme.muted_text_color)
    num = slide.shapes.add_textbox(Inches(12.25), Inches(7.08), Inches(0.65), Inches(0.24))
    p = num.text_frame.paragraphs[0]
    p.text = str(slide_number)
    p.alignment = PP_ALIGN.RIGHT
    p.font.size = Pt(8)
    p.font.color.rgb = _rgb(theme.muted_text_color)


def _add_contained_picture(slide: Any, image_path: Path, x: float, y: float, w: float, h: float) -> None:
    try:
        with Image.open(image_path) as image:
            width_px, height_px = image.size
    except Exception as exc:
        raise PowerPointReportError(f"Could not read plot image '{image_path}': {exc}") from exc
    if width_px <= 0 or height_px <= 0:
        raise PowerPointReportError(f"Plot image has invalid dimensions: {image_path}")
    image_ratio = width_px / height_px
    box_ratio = w / h
    if image_ratio >= box_ratio:
        final_w = w
        final_h = w / image_ratio
        final_x = x
        final_y = y + (h - final_h) / 2
    else:
        final_h = h
        final_w = h * image_ratio
        final_x = x + (w - final_w) / 2
        final_y = y
    slide.shapes.add_picture(
        str(image_path),
        Inches(final_x),
        Inches(final_y),
        width=Inches(final_w),
        height=Inches(final_h),
    )


def _plot_path(output_file: str, assets: Path) -> Path:
    candidate = Path(output_file)
    if candidate.is_absolute():
        return candidate
    return assets / candidate.name


def _write_metadata(result: PowerPointReportResult) -> None:
    used_stats = sorted({sid for slide in result.config.slides for sid in slide.statistic_ids})
    used_plots = sorted({pid for slide in result.config.slides for pid in slide.plot_ids})
    manifest = {
        "status": "PASS",
        "source_file": str(result.statistics_result.dataset.source_path),
        "source_sha256": sha256_file(result.statistics_result.dataset.source_path),
        "configuration_file": str(result.config_path),
        "configuration_sha256": sha256_file(result.config_path),
        "presentation": str(result.presentation_path),
        "sample_count": result.sample_count,
        "slide_count": result.slide_count,
        "statistic_count": result.statistic_count,
        "plot_count": result.plot_count,
        "statistics_used": used_stats,
        "plots_used": used_plots,
        "slides": [
            {
                "slide_id": slide.slide_id,
                "slide_type": slide.slide_type,
                "title": slide.title,
                "statistics": list(slide.statistic_ids),
                "plots": list(slide.plot_ids),
            }
            for slide in result.config.slides
        ],
    }
    result.manifest_path.write_text(json.dumps(manifest, indent=2), encoding="utf-8")
    lines = [
        "VSM POWERPOINT REPORT",
        "=====================",
        "Status: PASS",
        f"Source: {result.statistics_result.dataset.source_path}",
        f"Samples: {result.sample_count}",
        f"Slides: {result.slide_count}",
        f"Statistics used: {result.statistic_count}",
        f"Plots used: {result.plot_count}",
        f"Presentation: {result.presentation_path}",
        "",
        "Slides:",
    ]
    for index, slide in enumerate(result.config.slides, start=1):
        lines.append(
            f"{index:02d}. {slide.slide_id} | {slide.slide_type} | "
            f"statistics={len(slide.statistic_ids)} | plots={len(slide.plot_ids)}"
        )
    result.summary_path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def _load_theme(raw: object) -> PowerPointTheme:
    if raw is None:
        raw = {}
    if not isinstance(raw, dict):
        raise ConfigurationError("theme must be a YAML mapping")
    allowed = {
        "title_font_size", "subtitle_font_size", "kpi_label_font_size", "kpi_value_font_size",
        "body_font_size", "accent_fill", "accent_border", "text_color", "muted_text_color",
    }
    _reject_unknown_keys(raw, allowed, "theme")
    defaults = PowerPointTheme()
    numeric_fields = {
        "title_font_size": defaults.title_font_size,
        "subtitle_font_size": defaults.subtitle_font_size,
        "kpi_label_font_size": defaults.kpi_label_font_size,
        "kpi_value_font_size": defaults.kpi_value_font_size,
        "body_font_size": defaults.body_font_size,
    }
    values: dict[str, Any] = {}
    for key, default in numeric_fields.items():
        value = raw.get(key, default)
        if isinstance(value, bool) or not isinstance(value, (int, float)) or not 6 <= float(value) <= 60:
            raise ConfigurationError(f"theme.{key} must be a number from 6 to 60")
        values[key] = float(value)
    for key, default in {
        "accent_fill": defaults.accent_fill,
        "accent_border": defaults.accent_border,
        "text_color": defaults.text_color,
        "muted_text_color": defaults.muted_text_color,
    }.items():
        values[key] = _hex_color(raw.get(key, default), f"theme.{key}")
    return PowerPointTheme(**values)


def _load_slide(raw: object, index: int) -> PowerPointSlideDefinition:
    context = f"slides[{index}]"
    if not isinstance(raw, dict):
        raise ConfigurationError(f"{context} must be a YAML mapping")
    _reject_unknown_keys(raw, {"slide_id", "type", "title", "subtitle", "statistics", "plots", "footer"}, context)
    slide_id = _identifier(raw.get("slide_id"), f"{context}.slide_id")
    slide_type = _nonempty_string(raw.get("type"), f"{context}.type")
    if slide_type not in _ALLOWED_SLIDE_TYPES:
        raise ConfigurationError(f"{context}.type must be one of: " + ", ".join(sorted(_ALLOWED_SLIDE_TYPES)))
    title = _nonempty_string(raw.get("title"), f"{context}.title")
    subtitle = _optional_string(raw.get("subtitle"), f"{context}.subtitle")
    footer = _optional_string(raw.get("footer"), f"{context}.footer")
    statistic_ids = tuple(_string_list(raw.get("statistics", []), f"{context}.statistics", allow_empty=True))
    plot_ids = tuple(_string_list(raw.get("plots", []), f"{context}.plots", allow_empty=True))
    if slide_type == "summary":
        if plot_ids:
            raise ConfigurationError(f"{context}.plots must be empty for summary slides")
        if len(statistic_ids) > 6:
            raise ConfigurationError(f"{context}.statistics supports at most 6 KPI statistics")
    if slide_type == "plot_pair":
        if not 1 <= len(plot_ids) <= 2:
            raise ConfigurationError(f"{context}.plots must contain one or two plot IDs for plot_pair slides")
        if len(statistic_ids) > 6:
            raise ConfigurationError(f"{context}.statistics supports at most 6 KPI statistics")
    return PowerPointSlideDefinition(
        slide_id=slide_id,
        slide_type=slide_type,
        title=title,
        subtitle=subtitle,
        statistic_ids=statistic_ids,
        plot_ids=plot_ids,
        footer=footer,
    )


def _format_statistic(stat: StatisticResult) -> str:
    value = stat.value
    if not math.isfinite(value):
        value_text = "n/a"
    else:
        magnitude = abs(value)
        if magnitude >= 10000:
            value_text = f"{value:,.0f}"
        elif magnitude >= 1000:
            value_text = f"{value:,.1f}"
        elif magnitude >= 100:
            value_text = f"{value:.2f}"
        elif magnitude >= 10:
            value_text = f"{value:.2f}"
        else:
            value_text = f"{value:.3f}"
        value_text = value_text.rstrip("0").rstrip(".") if "." in value_text else value_text
    unit = stat.channel_unit
    return f"{value_text} {unit}" if unit else value_text


def _rgb(value: str) -> RGBColor:
    return RGBColor.from_string(value)


def _hex_color(value: object, context: str) -> str:
    if isinstance(value, int) and not isinstance(value, bool):
        text = str(value)
    elif isinstance(value, str):
        text = value.strip().lstrip("#").upper()
    else:
        raise ConfigurationError(f"{context} must be a six-digit hexadecimal color")
    if len(text) != 6 or any(ch not in "0123456789ABCDEF" for ch in text):
        raise ConfigurationError(f"{context} must be a six-digit hexadecimal color")
    return text


def _pptx_filename(value: object) -> str:
    text = _nonempty_string(value, "presentation.output_filename")
    if Path(text).name != text or not text.lower().endswith(".pptx"):
        raise ConfigurationError("presentation.output_filename must be a plain .pptx filename")
    return text


def _nonempty_string(value: object, context: str) -> str:
    if not isinstance(value, str) or not value.strip():
        raise ConfigurationError(f"{context} must be a non-empty string")
    return value.strip()


def _optional_string(value: object, context: str) -> str | None:
    if value is None:
        return None
    return _nonempty_string(value, context)


def _identifier(value: object, context: str) -> str:
    text = _nonempty_string(value, context)
    if not all(ch.isalnum() or ch in {"_", "-"} for ch in text):
        raise ConfigurationError(f"{context} may contain only letters, numbers, '_' and '-'")
    return text


def _string_list(value: object, context: str, *, allow_empty: bool) -> list[str]:
    if not isinstance(value, list):
        raise ConfigurationError(f"{context} must be a YAML list")
    if not allow_empty and not value:
        raise ConfigurationError(f"{context} must not be empty")
    result: list[str] = []
    for item in value:
        if not isinstance(item, str) or not item.strip():
            raise ConfigurationError(f"{context} must contain only non-empty strings")
        text = item.strip()
        if text in result:
            raise ConfigurationError(f"{context} contains duplicate values")
        result.append(text)
    return result


def _duplicates(values: Iterable[str]) -> list[str]:
    seen: set[str] = set()
    duplicates: list[str] = []
    for value in values:
        if value in seen and value not in duplicates:
            duplicates.append(value)
        seen.add(value)
    return duplicates


def _reject_unknown_keys(mapping: dict[str, Any], allowed: set[str], context: str) -> None:
    unknown = sorted(set(mapping) - allowed)
    if unknown:
        raise ConfigurationError(f"Unknown key(s) in {context}: " + ", ".join(unknown))
