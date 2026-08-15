from __future__ import annotations

import json
import math
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Mapping

import yaml
from pptx import Presentation

from .errors import PowerPointReportError
from .excel_report_engine import ProfileExcelReportResult, generate_profile_excel_report
from .importer import ImportOptions
from .powerpoint_report_engine import PowerPointReportResult, build_powerpoint_report
from .profile_statistics import ProfileKPIResult, ProfileStatisticResult
from .report_profile import ReportingProfile
from .statistics_engine import (
    StatisticResult,
    StatisticsConfig,
    StatisticsOutputOptions,
    StatisticsResult,
)
from .utils import client_display_filename, sha256_file
from .version import __version__


_PRESENTATION_SLIDE_ORDER = (
    "cover",
    "system_overview",
    "executive_results",
    "vehicle_operation",
    "battery_system",
    "battery_power_recovery",
    "range_extender_generator",
    "profile_energy_range",
    "agrochemical_battery",
    "traction_auxiliaries",
    "simulation_summary",
)
_HYBRID_ACTIVITY_IDS = (
    "engine_fuel_consumption_last",
    "fuel_flow_max",
    "engine_speed_max",
    "engine_torque_max",
    "engine_power_required_max",
    "engine_energy_delivered_sum",
    "generator_torque_1_max",
    "generator_power_1_max",
)


@dataclass(frozen=True)
class PowerPointReferenceSlideSpec:
    slide_number: int
    title: str | None
    shape_count: int
    picture_count: int
    text_shape_count: int


@dataclass(frozen=True)
class PowerPointReferenceDeckSpec:
    source_path: Path
    slide_count: int
    slide_width_in: float
    slide_height_in: float
    titles: tuple[str, ...]
    slides: tuple[PowerPointReferenceSlideSpec, ...]


@dataclass
class ProfilePowerPointReportResult:
    presentation_path: Path
    manifest_path: Path
    summary_path: Path
    config_path: Path
    plot_assets_dir: Path
    excel_result: ProfileExcelReportResult
    powerpoint_result: PowerPointReportResult

    @property
    def sample_count(self) -> int:
        return self.excel_result.sample_count

    @property
    def slide_count(self) -> int:
        return self.powerpoint_result.slide_count

    @property
    def plot_count(self) -> int:
        return self.powerpoint_result.plot_count

    @property
    def displayed_kpi_count(self) -> int:
        return self.powerpoint_result.displayed_kpi_count


def inspect_reference_powerpoint_layout(path: str | Path) -> PowerPointReferenceDeckSpec:
    """Return a concise structural specification for a reference PowerPoint deck."""

    source = Path(path).expanduser().resolve()
    if not source.exists() or not source.is_file():
        raise PowerPointReportError(f"Reference PowerPoint deck does not exist: {source}")
    try:
        prs = Presentation(source)
    except Exception as exc:
        raise PowerPointReportError(f"Could not open reference PowerPoint deck '{source}': {exc}") from exc

    slides: list[PowerPointReferenceSlideSpec] = []
    titles: list[str] = []
    for slide_number, slide in enumerate(prs.slides, start=1):
        text_items: list[str] = []
        picture_count = 0
        text_shape_count = 0
        for shape in slide.shapes:
            if shape.shape_type == 13:
                picture_count += 1
            if getattr(shape, "has_text_frame", False):
                text = " ".join(paragraph.text for paragraph in shape.text_frame.paragraphs).strip()
                if text:
                    text_shape_count += 1
                    text_items.append(text)
        title = text_items[0] if text_items else None
        if title:
            titles.append(title)
        slides.append(
            PowerPointReferenceSlideSpec(
                slide_number=slide_number,
                title=title,
                shape_count=len(slide.shapes),
                picture_count=picture_count,
                text_shape_count=text_shape_count,
            )
        )
    return PowerPointReferenceDeckSpec(
        source_path=source,
        slide_count=len(prs.slides),
        slide_width_in=round(prs.slide_width / 914400, 3),
        slide_height_in=round(prs.slide_height / 914400, 3),
        titles=tuple(titles),
        slides=tuple(slides),
    )


def generate_profile_powerpoint_report(
    input_file: str | Path,
    profile_file: str | Path,
    output_dir: str | Path,
    import_options: ImportOptions | None = None,
    *,
    output_filename: str | None = None,
) -> ProfilePowerPointReportResult:
    excel_result = generate_profile_excel_report(
        input_file,
        profile_file,
        output_dir,
        import_options or ImportOptions(strict=True),
        output_filename=_profile_report_filename_from_path(profile_file, ".xlsx"),
    )
    return build_profile_powerpoint_report(excel_result, output_dir, output_filename=output_filename)


def build_profile_powerpoint_report(
    excel_result: ProfileExcelReportResult,
    output_dir: str | Path,
    *,
    output_filename: str | None = None,
) -> ProfilePowerPointReportResult:
    profile = excel_result.profile
    if profile.presentation is None or not profile.presentation.slides:
        raise PowerPointReportError(f"Reporting profile '{profile.profile_id}' does not define presentation slides")

    destination = Path(output_dir).expanduser().resolve()
    destination.mkdir(parents=True, exist_ok=True)
    config_path = destination / "profile_powerpoint_report_config.yaml"
    config_path.write_text(
        yaml.safe_dump(
            _profile_powerpoint_config(
                excel_result,
                output_filename=output_filename or _profile_output_filename(profile, ".pptx"),
            ),
            sort_keys=False,
            allow_unicode=False,
        ),
        encoding="utf-8",
    )
    statistics_result = _profile_statistics_as_powerpoint_statistics(excel_result)
    powerpoint_result = build_powerpoint_report(
        statistics_result,
        excel_result.plotting_result,  # type: ignore[arg-type]
        config_path,
        destination,
        plot_assets_dir=excel_result.plot_assets_dir,
    )
    result = ProfilePowerPointReportResult(
        presentation_path=powerpoint_result.presentation_path,
        manifest_path=powerpoint_result.manifest_path,
        summary_path=powerpoint_result.summary_path,
        config_path=config_path,
        plot_assets_dir=excel_result.plot_assets_dir,
        excel_result=excel_result,
        powerpoint_result=powerpoint_result,
    )
    _write_profile_manifest(result)
    return result


def _profile_powerpoint_config(
    excel_result: ProfileExcelReportResult,
    *,
    output_filename: str,
) -> dict[str, Any]:
    profile = excel_result.profile
    powertrain = (profile.metadata.powertrain or "").strip().lower()
    is_hybrid = powertrain == "hybrid"
    is_active_hybrid = _hybrid_subsystem_active(excel_result) if is_hybrid else False
    slides_by_id = profile.presentation.slides_by_id() if profile.presentation else {}

    def slide_config(slide_id: str) -> Any:
        try:
            return slides_by_id[slide_id]
        except KeyError as exc:
            raise PowerPointReportError(
                f"Reporting profile '{profile.profile_id}' is missing presentation slide '{slide_id}'"
            ) from exc

    slide_7_id = "range_extender_generator" if is_hybrid else "profile_energy_range"
    ordered_ids = (
        "cover",
        "system_overview",
        "executive_results",
        "vehicle_operation",
        "battery_system",
        "battery_power_recovery",
        slide_7_id,
        "agrochemical_battery",
        "traction_auxiliaries",
        "simulation_summary",
    )
    for slide_id in ordered_ids:
        slide_config(slide_id)

    footer = (profile.presentation.footer if profile.presentation else None) or "VSM Engineering Post-Processing Tool"
    title = profile.metadata.name
    subtitle = "VSM Engineering - Feasibility Study"
    return {
        "version": 1,
        "presentation": {
            "title": title,
            "subtitle": subtitle,
            "output_filename": output_filename,
            "footer": footer,
            "author": "VSM Engineering",
            "subject": f"{profile.metadata.name} deterministic profile engineering report",
            "keywords": "VSM, RoboSprayer, profile-driven, engineering report",
            "comments": f"Generated from deterministic profile outputs by v{__version__}.",
        },
        "theme": {
            "title_font_size": 30,
            "subtitle_font_size": 13.5,
            "kpi_label_font_size": 9.5,
            "kpi_value_font_size": 20,
            "body_font_size": 10.5,
            "accent_fill": "F7F3E8",
            "accent_border": "9C7A2F",
            "text_color": "1E252B",
            "muted_text_color": "5D6770",
            "background_color": "FFFFFF",
            "rule_color": "B8C2CC",
        },
        "slides": [
            _slide(
                slide_config("cover"),
                "cover",
                title,
                subtitle,
                body=_cover_body(excel_result),
            ),
            _slide(
                slide_config("system_overview"),
                "overview",
                "System and Simulation Overview",
                _overview_subtitle(excel_result),
                body=_overview_body(excel_result),
            ),
            _slide(
                slide_config("executive_results"),
                "kpi_grid",
                "Executive Results",
                "Deterministic KPIs from the validated profile pipeline",
                body=("Calculations are sourced from the profile statistics and KPI results.",),
            ),
            _slide(
                slide_config("vehicle_operation"),
                "plot_full",
                "Vehicle Operation",
                "Speed trace from the validated profile plot output",
            ),
            _slide(
                slide_config("battery_system"),
                "plot_pair",
                "Battery and Electrical Energy System",
                "State of charge, stored energy and battery demand",
            ),
            _slide(
                slide_config("battery_power_recovery"),
                "plot_pair",
                "Battery Power and Energy Recovery",
                "Charge/discharge behavior and released/recuperated energy",
            ),
            _slide(
                slide_config(slide_7_id),
                "plot_pair",
                "Range Extender and Generator" if is_hybrid else "Energy Consumption and Estimated Range",
                _slide_7_subtitle(is_hybrid, is_active_hybrid),
                body=_slide_7_body(is_hybrid, is_active_hybrid),
            ),
            _slide(
                slide_config("agrochemical_battery"),
                "plot_pair",
                "Agrochemical and Battery Behaviour",
                "Profile-supported system interaction without reconstructed charging assumptions",
            ),
            _slide(
                slide_config("traction_auxiliaries"),
                "plot_pair",
                "Traction, EDU and Auxiliary Energy Demand",
                "Wheel, EDU, tyre and auxiliary demand from profile plots",
            ),
            _slide(
                slide_config("simulation_summary"),
                "conclusion",
                "Simulation Summary",
                "Deterministic profile result snapshot",
                body=_summary_body(excel_result, is_hybrid, is_active_hybrid),
            ),
        ],
    }


def _slide(
    definition: Any,
    slide_type: str,
    title: str,
    subtitle: str | None,
    *,
    body: tuple[str, ...] = (),
) -> dict[str, Any]:
    return {
        "slide_id": definition.slide_id,
        "type": slide_type,
        "title": title,
        "subtitle": subtitle,
        "statistics": list(definition.statistics),
        "plots": list(definition.plots),
        "body": list(body),
    }


def _profile_statistics_as_powerpoint_statistics(excel_result: ProfileExcelReportResult) -> StatisticsResult:
    statistics = [
        _profile_statistic_as_powerpoint_item(item)
        for item in excel_result.statistics_result.statistics
    ]
    statistics.extend(
        _profile_kpi_as_powerpoint_item(item, excel_result.sample_count, excel_result.profile)
        for item in excel_result.statistics_result.kpis
    )
    return StatisticsResult(
        dataset=excel_result.dataset,
        config_path=excel_result.report_path,
        config=StatisticsConfig(version=1, statistics=(), output=StatisticsOutputOptions()),
        channels_by_id=excel_result.plotting_result.channels_by_semantic_name,
        values_by_id=excel_result.plotting_result.values_by_semantic_name,
        statistics=statistics,
        math_result=None,
    )


def _profile_statistic_as_powerpoint_item(item: ProfileStatisticResult) -> StatisticResult:
    definition = item.definition
    display_name = definition.display_name or item.channel_display_name
    return StatisticResult(
        statistic_id=definition.statistic_id,
        channel_id=item.target_channel,
        channel_display_name=item.channel_display_name,
        channel_unit=item.channel_unit,
        channel_kind=item.channel_kind,
        operation=definition.operation,
        placement_group=definition.placement_group or "profile",
        nan_policy="error",
        value=item.value,
        sample_count=item.sample_count,
        used_sample_count=item.used_sample_count,
        omitted_sample_count=item.omitted_sample_count,
        display_name=display_name,
        description=definition.notes,
        comparison=None,
    )


def _profile_kpi_as_powerpoint_item(
    item: ProfileKPIResult,
    sample_count: int,
    profile: ReportingProfile,
) -> StatisticResult:
    definition = item.definition
    display_name = definition.display_name or definition.kpi_id.replace("_", " ").title()
    if (profile.metadata.powertrain or "").lower() == "electric" and definition.kpi_id == "range_85_battery_km":
        display_name = "Range for 85% Battery"
    return StatisticResult(
        statistic_id=definition.kpi_id,
        channel_id=definition.kpi_id,
        channel_display_name=display_name,
        channel_unit=definition.unit,
        channel_kind="KPI",
        operation="kpi",
        placement_group=definition.placement_group or "profile",
        nan_policy="error",
        value=item.value,
        sample_count=sample_count,
        used_sample_count=sample_count,
        omitted_sample_count=0,
        display_name=display_name,
        description=definition.notes,
        comparison=None,
    )


def _cover_body(result: ProfileExcelReportResult) -> tuple[str, ...]:
    profile = result.profile
    return (
        f"Source: {client_display_filename(result.dataset.source_path)}",
        f"Powertrain: {(profile.metadata.powertrain or profile.profile_id).title()}",
        f"Resolved raw channels: {len(result.resolution.resolved)}/{len(profile.raw_channels)}",
        f"Profile MATH/statistics/KPIs: {result.math_count}/{result.statistic_count}/{result.kpi_count}",
        f"Profile plots rendered: {result.plot_count}",
    )


def _overview_subtitle(result: ProfileExcelReportResult) -> str:
    powertrain = (result.profile.metadata.powertrain or "profile").title()
    return f"{powertrain} profile resolution and deterministic simulation scope"


def _overview_body(result: ProfileExcelReportResult) -> tuple[str, ...]:
    powertrain = (result.profile.metadata.powertrain or "profile").title()
    inactive_names = [
        item.definition.report_name.strip()
        for item in result.resolution.resolved.values()
        if item.is_all_zero
    ]
    left = (
        f"Selected reporting profile: {result.profile.metadata.name}",
        f"Powertrain type: {powertrain}",
        f"Imported samples: {result.sample_count:,}",
        f"Source channels: {result.source_raw_channel_count:,}",
        f"Resolved required raw channels: {len(result.resolution.resolved):,}",
    )
    right = (
        f"Calculated MATH channels: {result.math_count:,}",
        f"Calculated statistics: {result.statistic_count:,}",
        f"Calculated KPIs: {result.kpi_count:,}",
        f"Rendered profile plots: {result.plot_count:,}",
        _inactive_overview_line(result, inactive_names),
    )
    return (*left, *right)


def _inactive_overview_line(result: ProfileExcelReportResult, inactive_names: list[str]) -> str:
    powertrain = (result.profile.metadata.powertrain or "").lower()
    if powertrain == "hybrid":
        return "Hybrid engine/generator activity is reported from resolved channel values."
    if inactive_names:
        return f"Inactive/all-zero resolved channels: {len(inactive_names):,}"
    return "No inactive required subsystem channels were detected."


def _slide_7_subtitle(is_hybrid: bool, is_active_hybrid: bool) -> str:
    if not is_hybrid:
        return "Battery energy consumption and range indicators"
    if is_active_hybrid:
        return "Engine demand, fuel consumption and generator behavior"
    return "Resolved range-extender channels are inactive in this simulation"


def _slide_7_body(is_hybrid: bool, is_active_hybrid: bool) -> tuple[str, ...]:
    if not is_hybrid:
        return ("Range and consumption are deterministic profile KPI outputs.",)
    if is_active_hybrid:
        return ("Range-extender activity is derived from non-zero resolved engine/generator statistics.",)
    return ("ICE/generator channels resolved; no operating activity detected in this simulation.",)


def _summary_body(result: ProfileExcelReportResult, is_hybrid: bool, is_active_hybrid: bool) -> tuple[str, ...]:
    values = _statistic_values(result)
    sentences = [
        _sentence(
            "Deterministic run covered {distance} in {time}.",
            distance=_formatted_value(values, "distance_km_last", "km"),
            time=_formatted_value(values, "time_minutes_last", "min"),
        ),
        _sentence(
            "Final battery SOC is {soc}.",
            soc=_formatted_value(values, "battery_soc_last", "%"),
        ),
    ]
    if is_hybrid:
        sentences.append(
            _sentence(
                "Fuel consumption is {fuel}; maximum generator power is {generator}.",
                fuel=_formatted_value(values, "engine_fuel_consumption_last", "kg"),
                generator=_formatted_value(values, "generator_power_1_max", "kW"),
            )
        )
        if not is_active_hybrid:
            sentences.append("Resolved ICE/generator channels are inactive/all-zero for this run.")
    else:
        sentences.append(
            _sentence(
                "Battery capacity used is {used}; estimated range for 85 percent battery is {range_value}.",
                used=_formatted_value(values, "battery_capacity_used", "kWh"),
                range_value=_formatted_value(values, "range_85_battery_km", "km"),
            )
        )
    sentences.append("Results are deterministic simulation outputs; acceptance conclusions require explicit engineering criteria.")
    return tuple(item for item in sentences if item)


def _sentence(template: str, **values: str) -> str:
    if any(value == "n/a" for value in values.values()):
        return ""
    return template.format(**values)


def _statistic_values(result: ProfileExcelReportResult) -> dict[str, float]:
    values = {item.definition.statistic_id: item.value for item in result.statistics_result.statistics}
    values.update({item.definition.kpi_id: item.value for item in result.statistics_result.kpis})
    return values


def _formatted_value(values: Mapping[str, float], statistic_id: str, unit: str) -> str:
    value = values.get(statistic_id)
    if value is None or not math.isfinite(value):
        return "n/a"
    if unit == "rpm":
        return f"{value:,.0f} {unit}"
    if abs(value) >= 1000:
        return f"{value:,.1f} {unit}"
    return f"{value:.2f} {unit}"


def _hybrid_subsystem_active(result: ProfileExcelReportResult) -> bool:
    values = _statistic_values(result)
    if any(abs(values.get(statistic_id, 0.0)) > 1e-9 for statistic_id in _HYBRID_ACTIVITY_IDS):
        return True
    active_semantic_names = {
        "engine_fuel_consumption",
        "fuel_flow",
        "engine_speed",
        "engine_torque",
        "generator_torque_1",
    }
    return any(
        semantic_name in active_semantic_names and resolved.is_active and not resolved.is_all_zero
        for semantic_name, resolved in result.resolution.resolved.items()
    )


def _profile_report_filename_from_path(profile_file: str | Path, suffix: str) -> str:
    profile_name = Path(profile_file).stem.replace("robosprayer_", "RoboSprayer_").replace("_", " ")
    if "electric" in profile_name.lower():
        return "RoboSprayer_Electric_Engineering_Report" + suffix
    if "hybrid" in profile_name.lower():
        return "RoboSprayer_Hybrid_Engineering_Report" + suffix
    return _plain_report_filename(profile_name, suffix)


def _profile_output_filename(profile: ReportingProfile, suffix: str) -> str:
    return _plain_report_filename(profile.metadata.name, suffix)


def _plain_report_filename(name: str, suffix: str) -> str:
    stem = re.sub(r"[^A-Za-z0-9]+", "_", name).strip("_") or "VSM_Profile"
    return f"{stem}_Engineering_Report{suffix}"


def _write_profile_manifest(result: ProfilePowerPointReportResult) -> None:
    payload = {
        "status": "PASS",
        "profile_id": result.excel_result.profile.profile_id,
        "profile_name": result.excel_result.profile.metadata.name,
        "powertrain": result.excel_result.profile.metadata.powertrain,
        "source_file": str(result.excel_result.dataset.source_path),
        "source_sha256": sha256_file(result.excel_result.dataset.source_path),
        "presentation": str(result.presentation_path),
        "sample_count": result.sample_count,
        "slide_count": result.slide_count,
        "plot_count": result.plot_count,
        "displayed_kpi_count": result.displayed_kpi_count,
        "excel_report": str(result.excel_result.report_path),
        "plot_assets_dir": str(result.plot_assets_dir),
    }
    profile_manifest = result.presentation_path.with_name("profile_powerpoint_report_manifest.json")
    profile_manifest.write_text(json.dumps(payload, indent=2), encoding="utf-8")
