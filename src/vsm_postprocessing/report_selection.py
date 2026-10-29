"""User selections shared by profile reports, without a separate reporting engine."""
from collections import Counter
from dataclasses import dataclass, replace
from typing import Mapping

from .errors import ConfigurationError
from .math_engine import _compile_expression
from .report_profile import ReportingProfile


@dataclass(frozen=True)
class ProfileReportSelection:
    export_channels: tuple[str, ...] | None = None
    math_channels: tuple[str, ...] | None = None
    statistics: tuple[str, ...] | None = None
    kpis: tuple[str, ...] | None = None
    plots: tuple[str, ...] | None = None
    # Replace plots in existing slide slots; no additional slides or layout changes.
    powerpoint_plots: Mapping[str, tuple[str, ...]] | None = None
    powerpoint_statistics: tuple[str, ...] | None = None
    bottom_operations: tuple[str, ...] | None = None
    plot_series: Mapping[str, tuple[str, ...]] | None = None


def select_profile(profile: ReportingProfile, selection: ProfileReportSelection | None) -> ReportingProfile:
    if selection is None:
        return profile
    raw, math = profile.raw_by_semantic_name(), profile.math_by_semantic_name()
    stats, kpis, plots = profile.statistics_by_id(), profile.kpis_by_id(), profile.plots_by_id()

    def chosen(requested, known, label):
        if requested is None:
            return set(known)
        unknown = set(requested) - set(known)
        duplicates = sorted(name for name, count in Counter(requested).items() if count > 1)
        if unknown or duplicates:
            details = []
            if unknown:
                details.append(f"unknown identifiers: {sorted(unknown)}")
            if duplicates:
                details.append(f"duplicate identifiers: {duplicates}")
            raise ConfigurationError(f"Invalid {label} selection: {'; '.join(details)}")
        return set(requested)

    exported = chosen(selection.export_channels, raw.keys() | math.keys(), "export channels")
    selected_math = chosen(selection.math_channels, math, "MATH channels")
    selected_stats = chosen(selection.statistics, stats, "statistics")
    selected_kpis = chosen(selection.kpis, kpis, "KPIs")
    selected_plots = chosen(selection.plots, plots, "plots")
    if selection.plot_series is not None:
        unknown = set(selection.plot_series) - set(plots)
        if unknown:
            raise ConfigurationError(f"Unknown plot series selections: {sorted(unknown)}")
        for name, requested in selection.plot_series.items():
            definition = plots[name]
            allowed = chosen(requested, {s.semantic_name for s in definition.series}, "plot series")
            plots[name] = replace(definition, series=tuple(s for s in definition.series if s.semantic_name in allowed))
    chosen(selection.bottom_operations, {"max", "min", "first", "last"}, "Excel column statistics")
    presentation = profile.presentation
    if presentation and selection.powerpoint_plots is not None:
        unknown_slides = set(selection.powerpoint_plots) - set(presentation.slides_by_id())
        if unknown_slides:
            raise ConfigurationError(f"Unknown PowerPoint slides: {sorted(unknown_slides)}")
        slides = []
        for slide in presentation.slides:
            requested = selection.powerpoint_plots.get(slide.slide_id, slide.plots)
            chosen(requested, plots, "PowerPoint plots")
            if len(requested) > max(2, len(slide.plots)):
                raise ConfigurationError("PowerPoint plot selection exceeds the existing slide slots")
            selected_plots.update(requested)
            slides.append(replace(slide, plots=tuple(requested),
                                  fallback_mode="configured" if tuple(requested) != slide.plots else slide.fallback_mode))
        presentation = replace(presentation, slides=tuple(slides))
    if presentation:
        # Additional pages follow explicit chart inclusion, including the new maps.
        presentation = replace(presentation, slides=tuple(
            replace(slide, plots=tuple(p for p in slide.plots if p in selected_plots))
            if slide.mode == "additional_plot_pair" and any(
                plots[p].report_group == "torque_maps" for p in slide.plots if p in plots) else slide
            for slide in presentation.slides))
    if presentation and selection.powerpoint_statistics is not None:
        known = stats.keys() | kpis.keys() | {s for slide in presentation.slides for s in slide.statistics}
        allowed = chosen(selection.powerpoint_statistics, known, "PowerPoint statistics")
        selected_stats.update(allowed & stats.keys())
        selected_kpis.update(allowed & kpis.keys())
        presentation = replace(presentation, slides=tuple(
            replace(slide, statistics=tuple(s for s in slide.statistics if s in allowed)) for slide in presentation.slides))

    # Keep explicit KPI prerequisites even if they are not selected as separate
    # cards. They retain the same numerical definitions and dependency tracing.
    pending = list(selected_kpis)
    visited = set()
    while pending:
        name = pending.pop()
        if name in visited:
            continue
        visited.add(name)
        for dep in _compile_expression(kpis[name].expression, context=name).dependencies:
            if dep in stats:
                selected_stats.add(dep)
            elif dep in kpis:
                selected_kpis.add(dep); pending.append(dep)
            elif dep in math:
                selected_math.add(dep)
    needed = set(selected_math)
    # Shared optional reporting KPI needs remaining mass even when its plot or
    # individual math-channel export is deselected.
    if 'agrochemical_discharge' in math:
        needed.add('agrochemical_discharge')
    if selection.export_channels is not None:
        needed.update(exported & math.keys())
    needed.update(stats[name].target for name in selected_stats)
    for name in selected_plots:
        needed.add(plots[name].x)
        needed.update(s.semantic_name for s in plots[name].series)
        needed.update(s.x for s in plots[name].series if s.x)
    pending = list(needed)
    while pending:
        name = pending.pop()
        if name in math:
            for dep in _compile_expression(math[name].expression, context=name).dependencies:
                if dep not in needed:
                    needed.add(dep); pending.append(dep)
    return replace(profile,
                   math_channels=tuple(c for c in profile.math_channels if c.semantic_name in needed),
                   statistics=tuple(c for c in profile.statistics if c.statistic_id in selected_stats),
                   kpis=tuple(c for c in profile.kpis if c.kpi_id in selected_kpis),
                   plots=tuple(plots[c.plot_id] for c in profile.plots if c.plot_id in selected_plots),
                   presentation=presentation)
