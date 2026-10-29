from __future__ import annotations

import csv
import json
from dataclasses import dataclass, field, replace, asdict
from pathlib import Path
from typing import Any, Mapping

import numpy as np

from .errors import PlottingError
from .importer import ImportOptions, load_data_file
from .models import ChannelInfo, ImportedDataset
from .plotting_engine import (
    PlotDefinition,
    PlotDefaults,
    PlotSeriesDefinition,
    RenderedPlot,
    _render_one_plot,
)
from .profile_math import ProfileMathResult, calculate_profile_math_channels
from .report_profile import (
    ProfilePlotDefinition,
    ProfileResolutionResult,
    ReportingProfile,
    resolve_profile,
)
from .utils import sha256_file


@dataclass(frozen=True)
class UnavailableProfilePlot:
    definition: ProfilePlotDefinition
    reason: str
    missing_semantic_names: tuple[str, ...] = ()
    omitted_optional_series: tuple[str, ...] = ()


@dataclass(frozen=True)
class ProfilePlotSeriesSummary:
    semantic_name: str
    source_name: str
    unit: str | None
    axis: str
    is_constant: bool
    is_all_zero: bool
    min_value: float
    max_value: float
    x_semantic_name: str | None = None
    x_source_name: str | None = None
    x_unit: str | None = None
    component: str | None = None
    measurement_location: str | None = None


@dataclass
class ProfilePlottingResult:
    dataset: ImportedDataset
    profile: ReportingProfile
    resolution: ProfileResolutionResult
    math_result: ProfileMathResult
    rendered_plots: list[RenderedPlot]
    unavailable_plots: list[UnavailableProfilePlot]
    channels_by_semantic_name: dict[str, ChannelInfo]
    values_by_semantic_name: dict[str, np.ndarray]
    series_summaries: dict[str, tuple[ProfilePlotSeriesSummary, ...]]
    series_diagnostics: dict[str, tuple[dict[str, Any], ...]] = field(default_factory=dict)

    @property
    def sample_count(self) -> int:
        return self.dataset.quality.sample_count

    @property
    def configured_plot_count(self) -> int:
        return len(self.profile.plots)

    @property
    def rendered_plot_count(self) -> int:
        return len(self.rendered_plots)

    @property
    def unavailable_plot_count(self) -> int:
        return len(self.unavailable_plots)

    @property
    def series_count(self) -> int:
        return sum(len(item.primary_series_ids) + len(item.secondary_series_ids) for item in self.rendered_plots)


def render_profile_plot_file(
    input_file: str | Path,
    profile: ReportingProfile,
    output_dir: str | Path,
    import_options: ImportOptions | None = None,
    defaults: PlotDefaults | None = None,
) -> ProfilePlottingResult:
    dataset = load_data_file(input_file, import_options)
    return render_profile_plots(dataset, profile, output_dir, defaults=defaults)


def render_profile_plots(
    dataset: ImportedDataset,
    profile: ReportingProfile,
    output_dir: str | Path,
    resolution: ProfileResolutionResult | None = None,
    math_result: ProfileMathResult | None = None,
    defaults: PlotDefaults | None = None,
    *,
    reporting_agro=None,
) -> ProfilePlottingResult:
    if resolution is None:
        resolution = resolve_profile(dataset, profile)
    if math_result is None:
        math_result = calculate_profile_math_channels(dataset, profile, resolution)

    channels_by_name, values_by_name = _build_semantic_maps(dataset, resolution, math_result)
    if reporting_agro is not None:
        # A reporting view, not a mutation of the validated legacy MATH arrays.
        name = 'reporting_agro_mass'
        mass_label = 'Remaining Mass (' + reporting_agro.capacity.status + ')'
        if reporting_agro.mass_values is not None:
            prototype = next(iter(channels_by_name.values()))
            channels_by_name[name] = replace(prototype, channel_id=name, unit='kg', kind='reporting',
                source_name=reporting_agro.capacity.source, display_name=mass_label,
                provenance=reporting_agro.capacity.provenance)
            values_by_name[name] = reporting_agro.mass_values
        plots = []
        for plot in profile.plots:
            if any(s.semantic_name == 'agrochemical_discharge' for s in plot.series):
                plot = replace(plot, series=tuple(replace(s, semantic_name=name,
                    label=mass_label) if s.semantic_name == 'agrochemical_discharge' else s
                    for s in plot.series), primary_y_label=mass_label + ' [kg]')
            plots.append(plot)
        profile = replace(profile, plots=tuple(plots))
    destination = Path(output_dir).expanduser().resolve()
    destination.mkdir(parents=True, exist_ok=True)
    defaults = defaults or PlotDefaults()

    rendered: list[RenderedPlot] = []
    unavailable: list[UnavailableProfilePlot] = []
    summaries: dict[str, tuple[ProfilePlotSeriesSummary, ...]] = {}
    for definition in sorted(profile.plots, key=lambda item: item.order or 999_999):
        prepared = _prepare_plot_definition(definition, channels_by_name, values_by_name, dataset.quality.sample_count)
        if isinstance(prepared, UnavailableProfilePlot):
            unavailable.append(prepared)
            continue
        plotting_definition, series_summaries = prepared
        plot_defaults = defaults
        if definition.report_group == "torque_maps":
            # Keep the established 2:1 panels, with readable text at slide size.
            plot_defaults = replace(defaults, width_inches=7.2, height_inches=3.6,
                                    dpi=max(180, defaults.dpi))
        if (profile.metadata.powertrain == "diesel" and profile.metadata.excel_template
                and definition.plot_id in {"diesel_engine_oil_temperature", "diesel_engine_water_temperature"}):
            temperatures = [values_by_name[n] for n in ("engine_oil_temperature", "engine_water_temperature")
                            if n in values_by_name]
            low = min(float(np.nanmin(v)) for v in temperatures)
            high = max(float(np.nanmax(v)) for v in temperatures)
            padding = max(1.0, (high-low) * .05)
            colour = "#C65911" if definition.plot_id.endswith("oil_temperature") else "#0070C0"
            plotting_definition = replace(plotting_definition,
                primary_y_label="Temperature [°C]", primary_y_limits=(low-padding, high+padding),
                series=tuple(replace(s, colour=colour) for s in plotting_definition.series))
        rendered.append(
            _render_one_plot(
                plotting_definition,
                plot_defaults,
                destination,
                channels_by_name,
                values_by_name,
                dataset.quality.sample_count,
            )
        )
        summaries[definition.plot_id] = series_summaries

    result = ProfilePlottingResult(
        dataset=dataset,
        profile=profile,
        resolution=resolution,
        math_result=math_result,
        rendered_plots=rendered,
        unavailable_plots=unavailable,
        channels_by_semantic_name=channels_by_name,
        values_by_semantic_name=values_by_name,
        series_summaries=summaries,
        series_diagnostics={plot.plot_id: tuple(
            _paired_series_status(plot, item, channels_by_name, values_by_name, dataset.quality.sample_count)
            for item in plot.series) for plot in profile.plots if plot.render_mode == "scatter"},
    )
    _write_profile_plot_metadata(result, destination)
    return result


def _build_semantic_maps(
    dataset: ImportedDataset,
    resolution: ProfileResolutionResult,
    math_result: ProfileMathResult,
) -> tuple[dict[str, ChannelInfo], dict[str, np.ndarray]]:
    channels: dict[str, ChannelInfo] = {}
    values: dict[str, np.ndarray] = {}
    for semantic_name, resolved in resolution.resolved.items():
        index = dataset.channel_index(resolved.channel.channel_id)
        channels[semantic_name] = replace(
            resolved.engineering_channel,
            channel_id=semantic_name,
            display_name=resolved.definition.report_name,
            kind=resolved.definition.channel_type.lower(),
            provenance=f"profile:{resolution.profile.profile_id}:{semantic_name}:{resolved.channel.channel_id}"
                       + (f"; {resolved.unit_interpretation_note}" if resolved.unit_interpretation else ""),
        )
        values[semantic_name] = np.asarray(dataset.values[:, index], dtype=np.float64)

    calculated_by_semantic = {
        channel.channel_id.rsplit("__math__", 1)[-1]: channel for channel in math_result.calculated_channels
    }
    for semantic_name, series in math_result.values_by_semantic_name.items():
        definition = math_result.profile.math_by_semantic_name()[semantic_name]
        source_channel = calculated_by_semantic[semantic_name]
        channels[semantic_name] = replace(
            source_channel,
            channel_id=semantic_name,
            display_name=definition.report_name,
            provenance=f"profile:{math_result.profile.profile_id}:{semantic_name}:{source_channel.channel_id}",
        )
        values[semantic_name] = np.asarray(series, dtype=np.float64)
    return channels, values


def _prepare_plot_definition(
    definition: ProfilePlotDefinition,
    channels_by_name: Mapping[str, ChannelInfo],
    values_by_name: Mapping[str, np.ndarray],
    sample_count: int,
) -> tuple[PlotDefinition, tuple[ProfilePlotSeriesSummary, ...]] | UnavailableProfilePlot:
    if definition.render_mode != "scatter" and definition.x not in channels_by_name:
        return UnavailableProfilePlot(definition, f"x channel unavailable: {definition.x}", (definition.x,))
    x_values = values_by_name.get(definition.x, np.array([]))
    if definition.render_mode != "scatter" and x_values.size != sample_count:
        raise PlottingError(f"Profile plot '{definition.plot_id}' x channel '{definition.x}' length is not aligned")

    series: list[PlotSeriesDefinition] = []
    missing_required: list[str] = []
    missing_optional: list[str] = []
    series_summaries: list[ProfilePlotSeriesSummary] = []
    pair_reasons: list[str] = []
    for item in definition.series:
        x_name = item.x or definition.x
        if definition.render_mode == "scatter":
            state = _paired_series_status(definition, item, channels_by_name, values_by_name, sample_count)
            if not state["included"]:
                pair_reasons.append(f"{item.label or item.component or item.semantic_name}: {state['reason']}")
                missing_optional.append(item.semantic_name)
                if item.required and state["status"] != "inactive":
                    missing_required.extend([name for name in (x_name, item.semantic_name) if name not in channels_by_name])
                    if not missing_required:
                        missing_required.append(item.semantic_name)
                continue
            x_values = values_by_name[x_name]
        elif item.x:
            if x_name not in channels_by_name:
                (missing_required if item.required else missing_optional).append(x_name)
                continue
            x_values = values_by_name[x_name]
        if item.semantic_name not in channels_by_name:
            if item.required:
                missing_required.append(item.semantic_name)
            else:
                missing_optional.append(item.semantic_name)
            continue
        y_values = values_by_name[item.semantic_name]
        if y_values.size != sample_count or y_values.size != x_values.size:
            raise PlottingError(
                f"Profile plot '{definition.plot_id}' series '{item.semantic_name}' length is not aligned"
            )
        finite = y_values[np.isfinite(y_values)]
        if finite.size == 0:
            raise PlottingError(f"Profile plot '{definition.plot_id}' series '{item.semantic_name}' has no finite values")
        channel = channels_by_name[item.semantic_name]
        series.append(
            PlotSeriesDefinition(
                channel_id=item.semantic_name,
                axis=item.axis,
                label=item.label,
                x_channel_id=x_name if item.x is not None or definition.render_mode == "scatter" else None,
            )
        )
        series_summaries.append(
            ProfilePlotSeriesSummary(
                semantic_name=item.semantic_name,
                source_name=channel.source_name,
                unit=channel.unit,
                axis=item.axis,
                is_constant=bool(np.allclose(finite, finite[0], rtol=0.0, atol=1e-12)),
                is_all_zero=bool(np.allclose(finite, 0.0, rtol=0.0, atol=1e-12)),
                min_value=float(finite.min()),
                max_value=float(finite.max()),
                x_semantic_name=x_name,
                x_source_name=channels_by_name[x_name].source_name,
                x_unit=channels_by_name[x_name].unit,
                component=item.component,
                measurement_location=item.measurement_location,
            )
        )

    if missing_required:
        return UnavailableProfilePlot(
            definition,
            "required series unavailable: " + ", ".join(missing_required),
            tuple(missing_required),
            tuple(missing_optional),
        )
    if not series:
        return UnavailableProfilePlot(
            definition,
            "; ".join(pair_reasons) or "no renderable series",
            (),
            tuple(missing_optional),
        )

    return (
        PlotDefinition(
            plot_id=definition.plot_id,
            title=definition.title,
            x_channel_id=(series[0].x_channel_id or definition.x) if definition.render_mode == "scatter" else definition.x,
            series=tuple(series),
            output_filename=definition.output_filename or f"{definition.plot_id}.png",
            x_label=definition.x_label,
            primary_y_label=definition.primary_y_label,
            secondary_y_label=definition.secondary_y_label,
            reference_chart_number=definition.reference_chart_number,
            render_mode=definition.render_mode,
        ),
        tuple(series_summaries),
    )


def _paired_series_status(definition, item, channels, values, sample_count):
    """Classify complete pairs; never filter or manufacture individual samples."""
    x_name, y_name = item.x or definition.x, item.semantic_name
    record = {"component": item.component or y_name, "x": x_name, "y": y_name,
              "measurement_location": item.measurement_location, "included": False,
              "status": "unavailable", "reason": "", "sample_count": sample_count}
    missing = [name for name in (x_name, y_name) if name not in channels or name not in values]
    if missing:
        record["missing_channels"] = missing
        record["reason"] = "; ".join(label + " source unavailable" for name, label in
                                      ((x_name, "RPM"), (y_name, "Torque")) if name in missing)
        return record
    x, y = np.asarray(values[x_name]), np.asarray(values[y_name])
    record.update(x_source=channels[x_name].source_name, y_source=channels[y_name].source_name,
                  x_unit=channels[x_name].unit, y_unit=channels[y_name].unit)
    if x.ndim != 1 or y.ndim != 1 or x.size != sample_count or y.size != sample_count or not sample_count:
        record["reason"] = "X/Y samples are not aligned"
    elif not np.isfinite(x).all() or not np.isfinite(y).all():
        record["reason"] = "pair contains non-finite values; entire component unavailable, no samples removed"
    elif definition.report_group == "torque_maps" and (
            (channels[x_name].unit or "").lower() != "rpm" or (channels[y_name].unit or "").lower() != "nm"):
        record["reason"] = "torque map requires rpm and Nm; no unit conversion assumed"
    elif definition.report_group == "torque_maps" and "countertorque" in channels[y_name].source_name.lower().replace("_", ""):
        record["reason"] = "counter-torque is not an approved delivered-torque measurement"
    elif definition.report_group == "torque_maps" and np.all(x == 0) and np.all(y == 0):
        record.update(status="inactive", reason="both RPM and torque are exactly zero throughout the selected interval")
    else:
        record.update(included=True, status="constant_nonzero" if np.all(x == x[0]) and np.all(y == y[0]) else "varying",
                      reason="all original paired samples retained")
    return record


def _write_profile_plot_metadata(result: ProfilePlottingResult, output_dir: Path) -> None:
    catalogue_path = output_dir / "profile_plot_catalogue.csv"
    manifest_path = output_dir / "profile_plot_manifest.json"
    with catalogue_path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=[
                "plot_order",
                "plot_id",
                "title",
                "status",
                "rendered",
                "x",
                "primary_series",
                "secondary_series",
                "png_file",
                "unavailable_reason",
            ],
        )
        writer.writeheader()
        unavailable_by_id = {item.definition.plot_id: item for item in result.unavailable_plots}
        rendered_by_id = {item.plot_id: item for item in result.rendered_plots}
        for index, definition in enumerate(sorted(result.profile.plots, key=lambda item: item.order or 999_999), 1):
            rendered = rendered_by_id.get(definition.plot_id)
            unavailable = unavailable_by_id.get(definition.plot_id)
            writer.writerow(
                {
                    "plot_order": index,
                    "plot_id": definition.plot_id,
                    "title": definition.title,
                    "status": definition.status,
                    "rendered": rendered is not None,
                    "x": definition.x,
                    "primary_series": ";".join(rendered.primary_series_ids) if rendered else "",
                    "secondary_series": ";".join(rendered.secondary_series_ids) if rendered else "",
                    "png_file": rendered.png_file if rendered else "",
                    "unavailable_reason": unavailable.reason if unavailable else "",
                }
            )

    manifest: dict[str, Any] = {
        "configuration_version": result.profile.version,
        "profile_id": result.profile.profile_id,
        "source_file": str(result.dataset.source_path),
        "source_sha256": result.dataset.quality.source_sha256,
        "sample_count": result.sample_count,
        "source_data_start_row": result.dataset.quality.data_start_row,
        "source_data_end_row": result.dataset.quality.data_end_row,
        "scenario": result.dataset.source_path.name,
        "paired_series": result.series_diagnostics,
        "series_sources": {key: [asdict(item) for item in items] for key, items in result.series_summaries.items()},
        "configured_plot_count": result.configured_plot_count,
        "rendered_plot_count": result.rendered_plot_count,
        "unavailable_plot_count": result.unavailable_plot_count,
        "profile_source_known": False,
        "plots": [item.to_dict() for item in result.rendered_plots],
        "unavailable_plots": [
            {
                "plot_id": item.definition.plot_id,
                "title": item.definition.title,
                "reason": item.reason,
                "missing_semantic_names": list(item.missing_semantic_names),
                "omitted_optional_series": list(item.omitted_optional_series),
            }
            for item in result.unavailable_plots
        ],
    }
    source_path = Path(result.profile.metadata.profile_id)
    if source_path.exists():
        manifest["profile_sha256"] = sha256_file(source_path)
        manifest["profile_source_known"] = True
    manifest_path.write_text(json.dumps(manifest, indent=2), encoding="utf-8")
