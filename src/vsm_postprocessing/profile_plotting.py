from __future__ import annotations

import csv
import json
from dataclasses import dataclass, replace
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
) -> ProfilePlottingResult:
    if resolution is None:
        resolution = resolve_profile(dataset, profile)
    if math_result is None:
        math_result = calculate_profile_math_channels(dataset, profile, resolution)

    channels_by_name, values_by_name = _build_semantic_maps(dataset, resolution, math_result)
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
        rendered.append(
            _render_one_plot(
                plotting_definition,
                defaults,
                destination,
                channels_by_name,
                values_by_name,
                dataset.quality.sample_count,
                (),
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
            resolved.channel,
            channel_id=semantic_name,
            display_name=resolved.definition.report_name,
            provenance=f"profile:{resolution.profile.profile_id}:{semantic_name}:{resolved.channel.channel_id}",
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
    if definition.x not in channels_by_name:
        return UnavailableProfilePlot(definition, f"x channel unavailable: {definition.x}", (definition.x,))
    x_values = values_by_name[definition.x]
    if x_values.size != sample_count:
        raise PlottingError(f"Profile plot '{definition.plot_id}' x channel '{definition.x}' length is not aligned")

    series: list[PlotSeriesDefinition] = []
    missing_required: list[str] = []
    missing_optional: list[str] = []
    series_summaries: list[ProfilePlotSeriesSummary] = []
    for item in definition.series:
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
            "no renderable series",
            (),
            tuple(missing_optional),
        )

    return (
        PlotDefinition(
            plot_id=definition.plot_id,
            title=definition.title,
            x_channel_id=definition.x,
            series=tuple(series),
            output_filename=definition.output_filename or f"{definition.plot_id}.png",
            x_label=definition.x_label,
            primary_y_label=definition.primary_y_label,
            secondary_y_label=definition.secondary_y_label,
            reference_chart_number=definition.reference_chart_number,
        ),
        tuple(series_summaries),
    )


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
