"""Presentation-only tyre/soil metrics using validated report channels and helpers."""
from __future__ import annotations

import math
from typing import TYPE_CHECKING

import numpy as np

from .errors import MathChannelError, StatisticsError
from .math_engine import _time_average
from .profile_statistics import CanonicalReportMetric
from .statistics_engine import compute_statistic

if TYPE_CHECKING:
    from .excel_report_engine import ProfileExcelReportResult


def soil_kpi_cards(slide_id: str) -> tuple[tuple[str, str, str], ...]:
    if slide_id not in {"tyre_total", "soil_bulldozing", "soil_compaction"}:
        return ()
    return tuple((label, f"{slide_id}_{suffix}", unit) for label, suffix, unit in (
        ("MAX POWER", "max_power", "kW"),
        ("MEAN POWER", "mean_power", "kW"),
        ("RMS POWER", "rms_power", "kW"),
        ("TOTAL ENERGY", "total_energy", "kWh"),
        ("ENERGY PER KM", "energy_per_km", "kWh/km"),
        ("DISTANCE", "distance", "km"),
    ))


def build_soil_powerpoint_metrics(
    result: ProfileExcelReportResult,
    canonical_metrics: list[CanonicalReportMetric],
) -> list[CanonicalReportMetric]:
    """Do not add statistics to the Excel profile or change any channel values."""
    presentation = result.profile.presentation
    if not presentation or result.profile.metadata.powertrain not in {"electric", "hybrid"}:
        return []
    slides = tuple(s for s in presentation.slides if soil_kpi_cards(s.slide_id) and s.statistics)
    if not slides:
        return []
    canonical = {m.metric_id: m for m in canonical_metrics}
    values = result.plotting_result.values_by_semantic_name
    channels = result.plotting_result.channels_by_semantic_name
    count = result.sample_count
    metrics = []

    def series(name: str, unit: str) -> np.ndarray:
        if name not in values or name not in channels:
            raise StatisticsError(f"Source channel unavailable: {name}")
        if (channels[name].unit or "").lower() != unit.lower():
            raise StatisticsError(f"Invalid unit for {name}: expected {unit}")
        data = np.asarray(values[name], dtype=np.float64)
        if data.ndim != 1 or data.size != count or not data.size or not np.isfinite(data).all():
            raise StatisticsError(f"Source must be finite, nonempty and sample-aligned: {name}")
        return data

    for slide in slides:
        cards = soil_kpi_cards(slide.slide_id)
        if not cards:
            continue
        mechanism = slide.slide_id.removeprefix("soil_")
        power = "tyre_total_resistance_power" if slide.slide_id == "tyre_total" else f"total_{mechanism}_resistance_power"
        energy = "tyre_total_resistance_energy_accumulated" if slide.slide_id == "tyre_total" else f"tyre_{mechanism}_resistance_energy_accumulated"
        distance = canonical.get("distance_km_last")

        def distance_value() -> float:
            if distance is None or not math.isfinite(distance.value):
                raise StatisticsError("Validated distance_km_last KPI unavailable")
            if (distance.unit or "").lower() != "km" or distance.value < 0:
                raise StatisticsError("Validated distance must be nonnegative and expressed in km")
            return distance.value

        def energy_change() -> float:
            data = series(energy, "kWh")
            return compute_statistic(data, "last")[0] - compute_statistic(data, "first")[0]

        definitions = (
            (power, "max", f"MAX({power})", lambda: compute_statistic(series(power, "kW"), "max")[0]),
            (power, "time_average", f"time_average({power}, track_time)",
             lambda: _time_average(series(power, "kW"), series("track_time", "s"))),
            (power, "rms", f"RMS({power}); sqrt(mean(P^2)), all samples",
             lambda: compute_statistic(series(power, "kW"), "rms")[0]),
            (energy, "net_change", f"LAST({energy}) - FIRST({energy})", energy_change),
            (energy, "energy_per_km", f"(LAST({energy}) - FIRST({energy})) / distance_km_last", None),
            ("distance_km", "canonical", "Existing validated distance_km_last KPI", distance_value),
        )
        for (label, metric_id, unit), (source, operation, calculation, evaluate) in zip(cards, definitions):
            if metric_id not in slide.statistics:
                continue
            notes = None
            try:
                if operation == "energy_per_km":
                    travelled = distance_value()
                    if travelled <= 0:
                        raise StatisticsError("ENERGY PER KM requires positive, nonzero distance_km_last")
                    value = energy_change() / travelled
                else:
                    value = evaluate()
                if not math.isfinite(value):
                    raise StatisticsError("Calculated metric is not finite")
            except (MathChannelError, StatisticsError) as exc:
                value = math.nan
                notes = f"Unavailable: {exc}; no value fabricated."
            channel = channels.get(source)
            metrics.append(CanonicalReportMetric(
                metric_id=metric_id, excel_label=label, powerpoint_label=label,
                value=value, unit=unit, source_channel=source,
                source_name=channel.display_name if channel else source,
                source_kind=channel.kind if channel else "KPI", statistic=operation,
                calculation=calculation, placement_group=slide.slide_id, notes=notes,
                sample_count=count, used_sample_count=count if notes is None else 0,
                omitted_sample_count=0,
            ))
    return metrics
