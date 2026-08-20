from __future__ import annotations

import math
from dataclasses import dataclass, field
from typing import Any, Mapping

import numpy as np

from .battery import nominal_capacity_estimates_kwh
from .errors import MathChannelError, StatisticsError
from .math_engine import _compile_expression, _evaluate_node, _topological_order
from .models import ChannelInfo, ImportedDataset
from .profile_math import ProfileMathResult, calculate_profile_math_channels
from .report_profile import (
    KPIDefinition,
    ProfileResolutionResult,
    ReportingProfile,
    StatisticDefinition,
    resolve_profile,
)
from .statistics_engine import compute_statistic


@dataclass(frozen=True)
class ProfileStatisticResult:
    definition: StatisticDefinition
    target_channel: str
    channel_display_name: str
    channel_unit: str | None
    channel_kind: str
    value: float
    sample_count: int
    used_sample_count: int
    omitted_sample_count: int


@dataclass(frozen=True)
class ProfileKPIResult:
    definition: KPIDefinition
    value: float


@dataclass(frozen=True)
class UnavailableProfileStatistic:
    definition: StatisticDefinition
    reason: str


@dataclass(frozen=True)
class UnavailableProfileKPI:
    definition: KPIDefinition
    missing_dependencies: tuple[str, ...]
    unavailable_dependencies: tuple[str, ...] = ()


@dataclass
class ProfileStatisticsResult:
    dataset: ImportedDataset
    profile: ReportingProfile
    resolution: ProfileResolutionResult
    math_result: ProfileMathResult
    statistics: list[ProfileStatisticResult]
    kpis: list[ProfileKPIResult]
    unavailable_statistics: list[UnavailableProfileStatistic] = field(default_factory=list)
    unavailable_kpis: list[UnavailableProfileKPI] = field(default_factory=list)
    kpi_calculation_order: list[str] = field(default_factory=list)
    diagnostics: list[str] = field(default_factory=list)

    @property
    def configured_statistic_count(self) -> int:
        return len(self.profile.statistics)

    @property
    def configured_kpi_count(self) -> int:
        return len(self.profile.kpis)

    @property
    def calculated_statistic_count(self) -> int:
        return len(self.statistics)

    @property
    def calculated_kpi_count(self) -> int:
        return len(self.kpis)

    @property
    def unavailable_required_statistics(self) -> list[UnavailableProfileStatistic]:
        return [item for item in self.unavailable_statistics if item.definition.required]

    @property
    def unavailable_optional_statistics(self) -> list[UnavailableProfileStatistic]:
        return [item for item in self.unavailable_statistics if not item.definition.required]

    @property
    def unavailable_required_kpis(self) -> list[UnavailableProfileKPI]:
        return [item for item in self.unavailable_kpis if item.definition.required]

    @property
    def unavailable_optional_kpis(self) -> list[UnavailableProfileKPI]:
        return [item for item in self.unavailable_kpis if not item.definition.required]

    @property
    def is_complete(self) -> bool:
        return not self.unavailable_required_statistics and not self.unavailable_required_kpis


def calculate_profile_statistics(
    dataset: ImportedDataset,
    profile: ReportingProfile,
    resolution: ProfileResolutionResult | None = None,
    math_result: ProfileMathResult | None = None,
    constants: Mapping[str, float] | None = None,
) -> ProfileStatisticsResult:
    if resolution is None:
        resolution = resolve_profile(dataset, profile)
    if math_result is None:
        math_result = calculate_profile_math_channels(dataset, profile, resolution)

    values_by_name, channels_by_name = _build_semantic_channel_maps(dataset, resolution, math_result)
    time_values = values_by_name.get("track_time")
    statistics: list[ProfileStatisticResult] = []
    unavailable_statistics: list[UnavailableProfileStatistic] = []

    for definition in profile.statistics:
        if definition.target not in values_by_name:
            unavailable_statistics.append(
                UnavailableProfileStatistic(definition=definition, reason=f"target unavailable: {definition.target}")
            )
            continue
        channel = channels_by_name[definition.target]
        value, used_count, omitted_count = compute_statistic(
            values_by_name[definition.target],
            definition.operation,
            "error",
            time_values=time_values,
        )
        statistics.append(
            ProfileStatisticResult(
                definition=definition,
                target_channel=definition.target,
                channel_display_name=definition.display_name or channel.display_name,
                channel_unit=definition.unit or channel.unit,
                channel_kind=channel.kind,
                value=value,
                sample_count=dataset.quality.sample_count,
                used_sample_count=used_count,
                omitted_sample_count=omitted_count,
            )
        )

    diagnostics = _profile_diagnostics(values_by_name)
    kpis, unavailable_kpis, kpi_order = _calculate_kpis(profile, statistics, constants or {}, values_by_name)

    return ProfileStatisticsResult(
        dataset=dataset,
        profile=profile,
        resolution=resolution,
        math_result=math_result,
        statistics=statistics,
        kpis=kpis,
        unavailable_statistics=unavailable_statistics,
        unavailable_kpis=unavailable_kpis,
        kpi_calculation_order=kpi_order,
        diagnostics=diagnostics,
    )


def _build_semantic_channel_maps(
    dataset: ImportedDataset,
    resolution: ProfileResolutionResult,
    math_result: ProfileMathResult,
) -> tuple[dict[str, np.ndarray], dict[str, ChannelInfo]]:
    values: dict[str, np.ndarray] = {}
    channels: dict[str, ChannelInfo] = {}
    for semantic_name, resolved in resolution.resolved.items():
        index = dataset.channel_index(resolved.channel.channel_id)
        values[semantic_name] = np.asarray(dataset.values[:, index], dtype=np.float64)
        channels[semantic_name] = resolved.channel
    calculated_by_name = {channel.channel_id.rsplit("__math__", 1)[-1]: channel for channel in math_result.calculated_channels}
    for semantic_name, series in math_result.values_by_semantic_name.items():
        values[semantic_name] = series
        channels[semantic_name] = calculated_by_name[semantic_name]
    return values, channels


def _calculate_kpis(
    profile: ReportingProfile,
    statistics: list[ProfileStatisticResult],
    constants: Mapping[str, float],
    channel_values: Mapping[str, np.ndarray] | None = None,
) -> tuple[list[ProfileKPIResult], list[UnavailableProfileKPI], list[str]]:
    definitions_by_id = {definition.kpi_id: definition for definition in profile.kpis}
    statistic_values = {item.definition.statistic_id: item.value for item in statistics}
    channel_values = channel_values or {}
    allowed_names = set(statistic_values) | set(definitions_by_id) | set(constants) | set(channel_values)
    compiled = {}
    unknown: dict[str, list[str]] = {}

    for definition in profile.kpis:
        expression = _compile_expression(definition.expression, context=f"profile KPI '{definition.kpi_id}'")
        compiled[definition.kpi_id] = expression
        missing = sorted(set(expression.dependencies) - allowed_names)
        if missing:
            unknown[definition.kpi_id] = missing

    if unknown:
        required_unknown = [
            f"{kpi_id}: {', '.join(missing)}"
            for kpi_id, missing in unknown.items()
            if definitions_by_id[kpi_id].required
        ]
        if required_unknown:
            raise StatisticsError("Profile KPIs reference unknown required dependencies: " + "; ".join(required_unknown))

    order = _topological_order([definition.kpi_id for definition in profile.kpis], compiled)
    context: dict[str, Any] = {**channel_values, **statistic_values, **constants}
    values: dict[str, float] = {}
    unavailable_by_id: dict[str, UnavailableProfileKPI] = {}
    executed_order: list[str] = []

    for kpi_id in order:
        definition = definitions_by_id[kpi_id]
        expression = compiled[kpi_id]
        missing_dependencies = tuple(
            dependency
            for dependency in expression.dependencies
            if dependency not in context and dependency not in definitions_by_id
        )
        unavailable_dependencies = tuple(
            dependency for dependency in expression.dependencies if dependency in unavailable_by_id
        )
        if missing_dependencies or unavailable_dependencies:
            unavailable_by_id[kpi_id] = UnavailableProfileKPI(
                definition=definition,
                missing_dependencies=missing_dependencies,
                unavailable_dependencies=unavailable_dependencies,
            )
            continue
        try:
            with np.errstate(all="ignore"):
                raw_value = _evaluate_node(expression.tree.body, context)
        except (MathChannelError, ValueError) as exc:
            raise StatisticsError(f"Profile KPI '{kpi_id}' could not be evaluated: {exc}") from exc
        array = np.asarray(raw_value, dtype=np.float64)
        if array.ndim != 0:
            raise StatisticsError(f"Profile KPI '{kpi_id}' produced shape {array.shape}; expected a scalar")
        value = float(array)
        if not math.isfinite(value):
            raise StatisticsError(f"Profile KPI '{kpi_id}' produced a non-finite value")
        values[kpi_id] = value
        context[kpi_id] = value
        executed_order.append(kpi_id)

    for kpi_id, missing in unknown.items():
        if kpi_id not in unavailable_by_id:
            unavailable_by_id[kpi_id] = UnavailableProfileKPI(
                definition=definitions_by_id[kpi_id],
                missing_dependencies=tuple(missing),
            )

    results = [ProfileKPIResult(definition=definition, value=values[definition.kpi_id]) for definition in profile.kpis if definition.kpi_id in values]
    unavailable = [unavailable_by_id[definition.kpi_id] for definition in profile.kpis if definition.kpi_id in unavailable_by_id]
    return results, unavailable, executed_order


def _profile_diagnostics(values_by_name: Mapping[str, np.ndarray]) -> list[str]:
    energy = values_by_name.get("electricsystem_battery_energy")
    soc = values_by_name.get("electricsystem_battery_soc")
    if energy is None or soc is None:
        return []
    try:
        estimates = nominal_capacity_estimates_kwh(energy, soc)
    except ValueError as exc:
        return [f"Nominal battery capacity could not be inferred from Battery Energy/SOC: {exc}"]
    median = float(np.median(estimates))
    max_deviation = float(np.max(np.abs(estimates - median)))
    tolerance = max(0.1, abs(median) * 0.01)
    if max_deviation <= tolerance:
        return []
    return [
        "Nominal battery capacity estimates vary across Battery Energy/SOC samples: "
        f"median={median:.6g} kWh, max deviation={max_deviation:.6g} kWh, tolerance={tolerance:.6g} kWh"
    ]
