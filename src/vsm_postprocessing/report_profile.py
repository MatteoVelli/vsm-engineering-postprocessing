from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable, Mapping, Sequence

import numpy as np
import yaml

from .errors import ConfigurationError
from .models import ChannelInfo, ImportedDataset

_IDENTIFIER = re.compile(r"^[a-z][a-z0-9_]*$")
_NORMALIZE_PATTERN = re.compile(r"[^a-z0-9]+")


@dataclass(frozen=True)
class ProfileMetadata:
    profile_id: str
    name: str
    description: str | None = None
    powertrain: str | None = None
    extends: str | None = None


@dataclass(frozen=True)
class RawChannelDefinition:
    semantic_name: str
    source_name: str
    report_name: str
    channel_type: str
    unit: str | None = None
    aliases: tuple[str, ...] = ()
    required: bool = True
    for_plot: bool = False
    notes: str | None = None


@dataclass(frozen=True)
class MathChannelDefinition:
    semantic_name: str
    source_name: str
    report_name: str
    unit: str | None = None
    dependencies: tuple[str, ...] = ()
    expression: str | None = None
    formula: str | None = None
    required: bool = True
    for_plot: bool = False
    notes: str | None = None

    @property
    def channel_type(self) -> str:
        return "MATH"


@dataclass(frozen=True)
class StatisticDefinition:
    statistic_id: str
    target: str
    operation: str
    display_name: str | None = None
    unit: str | None = None
    placement_group: str | None = None
    required: bool = True
    notes: str | None = None


@dataclass(frozen=True)
class KPIDefinition:
    kpi_id: str
    expression: str
    dependencies: tuple[str, ...] = ()
    display_name: str | None = None
    unit: str | None = None
    placement_group: str | None = None
    required: bool = True
    notes: str | None = None


@dataclass(frozen=True)
class ProfilePlotSeriesDefinition:
    semantic_name: str
    axis: str = "primary"
    label: str | None = None
    required: bool = True


@dataclass(frozen=True)
class ProfilePlotDefinition:
    plot_id: str
    title: str
    x: str
    series: tuple[ProfilePlotSeriesDefinition, ...]
    output_filename: str | None = None
    x_label: str | None = None
    primary_y_label: str | None = None
    secondary_y_label: str | None = None
    reference_chart_number: int | None = None
    order: int | None = None
    status: str = "PASS"
    evidence: str | None = None
    notes: str | None = None


@dataclass(frozen=True)
class ReportingProfile:
    version: int
    metadata: ProfileMetadata
    raw_channels: tuple[RawChannelDefinition, ...]
    math_channels: tuple[MathChannelDefinition, ...]
    statistics: tuple[StatisticDefinition, ...] = ()
    kpis: tuple[KPIDefinition, ...] = ()
    plots: tuple[ProfilePlotDefinition, ...] = ()

    @property
    def profile_id(self) -> str:
        return self.metadata.profile_id

    def raw_by_semantic_name(self) -> dict[str, RawChannelDefinition]:
        return {channel.semantic_name: channel for channel in self.raw_channels}

    def math_by_semantic_name(self) -> dict[str, MathChannelDefinition]:
        return {channel.semantic_name: channel for channel in self.math_channels}

    def statistics_by_id(self) -> dict[str, StatisticDefinition]:
        return {definition.statistic_id: definition for definition in self.statistics}

    def kpis_by_id(self) -> dict[str, KPIDefinition]:
        return {definition.kpi_id: definition for definition in self.kpis}

    def plots_by_id(self) -> dict[str, ProfilePlotDefinition]:
        return {definition.plot_id: definition for definition in self.plots}


@dataclass(frozen=True)
class ResolvedChannel:
    definition: RawChannelDefinition
    channel: ChannelInfo
    match_type: str
    is_active: bool
    is_constant: bool
    is_all_zero: bool


@dataclass(frozen=True)
class UnresolvedChannel:
    definition: RawChannelDefinition
    reason: str


@dataclass(frozen=True)
class AmbiguousChannel:
    definition: RawChannelDefinition
    match_type: str
    candidates: tuple[ChannelInfo, ...]


@dataclass(frozen=True)
class UnitMismatch:
    definition: RawChannelDefinition
    channel: ChannelInfo
    expected_unit: str
    actual_unit: str
    match_type: str


@dataclass(frozen=True)
class MathDependencyResolution:
    definition: MathChannelDefinition
    resolved_dependencies: Mapping[str, str]
    math_dependencies: tuple[str, ...]
    missing_dependencies: tuple[str, ...]

    @property
    def is_resolved(self) -> bool:
        return not self.missing_dependencies


@dataclass
class ProfileResolutionResult:
    profile: ReportingProfile
    resolved: dict[str, ResolvedChannel] = field(default_factory=dict)
    missing_required: list[UnresolvedChannel] = field(default_factory=list)
    missing_optional: list[UnresolvedChannel] = field(default_factory=list)
    ambiguous: list[AmbiguousChannel] = field(default_factory=list)
    unit_mismatches: list[UnitMismatch] = field(default_factory=list)
    math_dependencies: list[MathDependencyResolution] = field(default_factory=list)

    @property
    def is_valid(self) -> bool:
        return not self.missing_required and not self.ambiguous and not self.unit_mismatches

    @property
    def resolved_channel_ids(self) -> dict[str, str]:
        return {semantic_name: item.channel.channel_id for semantic_name, item in self.resolved.items()}

    def summary_lines(self) -> list[str]:
        lines = [
            f"Profile: {self.profile.profile_id} ({self.profile.metadata.name})",
            f"Resolved raw channels: {len(self.resolved)}/{len(self.profile.raw_channels)}",
            f"Missing required: {len(self.missing_required)}",
            f"Missing optional: {len(self.missing_optional)}",
            f"Ambiguous: {len(self.ambiguous)}",
            f"Unit mismatches: {len(self.unit_mismatches)}",
            f"Math channels: {len(self.profile.math_channels)}",
        ]
        unresolved_math = [item for item in self.math_dependencies if not item.is_resolved]
        lines.append(f"Math dependency issues: {len(unresolved_math)}")
        return lines


def load_reporting_profile(path: str | Path) -> ReportingProfile:
    """Load a strict semantic reporting profile from YAML.

    Profiles can extend one parent profile by path relative to the child profile.
    Child raw/math channels are appended and may not redefine existing semantic names.
    """

    return _load_reporting_profile(Path(path).expanduser().resolve(), seen=())


def resolve_profile(dataset: ImportedDataset, profile: ReportingProfile) -> ProfileResolutionResult:
    """Resolve profile raw channels and math dependencies against an imported dataset."""

    exact_index: dict[str, list[ChannelInfo]] = {}
    normalized_index: dict[str, list[ChannelInfo]] = {}
    for channel in dataset.channels:
        exact_index.setdefault(channel.source_name, []).append(channel)
        normalized_index.setdefault(normalize_name(channel.source_name), []).append(channel)

    result = ProfileResolutionResult(profile=profile)
    for definition in profile.raw_channels:
        match = _resolve_raw_definition(definition, exact_index, normalized_index)
        if isinstance(match, tuple):
            channel, match_type = match
            mismatch = _unit_mismatch(definition.unit, channel.unit)
            if mismatch is not None:
                result.unit_mismatches.append(
                    UnitMismatch(
                        definition=definition,
                        channel=channel,
                        expected_unit=definition.unit or "",
                        actual_unit=channel.unit or "",
                        match_type=match_type,
                    )
                )
                continue
            result.resolved[definition.semantic_name] = ResolvedChannel(
                definition=definition,
                channel=channel,
                match_type=match_type,
                **_channel_activity(dataset, channel),
            )
        elif isinstance(match, AmbiguousChannel):
            result.ambiguous.append(match)
        else:
            unresolved = UnresolvedChannel(definition=definition, reason=match)
            if definition.required:
                result.missing_required.append(unresolved)
            else:
                result.missing_optional.append(unresolved)

    result.math_dependencies.extend(resolve_math_dependencies(profile, result))
    return result


def resolve_math_dependencies(
    profile: ReportingProfile, resolution: ProfileResolutionResult
) -> list[MathDependencyResolution]:
    """Resolve math dependencies by semantic name without calculating values."""

    raw_names = set(profile.raw_by_semantic_name())
    math_names = set(profile.math_by_semantic_name())
    resolved_raw = resolution.resolved_channel_ids
    items: list[MathDependencyResolution] = []
    for definition in profile.math_channels:
        resolved_dependencies: dict[str, str] = {}
        math_dependencies: list[str] = []
        missing_dependencies: list[str] = []
        for dependency in definition.dependencies:
            if dependency in resolved_raw:
                resolved_dependencies[dependency] = resolved_raw[dependency]
            elif dependency in math_names:
                math_dependencies.append(dependency)
            elif dependency in raw_names:
                missing_dependencies.append(dependency)
            else:
                missing_dependencies.append(dependency)
        items.append(
            MathDependencyResolution(
                definition=definition,
                resolved_dependencies=resolved_dependencies,
                math_dependencies=tuple(math_dependencies),
                missing_dependencies=tuple(missing_dependencies),
            )
        )
    return items


def normalize_name(value: str) -> str:
    return _NORMALIZE_PATTERN.sub("", value.strip().lower())


def _load_reporting_profile(path: Path, seen: tuple[Path, ...]) -> ReportingProfile:
    if path in seen:
        cycle = " -> ".join(str(item) for item in (*seen, path))
        raise ConfigurationError(f"Reporting profile inheritance cycle: {cycle}")
    if not path.exists():
        raise ConfigurationError(f"Reporting profile does not exist: {path}")
    if not path.is_file():
        raise ConfigurationError(f"Reporting profile is not a file: {path}")

    try:
        raw = yaml.safe_load(path.read_text(encoding="utf-8"))
    except UnicodeDecodeError as exc:
        raise ConfigurationError(f"Reporting profile must be UTF-8: {path}") from exc
    except yaml.YAMLError as exc:
        raise ConfigurationError(f"Invalid YAML in reporting profile '{path}': {exc}") from exc

    if not isinstance(raw, dict):
        raise ConfigurationError("Reporting profile root must be a YAML mapping")
    _reject_unknown_keys(raw, {"version", "profile", "channels", "statistics", "kpis", "plots"}, "root")
    version = raw.get("version")
    if version != 1:
        raise ConfigurationError("Reporting profile 'version' must be 1")

    metadata = _parse_metadata(raw.get("profile"))
    channels = raw.get("channels", {})
    if not isinstance(channels, dict):
        raise ConfigurationError("channels must be a YAML mapping")
    _reject_unknown_keys(channels, {"raw", "math"}, "channels")

    raw_channels = tuple(_parse_raw_channel(item, index) for index, item in enumerate(channels.get("raw", []), 1))
    math_channels = tuple(
        _parse_math_channel(item, index) for index, item in enumerate(channels.get("math", []), 1)
    )
    statistics = tuple(
        _parse_statistic(item, index) for index, item in enumerate(raw.get("statistics", []), 1)
    )
    kpis = tuple(_parse_kpi(item, index) for index, item in enumerate(raw.get("kpis", []), 1))
    plots = tuple(_parse_plot(item, index) for index, item in enumerate(raw.get("plots", []), 1))

    profile = ReportingProfile(
        version=version,
        metadata=metadata,
        raw_channels=raw_channels,
        math_channels=math_channels,
        statistics=statistics,
        kpis=kpis,
        plots=plots,
    )
    _validate_unique_semantic_names(profile)

    if metadata.extends is None:
        return profile

    parent_path = (path.parent / metadata.extends).resolve()
    parent = _load_reporting_profile(parent_path, seen=(*seen, path))
    return _merge_profiles(parent, profile)


def _parse_metadata(raw: Any) -> ProfileMetadata:
    if not isinstance(raw, dict):
        raise ConfigurationError("profile must be a YAML mapping")
    _reject_unknown_keys(raw, {"profile_id", "name", "description", "powertrain", "extends"}, "profile")
    profile_id = _required_string(raw, "profile.profile_id")
    _validate_identifier(profile_id, "profile.profile_id")
    name = _required_string(raw, "profile.name")
    return ProfileMetadata(
        profile_id=profile_id,
        name=name,
        description=_optional_string(raw.get("description"), "profile.description"),
        powertrain=_optional_string(raw.get("powertrain"), "profile.powertrain"),
        extends=_optional_string(raw.get("extends"), "profile.extends"),
    )


def _parse_raw_channel(raw: Any, index: int) -> RawChannelDefinition:
    context = f"channels.raw[{index}]"
    if not isinstance(raw, dict):
        raise ConfigurationError(f"{context} must be a YAML mapping")
    _reject_unknown_keys(
        raw,
        {
            "semantic_name",
            "source_name",
            "report_name",
            "channel_type",
            "unit",
            "aliases",
            "required",
            "for_plot",
            "notes",
        },
        context,
    )
    semantic_name = _required_string(raw, f"{context}.semantic_name")
    _validate_identifier(semantic_name, f"{context}.semantic_name")
    channel_type = _required_string(raw, f"{context}.channel_type")
    if channel_type not in {"VSM", "AVL"}:
        raise ConfigurationError(f"{context}.channel_type must be VSM or AVL")
    return RawChannelDefinition(
        semantic_name=semantic_name,
        source_name=_required_string(raw, f"{context}.source_name"),
        report_name=_required_string(raw, f"{context}.report_name"),
        channel_type=channel_type,
        unit=_optional_string(raw.get("unit"), f"{context}.unit"),
        aliases=_parse_string_tuple(raw.get("aliases", []), f"{context}.aliases"),
        required=_optional_bool(raw.get("required", True), f"{context}.required"),
        for_plot=_optional_bool(raw.get("for_plot", False), f"{context}.for_plot"),
        notes=_optional_string(raw.get("notes"), f"{context}.notes"),
    )


def _parse_math_channel(raw: Any, index: int) -> MathChannelDefinition:
    context = f"channels.math[{index}]"
    if not isinstance(raw, dict):
        raise ConfigurationError(f"{context} must be a YAML mapping")
    _reject_unknown_keys(
        raw,
        {
            "semantic_name",
            "source_name",
            "report_name",
            "unit",
            "dependencies",
            "expression",
            "formula",
            "required",
            "for_plot",
            "notes",
        },
        context,
    )
    semantic_name = _required_string(raw, f"{context}.semantic_name")
    _validate_identifier(semantic_name, f"{context}.semantic_name")
    return MathChannelDefinition(
        semantic_name=semantic_name,
        source_name=_required_string(raw, f"{context}.source_name"),
        report_name=_required_string(raw, f"{context}.report_name"),
        unit=_optional_string(raw.get("unit"), f"{context}.unit"),
        dependencies=_parse_string_tuple(raw.get("dependencies", []), f"{context}.dependencies"),
        expression=_optional_string(raw.get("expression"), f"{context}.expression"),
        formula=_optional_string(raw.get("formula"), f"{context}.formula"),
        required=_optional_bool(raw.get("required", True), f"{context}.required"),
        for_plot=_optional_bool(raw.get("for_plot", False), f"{context}.for_plot"),
        notes=_optional_string(raw.get("notes"), f"{context}.notes"),
    )


def _parse_statistic(raw: Any, index: int) -> StatisticDefinition:
    context = f"statistics[{index}]"
    if not isinstance(raw, dict):
        raise ConfigurationError(f"{context} must be a YAML mapping")
    _reject_unknown_keys(
        raw,
        {
            "statistic_id",
            "target",
            "operation",
            "display_name",
            "unit",
            "placement_group",
            "required",
            "notes",
        },
        context,
    )
    statistic_id = _required_string(raw, f"{context}.statistic_id")
    _validate_identifier(statistic_id, f"{context}.statistic_id")
    target = _required_string(raw, f"{context}.target")
    _validate_identifier(target, f"{context}.target")
    operation = _required_string(raw, f"{context}.operation")
    if operation not in {"rms", "time_weighted_rms", "max", "min", "first", "last", "sum"}:
        raise ConfigurationError(f"{context}.operation must be a supported statistics operation")
    return StatisticDefinition(
        statistic_id=statistic_id,
        target=target,
        operation=operation,
        display_name=_optional_string(raw.get("display_name"), f"{context}.display_name"),
        unit=_optional_string(raw.get("unit"), f"{context}.unit"),
        placement_group=_optional_string(raw.get("placement_group"), f"{context}.placement_group"),
        required=_optional_bool(raw.get("required", True), f"{context}.required"),
        notes=_optional_string(raw.get("notes"), f"{context}.notes"),
    )


def _parse_kpi(raw: Any, index: int) -> KPIDefinition:
    context = f"kpis[{index}]"
    if not isinstance(raw, dict):
        raise ConfigurationError(f"{context} must be a YAML mapping")
    _reject_unknown_keys(
        raw,
        {
            "kpi_id",
            "expression",
            "dependencies",
            "display_name",
            "unit",
            "placement_group",
            "required",
            "notes",
        },
        context,
    )
    kpi_id = _required_string(raw, f"{context}.kpi_id")
    _validate_identifier(kpi_id, f"{context}.kpi_id")
    return KPIDefinition(
        kpi_id=kpi_id,
        expression=_required_string(raw, f"{context}.expression"),
        dependencies=_parse_string_tuple(raw.get("dependencies", []), f"{context}.dependencies"),
        display_name=_optional_string(raw.get("display_name"), f"{context}.display_name"),
        unit=_optional_string(raw.get("unit"), f"{context}.unit"),
        placement_group=_optional_string(raw.get("placement_group"), f"{context}.placement_group"),
        required=_optional_bool(raw.get("required", True), f"{context}.required"),
        notes=_optional_string(raw.get("notes"), f"{context}.notes"),
    )


def _parse_plot(raw: Any, index: int) -> ProfilePlotDefinition:
    context = f"plots[{index}]"
    if not isinstance(raw, dict):
        raise ConfigurationError(f"{context} must be a YAML mapping")
    _reject_unknown_keys(
        raw,
        {
            "plot_id",
            "title",
            "x",
            "x_label",
            "primary_y_label",
            "secondary_y_label",
            "output_filename",
            "reference_chart_number",
            "order",
            "status",
            "evidence",
            "notes",
            "series",
        },
        context,
    )
    plot_id = _required_string(raw, f"{context}.plot_id")
    _validate_identifier(plot_id, f"{context}.plot_id")
    x = _required_string(raw, f"{context}.x")
    _validate_identifier(x, f"{context}.x")
    series_raw = raw.get("series")
    if not isinstance(series_raw, list) or not series_raw:
        raise ConfigurationError(f"{context}.series must be a non-empty YAML list")
    series = tuple(_parse_plot_series(item, context, item_index) for item_index, item in enumerate(series_raw, 1))
    output_filename = _optional_string(raw.get("output_filename"), f"{context}.output_filename")
    if output_filename is not None and (Path(output_filename).name != output_filename or not output_filename.endswith(".png")):
        raise ConfigurationError(f"{context}.output_filename must be a plain .png filename without directories")
    reference_chart_number = _optional_positive_int(raw.get("reference_chart_number"), f"{context}.reference_chart_number")
    order = _optional_positive_int(raw.get("order"), f"{context}.order")
    status = _required_string({"status": raw.get("status", "PASS")}, f"{context}.status")
    if status not in {"PASS", "RECONSTRUCTED", "REVIEW", "UNAVAILABLE", "INACTIVE"}:
        raise ConfigurationError(f"{context}.status must be PASS, RECONSTRUCTED, REVIEW, UNAVAILABLE, or INACTIVE")
    return ProfilePlotDefinition(
        plot_id=plot_id,
        title=_required_string(raw, f"{context}.title"),
        x=x,
        series=series,
        output_filename=output_filename,
        x_label=_optional_string(raw.get("x_label"), f"{context}.x_label"),
        primary_y_label=_optional_string(raw.get("primary_y_label"), f"{context}.primary_y_label"),
        secondary_y_label=_optional_string(raw.get("secondary_y_label"), f"{context}.secondary_y_label"),
        reference_chart_number=reference_chart_number,
        order=order,
        status=status,
        evidence=_optional_string(raw.get("evidence"), f"{context}.evidence"),
        notes=_optional_string(raw.get("notes"), f"{context}.notes"),
    )


def _parse_plot_series(raw: Any, plot_context: str, index: int) -> ProfilePlotSeriesDefinition:
    context = f"{plot_context}.series[{index}]"
    if not isinstance(raw, dict):
        raise ConfigurationError(f"{context} must be a YAML mapping")
    _reject_unknown_keys(raw, {"semantic_name", "axis", "label", "required"}, context)
    semantic_name = _required_string(raw, f"{context}.semantic_name")
    _validate_identifier(semantic_name, f"{context}.semantic_name")
    axis = _required_string({"axis": raw.get("axis", "primary")}, f"{context}.axis")
    if axis not in {"primary", "secondary"}:
        raise ConfigurationError(f"{context}.axis must be primary or secondary")
    return ProfilePlotSeriesDefinition(
        semantic_name=semantic_name,
        axis=axis,
        label=_optional_string(raw.get("label"), f"{context}.label"),
        required=_optional_bool(raw.get("required", True), f"{context}.required"),
    )


def _merge_profiles(parent: ReportingProfile, child: ReportingProfile) -> ReportingProfile:
    raw_names = {channel.semantic_name for channel in parent.raw_channels}
    raw_collisions = [channel.semantic_name for channel in child.raw_channels if channel.semantic_name in raw_names]
    if raw_collisions:
        raise ConfigurationError("Child profile redefines raw channels: " + ", ".join(raw_collisions))

    math_names = {channel.semantic_name for channel in parent.math_channels}
    math_collisions = [channel.semantic_name for channel in child.math_channels if channel.semantic_name in math_names]
    if math_collisions:
        raise ConfigurationError("Child profile redefines math channels: " + ", ".join(math_collisions))

    statistic_ids = {definition.statistic_id for definition in parent.statistics}
    statistic_collisions = [
        definition.statistic_id for definition in child.statistics if definition.statistic_id in statistic_ids
    ]
    if statistic_collisions:
        raise ConfigurationError("Child profile redefines statistics: " + ", ".join(statistic_collisions))

    kpi_ids = {definition.kpi_id for definition in parent.kpis}
    kpi_collisions = [definition.kpi_id for definition in child.kpis if definition.kpi_id in kpi_ids]
    if kpi_collisions:
        raise ConfigurationError("Child profile redefines KPIs: " + ", ".join(kpi_collisions))

    plot_ids = {definition.plot_id for definition in parent.plots}
    plot_collisions = [definition.plot_id for definition in child.plots if definition.plot_id in plot_ids]
    if plot_collisions:
        raise ConfigurationError("Child profile redefines plots: " + ", ".join(plot_collisions))

    return ReportingProfile(
        version=child.version,
        metadata=child.metadata,
        raw_channels=(*parent.raw_channels, *child.raw_channels),
        math_channels=(*parent.math_channels, *child.math_channels),
        statistics=(*parent.statistics, *child.statistics),
        kpis=(*parent.kpis, *child.kpis),
        plots=(*parent.plots, *child.plots),
    )


def _resolve_raw_definition(
    definition: RawChannelDefinition,
    exact_index: Mapping[str, Sequence[ChannelInfo]],
    normalized_index: Mapping[str, Sequence[ChannelInfo]],
) -> tuple[ChannelInfo, str] | AmbiguousChannel | str:
    stages = (
        ("exact", [definition.source_name], exact_index, lambda value: value),
        ("normalized", [definition.source_name], normalized_index, normalize_name),
        ("alias", definition.aliases, exact_index, lambda value: value),
        ("alias_normalized", definition.aliases, normalized_index, normalize_name),
    )
    for match_type, names, index, key_func in stages:
        for name in names:
            candidates = tuple(index.get(key_func(name), ()))
            if len(candidates) == 1:
                return candidates[0], match_type
            if len(candidates) > 1:
                return AmbiguousChannel(definition=definition, match_type=match_type, candidates=candidates)
    return "not found"


def _channel_activity(dataset: ImportedDataset, channel: ChannelInfo) -> dict[str, bool]:
    index = dataset.channel_index(channel.channel_id)
    values = np.asarray(dataset.values[:, index], dtype=np.float64)
    finite = values[np.isfinite(values)]
    if finite.size == 0:
        return {"is_active": False, "is_constant": False, "is_all_zero": False}
    is_constant = bool(np.allclose(finite, finite[0], rtol=0.0, atol=1e-12))
    is_all_zero = bool(np.allclose(finite, 0.0, rtol=0.0, atol=1e-12))
    return {
        "is_active": not is_constant,
        "is_constant": is_constant,
        "is_all_zero": is_all_zero,
    }


def _unit_mismatch(expected: str | None, actual: str | None) -> bool | None:
    if expected is None or not expected.strip():
        return None
    expected_normalized = _normalize_unit(expected)
    actual_normalized = _normalize_unit(actual or "")
    if expected_normalized != actual_normalized:
        return True
    return None


def _normalize_unit(unit: str) -> str:
    normalized = unit.strip().replace("Â°C", "°C").replace("deg C", "°C")
    return normalized.lower()


def _validate_unique_semantic_names(profile: ReportingProfile) -> None:
    _reject_duplicates(
        (channel.semantic_name for channel in profile.raw_channels),
        "channels.raw semantic names",
    )
    _reject_duplicates(
        (channel.semantic_name for channel in profile.math_channels),
        "channels.math semantic names",
    )
    collisions = sorted(set(profile.raw_by_semantic_name()) & set(profile.math_by_semantic_name()))
    if collisions:
        raise ConfigurationError("Raw and math semantic names collide: " + ", ".join(collisions))
    _reject_duplicates(
        (definition.statistic_id for definition in profile.statistics),
        "statistics IDs",
    )
    _reject_duplicates(
        (definition.kpi_id for definition in profile.kpis),
        "KPI IDs",
    )
    _reject_duplicates(
        (definition.plot_id for definition in profile.plots),
        "plot IDs",
    )
    statistic_kpi_collisions = sorted(set(profile.statistics_by_id()) & set(profile.kpis_by_id()))
    if statistic_kpi_collisions:
        raise ConfigurationError("Statistic and KPI IDs collide: " + ", ".join(statistic_kpi_collisions))


def _reject_duplicates(values: Iterable[str], context: str) -> None:
    seen: set[str] = set()
    duplicates: list[str] = []
    for value in values:
        if value in seen and value not in duplicates:
            duplicates.append(value)
        seen.add(value)
    if duplicates:
        raise ConfigurationError(f"{context} contain duplicate values: " + ", ".join(duplicates))


def _parse_string_tuple(raw: Any, context: str) -> tuple[str, ...]:
    if raw is None:
        return ()
    if not isinstance(raw, list):
        raise ConfigurationError(f"{context} must be a YAML list")
    values: list[str] = []
    for index, value in enumerate(raw, 1):
        if not isinstance(value, str) or not value.strip():
            raise ConfigurationError(f"{context}[{index}] must be a non-empty string")
        values.append(value.strip())
    _reject_duplicates(values, context)
    return tuple(values)


def _required_string(raw: Mapping[str, Any], key: str) -> str:
    field_name = key.rsplit(".", 1)[-1]
    value = raw.get(field_name)
    if not isinstance(value, str) or not value.strip():
        raise ConfigurationError(f"{key} must be a non-empty string")
    return value.strip()


def _optional_string(value: Any, context: str) -> str | None:
    if value is None:
        return None
    if not isinstance(value, str) or not value.strip():
        raise ConfigurationError(f"{context} must be null or a non-empty string")
    return value.strip()


def _optional_bool(value: Any, context: str) -> bool:
    if not isinstance(value, bool):
        raise ConfigurationError(f"{context} must be true or false")
    return value


def _optional_positive_int(value: Any, context: str) -> int | None:
    if value is None:
        return None
    if isinstance(value, bool) or not isinstance(value, int) or value < 1:
        raise ConfigurationError(f"{context} must be null or a positive integer")
    return value


def _validate_identifier(value: str, context: str) -> None:
    if not _IDENTIFIER.match(value):
        raise ConfigurationError(f"{context} must be a lower-case identifier")


def _reject_unknown_keys(mapping: Mapping[str, Any], allowed: set[str], context: str) -> None:
    unknown = sorted(set(mapping) - allowed)
    if unknown:
        raise ConfigurationError(f"Unknown key(s) in {context}: {', '.join(unknown)}")
