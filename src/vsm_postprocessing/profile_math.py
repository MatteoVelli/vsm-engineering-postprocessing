from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Mapping, Sequence

import numpy as np

from .errors import MathChannelError
from .math_engine import _compile_expression, _evaluate_node, _normalise_result, _topological_order
from .models import ChannelInfo, ImportedDataset
from .report_profile import MathChannelDefinition, ProfileResolutionResult, ReportingProfile, resolve_profile

DEFAULT_PROFILE_MATH_CONSTANTS: Mapping[str, float] = {
    "rpm_nm_to_kw_divisor": 9548.8,
}


@dataclass(frozen=True)
class UnavailableMathChannel:
    definition: MathChannelDefinition
    missing_dependencies: tuple[str, ...]
    unavailable_dependencies: tuple[str, ...] = ()


@dataclass
class ProfileMathResult:
    dataset: ImportedDataset
    profile: ReportingProfile
    resolution: ProfileResolutionResult
    calculated_channels: list[ChannelInfo]
    calculated_values: np.ndarray
    values_by_semantic_name: dict[str, np.ndarray]
    calculation_order: list[str]
    unavailable: list[UnavailableMathChannel] = field(default_factory=list)

    @property
    def sample_count(self) -> int:
        return int(self.dataset.quality.sample_count)

    @property
    def configured_math_count(self) -> int:
        return len(self.profile.math_channels)

    @property
    def calculated_math_count(self) -> int:
        return len(self.calculated_channels)

    @property
    def unavailable_required(self) -> list[UnavailableMathChannel]:
        return [item for item in self.unavailable if item.definition.required]

    @property
    def unavailable_optional(self) -> list[UnavailableMathChannel]:
        return [item for item in self.unavailable if not item.definition.required]

    @property
    def is_complete(self) -> bool:
        return not self.unavailable_required


def calculate_profile_math_channels(
    dataset: ImportedDataset,
    profile: ReportingProfile,
    resolution: ProfileResolutionResult | None = None,
    constants: Mapping[str, float] | None = None,
) -> ProfileMathResult:
    """Calculate available profile MATH channels using semantic dependencies.

    The profile expressions are compiled and evaluated by the existing deterministic
    math-engine internals; this adapter only maps semantic names to arrays and
    records unavailable profile channels.
    """

    if resolution is None:
        resolution = resolve_profile(dataset, profile)

    constants = {**DEFAULT_PROFILE_MATH_CONSTANTS, **(constants or {})}
    definitions_by_name = {definition.semantic_name: definition for definition in profile.math_channels}
    raw_values = _resolved_raw_values(dataset, resolution)
    raw_definition_names = set(profile.raw_by_semantic_name())
    available_names = raw_definition_names | set(definitions_by_name) | set(constants)

    compiled = {}
    missing_by_math: dict[str, list[str]] = {}
    for definition in profile.math_channels:
        expression = definition.expression
        if expression is None:
            missing_by_math[definition.semantic_name] = ["expression"]
            continue
        compiled_expression = _compile_expression(
            expression,
            context=f"profile math channel '{definition.semantic_name}'",
        )
        compiled[definition.semantic_name] = compiled_expression
        missing = sorted(set(compiled_expression.dependencies) - available_names)
        if missing:
            missing_by_math[definition.semantic_name] = missing

    if missing_by_math:
        required_missing = [
            f"{channel}: {', '.join(missing)}"
            for channel, missing in missing_by_math.items()
            if definitions_by_name[channel].required
        ]
        if required_missing:
            raise MathChannelError(
                "Profile math channels reference unknown required dependencies: " + "; ".join(required_missing)
            )

    calculation_candidates = [name for name in definitions_by_name if name in compiled]
    calculation_order = _topological_order(calculation_candidates, compiled)
    context: dict[str, Any] = {**raw_values, **constants}
    calculated_by_name: dict[str, np.ndarray] = {}
    unavailable_by_name: dict[str, UnavailableMathChannel] = {}
    executed_order: list[str] = []

    for semantic_name in calculation_order:
        definition = definitions_by_name[semantic_name]
        expression = compiled[semantic_name]
        missing_dependencies = tuple(
            dependency
            for dependency in expression.dependencies
            if dependency not in context and dependency not in definitions_by_name
        )
        unavailable_dependencies = tuple(
            dependency for dependency in expression.dependencies if dependency in unavailable_by_name
        )

        if missing_dependencies or unavailable_dependencies:
            unavailable_by_name[semantic_name] = UnavailableMathChannel(
                definition=definition,
                missing_dependencies=missing_dependencies,
                unavailable_dependencies=unavailable_dependencies,
            )
            continue

        with np.errstate(all="ignore"):
            raw_value = _evaluate_node(expression.tree.body, context)
        values = _normalise_result(raw_value, dataset.quality.sample_count, semantic_name)
        non_finite = ~np.isfinite(values)
        if non_finite.any():
            first_index = int(np.flatnonzero(non_finite)[0])
            raise MathChannelError(
                f"Profile math channel '{semantic_name}' produced {int(non_finite.sum())} non-finite values; "
                f"first failure at sample index {first_index}"
            )
        calculated_by_name[semantic_name] = values
        context[semantic_name] = values
        executed_order.append(semantic_name)

    for semantic_name, missing in missing_by_math.items():
        if semantic_name not in unavailable_by_name:
            unavailable_by_name[semantic_name] = UnavailableMathChannel(
                definition=definitions_by_name[semantic_name],
                missing_dependencies=tuple(missing),
            )

    unavailable = [
        unavailable_by_name[name]
        for name in definitions_by_name
        if name in unavailable_by_name
    ]
    calculated_channels = [
        _calculated_channel(dataset, profile, definitions_by_name[name], index)
        for index, name in enumerate(definitions_by_name, start=1)
        if name in calculated_by_name
    ]
    calculated_values = (
        np.column_stack([calculated_by_name[channel.semantic_name] for channel in profile.math_channels if channel.semantic_name in calculated_by_name])
        if calculated_by_name
        else np.empty((dataset.quality.sample_count, 0), dtype=np.float64)
    )

    return ProfileMathResult(
        dataset=dataset,
        profile=profile,
        resolution=resolution,
        calculated_channels=calculated_channels,
        calculated_values=calculated_values,
        values_by_semantic_name=calculated_by_name,
        calculation_order=executed_order,
        unavailable=unavailable,
    )


def calculate_profile_math_file(
    input_file: str | Path,
    profile: ReportingProfile,
    import_options: Any = None,
    constants: Mapping[str, float] | None = None,
) -> ProfileMathResult:
    from .importer import load_data_file

    dataset = load_data_file(input_file, import_options)
    return calculate_profile_math_channels(dataset, profile, constants=constants)


def _resolved_raw_values(dataset: ImportedDataset, resolution: ProfileResolutionResult) -> dict[str, np.ndarray]:
    values: dict[str, np.ndarray] = {}
    for semantic_name, resolved in resolution.resolved.items():
        index = dataset.channel_index(resolved.channel.channel_id)
        values[semantic_name] = np.asarray(dataset.values[:, index], dtype=np.float64)
    return values


def _calculated_channel(
    dataset: ImportedDataset,
    profile: ReportingProfile,
    definition: MathChannelDefinition,
    index: int,
) -> ChannelInfo:
    dependencies = tuple(definition.dependencies)
    return ChannelInfo(
        channel_id=f"{profile.profile_id}__math__{definition.semantic_name}",
        source_name=definition.source_name,
        display_name=definition.report_name,
        unit=definition.unit,
        source_column_index=dataset.quality.channel_count + index,
        source_column_label=f"PROFILE_MATH{index:03d}",
        kind="math",
        dtype="float64",
        provenance=f"profile:{profile.profile_id}:{definition.semantic_name}",
        dependencies=dependencies,
        formula_example=definition.expression,
    )
