from __future__ import annotations

import csv
from pathlib import Path

from ..errors import ConfigurationError, DataValidationError
from ..models import ImportedDataset
from .models import (
    DutyCycleComposition,
    DutyCycleCompositionPlan,
    DutyCycleRowProvenance,
    DutyCycleScenario,
    DutyCycleSourceValidation,
)


_RESOLVED_PROFILE_SOURCES = {
    "raw_vsm_source",
    "raw_vsm_source_with_mission_offsets",
    "full_raw_vsm_source_replay_with_mission_offsets",
    "synthetic_scenario_action",
}

_GENERATION_MODES = {
    "raw_vsm_source": "source_aligned",
    "raw_vsm_source_with_mission_offsets": "source_aligned_with_mission_offsets",
    "full_raw_vsm_source_replay_with_mission_offsets": "source_replay_with_mission_offsets",
    "synthetic_scenario_action": "synthetic_scenario_action",
    "sergio_reference_profile_after_generator_activation": "unresolved_reference_variant",
    "unresolved_sergio_reference_profile": "unresolved_reference_profile",
}


def build_composition_plan(
    scenario: DutyCycleScenario,
    profile_provider=None,
) -> DutyCycleCompositionPlan:
    """Build the full deterministic row/provenance plan without inventing missing physics.

    This is the first Milestone 13B composition layer. It establishes exact row,
    time, phase, source-alignment and resolution status for every output sample.
    Numerical materialisation is deliberately gated until each phase profile has
    a deterministic implementation or an explicit external reference provider.
    """

    rows: list[DutyCycleRowProvenance] = []
    sample_period_s = 1.0 / scenario.sample_rate_hz

    for phase in scenario.phases:
        provider_resolves = profile_provider is not None and profile_provider.supports(phase)
        if provider_resolves:
            generation_mode = f"external_profile:{profile_provider.provider_id}"
            profile_definition_resolved = True
        else:
            generation_mode = _GENERATION_MODES.get(phase.profile_source, phase.profile_source)
            profile_definition_resolved = phase.profile_source in _RESOLVED_PROFILE_SOURCES

        for local_index in range(phase.output_samples):
            output_index = len(rows)
            track_time_s = scenario.first_timestamp_s + output_index * sample_period_s
            report_row = phase.report_rows.start_row + local_index if phase.report_rows is not None else None

            if phase.source_rows_aligned is not None:
                source_row = phase.source_rows_aligned.start_row + local_index
                source_sample_index = source_row - scenario.source_data_start_row
            else:
                source_row = None
                source_sample_index = None

            if profile_definition_resolved:
                notes = ""
            elif phase.profile_source == "sergio_reference_profile_after_generator_activation":
                notes = "Numerical profile requires the unresolved Sergio range-extender/reference variant."
            elif phase.profile_source == "unresolved_sergio_reference_profile":
                notes = "Numerical profile requires a road/reference source not present in the supplied VSM field workbook."
            else:
                notes = f"Numerical profile source '{phase.profile_source}' has no materialiser yet."

            rows.append(
                DutyCycleRowProvenance(
                    output_sample_index=output_index,
                    report_row=report_row,
                    track_time_s=track_time_s,
                    phase_id=phase.phase_id,
                    phase_type=phase.phase_type,
                    phase_local_index=local_index,
                    generation_mode=generation_mode,
                    source_row_aligned=source_row,
                    source_sample_index_aligned=source_sample_index,
                    profile_definition_resolved=profile_definition_resolved,
                    profile_provider_id=(profile_provider.provider_id if provider_resolves else None),
                    profile_provider_source=(profile_provider.source_file if provider_resolves else None),
                    notes=notes,
                )
            )

    if len(rows) != scenario.expected_sample_count:
        raise ConfigurationError(
            f"Composition plan produced {len(rows)} rows but scenario expects {scenario.expected_sample_count}"
        )
    if rows and abs(rows[-1].track_time_s - scenario.last_timestamp_s) > 1e-9:
        raise ConfigurationError(
            f"Composition plan ends at {rows[-1].track_time_s}s but scenario expects {scenario.last_timestamp_s}s"
        )

    return DutyCycleCompositionPlan(scenario=scenario, rows=tuple(rows))


def export_composition_plan(plan: DutyCycleCompositionPlan, path: str | Path) -> Path:
    """Export row-level composition provenance as a deterministic CSV."""

    destination = Path(path).expanduser().resolve()
    destination.parent.mkdir(parents=True, exist_ok=True)
    fieldnames = [
        "output_sample_index",
        "report_row",
        "track_time_s",
        "phase_id",
        "phase_type",
        "phase_local_index",
        "generation_mode",
        "source_row_aligned",
        "source_sample_index_aligned",
        "profile_definition_resolved",
        "profile_provider_id",
        "profile_provider_source",
        "notes",
    ]
    with destination.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        for row in plan.rows:
            writer.writerow(row.to_dict())
    return destination


def validate_source_dataset(
    scenario: DutyCycleScenario, dataset: ImportedDataset
) -> DutyCycleSourceValidation:
    """Validate that an imported source workbook can satisfy all aligned source rows.

    This does not claim that unresolved Sergio/reference phases can be generated.
    It only proves that every configured source-row alignment is addressable and
    that the source timing is compatible with the scenario sample rate.
    """

    aligned_indices = [
        row_index
        for phase in scenario.phases
        if phase.source_rows_aligned is not None
        for row_index in (
            phase.source_rows_aligned.start_row - scenario.source_data_start_row,
            phase.source_rows_aligned.end_row - scenario.source_data_start_row,
        )
    ]
    required_max_index = max(aligned_indices, default=-1)
    if required_max_index >= dataset.quality.sample_count:
        raise DataValidationError(
            f"Duty-cycle scenario requires source sample index {required_max_index}, "
            f"but dataset contains only {dataset.quality.sample_count} samples"
        )

    if dataset.quality.data_start_row != scenario.source_data_start_row:
        raise DataValidationError(
            f"Duty-cycle scenario expects source data to start at row {scenario.source_data_start_row}, "
            f"but imported dataset starts at row {dataset.quality.data_start_row}"
        )

    sample_period_s = 1.0 / scenario.sample_rate_hz
    source_dt = dataset.quality.nominal_time_step
    if source_dt is None:
        raise DataValidationError("Source dataset has no validated nominal time step")
    if abs(source_dt - sample_period_s) > 1e-9:
        raise DataValidationError(
            f"Duty-cycle scenario sample period is {sample_period_s}s, but source dataset nominal time step is {source_dt}s"
        )

    return DutyCycleSourceValidation(
        source_sample_count=dataset.quality.sample_count,
        required_max_source_sample_index=required_max_index,
        source_data_start_row=dataset.quality.data_start_row,
        scenario_source_data_start_row=scenario.source_data_start_row,
        source_nominal_time_step_s=source_dt,
        scenario_sample_period_s=sample_period_s,
    )


_SOURCE_PROFILE_MODES = {
    "raw_vsm_source",
    "raw_vsm_source_with_mission_offsets",
    "full_raw_vsm_source_replay_with_mission_offsets",
}

_REQUIRED_NUMERICAL_ROLES = {
    "track_time_s",
    "time_minutes",
    "distance_m",
    "distance_km",
    "edu_elect_power_rl",
    "edu_elect_power_rr",
    "total_edu_elect_power",
    "battery_power_kw",
    "battery_power_squared",
    "battery_heatflow_kw",
    "battery_heatflow_squared",
    "battery_energy_kwh",
    "battery_soc_pct",
    "energy_recuperated_kwh",
    "energy_released_kwh",
    "edu_speed_rl",
    "edu_torque_rl",
    "edu_mech_power_rl",
    "edu_speed_rr",
    "edu_torque_rr",
    "edu_mech_power_rr",
    "total_edu_mech_power",
    "fuel_consumption_kg",
    "engine_specific_fuel_consumption",
    "fuel_flow_lph",
    "engine_speed_rpm",
    "engine_torque_nm",
    "engine_power_kw",
    "engine_energy_kwh",
    "tank_force_n",
    "tank_mass_kg",
    "rr_power_fl",
    "rr_power_fr",
    "rr_power_rl",
    "rr_power_rr",
    "rr_power_total",
    "rr_energy_kwh",
    "rr_energy_wh",
    "rr_energy_accum_kwh",
    "wheel_speed_rl",
    "wheel_speed_rr",
    "driveshaft_torque_rl",
    "driveshaft_torque_rr",
    "wheel_total_torque",
    "wheel_power_rl",
    "wheel_power_rr",
    "wheel_power_total",
    "aux_low_voltage_kw",
    "aux_high_voltage_kw",
    "aux_total_kw",
    "aux_energy_kwh",
    "aux_energy_accum_kwh",
    "generator_1_torque_nm",
    "generator_1_power_kw",
    "generator_2_torque_nm",
    "generator_2_power_kw",
    "generator_total_power_kw",
}


def _resolve_role_indexes(scenario: DutyCycleScenario, dataset: ImportedDataset) -> dict[str, int]:
    missing_roles = sorted(role for role in _REQUIRED_NUMERICAL_ROLES if role not in scenario.channel_roles)
    if missing_roles:
        raise ConfigurationError(
            "Duty-cycle numerical materialisation requires channel_roles entries for: " + ", ".join(missing_roles)
        )

    indexes: dict[str, int] = {}
    for role, channel_id in scenario.channel_roles.items():
        try:
            indexes[role] = dataset.channel_index(channel_id)
        except KeyError as exc:
            raise DataValidationError(
                f"Duty-cycle channel role '{role}' refers to missing source channel_id '{channel_id}'"
            ) from exc
    return indexes


def _require_opportunity_charge_parameters(scenario: DutyCycleScenario) -> None:
    cfg = scenario.opportunity_charge
    required = {
        "battery_power_kw": cfg.battery_power_kw,
        "battery_heatflow_kw": cfg.battery_heatflow_kw,
        "engine_speed_rpm": cfg.engine_speed_rpm,
        "engine_torque_nm": cfg.engine_torque_nm,
        "engine_specific_fuel_consumption_g_per_kwh": cfg.engine_specific_fuel_consumption_g_per_kwh,
        "fuel_flow_lph": cfg.fuel_flow_lph,
        "fuel_density_kg_per_l": cfg.fuel_density_kg_per_l,
        "tank_load_rate_kg_s": cfg.tank_load_rate_kg_s,
        "low_voltage_aux_kw": cfg.low_voltage_aux_kw,
        "high_voltage_aux_kw": cfg.high_voltage_aux_kw,
        "generator_1_torque_nm": cfg.generator_1_torque_nm,
        "generator_2_torque_nm": cfg.generator_2_torque_nm,
    }
    missing = [name for name, value in required.items() if value is None]
    if missing:
        raise ConfigurationError(
            "Duty-cycle loading materialisation requires controls.opportunity_charge values for: "
            + ", ".join(missing)
        )
    if scenario.battery_nominal_capacity_kwh is None:
        raise ConfigurationError("Duty-cycle numerical materialisation requires battery.mission_nominal_capacity_kwh")


def _recalculate_derived_channels(
    values,
    indexes: dict[str, int],
    sample_period_s: float,
    previous_row,
) -> None:
    import numpy as np

    i = indexes
    values[:, i["time_minutes"]] = values[:, i["track_time_s"]] / 60.0
    values[:, i["distance_km"]] = values[:, i["distance_m"]] / 1000.0
    values[:, i["total_edu_elect_power"]] = (
        values[:, i["edu_elect_power_rl"]] + values[:, i["edu_elect_power_rr"]]
    )
    values[:, i["battery_power_squared"]] = values[:, i["battery_power_kw"]] ** 2
    values[:, i["battery_heatflow_squared"]] = values[:, i["battery_heatflow_kw"]] ** 2
    values[:, i["edu_mech_power_rl"]] = (
        values[:, i["edu_speed_rl"]] * values[:, i["edu_torque_rl"]] / 9548.8
    )
    values[:, i["edu_mech_power_rr"]] = (
        values[:, i["edu_speed_rr"]] * values[:, i["edu_torque_rr"]] / 9548.8
    )
    values[:, i["total_edu_mech_power"]] = (
        values[:, i["edu_mech_power_rl"]] + values[:, i["edu_mech_power_rr"]]
    )
    values[:, i["engine_power_kw"]] = (
        values[:, i["engine_speed_rpm"]] * values[:, i["engine_torque_nm"]] / 9548.8
    )
    values[:, i["engine_energy_kwh"]] = values[:, i["engine_power_kw"]] * sample_period_s / 3600.0
    values[:, i["tank_mass_kg"]] = -values[:, i["tank_force_n"]] / 9.81
    values[:, i["rr_power_total"]] = (
        values[:, i["rr_power_fl"]]
        + values[:, i["rr_power_fr"]]
        + values[:, i["rr_power_rl"]]
        + values[:, i["rr_power_rr"]]
    )
    values[:, i["rr_energy_kwh"]] = values[:, i["rr_power_total"]] * sample_period_s / 3600.0
    values[:, i["rr_energy_wh"]] = values[:, i["rr_energy_kwh"]] * 1000.0

    previous_rr = 0.0 if previous_row is None else float(previous_row[i["rr_energy_accum_kwh"]])
    values[:, i["rr_energy_accum_kwh"]] = previous_rr + np.cumsum(values[:, i["rr_energy_kwh"]])

    values[:, i["wheel_total_torque"]] = (
        values[:, i["driveshaft_torque_rl"]] + values[:, i["driveshaft_torque_rr"]]
    )
    values[:, i["wheel_power_rl"]] = (
        values[:, i["wheel_speed_rl"]] * values[:, i["driveshaft_torque_rl"]] / 9548.8
    )
    values[:, i["wheel_power_rr"]] = (
        values[:, i["wheel_speed_rr"]] * values[:, i["driveshaft_torque_rr"]] / 9548.8
    )
    values[:, i["wheel_power_total"]] = values[:, i["wheel_power_rl"]] + values[:, i["wheel_power_rr"]]

    values[:, i["aux_total_kw"]] = values[:, i["aux_low_voltage_kw"]] + values[:, i["aux_high_voltage_kw"]]
    values[:, i["aux_energy_kwh"]] = values[:, i["aux_total_kw"]] * sample_period_s / 3600.0
    previous_aux = 0.0 if previous_row is None else float(previous_row[i["aux_energy_accum_kwh"]])
    values[:, i["aux_energy_accum_kwh"]] = previous_aux + np.cumsum(values[:, i["aux_energy_kwh"]])

    values[:, i["generator_1_power_kw"]] = (
        values[:, i["engine_speed_rpm"]] * values[:, i["generator_1_torque_nm"]] / 9548.8
    )
    values[:, i["generator_2_power_kw"]] = (
        values[:, i["engine_speed_rpm"]] * values[:, i["generator_2_torque_nm"]] / 9548.8
    )
    values[:, i["generator_total_power_kw"]] = (
        values[:, i["generator_1_power_kw"]] + values[:, i["generator_2_power_kw"]]
    )


def _materialize_source_phase(
    scenario: DutyCycleScenario,
    phase,
    source: ImportedDataset,
    phase_plan_rows,
    indexes: dict[str, int],
    previous_row,
):
    import numpy as np

    if phase.source_rows_aligned is None:
        raise ConfigurationError(f"Phase {phase.phase_id}: source-backed phase has no source_rows_aligned")
    if phase.initial_soc_pct is None:
        raise ConfigurationError(f"Phase {phase.phase_id}: source-backed phase requires initial_soc_pct")
    if scenario.battery_nominal_capacity_kwh is None:
        raise ConfigurationError("Duty-cycle numerical materialisation requires battery.mission_nominal_capacity_kwh")

    start = phase.source_rows_aligned.start_row - scenario.source_data_start_row
    stop = start + phase.output_samples
    values = np.array(source.values[start:stop, :], dtype=np.float64, copy=True)
    if values.shape[0] != phase.output_samples:
        raise DataValidationError(
            f"Phase {phase.phase_id}: source slice contains {values.shape[0]} rows, expected {phase.output_samples}"
        )

    i = indexes
    sample_period_s = 1.0 / scenario.sample_rate_hz
    values[:, i["track_time_s"]] = np.array([row.track_time_s for row in phase_plan_rows], dtype=np.float64)

    if previous_row is not None:
        distance_offset = float(previous_row[i["distance_m"]]) - float(values[0, i["distance_m"]])
        values[:, i["distance_m"]] += distance_offset

    target_energy_kwh = scenario.battery_nominal_capacity_kwh * phase.initial_soc_pct / 100.0
    energy_offset = target_energy_kwh - float(values[0, i["battery_energy_kwh"]])
    soc_offset = phase.initial_soc_pct - float(values[0, i["battery_soc_pct"]])
    values[:, i["battery_energy_kwh"]] += energy_offset
    values[:, i["battery_soc_pct"]] += soc_offset

    for role in ("energy_recuperated_kwh", "energy_released_kwh", "fuel_consumption_kg"):
        base = float(values[0, i[role]])
        target = 0.0 if previous_row is None else float(previous_row[i[role]])
        values[:, i[role]] += target - base

    for role in scenario.source_restart_zero_channel_roles:
        if role not in i:
            raise ConfigurationError(f"source_phase_rules refers to unknown channel role '{role}'")
        values[0, i[role]] = 0.0

    _recalculate_derived_channels(values, i, sample_period_s, previous_row)
    return values


def _materialize_loading_phase(
    scenario: DutyCycleScenario,
    phase,
    source: ImportedDataset,
    phase_plan_rows,
    indexes: dict[str, int],
    previous_row,
):
    import numpy as np

    if previous_row is None:
        raise ConfigurationError(f"Phase {phase.phase_id}: loading phase requires a previous materialised phase")
    _require_opportunity_charge_parameters(scenario)
    cfg = scenario.opportunity_charge
    assert scenario.battery_nominal_capacity_kwh is not None

    n = phase.output_samples
    dt = 1.0 / scenario.sample_rate_hz
    values = np.zeros((n, source.quality.channel_count), dtype=np.float64)
    i = indexes

    values[:, i["track_time_s"]] = np.array([row.track_time_s for row in phase_plan_rows], dtype=np.float64)
    values[:, i["distance_m"]] = float(previous_row[i["distance_m"]])

    values[:, i["battery_power_kw"]] = float(cfg.battery_power_kw)
    values[:, i["battery_heatflow_kw"]] = float(cfg.battery_heatflow_kw)
    battery_energy_increment = float(cfg.battery_power_kw) * dt / 3600.0
    increments = np.arange(1, n + 1, dtype=np.float64)
    values[:, i["battery_energy_kwh"]] = float(previous_row[i["battery_energy_kwh"]]) + increments * battery_energy_increment
    values[:, i["battery_soc_pct"]] = (
        values[:, i["battery_energy_kwh"]] / scenario.battery_nominal_capacity_kwh * 100.0
    )
    values[:, i["energy_recuperated_kwh"]] = float(previous_row[i["energy_recuperated_kwh"]])
    values[:, i["energy_released_kwh"]] = float(previous_row[i["energy_released_kwh"]])

    values[:, i["engine_specific_fuel_consumption"]] = float(cfg.engine_specific_fuel_consumption_g_per_kwh)
    values[:, i["fuel_flow_lph"]] = float(cfg.fuel_flow_lph)
    fuel_mass_increment = float(cfg.fuel_flow_lph) * float(cfg.fuel_density_kg_per_l) * dt / 3600.0
    values[:, i["fuel_consumption_kg"]] = float(previous_row[i["fuel_consumption_kg"]]) + increments * fuel_mass_increment
    values[:, i["engine_speed_rpm"]] = float(cfg.engine_speed_rpm)
    values[:, i["engine_torque_nm"]] = float(cfg.engine_torque_nm)

    tank_mass = increments * float(cfg.tank_load_rate_kg_s) * dt
    values[:, i["tank_force_n"]] = -tank_mass * 9.81

    values[:, i["aux_low_voltage_kw"]] = float(cfg.low_voltage_aux_kw)
    values[:, i["aux_high_voltage_kw"]] = float(cfg.high_voltage_aux_kw)
    values[:, i["generator_1_torque_nm"]] = float(cfg.generator_1_torque_nm)
    values[:, i["generator_2_torque_nm"]] = float(cfg.generator_2_torque_nm)

    _recalculate_derived_channels(values, i, dt, previous_row)
    return values


def compose_supported_prefix(
    scenario: DutyCycleScenario,
    source: ImportedDataset,
):
    """Materialise the longest deterministic prefix before the first unresolved phase.

    The function intentionally stops rather than filling unresolved phases with
    NaNs, copied source rows or guessed profiles. This makes the Milestone 13B.1
    implementation useful and testable while preserving the fidelity gate.
    """

    import numpy as np
    from .models import DutyCyclePartialComposition

    validate_source_dataset(scenario, source)
    indexes = _resolve_role_indexes(scenario, source)
    plan = build_composition_plan(scenario)

    segments = []
    provenance = []
    completed: list[str] = []
    previous_row = None
    plan_offset = 0
    stopped_before = None

    for phase in scenario.phases:
        phase_rows = plan.rows[plan_offset : plan_offset + phase.output_samples]
        if not phase_rows or phase_rows[0].phase_id != phase.phase_id:
            raise ConfigurationError(f"Composition plan is not aligned with phase {phase.phase_id}")

        if not phase_rows[0].profile_definition_resolved:
            stopped_before = phase.phase_id
            break

        if phase.profile_source in _SOURCE_PROFILE_MODES:
            values = _materialize_source_phase(
                scenario, phase, source, phase_rows, indexes, previous_row
            )
        elif phase.profile_source == "synthetic_scenario_action":
            values = _materialize_loading_phase(
                scenario, phase, source, phase_rows, indexes, previous_row
            )
        else:
            raise ConfigurationError(
                f"Phase {phase.phase_id}: no numerical materialiser for profile_source '{phase.profile_source}'"
            )

        segments.append(values)
        provenance.extend(phase_rows)
        completed.append(phase.phase_id)
        previous_row = values[-1, :]
        plan_offset += phase.output_samples

    combined = (
        np.vstack(segments)
        if segments
        else np.empty((0, source.quality.channel_count), dtype=np.float64)
    )
    return DutyCyclePartialComposition(
        scenario=scenario,
        channels=list(source.channels),
        values=combined,
        provenance=tuple(provenance),
        completed_phase_ids=tuple(completed),
        stopped_before_phase_id=stopped_before,
    )


def compose_duty_cycle(
    scenario: DutyCycleScenario,
    source: ImportedDataset,
    profile_provider,
) -> DutyCycleComposition:
    """Materialise the complete scenario using explicit external profiles where required.

    Unresolved phases are never guessed.  The supplied provider must explicitly
    support every phase that lacks a native source/synthetic materialiser and
    must pass channel/timing/provenance validation before composition starts.
    """

    import numpy as np

    validate_source_dataset(scenario, source)
    indexes = _resolve_role_indexes(scenario, source)
    provider_validation = profile_provider.validate(scenario, source)
    plan = build_composition_plan(scenario, profile_provider=profile_provider)

    if plan.unresolved_phase_ids:
        raise ConfigurationError(
            "Duty-cycle full composition still has unresolved phase profiles: "
            + ", ".join(plan.unresolved_phase_ids)
        )

    segments = []
    provenance = []
    completed: list[str] = []
    profile_provenance = []
    previous_row = None
    plan_offset = 0

    for phase in scenario.phases:
        phase_rows = plan.rows[plan_offset : plan_offset + phase.output_samples]
        if not phase_rows or phase_rows[0].phase_id != phase.phase_id:
            raise ConfigurationError(f"Composition plan is not aligned with phase {phase.phase_id}")

        if phase.profile_source in _SOURCE_PROFILE_MODES:
            values = _materialize_source_phase(
                scenario, phase, source, phase_rows, indexes, previous_row
            )
        elif phase.profile_source == "synthetic_scenario_action":
            values = _materialize_loading_phase(
                scenario, phase, source, phase_rows, indexes, previous_row
            )
        elif profile_provider.supports(phase):
            values, phase_profile_provenance = profile_provider.materialize_phase(
                scenario, phase, source, phase_rows
            )
            profile_provenance.append(phase_profile_provenance)
        else:
            raise ConfigurationError(
                f"Phase {phase.phase_id}: no native materialiser and profile provider "
                f"'{profile_provider.provider_id}' does not support it"
            )

        if values.shape != (phase.output_samples, source.quality.channel_count):
            raise DataValidationError(
                f"Phase {phase.phase_id}: materialiser returned shape {values.shape}; expected "
                f"({phase.output_samples}, {source.quality.channel_count})"
            )

        segments.append(values)
        provenance.extend(phase_rows)
        completed.append(phase.phase_id)
        previous_row = values[-1, :]
        plan_offset += phase.output_samples

    combined = np.vstack(segments)
    if combined.shape[0] != scenario.expected_sample_count:
        raise DataValidationError(
            f"Full composition produced {combined.shape[0]} samples; expected {scenario.expected_sample_count}"
        )
    if len(provenance) != scenario.expected_sample_count:
        raise DataValidationError(
            f"Full composition provenance contains {len(provenance)} rows; expected {scenario.expected_sample_count}"
        )

    # Touch the validation result to make the precondition explicit and protect
    # against accidental removal as the provider API evolves.
    if not provider_validation.channel_layout_compatible:
        raise DataValidationError("Profile-provider validation did not confirm channel compatibility")

    return DutyCycleComposition(
        scenario=scenario,
        channels=list(source.channels),
        values=combined,
        provenance=tuple(provenance),
        completed_phase_ids=tuple(completed),
        profile_provenance=tuple(profile_provenance),
    )


def export_pipeline_dataset(composition: DutyCycleComposition, path: str | Path) -> Path:
    """Export the composed mission as a normal numeric VSM-style CSV dataset.

    This export deliberately excludes provenance/string columns so the existing
    importer, math, statistics, plotting and report engines can consume the full
    mission without special-case code. Source channel names and units are kept in
    their original column order, which preserves the stable ``channel_id`` values
    regenerated by the CSV importer.
    """

    destination = Path(path).expanduser().resolve()
    destination.parent.mkdir(parents=True, exist_ok=True)
    with destination.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow([channel.source_name for channel in composition.channels])
        writer.writerow([channel.unit or "" for channel in composition.channels])
        for values in composition.values:
            writer.writerow(values.tolist())
    return destination


def export_duty_cycle_composition(composition: DutyCycleComposition, path: str | Path) -> Path:
    """Export a complete materialised duty cycle with per-row provenance columns."""

    destination = Path(path).expanduser().resolve()
    destination.parent.mkdir(parents=True, exist_ok=True)
    headers = [
        "output_sample_index",
        "phase_id",
        "phase_local_index",
        "generation_mode",
        "profile_provider_id",
        "profile_provider_source",
    ] + [channel.channel_id for channel in composition.channels]
    with destination.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(headers)
        for row_index, row_provenance in enumerate(composition.provenance):
            writer.writerow(
                [
                    row_provenance.output_sample_index,
                    row_provenance.phase_id,
                    row_provenance.phase_local_index,
                    row_provenance.generation_mode,
                    row_provenance.profile_provider_id or "",
                    row_provenance.profile_provider_source or "",
                    *composition.values[row_index, :].tolist(),
                ]
            )
    return destination


def export_partial_composition(partial, path: str | Path) -> Path:
    """Export a materialised duty-cycle prefix using unique channel IDs."""

    destination = Path(path).expanduser().resolve()
    destination.parent.mkdir(parents=True, exist_ok=True)
    headers = ["output_sample_index", "phase_id", "phase_local_index"] + [
        channel.channel_id for channel in partial.channels
    ]
    with destination.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(headers)
        for row_index, provenance in enumerate(partial.provenance):
            writer.writerow(
                [
                    provenance.output_sample_index,
                    provenance.phase_id,
                    provenance.phase_local_index,
                    *partial.values[row_index, :].tolist(),
                ]
            )
    return destination
