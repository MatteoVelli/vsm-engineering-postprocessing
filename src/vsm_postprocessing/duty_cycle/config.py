from __future__ import annotations

from pathlib import Path
from typing import Any

import yaml

from ..errors import ConfigurationError
from .models import (
    DutyCyclePhase,
    DutyCycleScenario,
    OpportunityChargeConfig,
    ReportRowRange,
    SourceRowRange,
)


def _require_mapping(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise ConfigurationError(f"{label} must be a mapping")
    return value


def _require_positive_int(value: Any, label: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value <= 0:
        raise ConfigurationError(f"{label} must be a positive integer")
    return value


def _parse_inclusive_range(value: Any, label: str, range_type):
    if value is None:
        return None
    if not isinstance(value, list) or len(value) != 2:
        raise ConfigurationError(f"{label} must be a two-item inclusive [start, end] list")
    start, end = value
    if isinstance(start, bool) or isinstance(end, bool) or not isinstance(start, int) or not isinstance(end, int):
        raise ConfigurationError(f"{label} values must be integers")
    if start < 1 or end < start:
        raise ConfigurationError(f"{label} must satisfy 1 <= start <= end")
    return range_type(start, end)


def load_duty_cycle_config(path: str | Path) -> DutyCycleScenario:
    """Load and strictly validate a duty-cycle YAML scenario definition.

    The loader accepts the compact runtime configuration introduced in
    Milestone 13B as well as the Phase 13A.1 formal specification structure.
    """

    config_path = Path(path).expanduser().resolve()
    if not config_path.exists():
        raise ConfigurationError(f"Duty-cycle configuration does not exist: {config_path}")

    try:
        raw = yaml.safe_load(config_path.read_text(encoding="utf-8"))
    except Exception as exc:
        raise ConfigurationError(f"Could not read duty-cycle configuration '{config_path}': {exc}") from exc
    root = _require_mapping(raw, "Duty-cycle configuration")

    scenario_block = _require_mapping(root.get("scenario"), "scenario")
    scenario_id = str(scenario_block.get("id", "")).strip()
    if not scenario_id:
        raise ConfigurationError("scenario.id must be a non-empty string")

    sample_rate_raw = scenario_block.get("sample_rate_hz")
    try:
        sample_rate_hz = float(sample_rate_raw)
    except (TypeError, ValueError) as exc:
        raise ConfigurationError("scenario.sample_rate_hz must be numeric") from exc
    if sample_rate_hz <= 0:
        raise ConfigurationError("scenario.sample_rate_hz must be greater than zero")

    expected_sample_count = _require_positive_int(
        scenario_block.get("expected_sample_count"), "scenario.expected_sample_count"
    )

    timeline = _require_mapping(root.get("timeline", {}), "timeline")
    first_timestamp_s = float(timeline.get("first_timestamp_s", 0.0))
    if "last_timestamp_s" in timeline:
        last_timestamp_s = float(timeline["last_timestamp_s"])
    else:
        last_timestamp_s = first_timestamp_s + (expected_sample_count - 1) / sample_rate_hz

    source_block = root.get("source_field_cycle")
    if isinstance(source_block, dict) and source_block.get("data_rows") is not None:
        source_rows = _parse_inclusive_range(source_block["data_rows"], "source_field_cycle.data_rows", SourceRowRange)
        assert source_rows is not None
        source_data_start_row = source_rows.start_row
    else:
        source_runtime = _require_mapping(root.get("source", {}), "source")
        source_data_start_row = _require_positive_int(
            source_runtime.get("data_start_row"), "source.data_start_row"
        )

    phases_raw = root.get("phases")
    if not isinstance(phases_raw, list) or not phases_raw:
        raise ConfigurationError("phases must be a non-empty list")

    phases: list[DutyCyclePhase] = []
    seen_ids: set[str] = set()
    previous_report_end: int | None = None
    for index, phase_raw_any in enumerate(phases_raw, start=1):
        phase_raw = _require_mapping(phase_raw_any, f"phases[{index}]")
        phase_id = str(phase_raw.get("id", "")).strip()
        if not phase_id:
            raise ConfigurationError(f"phases[{index}].id must be a non-empty string")
        if phase_id in seen_ids:
            raise ConfigurationError(f"Duplicate duty-cycle phase id '{phase_id}'")
        seen_ids.add(phase_id)

        phase_type = str(phase_raw.get("type", "")).strip()
        profile_source = str(phase_raw.get("profile_source", "")).strip()
        if not phase_type:
            raise ConfigurationError(f"Phase {phase_id}: type must be a non-empty string")
        if not profile_source:
            raise ConfigurationError(f"Phase {phase_id}: profile_source must be a non-empty string")

        output_samples = _require_positive_int(phase_raw.get("output_samples"), f"Phase {phase_id}.output_samples")
        report_rows = _parse_inclusive_range(phase_raw.get("report_rows"), f"Phase {phase_id}.report_rows", ReportRowRange)
        source_rows_aligned = _parse_inclusive_range(
            phase_raw.get("source_rows_aligned"), f"Phase {phase_id}.source_rows_aligned", SourceRowRange
        )

        if report_rows is not None:
            if report_rows.sample_count != output_samples:
                raise ConfigurationError(
                    f"Phase {phase_id}: report row count {report_rows.sample_count} does not match "
                    f"output_samples {output_samples}"
                )
            if previous_report_end is not None and report_rows.start_row != previous_report_end + 1:
                raise ConfigurationError(
                    f"Phase {phase_id}: report rows are not contiguous with previous phase "
                    f"({report_rows.start_row} != {previous_report_end + 1})"
                )
            previous_report_end = report_rows.end_row

        if source_rows_aligned is not None and source_rows_aligned.sample_count != output_samples:
            raise ConfigurationError(
                f"Phase {phase_id}: aligned source row count {source_rows_aligned.sample_count} does not match "
                f"output_samples {output_samples}"
            )
        if source_rows_aligned is not None and source_rows_aligned.start_row < source_data_start_row:
            raise ConfigurationError(
                f"Phase {phase_id}: source alignment starts before source data row {source_data_start_row}"
            )

        known_keys = {
            "id",
            "type",
            "report_rows",
            "output_samples",
            "source_rows_aligned",
            "initial_soc_pct",
            "profile_source",
            "drive_controller_initial_state",
            "generator_start_threshold_pct",
            "generator_stop_threshold_pct",
        }
        metadata = {key: value for key, value in phase_raw.items() if key not in known_keys}
        phases.append(
            DutyCyclePhase(
                phase_id=phase_id,
                phase_type=phase_type,
                output_samples=output_samples,
                profile_source=profile_source,
                report_rows=report_rows,
                source_rows_aligned=source_rows_aligned,
                initial_soc_pct=(float(phase_raw["initial_soc_pct"]) if phase_raw.get("initial_soc_pct") is not None else None),
                generator_start_threshold_pct=(
                    float(phase_raw["generator_start_threshold_pct"])
                    if phase_raw.get("generator_start_threshold_pct") is not None
                    else None
                ),
                generator_stop_threshold_pct=(
                    float(phase_raw["generator_stop_threshold_pct"])
                    if phase_raw.get("generator_stop_threshold_pct") is not None
                    else None
                ),
                drive_controller_initial_state=(
                    str(phase_raw["drive_controller_initial_state"])
                    if phase_raw.get("drive_controller_initial_state") is not None
                    else None
                ),
                metadata=metadata,
            )
        )

    configured_sample_count = sum(phase.output_samples for phase in phases)
    if configured_sample_count != expected_sample_count:
        raise ConfigurationError(
            f"Scenario '{scenario_id}' phases contain {configured_sample_count} samples; "
            f"expected {expected_sample_count}"
        )

    expected_last_timestamp = first_timestamp_s + (expected_sample_count - 1) / sample_rate_hz
    if abs(last_timestamp_s - expected_last_timestamp) > 1e-9:
        raise ConfigurationError(
            f"Scenario '{scenario_id}' last_timestamp_s={last_timestamp_s} is inconsistent with "
            f"{expected_sample_count} samples at {sample_rate_hz} Hz; expected {expected_last_timestamp}"
        )

    battery_block = root.get("battery", {})
    if battery_block is None:
        battery_block = {}
    battery_block = _require_mapping(battery_block, "battery")
    battery_nominal_capacity_kwh = battery_block.get("mission_nominal_capacity_kwh")
    if battery_nominal_capacity_kwh is not None:
        battery_nominal_capacity_kwh = float(battery_nominal_capacity_kwh)
        if battery_nominal_capacity_kwh <= 0:
            raise ConfigurationError("battery.mission_nominal_capacity_kwh must be greater than zero")

    channel_roles_raw = root.get("channel_roles", {})
    if channel_roles_raw is None:
        channel_roles_raw = {}
    channel_roles_map = _require_mapping(channel_roles_raw, "channel_roles")
    channel_roles: dict[str, str] = {}
    for role, channel_id in channel_roles_map.items():
        role_name = str(role).strip()
        channel_name = str(channel_id).strip()
        if not role_name or not channel_name:
            raise ConfigurationError("channel_roles entries must use non-empty role names and channel ids")
        channel_roles[role_name] = channel_name

    phase_rules = root.get("source_phase_rules", {})
    if phase_rules is None:
        phase_rules = {}
    phase_rules = _require_mapping(phase_rules, "source_phase_rules")
    restart_zero_raw = phase_rules.get("restart_zero_channel_roles", [])
    if not isinstance(restart_zero_raw, list):
        raise ConfigurationError("source_phase_rules.restart_zero_channel_roles must be a list")
    source_restart_zero_channel_roles = tuple(str(item).strip() for item in restart_zero_raw if str(item).strip())

    controls = root.get("controls", {})
    if controls is None:
        controls = {}
    controls = _require_mapping(controls, "controls")
    opportunity_raw = controls.get("opportunity_charge", {})
    if opportunity_raw is None:
        opportunity_raw = {}
    opportunity_raw = _require_mapping(opportunity_raw, "controls.opportunity_charge")

    def optional_float(name: str) -> float | None:
        value = opportunity_raw.get(name)
        return None if value is None else float(value)

    opportunity_charge = OpportunityChargeConfig(
        battery_power_kw=optional_float("battery_power_kw"),
        battery_heatflow_kw=optional_float("battery_heatflow_kw"),
        engine_speed_rpm=optional_float("generator_speed_rpm"),
        engine_torque_nm=optional_float("engine_torque_nm"),
        engine_specific_fuel_consumption_g_per_kwh=optional_float("engine_specific_fuel_consumption_g_per_kwh"),
        fuel_flow_lph=optional_float("fuel_flow_lph_nominal"),
        fuel_density_kg_per_l=optional_float("fuel_density_kg_per_l"),
        tank_load_rate_kg_s=optional_float("tank_load_rate_kg_s"),
        low_voltage_aux_kw=optional_float("low_voltage_aux_kw"),
        high_voltage_aux_kw=optional_float("high_voltage_aux_kw"),
        generator_1_torque_nm=optional_float("generator_1_torque_nm"),
        generator_2_torque_nm=optional_float("generator_2_torque_nm"),
    )

    metadata = {
        key: value
        for key, value in root.items()
        if key not in {
            "scenario", "timeline", "source", "source_field_cycle", "phases",
            "battery", "channel_roles", "source_phase_rules", "controls"
        }
    }
    return DutyCycleScenario(
        scenario_id=scenario_id,
        sample_rate_hz=sample_rate_hz,
        expected_sample_count=expected_sample_count,
        first_timestamp_s=first_timestamp_s,
        last_timestamp_s=last_timestamp_s,
        source_data_start_row=source_data_start_row,
        phases=tuple(phases),
        battery_nominal_capacity_kwh=battery_nominal_capacity_kwh,
        channel_roles=channel_roles,
        source_restart_zero_channel_roles=source_restart_zero_channel_roles,
        opportunity_charge=opportunity_charge,
        metadata=metadata,
    )


def load_profile_provider_config(path: str | Path):
    """Load and validate a workbook-backed duty-cycle profile-provider YAML."""

    from .profile_provider import WorkbookProfileProviderConfig

    config_path = Path(path).expanduser().resolve()
    if not config_path.exists():
        raise ConfigurationError(f"Profile-provider configuration does not exist: {config_path}")
    try:
        raw = yaml.safe_load(config_path.read_text(encoding="utf-8"))
    except Exception as exc:
        raise ConfigurationError(
            f"Could not read profile-provider configuration '{config_path}': {exc}"
        ) from exc
    root = _require_mapping(raw, "Profile-provider configuration")
    provider = _require_mapping(root.get("provider"), "provider")

    known_provider_keys = {
        "id",
        "type",
        "value_policy",
        "expected_filename",
        "expected_sha256",
        "phase_ids",
        "channel_alignment",
        "phase_rows",
        "import",
    }
    unknown = sorted(set(provider) - known_provider_keys)
    if unknown:
        raise ConfigurationError(
            "Unknown profile-provider configuration key(s): " + ", ".join(unknown)
        )

    provider_id = str(provider.get("id", "")).strip()
    provider_type = str(provider.get("type", "")).strip()
    value_policy = str(provider.get("value_policy", "absolute_reference")).strip()
    channel_alignment = str(provider.get("channel_alignment", "channel_id")).strip()
    if not provider_id:
        raise ConfigurationError("provider.id must be a non-empty string")
    if not provider_type:
        raise ConfigurationError("provider.type must be a non-empty string")

    phase_ids_raw = provider.get("phase_ids")
    if not isinstance(phase_ids_raw, list) or not phase_ids_raw:
        raise ConfigurationError("provider.phase_ids must be a non-empty list")
    phase_ids = tuple(str(item).strip() for item in phase_ids_raw)
    if any(not item for item in phase_ids):
        raise ConfigurationError("provider.phase_ids entries must be non-empty strings")
    if len(set(phase_ids)) != len(phase_ids):
        raise ConfigurationError("provider.phase_ids must not contain duplicates")

    expected_filename_raw = provider.get("expected_filename")
    expected_filename = None if expected_filename_raw is None else str(expected_filename_raw).strip()
    expected_sha_raw = provider.get("expected_sha256")
    expected_sha256 = None if expected_sha_raw is None else str(expected_sha_raw).strip().lower()
    if expected_sha256 is not None:
        if len(expected_sha256) != 64 or any(c not in "0123456789abcdef" for c in expected_sha256):
            raise ConfigurationError("provider.expected_sha256 must be a 64-character hexadecimal SHA-256")

    import_raw = provider.get("import", {})
    if import_raw is None:
        import_raw = {}
    import_raw = _require_mapping(import_raw, "provider.import")
    allowed_import = {
        "sheet_name",
        "header_row",
        "unit_row",
        "data_start_row",
        "data_end_row",
        "last_channel_column",
    }
    unknown_import = sorted(set(import_raw) - allowed_import)
    if unknown_import:
        raise ConfigurationError(
            "Unknown provider.import configuration key(s): " + ", ".join(unknown_import)
        )

    def optional_int(name: str) -> int | None:
        value = import_raw.get(name)
        if value is None:
            return None
        if isinstance(value, bool) or not isinstance(value, int) or value <= 0:
            raise ConfigurationError(f"provider.import.{name} must be a positive integer")
        return value

    sheet_name_raw = import_raw.get("sheet_name")
    sheet_name = None if sheet_name_raw is None else str(sheet_name_raw).strip()
    if sheet_name == "":
        sheet_name = None

    data_start_row = optional_int("data_start_row")
    data_end_row = optional_int("data_end_row")
    if data_start_row is not None and data_end_row is not None and data_end_row < data_start_row:
        raise ConfigurationError("provider.import.data_end_row must be >= data_start_row")

    phase_rows_raw = provider.get("phase_rows", {})
    if phase_rows_raw is None:
        phase_rows_raw = {}
    phase_rows_raw = _require_mapping(phase_rows_raw, "provider.phase_rows")
    phase_rows: dict[str, tuple[int, int]] = {}
    for phase_id, row_range in phase_rows_raw.items():
        phase_name = str(phase_id).strip()
        parsed = _parse_inclusive_range(row_range, f"provider.phase_rows.{phase_name}", SourceRowRange)
        assert parsed is not None
        phase_rows[phase_name] = (parsed.start_row, parsed.end_row)
    unknown_phase_rows = sorted(set(phase_rows) - set(phase_ids))
    if unknown_phase_rows:
        raise ConfigurationError(
            "provider.phase_rows contains phase ids not listed in provider.phase_ids: "
            + ", ".join(unknown_phase_rows)
        )

    return WorkbookProfileProviderConfig(
        provider_id=provider_id,
        provider_type=provider_type,
        phase_ids=phase_ids,
        expected_filename=expected_filename,
        expected_sha256=expected_sha256,
        sheet_name=sheet_name,
        header_row=optional_int("header_row"),
        unit_row=optional_int("unit_row"),
        data_start_row=data_start_row,
        data_end_row=data_end_row,
        last_channel_column=optional_int("last_channel_column"),
        value_policy=value_policy,
        channel_alignment=channel_alignment,
        phase_rows=phase_rows,
    )
