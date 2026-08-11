from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Any


@dataclass(frozen=True)
class SourceRowRange:
    """Inclusive Excel/source row range aligned to one composed phase."""

    start_row: int
    end_row: int

    @property
    def sample_count(self) -> int:
        return self.end_row - self.start_row + 1


@dataclass(frozen=True)
class ReportRowRange:
    """Inclusive Sergio-reference report row range occupied by one phase."""

    start_row: int
    end_row: int

    @property
    def sample_count(self) -> int:
        return self.end_row - self.start_row + 1


@dataclass(frozen=True)
class DutyCyclePhase:
    """Configuration for one deterministic duty-cycle phase."""

    phase_id: str
    phase_type: str
    output_samples: int
    profile_source: str
    report_rows: ReportRowRange | None = None
    source_rows_aligned: SourceRowRange | None = None
    initial_soc_pct: float | None = None
    generator_start_threshold_pct: float | None = None
    generator_stop_threshold_pct: float | None = None
    drive_controller_initial_state: str | None = None
    metadata: dict[str, Any] = field(default_factory=dict)

    @property
    def has_source_alignment(self) -> bool:
        return self.source_rows_aligned is not None


@dataclass(frozen=True)
class OpportunityChargeConfig:
    """Numerical parameters for one stationary loading/opportunity-charge action."""

    battery_power_kw: float | None = None
    battery_heatflow_kw: float | None = None
    engine_speed_rpm: float | None = None
    engine_torque_nm: float | None = None
    engine_specific_fuel_consumption_g_per_kwh: float | None = None
    fuel_flow_lph: float | None = None
    fuel_density_kg_per_l: float | None = None
    tank_load_rate_kg_s: float | None = None
    low_voltage_aux_kw: float | None = None
    high_voltage_aux_kw: float | None = None
    generator_1_torque_nm: float | None = None
    generator_2_torque_nm: float | None = None


@dataclass(frozen=True)
class DutyCycleScenario:
    """Validated scenario definition loaded from YAML."""

    scenario_id: str
    sample_rate_hz: float
    expected_sample_count: int
    first_timestamp_s: float
    last_timestamp_s: float
    source_data_start_row: int
    phases: tuple[DutyCyclePhase, ...]
    battery_nominal_capacity_kwh: float | None = None
    channel_roles: dict[str, str] = field(default_factory=dict)
    source_restart_zero_channel_roles: tuple[str, ...] = ()
    opportunity_charge: OpportunityChargeConfig = field(default_factory=OpportunityChargeConfig)
    metadata: dict[str, Any] = field(default_factory=dict)

    @property
    def configured_sample_count(self) -> int:
        return sum(phase.output_samples for phase in self.phases)




@dataclass(frozen=True)
class DutyCycleSourceValidation:
    """Compatibility summary between a scenario and one imported source dataset."""

    source_sample_count: int
    required_max_source_sample_index: int
    source_data_start_row: int
    scenario_source_data_start_row: int
    source_nominal_time_step_s: float | None
    scenario_sample_period_s: float

@dataclass
class DutyCyclePartialComposition:
    """Numerically materialised prefix of a scenario, stopped before unresolved phases."""

    scenario: DutyCycleScenario
    channels: list[Any]
    values: Any
    provenance: tuple[Any, ...]
    completed_phase_ids: tuple[str, ...]
    stopped_before_phase_id: str | None

    @property
    def sample_count(self) -> int:
        return int(self.values.shape[0])


@dataclass(frozen=True)
class DutyCycleRowProvenance:
    """Traceability record for exactly one planned output sample."""

    output_sample_index: int
    report_row: int | None
    track_time_s: float
    phase_id: str
    phase_type: str
    phase_local_index: int
    generation_mode: str
    source_row_aligned: int | None
    source_sample_index_aligned: int | None
    profile_definition_resolved: bool
    profile_provider_id: str | None = None
    profile_provider_source: str | None = None
    notes: str = ""

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)




@dataclass
class DutyCycleComposition:
    """Complete numerically materialised duty cycle with profile provenance."""

    scenario: DutyCycleScenario
    channels: list[Any]
    values: Any
    provenance: tuple[DutyCycleRowProvenance, ...]
    completed_phase_ids: tuple[str, ...]
    profile_provenance: tuple[Any, ...] = ()

    @property
    def sample_count(self) -> int:
        return int(self.values.shape[0])

@dataclass(frozen=True)
class DutyCycleCompositionPlan:
    """Deterministic row-level plan generated before numerical composition."""

    scenario: DutyCycleScenario
    rows: tuple[DutyCycleRowProvenance, ...]

    @property
    def sample_count(self) -> int:
        return len(self.rows)

    @property
    def unresolved_phase_ids(self) -> tuple[str, ...]:
        unresolved: list[str] = []
        seen: set[str] = set()
        for row in self.rows:
            if not row.profile_definition_resolved and row.phase_id not in seen:
                seen.add(row.phase_id)
                unresolved.append(row.phase_id)
        return tuple(unresolved)

    def rows_for_phase(self, phase_id: str) -> tuple[DutyCycleRowProvenance, ...]:
        return tuple(row for row in self.rows if row.phase_id == phase_id)
