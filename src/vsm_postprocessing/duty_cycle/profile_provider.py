from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Protocol

import numpy as np

from ..errors import ConfigurationError, DataValidationError
from ..importer import ImportOptions, load_data_file
from ..models import ImportedDataset
from ..utils import sha256_file
from .models import DutyCyclePhase, DutyCycleRowProvenance, DutyCycleScenario


@dataclass(frozen=True)
class WorkbookProfileProviderConfig:
    """Portable definition of a workbook-backed phase-profile provider."""

    provider_id: str
    provider_type: str
    phase_ids: tuple[str, ...]
    expected_filename: str | None = None
    expected_sha256: str | None = None
    sheet_name: str | None = None
    header_row: int | None = None
    unit_row: int | None = None
    data_start_row: int | None = None
    data_end_row: int | None = None
    last_channel_column: int | None = None
    value_policy: str = "absolute_reference"
    channel_alignment: str = "channel_id"
    phase_rows: dict[str, tuple[int, int]] | None = None


@dataclass(frozen=True)
class PhaseProfileProvenance:
    """Traceability metadata for one external numerical phase profile."""

    provider_id: str
    phase_id: str
    provider_type: str
    value_policy: str
    source_file: str
    source_sha256: str
    source_start_row: int
    source_end_row: int
    sample_count: int
    channel_count: int


@dataclass(frozen=True)
class ProfileProviderValidation:
    """Validation summary proving a provider is compatible with a scenario/source dataset."""

    provider_id: str
    provider_type: str
    source_file: str
    source_sha256: str
    supported_phase_ids: tuple[str, ...]
    sample_count: int
    channel_count: int
    channel_alignment: str
    channel_ids_match: bool
    channel_layout_compatible: bool
    nominal_time_step_s: float | None


class PhaseProfileProvider(Protocol):
    """Minimal interface required by the duty-cycle composer."""

    @property
    def provider_id(self) -> str: ...

    @property
    def source_file(self) -> str: ...

    def supports(self, phase: DutyCyclePhase) -> bool: ...

    def validate(
        self, scenario: DutyCycleScenario, source: ImportedDataset
    ) -> ProfileProviderValidation: ...

    def materialize_phase(
        self,
        scenario: DutyCycleScenario,
        phase: DutyCyclePhase,
        source: ImportedDataset,
        phase_plan_rows: tuple[DutyCycleRowProvenance, ...] | list[DutyCycleRowProvenance],
    ) -> tuple[np.ndarray, PhaseProfileProvenance]: ...


class WorkbookRowProfileProvider:
    """Supply selected phases from explicit worksheet row ranges.

    The provider is intentionally a *data provider*, not a physics model.  For
    the Sergio fidelity scenario it can replay the unresolved rows from the
    authoritative reference workbook while preserving explicit provenance.
    Future independent VSM road/range-extender workbooks can use the same
    interface without changing ``composer.py``.

    ``absolute_reference`` means the supplied rows already contain mission-level
    time, distance and cumulative quantities and must not be silently offset or
    re-integrated by the provider.
    """

    SUPPORTED_VALUE_POLICIES = {"absolute_reference"}
    SUPPORTED_CHANNEL_ALIGNMENTS = {"channel_id", "column_position"}

    def __init__(self, config: WorkbookProfileProviderConfig, workbook_path: str | Path):
        if config.provider_type not in {"workbook_phase_rows", "workbook_report_rows"}:
            raise ConfigurationError(
                f"Unsupported profile provider type '{config.provider_type}'; expected 'workbook_phase_rows'"
            )
        if config.value_policy not in self.SUPPORTED_VALUE_POLICIES:
            raise ConfigurationError(
                f"Unsupported profile provider value_policy '{config.value_policy}'"
            )
        if config.channel_alignment not in self.SUPPORTED_CHANNEL_ALIGNMENTS:
            raise ConfigurationError(
                f"Unsupported profile provider channel_alignment '{config.channel_alignment}'"
            )
        if config.channel_alignment == "column_position" and config.expected_sha256 is None:
            raise ConfigurationError(
                "column_position profile alignment requires expected_sha256 for deterministic provenance"
            )
        if not config.phase_ids:
            raise ConfigurationError("Profile provider must configure at least one phase_id")

        path = Path(workbook_path).expanduser().resolve()
        if not path.exists() or not path.is_file():
            raise ConfigurationError(f"Profile-provider workbook does not exist: {path}")
        if config.expected_filename is not None and path.name != config.expected_filename:
            raise DataValidationError(
                f"Profile-provider workbook filename '{path.name}' does not match expected '{config.expected_filename}'"
            )
        actual_hash = sha256_file(path)
        if config.expected_sha256 is not None and actual_hash.lower() != config.expected_sha256.lower():
            raise DataValidationError(
                "Profile-provider workbook SHA-256 does not match configuration: "
                f"expected {config.expected_sha256}, got {actual_hash}"
            )

        options = ImportOptions(
            sheet_name=config.sheet_name,
            header_row=config.header_row,
            unit_row=config.unit_row,
            data_start_row=config.data_start_row,
            data_end_row=config.data_end_row,
            last_channel_column=config.last_channel_column,
            strict=True,
        )
        self.config = config
        self.workbook_path = path
        self.dataset = load_data_file(path, options)
        self._sha256 = actual_hash
        self._phase_ids = frozenset(config.phase_ids)

    @property
    def provider_id(self) -> str:
        return self.config.provider_id

    @property
    def source_file(self) -> str:
        return self.workbook_path.name

    @property
    def source_sha256(self) -> str:
        return self._sha256

    def supports(self, phase: DutyCyclePhase) -> bool:
        return phase.phase_id in self._phase_ids

    def _rows_for_phase(self, phase: DutyCyclePhase) -> tuple[int, int]:
        configured = (self.config.phase_rows or {}).get(phase.phase_id)
        if configured is not None:
            return configured
        if phase.report_rows is not None:
            return phase.report_rows.start_row, phase.report_rows.end_row
        raise ConfigurationError(
            f"Phase {phase.phase_id}: workbook profile provider requires an explicit phase_rows mapping "
            "or scenario report_rows"
        )

    def validate(
        self, scenario: DutyCycleScenario, source: ImportedDataset
    ) -> ProfileProviderValidation:
        scenario_ids = {phase.phase_id for phase in scenario.phases}
        unknown = sorted(self._phase_ids - scenario_ids)
        if unknown:
            raise ConfigurationError(
                "Profile provider refers to phase ids not present in scenario: " + ", ".join(unknown)
            )

        source_ids = [channel.channel_id for channel in source.channels]
        provider_ids = [channel.channel_id for channel in self.dataset.channels]
        if len(provider_ids) != len(source_ids):
            raise DataValidationError(
                "Profile-provider channel layout does not match the source dataset "
                f"({len(provider_ids)} provider columns vs {len(source_ids)} source columns)"
            )
        if self.config.channel_alignment == "channel_id" and provider_ids != source_ids:
            first_difference = next(
                (
                    index
                    for index, pair in enumerate(zip(source_ids, provider_ids, strict=True))
                    if pair[0] != pair[1]
                ),
                None,
            )
            assert first_difference is not None
            raise DataValidationError(
                "Profile-provider channel IDs do not match the source dataset; "
                f"first difference at column {first_difference + 1}: "
                f"source='{source_ids[first_difference]}', provider='{provider_ids[first_difference]}'"
            )

        sample_period_s = 1.0 / scenario.sample_rate_hz
        provider_dt = self.dataset.quality.nominal_time_step
        if provider_dt is None:
            raise DataValidationError("Profile-provider dataset has no validated nominal time step")
        if abs(provider_dt - sample_period_s) > 1e-9:
            raise DataValidationError(
                f"Profile-provider nominal time step is {provider_dt}s but scenario requires {sample_period_s}s"
            )

        for phase in scenario.phases:
            if not self.supports(phase):
                continue
            provider_start_row, provider_end_row = self._rows_for_phase(phase)
            start_index = provider_start_row - self.dataset.quality.data_start_row
            end_index = provider_end_row - self.dataset.quality.data_start_row
            if start_index < 0 or end_index >= self.dataset.quality.sample_count:
                raise DataValidationError(
                    f"Phase {phase.phase_id}: provider rows {provider_start_row}:{provider_end_row} "
                    f"fall outside provider data rows {self.dataset.quality.data_start_row}:{self.dataset.quality.data_end_row}"
                )
            if end_index - start_index + 1 != phase.output_samples:
                raise DataValidationError(
                    f"Phase {phase.phase_id}: provider row count does not equal output_samples"
                )

        return ProfileProviderValidation(
            provider_id=self.provider_id,
            provider_type=self.config.provider_type,
            source_file=self.source_file,
            source_sha256=self.source_sha256,
            supported_phase_ids=tuple(
                phase.phase_id for phase in scenario.phases if self.supports(phase)
            ),
            sample_count=self.dataset.quality.sample_count,
            channel_count=self.dataset.quality.channel_count,
            channel_alignment=self.config.channel_alignment,
            channel_ids_match=(provider_ids == source_ids),
            channel_layout_compatible=True,
            nominal_time_step_s=provider_dt,
        )

    def materialize_phase(
        self,
        scenario: DutyCycleScenario,
        phase: DutyCyclePhase,
        source: ImportedDataset,
        phase_plan_rows: tuple[DutyCycleRowProvenance, ...] | list[DutyCycleRowProvenance],
    ) -> tuple[np.ndarray, PhaseProfileProvenance]:
        if not self.supports(phase):
            raise ConfigurationError(
                f"Profile provider '{self.provider_id}' does not support phase {phase.phase_id}"
            )
        provider_start_row, provider_end_row = self._rows_for_phase(phase)
        if len(phase_plan_rows) != phase.output_samples:
            raise ConfigurationError(
                f"Phase {phase.phase_id}: plan contains {len(phase_plan_rows)} rows; expected {phase.output_samples}"
            )

        # Ensure the provider is compatible before returning any numerical data.
        self.validate(scenario, source)

        start_index = provider_start_row - self.dataset.quality.data_start_row
        stop_index = provider_end_row - self.dataset.quality.data_start_row + 1
        values = np.array(self.dataset.values[start_index:stop_index, :], dtype=np.float64, copy=True)
        if values.shape != (phase.output_samples, source.quality.channel_count):
            raise DataValidationError(
                f"Phase {phase.phase_id}: provider returned shape {values.shape}; expected "
                f"({phase.output_samples}, {source.quality.channel_count})"
            )

        # Absolute-reference providers must already agree with the scenario's
        # global timeline.  We verify rather than silently rewriting it.
        track_time_id = scenario.channel_roles.get("track_time_s")
        if not track_time_id:
            raise ConfigurationError("Duty-cycle scenario has no track_time_s channel role")
        track_index = source.channel_index(track_time_id)
        planned_time = np.asarray([row.track_time_s for row in phase_plan_rows], dtype=np.float64)
        if not np.allclose(values[:, track_index], planned_time, rtol=0.0, atol=1e-9):
            max_error = float(np.max(np.abs(values[:, track_index] - planned_time)))
            raise DataValidationError(
                f"Phase {phase.phase_id}: provider timeline does not match composition plan; max error {max_error}s"
            )

        provenance = PhaseProfileProvenance(
            provider_id=self.provider_id,
            phase_id=phase.phase_id,
            provider_type=self.config.provider_type,
            value_policy=self.config.value_policy,
            source_file=self.source_file,
            source_sha256=self.source_sha256,
            source_start_row=provider_start_row,
            source_end_row=provider_end_row,
            sample_count=phase.output_samples,
            channel_count=source.quality.channel_count,
        )
        return values, provenance


def provider_supports_all(
    provider: PhaseProfileProvider | None,
    phases: Iterable[DutyCyclePhase],
) -> bool:
    if provider is None:
        return False
    return all(provider.supports(phase) for phase in phases)
