"""Source invariants for the explicit Diesel analysis configuration only."""
from __future__ import annotations

from typing import TYPE_CHECKING

import numpy as np

from .errors import DataValidationError
from .models import ImportedDataset

if TYPE_CHECKING:
    from .report_profile import ProfileResolutionResult


def validate_diesel_inputs(dataset: ImportedDataset, resolution: ProfileResolutionResult) -> None:
    """Reject unresolved or unsafe inputs before producing Diesel numerical results.

    Missing optional channels are allowed; present channels must have finite data.
    Fuel and distance counters may start above zero, but resets are unsupported.
    Signed engine torque/load and hence mechanical power are deliberately retained.
    """
    if not resolution.is_valid:
        missing = [item.definition.source_name for item in resolution.missing_required]
        ambiguous = [item.definition.source_name for item in resolution.ambiguous]
        units = [item.definition.source_name for item in resolution.unit_mismatches]
        raise DataValidationError(
            f"Diesel source validation failed: missing={missing}; ambiguous={ambiguous}; unit_mismatches={units}"
        )
    values = {}
    for name, item in resolution.resolved.items():
        series = dataset.values[:, dataset.channel_index(item.channel.channel_id)]
        if not np.isfinite(series).all():
            raise DataValidationError(f"Diesel channel '{item.channel.source_name}' contains non-finite values")
        values[name] = series
    time = values["track_time"]
    if time.size < 2 or np.any(np.diff(time) <= 0):
        raise DataValidationError("Diesel Track_Time requires at least two strictly increasing samples")
    for name in ("engine_fuel_consumption", "track_distance"):
        series = values[name]
        if np.any(series < 0) or np.any(np.diff(series) < 0):
            source = resolution.resolved[name].channel.source_name
            raise DataValidationError(f"Diesel cumulative channel '{source}' must be nonnegative and nondecreasing; resets are unsupported")
    for name in ("engine_speed", "fuel_flow"):
        if name in values and np.any(values[name] < 0):
            raise DataValidationError(f"Diesel channel '{resolution.resolved[name].channel.source_name}' must be nonnegative")
