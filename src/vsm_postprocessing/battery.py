from __future__ import annotations

import math
from typing import Sequence

import numpy as np


SOC_FRACTION_EPSILON = 1e-9


def nominal_capacity_estimates_kwh(
    battery_energy_kwh: Sequence[float] | np.ndarray,
    battery_soc_percent: Sequence[float] | np.ndarray,
    *,
    soc_epsilon: float = SOC_FRACTION_EPSILON,
) -> np.ndarray:
    energy = np.asarray(battery_energy_kwh, dtype=np.float64)
    soc = np.asarray(battery_soc_percent, dtype=np.float64)
    if energy.shape != soc.shape:
        raise ValueError(f"Battery Energy and SOC must have identical shapes; got {energy.shape} and {soc.shape}")
    if energy.ndim != 1:
        raise ValueError("Battery capacity inference expects one-dimensional channels")
    soc_fraction = soc / 100.0
    valid = np.isfinite(energy) & np.isfinite(soc_fraction) & (soc_fraction > soc_epsilon)
    if not valid.any():
        raise ValueError("Battery capacity inference found no finite samples with SOC above zero")
    estimates = energy[valid] / soc_fraction[valid]
    estimates = estimates[np.isfinite(estimates)]
    if estimates.size == 0:
        raise ValueError("Battery capacity inference produced no finite capacity estimates")
    return estimates


def infer_nominal_battery_capacity_kwh(
    battery_energy_kwh: Sequence[float] | np.ndarray,
    battery_soc_percent: Sequence[float] | np.ndarray,
) -> float:
    estimates = nominal_capacity_estimates_kwh(battery_energy_kwh, battery_soc_percent)
    value = float(np.median(estimates))
    if not math.isfinite(value):
        raise ValueError("Battery capacity inference produced a non-finite median")
    return value


def max_charging_power_kw(battery_power_kw: Sequence[float] | np.ndarray) -> float:
    power = np.asarray(battery_power_kw, dtype=np.float64)
    finite = power[np.isfinite(power)]
    if finite.size == 0:
        raise ValueError("Battery charging-power statistic found no finite samples")
    return float(np.max(np.maximum(finite, 0.0)))
