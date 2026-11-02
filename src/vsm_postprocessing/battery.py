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
    valid = (np.isfinite(energy) & np.isfinite(soc_fraction) & (energy >= 0)
             & (soc_fraction > soc_epsilon) & (soc_fraction <= 1))
    if not valid.any():
        raise ValueError("Battery capacity inference found no finite nonnegative energy samples with SOC in (0, 100]")
    estimates = energy[valid] / soc_fraction[valid]
    estimates = estimates[np.isfinite(estimates)]
    if estimates.size == 0:
        raise ValueError("Battery capacity inference produced no finite capacity estimates")
    return estimates


def infer_nominal_battery_capacity_kwh(
    battery_energy_kwh: Sequence[float] | np.ndarray,
    battery_soc_percent: Sequence[float] | np.ndarray,
) -> float:
    # Legacy expression name: this is an energy-equivalent approximation,
    # not a measurement of the nominal battery specification.
    estimates = nominal_capacity_estimates_kwh(battery_energy_kwh, battery_soc_percent)
    value = float(np.median(estimates))
    if not math.isfinite(value):
        raise ValueError("Battery capacity inference produced a non-finite median")
    return value


def battery_capacity_reference(energy, soc, inputs):
    """Keep the declared nominal reference distinct from Energy/SOC estimates."""
    nominal = inputs.battery_nominal_capacity_kwh
    if energy is None and soc is None and nominal is None:
        return {}
    record = {
        'nominal_capacity_kwh': nominal,
        'nominal_capacity_status': inputs.battery_nominal_capacity_status,
        'nominal_capacity_provenance': inputs.battery_nominal_capacity_provenance,
        'reference_capacity_kwh': nominal,
        'reference_status': inputs.battery_nominal_capacity_status if nominal is not None else 'unavailable',
        'range_energy_fraction': 0.85,
        'range_policy': '85% of reference nominal energy / observed net stored-energy depletion per km',
        'limitations': 'Reference-energy extrapolation, not a validated usable SOC window; no reserve, cutoff or temperature law is assumed.',
    }
    if energy is not None and soc is not None:
        try:
            estimates = nominal_capacity_estimates_kwh(energy, soc)
        except ValueError as exc:
            record['inference_unavailable_reason'] = str(exc)
        else:
            median = float(np.median(estimates))
            record['energy_soc_equivalent_kwh'] = {
                'median': median, 'minimum': float(estimates.min()), 'maximum': float(estimates.max()),
                'max_deviation': float(np.max(np.abs(estimates-median))),
                'tolerance': max(0.1, abs(median)*0.01), 'valid_samples': len(estimates),
                'tolerance_basis': 'Legacy diagnostic heuristic: max(0.1 kWh, 1% of median); not a physical acceptance limit or KPI reliability criterion',
                'excluded_samples': len(energy)-len(estimates),
                'interpretation': 'Energy / SOC fraction; not an independently confirmed nominal capacity',
            }
            if nominal is None:
                record.update(reference_capacity_kwh=median, reference_status='inferred')
        if len(energy) and nominal is not None:
            record['initial_energy_minus_nominal_times_soc_kwh'] = float(energy[0]-nominal*soc[0]/100)
    return record


def max_charging_power_kw(battery_power_kw: Sequence[float] | np.ndarray) -> float:
    power = np.asarray(battery_power_kw, dtype=np.float64)
    finite = power[np.isfinite(power)]
    if finite.size == 0:
        raise ValueError("Battery charging-power statistic found no finite samples")
    return float(np.max(np.maximum(finite, 0.0)))
