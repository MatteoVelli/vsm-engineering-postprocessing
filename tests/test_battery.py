from __future__ import annotations

import numpy as np
import pytest

from vsm_postprocessing.battery import (
    infer_nominal_battery_capacity_kwh,
    max_charging_power_kw,
    nominal_capacity_estimates_kwh,
)


def test_nominal_capacity_infers_50_kwh_from_80_percent_soc() -> None:
    assert infer_nominal_battery_capacity_kwh([40.0], [80.0]) == pytest.approx(50.0)


def test_nominal_capacity_infers_100_kwh_from_57_4_percent_soc() -> None:
    assert infer_nominal_battery_capacity_kwh([57.4], [57.4]) == pytest.approx(100.0)


def test_nominal_capacity_is_not_tied_to_fixed_initial_soc() -> None:
    assert infer_nominal_battery_capacity_kwh([35.0], [70.0]) == pytest.approx(50.0)


def test_nominal_capacity_ignores_invalid_zero_soc_and_nan_samples() -> None:
    estimates = nominal_capacity_estimates_kwh(
        [10.0, np.nan, 40.0, 35.0],
        [0.0, 80.0, 80.0, 70.0],
    )

    np.testing.assert_allclose(estimates, [50.0, 50.0])
    assert infer_nominal_battery_capacity_kwh([10.0, np.nan, 40.0, 35.0], [0.0, 80.0, 80.0, 70.0]) == pytest.approx(50.0)


def test_nominal_capacity_uses_median_across_consistent_samples() -> None:
    assert infer_nominal_battery_capacity_kwh([40.0, 37.5, 35.0], [80.0, 75.0, 70.0]) == pytest.approx(50.0)


def test_max_charging_power_is_zero_when_no_positive_samples_exist() -> None:
    assert max_charging_power_kw([-40.0, -30.0, -20.0]) == pytest.approx(0.0)


def test_max_charging_power_uses_positive_charging_samples() -> None:
    assert max_charging_power_kw([-30.0, 5.0, 12.0, -10.0]) == pytest.approx(12.0)
