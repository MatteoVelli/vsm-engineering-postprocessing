from __future__ import annotations

import math
from pathlib import Path

import numpy as np
import pytest

from vsm_postprocessing.errors import ConfigurationError, StatisticsError
from vsm_postprocessing.importer import ImportOptions
from vsm_postprocessing.statistics_engine import (
    calculate_statistics,
    compute_statistic,
    load_statistics_config,
)
from conftest import (
    CAIMAN_PROFILE_REFERENCE_DESCRIPTION,
    CAIMAN_PROFILE_REFERENCE_XLSX,
    CAIMAN_REFERENCE_DESCRIPTION,
    CAIMAN_REFERENCE_XLSX,
    require_private_reference_file,
)


def _write_csv(path: Path) -> None:
    path.write_text(
        "Time,Power,Torque\n"
        "s,kW,Nm\n"
        "0,3,10\n"
        "1,4,20\n"
        "3,0,30\n",
        encoding="utf-8",
    )


def _write_statistics_config(path: Path, definitions: str) -> None:
    path.write_text(
        f"""version: 1
statistics:
{definitions}
output:
  results_filename: results.csv
  wide_filename: wide.csv
  float_precision: 12
""",
        encoding="utf-8",
    )


def test_basic_statistics_are_calculated_deterministically(tmp_path: Path) -> None:
    data_path = tmp_path / "data.csv"
    config_path = tmp_path / "statistics.yaml"
    _write_csv(data_path)
    _write_statistics_config(
        config_path,
        """  - statistic_id: power_rms
    channel_id: power__col_002
    operation: rms
    placement_group: top_rms
  - statistic_id: power_max
    channel_id: power__col_002
    operation: max
    placement_group: bottom_channel
  - statistic_id: power_min
    channel_id: power__col_002
    operation: min
    placement_group: bottom_channel
  - statistic_id: power_last
    channel_id: power__col_002
    operation: last
    placement_group: bottom_channel
  - statistic_id: power_sum
    channel_id: power__col_002
    operation: sum
    placement_group: kpi_block
""",
    )

    result = calculate_statistics(data_path, config_path)
    values = {item.statistic_id: item.value for item in result.statistics}

    assert result.sample_count == 3
    assert values["power_rms"] == pytest.approx(math.sqrt(25 / 3))
    assert values["power_max"] == pytest.approx(4.0)
    assert values["power_min"] == pytest.approx(0.0)
    assert values["power_last"] == pytest.approx(0.0)
    assert values["power_sum"] == pytest.approx(7.0)


def test_time_weighted_rms_uses_trapezoidal_integration(tmp_path: Path) -> None:
    data_path = tmp_path / "data.csv"
    config_path = tmp_path / "statistics.yaml"
    _write_csv(data_path)
    _write_statistics_config(
        config_path,
        """  - statistic_id: power_time_weighted_rms
    channel_id: power__col_002
    operation: time_weighted_rms
    placement_group: top_rms
""",
    )

    result = calculate_statistics(data_path, config_path)
    expected = math.sqrt(np.trapezoid(np.square([3.0, 4.0, 0.0]), [0.0, 1.0, 3.0]) / 3.0)
    assert result.statistics[0].value == pytest.approx(expected)


def test_nan_policy_omit_removes_non_finite_samples() -> None:
    value, used, omitted = compute_statistic([1.0, math.nan, 3.0], "rms", "omit")
    assert value == pytest.approx(math.sqrt(5.0))
    assert used == 2
    assert omitted == 1


def test_nan_policy_error_rejects_non_finite_samples() -> None:
    with pytest.raises(StatisticsError, match="encountered 1 non-finite values"):
        compute_statistic([1.0, math.nan, 3.0], "max", "error")


def test_nan_policy_propagate_returns_nan() -> None:
    value, used, omitted = compute_statistic([1.0, math.nan, 3.0], "sum", "propagate")
    assert math.isnan(value)
    assert used == 3
    assert omitted == 0


def test_missing_channel_reports_suggestions(tmp_path: Path) -> None:
    data_path = tmp_path / "data.csv"
    config_path = tmp_path / "statistics.yaml"
    _write_csv(data_path)
    _write_statistics_config(
        config_path,
        """  - statistic_id: missing
    channel_id: powr__col_002
    operation: max
    placement_group: bottom_channel
""",
    )

    with pytest.raises(StatisticsError, match="Configured statistic channel IDs were not found") as exc_info:
        calculate_statistics(data_path, config_path)
    assert "power__col_002" in str(exc_info.value)


def test_duplicate_statistic_ids_are_rejected(tmp_path: Path) -> None:
    config_path = tmp_path / "statistics.yaml"
    _write_statistics_config(
        config_path,
        """  - statistic_id: duplicate
    channel_id: power__col_002
    operation: max
    placement_group: bottom_channel
  - statistic_id: duplicate
    channel_id: power__col_002
    operation: min
    placement_group: bottom_channel
""",
    )

    with pytest.raises(ConfigurationError, match="duplicate statistic IDs"):
        load_statistics_config(config_path)


def test_unknown_configuration_keys_are_rejected(tmp_path: Path) -> None:
    config_path = tmp_path / "statistics.yaml"
    _write_statistics_config(
        config_path,
        """  - statistic_id: power_max
    channel_id: power__col_002
    operation: max
    placement_group: bottom_channel
    typo_field: true
""",
    )

    with pytest.raises(ConfigurationError, match="Unknown key"):
        load_statistics_config(config_path)


def test_invalid_operation_is_rejected(tmp_path: Path) -> None:
    config_path = tmp_path / "statistics.yaml"
    _write_statistics_config(
        config_path,
        """  - statistic_id: invalid
    channel_id: power__col_002
    operation: average
    placement_group: bottom_channel
""",
    )

    with pytest.raises(ConfigurationError, match="operation must be one of"):
        load_statistics_config(config_path)


PROJECT_ROOT = Path(__file__).resolve().parents[1]
MATH_CONFIG = PROJECT_ROOT / "config" / "math_channels_example.yaml"
STATISTICS_CONFIG = PROJECT_ROOT / "config" / "statistics_example.yaml"
REPORT_STATISTICS_CONFIG = PROJECT_ROOT / "config" / "statistics_reference_report.yaml"


def test_supplied_source_workbook_statistics_acceptance() -> None:
    source_workbook = require_private_reference_file(CAIMAN_REFERENCE_XLSX, CAIMAN_REFERENCE_DESCRIPTION)
    result = calculate_statistics(
        source_workbook,
        STATISTICS_CONFIG,
        ImportOptions(strict=True),
        math_config_file=MATH_CONFIG,
    )

    assert result.sample_count == 1866
    assert result.statistic_count == 14
    assert len(result.channels_by_id) == 83
    values = {item.statistic_id: item.value for item in result.statistics}
    assert values["chassis_speed_max"] == pytest.approx(30.0717)
    assert values["battery_power_min"] == pytest.approx(-149.18)
    assert values["vehicle_distance_last"] == pytest.approx(12000.0)
    assert values["auxiliary_energy_sum"] == pytest.approx(values["auxiliary_energy_accumulated_last"])


def test_supplied_report_statistics_reference_acceptance() -> None:
    report_workbook = require_private_reference_file(
        CAIMAN_PROFILE_REFERENCE_XLSX,
        CAIMAN_PROFILE_REFERENCE_DESCRIPTION,
    )
    result = calculate_statistics(
        report_workbook,
        REPORT_STATISTICS_CONFIG,
        ImportOptions(
            header_row=3,
            unit_row=4,
            data_start_row=5,
            data_end_row=17422,
            last_channel_column=70,
            strict=True,
        ),
    )

    assert result.sample_count == 17418
    assert result.statistic_count == 31
    assert result.comparison_count == 31
    assert result.required_comparisons_passed
    assert all(item.comparison is not None and item.comparison.passed for item in result.statistics)
    rms = {item.statistic_id: item for item in result.statistics}["report_battery_power_rms"]
    assert rms.value == pytest.approx(55.131048462310204)
    assert rms.comparison is not None
    assert rms.comparison.absolute_error == pytest.approx(0.0015826565179608565)
