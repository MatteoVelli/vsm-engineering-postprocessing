from __future__ import annotations

import csv
import json
from pathlib import Path

import numpy as np
import pytest

from vsm_postprocessing.errors import ConfigurationError, MathChannelError
from vsm_postprocessing.importer import ImportOptions
from vsm_postprocessing.math_engine import (
    calculate_math_channels,
    export_math_channels,
    load_math_config,
)


def _write_csv(path: Path) -> None:
    path.write_text(
        "Track_Time,Power,Torque\n"
        "s,kW,Nm\n"
        "0,10,2\n"
        "1,20,4\n"
        "2,30,6\n",
        encoding="utf-8",
    )


def _write_config(path: Path, math_block: str, source_channels: str = "    - power__col_002\n") -> None:
    path.write_text(
        "version: 1\n"
        "selection:\n"
        "  include_time: true\n"
        "  export_source_channels:\n"
        f"{source_channels}"
        "constants:\n"
        "  scale: 2\n"
        "math_channels:\n"
        f"{math_block}"
        "output:\n"
        "  data_filename: result.csv\n"
        "  include_units_row: true\n"
        "  float_precision: 12\n",
        encoding="utf-8",
    )


def test_calculates_dependencies_in_topological_order_and_exports(tmp_path: Path) -> None:
    data_path = tmp_path / "data.csv"
    config_path = tmp_path / "math.yaml"
    output_dir = tmp_path / "out"
    _write_csv(data_path)
    _write_config(
        config_path,
        """  - channel_id: calc_accumulated_energy
    display_name: Accumulated Energy
    unit: kWh
    expression: cumulative_sum(calc_sample_energy)
  - channel_id: calc_scaled_power
    display_name: Scaled Power
    unit: kW
    expression: power__col_002 * scale
  - channel_id: calc_sample_energy
    display_name: Sample Energy
    unit: kWh
    expression: sample_energy_kwh(power__col_002, track_time__col_001)
""",
    )

    result = calculate_math_channels(data_path, config_path)
    outputs = export_math_channels(result, output_dir)

    assert result.calculation_order == [
        "calc_sample_energy",
        "calc_accumulated_energy",
        "calc_scaled_power",
    ]
    assert [channel.channel_id for channel in result.output_channels] == [
        "track_time__col_001",
        "power__col_002",
        "calc_accumulated_energy",
        "calc_scaled_power",
        "calc_sample_energy",
    ]
    np.testing.assert_allclose(
        result.calculated_values,
        np.array(
            [
                [10 / 3600, 20, 10 / 3600],
                [30 / 3600, 40, 20 / 3600],
                [60 / 3600, 60, 30 / 3600],
            ]
        ),
    )

    with outputs["output_data"].open(encoding="utf-8", newline="") as handle:
        rows = list(csv.reader(handle))
    assert rows[0][-3:] == ["calc_accumulated_energy", "calc_scaled_power", "calc_sample_energy"]
    assert rows[1][-3:] == ["kWh", "kW", "kWh"]
    manifest = json.loads(outputs["math_manifest"].read_text(encoding="utf-8"))
    assert manifest["math_channel_count"] == 3
    assert manifest["calculation_order"][0] == "calc_sample_energy"


def test_circular_dependencies_are_rejected(tmp_path: Path) -> None:
    data_path = tmp_path / "data.csv"
    config_path = tmp_path / "math.yaml"
    _write_csv(data_path)
    _write_config(
        config_path,
        """  - channel_id: calc_a
    display_name: A
    unit: kW
    expression: calc_b + 1
  - channel_id: calc_b
    display_name: B
    unit: kW
    expression: calc_a + 1
""",
    )

    with pytest.raises(MathChannelError, match="Circular math-channel dependency"):
        calculate_math_channels(data_path, config_path)


def test_unknown_dependency_has_actionable_message(tmp_path: Path) -> None:
    data_path = tmp_path / "data.csv"
    config_path = tmp_path / "math.yaml"
    _write_csv(data_path)
    _write_config(
        config_path,
        """  - channel_id: calc_power
    display_name: Power
    unit: kW
    expression: powr__col_002 * 2
""",
    )

    with pytest.raises(MathChannelError, match="unknown dependencies") as exc_info:
        calculate_math_channels(data_path, config_path)
    assert "power__col_002" in str(exc_info.value)


def test_non_finite_result_from_division_by_zero_is_rejected(tmp_path: Path) -> None:
    data_path = tmp_path / "data.csv"
    config_path = tmp_path / "math.yaml"
    _write_csv(data_path)
    _write_config(
        config_path,
        """  - channel_id: calc_bad
    display_name: Bad
    unit: kW
    expression: power__col_002 / (torque__col_003 - torque__col_003)
""",
    )

    with pytest.raises(MathChannelError, match="produced 3 non-finite values"):
        calculate_math_channels(data_path, config_path)


def test_unsafe_expression_is_rejected_during_configuration_load(tmp_path: Path) -> None:
    config_path = tmp_path / "math.yaml"
    _write_config(
        config_path,
        """  - channel_id: calc_bad
    display_name: Bad
    unit: kW
    expression: __import__('os').system('echo unsafe')
""",
    )

    with pytest.raises(ConfigurationError, match="Unsupported function call"):
        load_math_config(config_path)


def test_required_source_comparison_failure_stops_run(tmp_path: Path) -> None:
    data_path = tmp_path / "data.csv"
    config_path = tmp_path / "math.yaml"
    _write_csv(data_path)
    _write_config(
        config_path,
        """  - channel_id: calc_wrong_power
    display_name: Wrong Power
    unit: kW
    expression: power__col_002 + 1
    compare_to:
      channel_id: power__col_002
      absolute_tolerance: 0
      relative_tolerance: 0
""",
    )

    with pytest.raises(MathChannelError, match="Required math-channel comparison failed"):
        calculate_math_channels(data_path, config_path)


def test_math_channel_id_cannot_collide_with_source_channel(tmp_path: Path) -> None:
    data_path = tmp_path / "data.csv"
    config_path = tmp_path / "math.yaml"
    _write_csv(data_path)
    _write_config(
        config_path,
        """  - channel_id: power__col_002
    display_name: Duplicate
    unit: kW
    expression: power__col_002 * 2
""",
    )

    with pytest.raises(MathChannelError, match="collide with source channel IDs"):
        calculate_math_channels(data_path, config_path)


def test_unknown_configuration_keys_are_rejected(tmp_path: Path) -> None:
    config_path = tmp_path / "math.yaml"
    config_path.write_text(
        """version: 1
selection:
  include_time: true
  export_source_channels: []
math_channels:
  - channel_id: calc_one
    display_name: One
    unit: '-'
    expression: 1
    typo_field: true
""",
        encoding="utf-8",
    )

    with pytest.raises(ConfigurationError, match="Unknown key"):
        load_math_config(config_path)


SOURCE_WORKBOOK = Path(__file__).resolve().parents[1] / "reference_files" / (
    "Sprayer_Caiman_SP_9300Kg_Hybrid_Gen80kW_30kph_74Ht_4000KgAQ_57-4pcSOC_5-80_1C2G_02.xlsx"
)
EXAMPLE_CONFIG = Path(__file__).resolve().parents[1] / "config" / "math_channels_example.yaml"


@pytest.mark.skipif(not SOURCE_WORKBOOK.exists(), reason="Client reference workbook is not present")
def test_supplied_workbook_math_channels_acceptance() -> None:
    result = calculate_math_channels(SOURCE_WORKBOOK, EXAMPLE_CONFIG, ImportOptions(strict=True))

    assert result.sample_count == 1866
    assert result.source_channel_count == 12
    assert result.math_channel_count == 11
    assert result.output_channel_count == 23
    assert len(result.comparisons) == 10
    assert all(comparison.passed for comparison in result.comparisons)
    assert result.output_channels[0].channel_id == "track_time__col_001"
    assert result.calculated_channels[-1].channel_id == "calc_total_generator_power"
    assert result.calculated_values[0, 0] == pytest.approx(0.0)
    assert result.calculated_values[-1, 0] == pytest.approx(1865 / 60)
    assert result.calculated_values[-1, 7] > result.calculated_values[0, 7]
