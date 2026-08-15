from __future__ import annotations

from pathlib import Path

import numpy as np
import pytest

from vsm_postprocessing.errors import MathChannelError
from vsm_postprocessing.importer import ImportOptions, load_data_file
from vsm_postprocessing.math_engine import calculate_math_channels
from vsm_postprocessing.models import ChannelInfo, DataQualityReport, ImportedDataset
from vsm_postprocessing.profile_math import calculate_profile_math_channels
from vsm_postprocessing.report_profile import (
    MathChannelDefinition,
    ProfileMetadata,
    RawChannelDefinition,
    ReportingProfile,
    load_reporting_profile,
)


REFERENCE_CSV = Path(
    "reference_files/RoboSprayer_3500Kg_Electric_12kph_Batt_50kW_Motor_63RPM_Susp_Cool_Rough_Crop_Field_05.csv"
)
ELECTRIC_PROFILE = Path("config/report_profiles/robosprayer_electric.yaml")
HYBRID_PROFILE = Path("config/report_profiles/robosprayer_hybrid.yaml")


def test_profile_math_executes_semantic_raw_dependency() -> None:
    dataset = _dataset([_channel("power__col_002", "Power", "kW")], [[10.0], [20.0], [30.0]])
    profile = _profile(
        raw_channels=[RawChannelDefinition("power", "Power", "Power", "VSM", unit="kW")],
        math_channels=[
            MathChannelDefinition("double_power", "Double Power", "Double Power", "kW", ("power",), "power * 2")
        ],
    )

    result = calculate_profile_math_channels(dataset, profile)

    assert result.calculation_order == ["double_power"]
    np.testing.assert_allclose(result.values_by_semantic_name["double_power"], [20.0, 40.0, 60.0])
    assert result.calculated_channels[0].kind == "math"


def test_profile_math_executes_math_on_math_dependency() -> None:
    dataset = _dataset([_channel("power__col_002", "Power", "kW")], [[10.0], [20.0], [30.0]])
    profile = _profile(
        raw_channels=[RawChannelDefinition("power", "Power", "Power", "VSM", unit="kW")],
        math_channels=[
            MathChannelDefinition("double_power", "Double Power", "Double Power", "kW", ("power",), "power * 2"),
            MathChannelDefinition(
                "quad_power",
                "Quad Power",
                "Quad Power",
                "kW",
                ("double_power",),
                "double_power * 2",
            ),
        ],
    )

    result = calculate_profile_math_channels(dataset, profile)

    assert result.calculation_order == ["double_power", "quad_power"]
    np.testing.assert_allclose(result.values_by_semantic_name["quad_power"], [40.0, 80.0, 120.0])


def test_profile_math_dependency_order_is_topological() -> None:
    dataset = _dataset([_channel("power__col_002", "Power", "kW")], [[10.0], [20.0], [30.0]])
    profile = _profile(
        raw_channels=[RawChannelDefinition("power", "Power", "Power", "VSM", unit="kW")],
        math_channels=[
            MathChannelDefinition("third", "Third", "Third", "kW", ("second",), "second + 1"),
            MathChannelDefinition("second", "Second", "Second", "kW", ("first",), "first + 1"),
            MathChannelDefinition("first", "First", "First", "kW", ("power",), "power + 1"),
        ],
    )

    result = calculate_profile_math_channels(dataset, profile)

    assert result.calculation_order == ["first", "second", "third"]
    np.testing.assert_allclose(result.values_by_semantic_name["third"], [13.0, 23.0, 33.0])


def test_profile_math_dependency_cycle_is_rejected() -> None:
    profile = _profile(
        raw_channels=[],
        math_channels=[
            MathChannelDefinition("first", "First", "First", "kW", ("second",), "second + 1"),
            MathChannelDefinition("second", "Second", "Second", "kW", ("first",), "first + 1"),
        ],
    )

    with pytest.raises(MathChannelError, match="Circular math-channel dependency"):
        calculate_profile_math_channels(_dataset([], [[], [], []]), profile)


def test_profile_math_missing_required_dependency_is_unavailable() -> None:
    profile = _profile(
        raw_channels=[RawChannelDefinition("power", "Power", "Power", "VSM", unit="kW")],
        math_channels=[
            MathChannelDefinition("double_power", "Double Power", "Double Power", "kW", ("power",), "power * 2")
        ],
    )

    result = calculate_profile_math_channels(_dataset([], [[], [], []]), profile)

    assert result.calculated_math_count == 0
    assert [(item.definition.semantic_name, item.missing_dependencies) for item in result.unavailable_required] == [
        ("double_power", ("power",))
    ]


def test_profile_math_optional_unavailable_result_does_not_block_unrelated_math() -> None:
    dataset = _dataset([_channel("power__col_002", "Power", "kW")], [[10.0], [20.0], [30.0]])
    profile = _profile(
        raw_channels=[
            RawChannelDefinition("power", "Power", "Power", "VSM", unit="kW"),
            RawChannelDefinition("optional_force", "Missing Force", "Missing Force", "VSM", unit="N", required=False),
        ],
        math_channels=[
            MathChannelDefinition("double_power", "Double Power", "Double Power", "kW", ("power",), "power * 2"),
            MathChannelDefinition(
                "optional_mass",
                "Optional Mass",
                "Optional Mass",
                "kg",
                ("optional_force",),
                "optional_force / 9.81",
                required=False,
            ),
        ],
    )

    result = calculate_profile_math_channels(dataset, profile)

    assert result.calculated_math_count == 1
    assert result.unavailable_optional[0].definition.semantic_name == "optional_mass"
    np.testing.assert_allclose(result.values_by_semantic_name["double_power"], [20.0, 40.0, 60.0])


def test_profile_math_explicit_constant_zero_channel() -> None:
    profile = _profile(
        raw_channels=[],
        math_channels=[
            MathChannelDefinition("zero_torque", "Zero Torque", "Zero Torque", "Nm", (), "0"),
        ],
    )

    result = calculate_profile_math_channels(_dataset([], [[], [], []]), profile)

    np.testing.assert_allclose(result.values_by_semantic_name["zero_torque"], [0.0, 0.0, 0.0])


def test_profile_math_rpm_nm_to_kw_calculation() -> None:
    dataset = _dataset(
        [
            _channel("speed__col_001", "Speed", "rpm"),
            _channel("torque__col_002", "Torque", "Nm", column=2),
        ],
        [[9548.8, 2.0], [4774.4, 4.0], [0.0, 10.0]],
    )
    profile = _profile(
        raw_channels=[
            RawChannelDefinition("speed", "Speed", "Speed", "VSM", unit="rpm"),
            RawChannelDefinition("torque", "Torque", "Torque", "VSM", unit="Nm"),
        ],
        math_channels=[
            MathChannelDefinition("power", "Power", "Power", "kW", ("speed", "torque"), "speed * torque / rpm_nm_to_kw_divisor")
        ],
    )

    result = calculate_profile_math_channels(dataset, profile)

    np.testing.assert_allclose(result.values_by_semantic_name["power"], [2.0, 2.0, 0.0])


def test_profile_math_sample_time_based_energy() -> None:
    dataset = _dataset(
        [
            _channel("time__col_001", "Track_Time", "s"),
            _channel("power__col_002", "Power", "kW", column=2),
        ],
        [[0.0, 10.0], [2.0, 20.0], [5.0, 30.0]],
    )
    profile = _profile(
        raw_channels=[
            RawChannelDefinition("track_time", "Track_Time", "Time", "VSM", unit="s"),
            RawChannelDefinition("power", "Power", "Power", "VSM", unit="kW"),
        ],
        math_channels=[
            MathChannelDefinition(
                "sample_energy",
                "Sample Energy",
                "Sample Energy",
                "kWh",
                ("power", "track_time"),
                "sample_energy_kwh(power, track_time)",
            )
        ],
    )

    result = calculate_profile_math_channels(dataset, profile)

    np.testing.assert_allclose(result.values_by_semantic_name["sample_energy"], [20 / 3600, 40 / 3600, 90 / 3600])


def test_profile_math_accumulated_energy() -> None:
    dataset = _dataset([_channel("energy__col_001", "Energy", "kWh")], [[1.0], [2.0], [3.0]])
    profile = _profile(
        raw_channels=[RawChannelDefinition("energy", "Energy", "Energy", "VSM", unit="kWh")],
        math_channels=[
            MathChannelDefinition(
                "accumulated_energy",
                "Accumulated Energy",
                "Accumulated Energy",
                "kWh",
                ("energy",),
                "cumulative_sum(energy)",
            )
        ],
    )

    result = calculate_profile_math_channels(dataset, profile)

    np.testing.assert_allclose(result.values_by_semantic_name["accumulated_energy"], [1.0, 3.0, 6.0])


def test_electric_profile_full_math_execution_against_reference_csv() -> None:
    dataset = load_data_file(REFERENCE_CSV, ImportOptions())
    result = calculate_profile_math_channels(dataset, load_reporting_profile(ELECTRIC_PROFILE))

    assert result.configured_math_count == 29
    assert result.calculated_math_count == 28
    assert [(item.definition.semantic_name, item.missing_dependencies) for item in result.unavailable_optional] == [
        ("agrochemical_discharge", ("agrochemical_discharge_force",))
    ]
    assert not result.unavailable_required


def test_electric_profile_representative_numeric_regression() -> None:
    dataset = load_data_file(REFERENCE_CSV, ImportOptions())
    result = calculate_profile_math_channels(dataset, load_reporting_profile(ELECTRIC_PROFILE))

    assert result.values_by_semantic_name["time_minutes"][-1] == pytest.approx(64.2)
    assert result.values_by_semantic_name["distance_km"][-1] == pytest.approx(11.9996)
    assert result.values_by_semantic_name["total_edu_mech_power"].max() == pytest.approx(15.165330139630111)
    assert result.values_by_semantic_name["total_auxiliary_power"][0] == pytest.approx(10.7527)
    assert result.values_by_semantic_name["auxiliary_energy_consumption_accumulated"][-1] == pytest.approx(
        11.508375861111489
    )
    assert np.isfinite(result.calculated_values).all()


def test_hybrid_engine_power_is_zero_against_electric_reference_csv() -> None:
    dataset = load_data_file(REFERENCE_CSV, ImportOptions())
    result = calculate_profile_math_channels(dataset, load_reporting_profile(HYBRID_PROFILE))

    np.testing.assert_allclose(result.values_by_semantic_name["engine_power_required"], 0.0)
    np.testing.assert_allclose(result.values_by_semantic_name["engine_energy_delivered_kwh"], 0.0)


def test_hybrid_generator_power_computes_zero_against_electric_reference_csv() -> None:
    dataset = load_data_file(REFERENCE_CSV, ImportOptions())
    result = calculate_profile_math_channels(dataset, load_reporting_profile(HYBRID_PROFILE))

    assert result.configured_math_count == 32
    assert result.calculated_math_count == 31
    assert [(item.definition.semantic_name, item.missing_dependencies) for item in result.unavailable_optional] == [
        ("agrochemical_discharge", ("agrochemical_discharge_force",))
    ]
    assert not result.unavailable_required
    np.testing.assert_allclose(result.values_by_semantic_name["generator_power_1"], 0.0)


def test_legacy_caiman_math_path_remains_unchanged(tmp_path: Path) -> None:
    data_path = tmp_path / "data.csv"
    config_path = tmp_path / "math.yaml"
    data_path.write_text("Track_Time,Power\ns,kW\n0,10\n1,20\n2,30\n", encoding="utf-8")
    config_path.write_text(
        "version: 1\n"
        "selection:\n"
        "  include_time: true\n"
        "  export_source_channels:\n"
        "    - power__col_002\n"
        "constants: {}\n"
        "math_channels:\n"
        "  - channel_id: calc_power\n"
        "    display_name: Calculated Power\n"
        "    unit: kW\n"
        "    expression: power__col_002 * 2\n"
        "output:\n"
        "  data_filename: result.csv\n",
        encoding="utf-8",
    )

    result = calculate_math_channels(data_path, config_path)

    assert result.calculation_order == ["calc_power"]
    np.testing.assert_allclose(result.calculated_values[:, 0], [20.0, 40.0, 60.0])


def _profile(
    raw_channels: list[RawChannelDefinition],
    math_channels: list[MathChannelDefinition],
) -> ReportingProfile:
    return ReportingProfile(
        version=1,
        metadata=ProfileMetadata(profile_id="test_profile", name="Test Profile"),
        raw_channels=tuple(raw_channels),
        math_channels=tuple(math_channels),
    )


def _channel(channel_id: str, source_name: str, unit: str | None, column: int = 1) -> ChannelInfo:
    return ChannelInfo(
        channel_id=channel_id,
        source_name=source_name,
        display_name=source_name,
        unit=unit,
        source_column_index=column,
        source_column_label=str(column),
        kind="raw",
        dtype="float64",
        provenance="test",
    )


def _dataset(channels: list[ChannelInfo], values: list[list[float]]) -> ImportedDataset:
    array = np.asarray(values, dtype=np.float64)
    if array.ndim == 1:
        array = array.reshape((-1, len(channels)))
    quality = DataQualityReport(
        source_file="test.csv",
        source_sha256="sha",
        file_type="csv",
        sheet_name=None,
        header_row=1,
        unit_row=2,
        data_start_row=3,
        data_end_row=2 + int(array.shape[0]),
        sample_count=int(array.shape[0]),
        channel_count=len(channels),
        raw_channel_count=len(channels),
        math_channel_count=0,
        time_channel_id=None,
        time_channel_name=None,
        time_unit=None,
        time_start=None,
        time_end=None,
        nominal_time_step=None,
        time_is_strictly_increasing=True,
        duplicate_timestamp_count=0,
        missing_cell_count=0,
        invalid_numeric_cell_count=0,
        non_finite_cell_count=0,
    )
    return ImportedDataset(
        source_path=Path("test.csv"),
        channels=channels,
        quality=quality,
        values=array,
    )
