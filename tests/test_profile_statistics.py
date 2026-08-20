from __future__ import annotations

import math
from pathlib import Path

import numpy as np
import pytest

from vsm_postprocessing.importer import ImportOptions, load_data_file
from vsm_postprocessing.math_engine import calculate_math_channels
from vsm_postprocessing.models import ChannelInfo, DataQualityReport, ImportedDataset
from vsm_postprocessing.profile_math import calculate_profile_math_channels
from vsm_postprocessing.profile_statistics import calculate_profile_statistics
from vsm_postprocessing.report_profile import (
    KPIDefinition,
    MathChannelDefinition,
    ProfileMetadata,
    RawChannelDefinition,
    ReportingProfile,
    StatisticDefinition,
    load_reporting_profile,
    resolve_profile,
)
from vsm_postprocessing.statistics_engine import calculate_statistics

from conftest import (
    ROBOSPRAYER_LATEST_ELECTRIC_CSV,
    ROBOSPRAYER_LATEST_ELECTRIC_DESCRIPTION,
    ROBOSPRAYER_LATEST_HYBRID_CSV,
    ROBOSPRAYER_LATEST_HYBRID_DESCRIPTION,
    ROBOSPRAYER_REFERENCE_CSV,
    ROBOSPRAYER_REFERENCE_DESCRIPTION,
    require_private_reference_file,
)
ELECTRIC_PROFILE = Path("config/report_profiles/robosprayer_electric.yaml")
HYBRID_PROFILE = Path("config/report_profiles/robosprayer_hybrid.yaml")


def _robosprayer_csv() -> Path:
    return require_private_reference_file(ROBOSPRAYER_REFERENCE_CSV, ROBOSPRAYER_REFERENCE_DESCRIPTION)


def _latest_electric_csv() -> Path:
    return require_private_reference_file(ROBOSPRAYER_LATEST_ELECTRIC_CSV, ROBOSPRAYER_LATEST_ELECTRIC_DESCRIPTION)


def _latest_hybrid_csv() -> Path:
    return require_private_reference_file(ROBOSPRAYER_LATEST_HYBRID_CSV, ROBOSPRAYER_LATEST_HYBRID_DESCRIPTION)


def test_profile_statistic_uses_semantic_raw_channel() -> None:
    dataset = _dataset([_channel("power__col_002", "Power", "kW")], [[1.0], [3.0], [2.0]])
    profile = _profile(
        raw_channels=[RawChannelDefinition("power", "Power", "Power", "VSM", unit="kW")],
        statistics=[StatisticDefinition("power_max", "power", "max")],
    )

    result = calculate_profile_statistics(dataset, profile)

    assert result.statistics[0].value == pytest.approx(3.0)
    assert result.statistics[0].target_channel == "power"


def test_profile_statistic_uses_semantic_math_channel() -> None:
    dataset = _dataset([_channel("power__col_002", "Power", "kW")], [[1.0], [3.0], [2.0]])
    profile = _profile(
        raw_channels=[RawChannelDefinition("power", "Power", "Power", "VSM", unit="kW")],
        math_channels=[MathChannelDefinition("double_power", "Double Power", "Double Power", "kW", ("power",), "power * 2")],
        statistics=[StatisticDefinition("double_power_last", "double_power", "last")],
    )

    result = calculate_profile_statistics(dataset, profile)

    assert result.statistics[0].value == pytest.approx(4.0)
    assert result.statistics[0].channel_kind == "math"


def test_profile_statistics_operations_cover_rms_max_min_first_last_sum() -> None:
    dataset = _dataset([_channel("power__col_002", "Power", "kW")], [[1.0], [3.0], [2.0]])
    profile = _profile(
        raw_channels=[RawChannelDefinition("power", "Power", "Power", "VSM", unit="kW")],
        statistics=[
            StatisticDefinition("power_rms", "power", "rms"),
            StatisticDefinition("power_max", "power", "max"),
            StatisticDefinition("power_min", "power", "min"),
            StatisticDefinition("power_first", "power", "first"),
            StatisticDefinition("power_last", "power", "last"),
            StatisticDefinition("power_sum", "power", "sum"),
        ],
    )

    result = calculate_profile_statistics(dataset, profile)
    values = {item.definition.statistic_id: item.value for item in result.statistics}

    assert values["power_rms"] == pytest.approx(math.sqrt(14 / 3))
    assert values["power_max"] == pytest.approx(3.0)
    assert values["power_min"] == pytest.approx(1.0)
    assert values["power_first"] == pytest.approx(1.0)
    assert values["power_last"] == pytest.approx(2.0)
    assert values["power_sum"] == pytest.approx(6.0)


def test_profile_statistics_positive_max_reports_zero_when_no_charging() -> None:
    dataset = _dataset([_channel("power__col_002", "Battery Power", "kW")], [[-40.0], [-30.0], [-20.0]])
    profile = _profile(
        raw_channels=[RawChannelDefinition("battery_power", "Battery Power", "Battery Power", "VSM", unit="kW")],
        statistics=[
            StatisticDefinition(
                "battery_power_max",
                "battery_power",
                "positive_max",
                display_name="Max Battery Charging Power",
                unit="kW",
            )
        ],
    )

    result = calculate_profile_statistics(dataset, profile)

    assert result.statistics[0].value == pytest.approx(0.0)


def test_profile_statistics_positive_max_uses_positive_charging_samples() -> None:
    dataset = _dataset([_channel("power__col_002", "Battery Power", "kW")], [[-30.0], [5.0], [12.0], [-10.0]])
    profile = _profile(
        raw_channels=[RawChannelDefinition("battery_power", "Battery Power", "Battery Power", "VSM", unit="kW")],
        statistics=[
            StatisticDefinition(
                "battery_power_max",
                "battery_power",
                "positive_max",
                display_name="Max Battery Charging Power",
                unit="kW",
            )
        ],
    )

    result = calculate_profile_statistics(dataset, profile)

    assert result.statistics[0].value == pytest.approx(12.0)


def test_profile_kpi_uses_statistic_dependencies() -> None:
    dataset = _dataset([_channel("energy__col_001", "Energy", "kWh")], [[40.0], [25.0], [10.0]])
    profile = _profile(
        raw_channels=[RawChannelDefinition("energy", "Energy", "Energy", "VSM", unit="kWh")],
        statistics=[
            StatisticDefinition("energy_first", "energy", "first"),
            StatisticDefinition("energy_last", "energy", "last"),
        ],
        kpis=[
            KPIDefinition(
                kpi_id="energy_used",
                expression="energy_first - energy_last",
                dependencies=("energy_first", "energy_last"),
                unit="kWh",
            )
        ],
    )

    result = calculate_profile_statistics(dataset, profile)

    assert result.kpis[0].value == pytest.approx(30.0)


def test_profile_kpi_nominal_capacity_uses_actual_energy_soc_arrays() -> None:
    result = calculate_profile_statistics(
        _battery_profile_dataset([40.0, 37.5, 35.0], [80.0, 75.0, 70.0]),
        _battery_capacity_profile(),
    )
    kpis = {item.definition.kpi_id: item.value for item in result.kpis}

    assert kpis["battery_capacity_100"] == pytest.approx(50.0)
    assert not result.diagnostics


def test_profile_kpi_nominal_capacity_ignores_invalid_soc_samples() -> None:
    result = calculate_profile_statistics(
        _battery_profile_dataset([10.0, np.nan, 40.0, 35.0], [0.0, 80.0, 80.0, 70.0]),
        _battery_capacity_profile(),
    )
    kpis = {item.definition.kpi_id: item.value for item in result.kpis}

    assert kpis["battery_capacity_100"] == pytest.approx(50.0)


def test_profile_kpi_nominal_capacity_emits_consistency_diagnostic() -> None:
    result = calculate_profile_statistics(
        _battery_profile_dataset([40.0, 42.0, 35.0], [80.0, 75.0, 70.0]),
        _battery_capacity_profile(),
    )

    assert result.diagnostics
    assert "Nominal battery capacity estimates vary" in result.diagnostics[0]


def test_range_85_uses_nominal_capacity_not_fixed_starting_soc() -> None:
    profile = _battery_range_profile()
    run_a = calculate_profile_statistics(
        _battery_profile_dataset([40.0, 35.0], [80.0, 70.0], distance_km=1.0),
        profile,
    )
    run_b = calculate_profile_statistics(
        _battery_profile_dataset([47.5, 42.5], [95.0, 85.0], distance_km=1.0),
        profile,
    )
    kpis_a = {item.definition.kpi_id: item.value for item in run_a.kpis}
    kpis_b = {item.definition.kpi_id: item.value for item in run_b.kpis}

    assert kpis_a["battery_capacity_100"] == pytest.approx(50.0)
    assert kpis_b["battery_capacity_100"] == pytest.approx(50.0)
    assert kpis_a["range_85_battery_km"] == pytest.approx(8.5)
    assert kpis_b["range_85_battery_km"] == pytest.approx(kpis_a["range_85_battery_km"])


def test_optional_unavailable_statistic_does_not_block_unrelated_results() -> None:
    dataset = _dataset([_channel("power__col_002", "Power", "kW")], [[1.0], [3.0], [2.0]])
    profile = _profile(
        raw_channels=[
            RawChannelDefinition("power", "Power", "Power", "VSM", unit="kW"),
            RawChannelDefinition("optional_force", "Force", "Force", "VSM", unit="N", required=False),
        ],
        statistics=[
            StatisticDefinition("power_max", "power", "max"),
            StatisticDefinition("force_max", "optional_force", "max", required=False),
        ],
    )

    result = calculate_profile_statistics(dataset, profile)

    assert result.is_complete
    assert result.calculated_statistic_count == 1
    assert [(item.definition.statistic_id, item.reason) for item in result.unavailable_optional_statistics] == [
        ("force_max", "target unavailable: optional_force")
    ]


def test_required_missing_statistic_marks_result_incomplete() -> None:
    profile = _profile(
        raw_channels=[RawChannelDefinition("power", "Power", "Power", "VSM", unit="kW")],
        statistics=[StatisticDefinition("power_max", "power", "max")],
    )

    result = calculate_profile_statistics(_dataset([], [[], [], []]), profile)

    assert not result.is_complete
    assert [(item.definition.statistic_id, item.reason) for item in result.unavailable_required_statistics] == [
        ("power_max", "target unavailable: power")
    ]


def test_profile_statistics_are_independent_from_runtime_column_position() -> None:
    profile = _profile(
        raw_channels=[RawChannelDefinition("speed", "Speed", "Speed", "VSM", unit="kph")],
        statistics=[StatisticDefinition("speed_max", "speed", "max")],
    )
    first = calculate_profile_statistics(
        _dataset([_channel("speed__col_001", "Speed", "kph")], [[1.0], [4.0], [2.0]]),
        profile,
    )
    reordered = calculate_profile_statistics(
        _dataset([_channel("speed__col_039", "Speed", "kph", column=39)], [[1.0], [4.0], [2.0]]),
        profile,
    )

    assert first.statistics[0].value == reordered.statistics[0].value == pytest.approx(4.0)


def test_electric_profile_full_statistics_execution_against_reference_csv() -> None:
    dataset = load_data_file(_robosprayer_csv(), ImportOptions())
    result = calculate_profile_statistics(dataset, load_reporting_profile(ELECTRIC_PROFILE))

    assert result.configured_statistic_count == 27
    assert result.calculated_statistic_count == 27
    assert result.configured_kpi_count == 9
    assert result.calculated_kpi_count == 9
    assert result.is_complete
    assert not result.unavailable_optional_statistics
    assert not result.unavailable_required_statistics


def test_electric_profile_representative_statistics_regression() -> None:
    dataset = load_data_file(_robosprayer_csv(), ImportOptions())
    result = calculate_profile_statistics(dataset, load_reporting_profile(ELECTRIC_PROFILE))
    stats = {item.definition.statistic_id: item.value for item in result.statistics}
    kpis = {item.definition.kpi_id: item.value for item in result.kpis}

    assert stats["time_minutes_last"] == pytest.approx(64.2)
    assert stats["distance_km_last"] == pytest.approx(11.9996)
    assert stats["battery_soc_last"] == pytest.approx(12.6927)
    assert stats["battery_power_rms"] == pytest.approx(28.716636645770492)
    assert stats["battery_heatflow_rms"] == pytest.approx(2.840107066457429)
    assert stats["chassis_speed_max"] == pytest.approx(11.9858)
    assert stats["total_edu_mech_power_max"] == pytest.approx(15.165330139630111)
    assert stats["total_edu_elect_power_min"] == pytest.approx(-19.101329999999997)
    assert stats["total_rolling_resistance_power_max"] == pytest.approx(13.471919999999999)
    assert stats["agrochemical_discharge_max"] == pytest.approx(0.0)
    assert stats["auxiliary_energy_accumulated_last"] == pytest.approx(11.508375861111489)
    assert stats["tyre_rr_energy_accumulated_last"] == pytest.approx(13.355223967066818)
    assert kpis["battery_capacity_used"] == pytest.approx(33.65367)
    assert kpis["battery_energy_consumption_wh_per_km"] == pytest.approx(2804.565985532851)
    assert kpis["battery_capacity_100"] == pytest.approx(50.0)
    assert kpis["range_85_battery_km"] == pytest.approx(15.153859891060916)


def test_electric_rms_does_not_reproduce_stale_17417_denominator() -> None:
    dataset = load_data_file(_robosprayer_csv(), ImportOptions())
    result = calculate_profile_statistics(dataset, load_reporting_profile(ELECTRIC_PROFILE))
    stats = {item.definition.statistic_id: item.value for item in result.statistics}

    assert stats["battery_power_rms"] == pytest.approx(28.716636645770492)
    assert stats["battery_power_rms"] != pytest.approx(13.506611297860566)
    assert stats["battery_heatflow_rms"] == pytest.approx(2.840107066457429)
    assert stats["battery_heatflow_rms"] != pytest.approx(1.3358187681981721)


def test_hybrid_engine_statistics_are_zero_against_electric_reference_csv() -> None:
    dataset = load_data_file(_robosprayer_csv(), ImportOptions())
    result = calculate_profile_statistics(dataset, load_reporting_profile(HYBRID_PROFILE))
    stats = {item.definition.statistic_id: item.value for item in result.statistics}

    assert stats["engine_fuel_consumption_last"] == pytest.approx(0.0)
    assert stats["fuel_flow_max"] == pytest.approx(0.0)
    assert stats["engine_speed_max"] == pytest.approx(0.0)
    assert stats["engine_torque_max"] == pytest.approx(0.0)
    assert stats["engine_power_required_max"] == pytest.approx(0.0)
    assert stats["engine_energy_delivered_sum"] == pytest.approx(0.0)


def test_hybrid_generator_statistics_are_zero_against_electric_reference_csv() -> None:
    dataset = load_data_file(_robosprayer_csv(), ImportOptions())
    result = calculate_profile_statistics(dataset, load_reporting_profile(HYBRID_PROFILE))
    stats = {item.definition.statistic_id: item.value for item in result.statistics}

    assert result.configured_statistic_count == 36
    assert result.calculated_statistic_count == 36
    assert result.configured_kpi_count == 9
    assert result.calculated_kpi_count == 9
    assert result.is_complete
    assert not result.unavailable_optional_statistics
    assert not result.unavailable_required_statistics
    assert stats["agrochemical_discharge_max"] == pytest.approx(0.0)
    assert stats["generator_torque_1_max"] == pytest.approx(0.0)
    assert stats["generator_power_1_max"] == pytest.approx(0.0)


def test_latest_electric_profile_derives_nominal_capacity_range_and_zero_charging_power() -> None:
    dataset = load_data_file(_latest_electric_csv(), ImportOptions(strict=True))
    result = calculate_profile_statistics(dataset, load_reporting_profile(ELECTRIC_PROFILE))
    stats = {item.definition.statistic_id: item.value for item in result.statistics}
    kpis = {item.definition.kpi_id: item.value for item in result.kpis}

    assert stats["battery_energy_first"] == pytest.approx(40.0)
    assert stats["battery_soc_first"] == pytest.approx(80.0)
    assert stats["battery_power_min"] == pytest.approx(-47.9221)
    assert stats["battery_power_max"] == pytest.approx(0.0)
    assert kpis["battery_capacity_100"] == pytest.approx(50.0)
    assert kpis["usable_battery_capacity_85"] == pytest.approx(42.5)
    assert kpis["battery_energy_consumption_wh_per_km"] == pytest.approx(3778.0829156591767)
    assert kpis["range_85_battery_km"] == pytest.approx(11.24909139072848)
    assert not result.diagnostics


def test_latest_hybrid_profile_derives_nominal_capacity_without_adding_range_surface() -> None:
    dataset = load_data_file(_latest_hybrid_csv(), ImportOptions(strict=True))
    result = calculate_profile_statistics(dataset, load_reporting_profile(HYBRID_PROFILE))
    stats = {item.definition.statistic_id: item.value for item in result.statistics}
    kpis = {item.definition.kpi_id: item.value for item in result.kpis}

    assert stats["battery_energy_first"] == pytest.approx(57.4)
    assert stats["battery_soc_first"] == pytest.approx(57.4)
    assert kpis["battery_capacity_100"] == pytest.approx(100.0)
    assert stats["wheel_power_total_max"] == pytest.approx(106.80434160103889)
    assert stats["generator_power_1_max"] == pytest.approx(80.00422042560322)
    assert not result.diagnostics


def test_generic_statistics_path_remains_available(tmp_path: Path) -> None:
    data_path = tmp_path / "data.csv"
    config_path = tmp_path / "statistics.yaml"
    data_path.write_text("Time,Power\ns,kW\n0,1\n1,3\n2,2\n", encoding="utf-8")
    config_path.write_text(
        "version: 1\n"
        "statistics:\n"
        "  - statistic_id: power_max\n"
        "    channel_id: power__col_002\n"
        "    operation: max\n"
        "    placement_group: bottom_channel\n"
        "output:\n"
        "  results_filename: results.csv\n"
        "  wide_filename: wide.csv\n",
        encoding="utf-8",
    )

    result = calculate_statistics(data_path, config_path)

    assert result.statistics[0].value == pytest.approx(3.0)


def test_generic_math_path_still_available_for_statistics(tmp_path: Path) -> None:
    data_path = tmp_path / "data.csv"
    math_path = tmp_path / "math.yaml"
    stats_path = tmp_path / "statistics.yaml"
    data_path.write_text("Track_Time,Power\ns,kW\n0,1\n1,3\n2,2\n", encoding="utf-8")
    math_path.write_text(
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
    stats_path.write_text(
        "version: 1\n"
        "statistics:\n"
        "  - statistic_id: calc_power_max\n"
        "    channel_id: calc_power\n"
        "    operation: max\n"
        "    placement_group: bottom_channel\n"
        "output:\n"
        "  results_filename: results.csv\n"
        "  wide_filename: wide.csv\n",
        encoding="utf-8",
    )

    math_result = calculate_math_channels(data_path, math_path)
    stats_result = calculate_statistics(data_path, stats_path, math_config_file=math_path)

    assert math_result.calculated_values[:, 0].tolist() == [2.0, 6.0, 4.0]
    assert stats_result.statistics[0].value == pytest.approx(6.0)


def _profile(
    raw_channels: list[RawChannelDefinition],
    statistics: list[StatisticDefinition],
    math_channels: list[MathChannelDefinition] | None = None,
    kpis: list[KPIDefinition] | None = None,
) -> ReportingProfile:
    return ReportingProfile(
        version=1,
        metadata=ProfileMetadata(profile_id="test_profile", name="Test Profile"),
        raw_channels=tuple(raw_channels),
        math_channels=tuple(math_channels or []),
        statistics=tuple(statistics),
        kpis=tuple(kpis or []),
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


def _battery_profile_dataset(
    energy_kwh: list[float],
    soc_percent: list[float],
    *,
    distance_km: float = 1.0,
) -> ImportedDataset:
    rows = []
    for index, (energy, soc) in enumerate(zip(energy_kwh, soc_percent), start=1):
        rows.append([energy, soc, distance_km * index / len(energy_kwh)])
    return _dataset(
        [
            _channel("battery_energy__col_001", "Battery Energy", "kWh"),
            _channel("battery_soc__col_002", "Battery SOC", "%", column=2),
            _channel("distance__col_003", "Distance", "km", column=3),
        ],
        rows,
    )


def _battery_capacity_profile() -> ReportingProfile:
    return _profile(
        raw_channels=[
            RawChannelDefinition("electricsystem_battery_energy", "Battery Energy", "Battery Energy", "VSM", unit="kWh"),
            RawChannelDefinition("electricsystem_battery_soc", "Battery SOC", "Battery SOC", "VSM", unit="%"),
        ],
        statistics=[],
        kpis=[
            KPIDefinition(
                kpi_id="battery_capacity_100",
                expression="nominal_battery_capacity_kwh(electricsystem_battery_energy, electricsystem_battery_soc)",
                dependencies=("electricsystem_battery_energy", "electricsystem_battery_soc"),
                display_name="100% Battery Capacity",
                unit="kWh",
            )
        ],
    )


def _battery_range_profile() -> ReportingProfile:
    return _profile(
        raw_channels=[
            RawChannelDefinition("electricsystem_battery_energy", "Battery Energy", "Battery Energy", "VSM", unit="kWh"),
            RawChannelDefinition("electricsystem_battery_soc", "Battery SOC", "Battery SOC", "VSM", unit="%"),
            RawChannelDefinition("distance_km", "Distance", "Distance", "VSM", unit="km"),
        ],
        statistics=[
            StatisticDefinition("battery_energy_first", "electricsystem_battery_energy", "first"),
            StatisticDefinition("battery_energy_last", "electricsystem_battery_energy", "last"),
            StatisticDefinition("distance_km_last", "distance_km", "last"),
        ],
        kpis=[
            KPIDefinition(
                "battery_capacity_used",
                "battery_energy_first - battery_energy_last",
                dependencies=("battery_energy_first", "battery_energy_last"),
                unit="kWh",
            ),
            KPIDefinition(
                "battery_energy_consumption_wh_per_km",
                "1000 * battery_capacity_used / distance_km_last",
                dependencies=("battery_capacity_used", "distance_km_last"),
                unit="Wh/Km",
            ),
            KPIDefinition(
                "battery_capacity_100",
                "nominal_battery_capacity_kwh(electricsystem_battery_energy, electricsystem_battery_soc)",
                dependencies=("electricsystem_battery_energy", "electricsystem_battery_soc"),
                unit="kWh",
            ),
            KPIDefinition(
                "usable_battery_capacity_85",
                "battery_capacity_100 * 0.85",
                dependencies=("battery_capacity_100",),
                unit="kWh",
            ),
            KPIDefinition(
                "range_85_battery_km",
                "1000 * usable_battery_capacity_85 / battery_energy_consumption_wh_per_km",
                dependencies=("usable_battery_capacity_85", "battery_energy_consumption_wh_per_km"),
                unit="Km",
            ),
        ],
    )
