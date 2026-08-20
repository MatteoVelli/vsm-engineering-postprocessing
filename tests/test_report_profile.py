from __future__ import annotations

from pathlib import Path

import numpy as np

from vsm_postprocessing.importer import ImportOptions, load_data_file
from vsm_postprocessing.models import ChannelInfo, DataQualityReport, ImportedDataset
from vsm_postprocessing.report_profile import (
    MathChannelDefinition,
    ProfileMetadata,
    RawChannelDefinition,
    ReportingProfile,
    load_reporting_profile,
    resolve_profile,
)

from conftest import (
    ROBOSPRAYER_LATEST_ELECTRIC_CSV,
    ROBOSPRAYER_LATEST_ELECTRIC_DESCRIPTION,
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


def test_exact_semantic_name_resolution() -> None:
    dataset = _dataset([_channel("chassis_speed__col_001", "Chassis_Speed", "kph")])
    profile = _profile(
        raw_channels=[
            RawChannelDefinition(
                semantic_name="chassis_speed",
                source_name="Chassis_Speed",
                report_name="Speed",
                channel_type="VSM",
                unit="kph",
            )
        ]
    )

    result = resolve_profile(dataset, profile)

    assert result.is_valid
    assert result.resolved_channel_ids == {"chassis_speed": "chassis_speed__col_001"}
    assert result.resolved["chassis_speed"].match_type == "exact"


def test_normalized_name_resolution() -> None:
    dataset = _dataset([_channel("engine_speed__col_007", "Engine_Speed", "rpm")])
    profile = _profile(
        raw_channels=[
            RawChannelDefinition(
                semantic_name="engine_speed",
                source_name="Engine Speed",
                report_name="Engine Speed",
                channel_type="VSM",
                unit="rpm",
            )
        ]
    )

    result = resolve_profile(dataset, profile)

    assert result.is_valid
    assert result.resolved_channel_ids["engine_speed"] == "engine_speed__col_007"
    assert result.resolved["engine_speed"].match_type == "normalized"


def test_alias_resolution() -> None:
    dataset = _dataset([_channel("engine_fuelconsumption_volumeflow__col_140", "Engine_FuelConsumption_volumeflow", "l/h")])
    profile = _profile(
        raw_channels=[
            RawChannelDefinition(
                semantic_name="fuel_flow",
                source_name="Fuel Flow",
                report_name="Fuel Flow",
                channel_type="VSM",
                unit="l/h",
                aliases=("Engine_FuelConsumption_volumeflow",),
            )
        ]
    )

    result = resolve_profile(dataset, profile)

    assert result.is_valid
    assert result.resolved_channel_ids["fuel_flow"] == "engine_fuelconsumption_volumeflow__col_140"
    assert result.resolved["fuel_flow"].match_type == "alias"


def test_missing_required_channel_is_reported() -> None:
    result = resolve_profile(
        _dataset([]),
        _profile(
            raw_channels=[
                RawChannelDefinition(
                    semantic_name="battery_power",
                    source_name="ElectricSystem_Battery_Power",
                    report_name="Battery Power",
                    channel_type="VSM",
                    unit="kW",
                )
            ]
        ),
    )

    assert not result.is_valid
    assert [item.definition.semantic_name for item in result.missing_required] == ["battery_power"]
    assert not result.missing_optional


def test_missing_optional_channel_is_reported_without_invalidating_profile() -> None:
    result = resolve_profile(
        _dataset([]),
        _profile(
            raw_channels=[
                RawChannelDefinition(
                    semantic_name="agrochemical_discharge_force",
                    source_name="Agrochemical Discharge",
                    report_name="Agrochemical Discharge",
                    channel_type="VSM",
                    unit="N",
                    required=False,
                )
            ]
        ),
    )

    assert result.is_valid
    assert [item.definition.semantic_name for item in result.missing_optional] == ["agrochemical_discharge_force"]


def test_unit_mismatch_is_reported() -> None:
    dataset = _dataset([_channel("battery_power__col_001", "ElectricSystem_Battery_Power", "W")])
    profile = _profile(
        raw_channels=[
            RawChannelDefinition(
                semantic_name="battery_power",
                source_name="ElectricSystem_Battery_Power",
                report_name="Battery Power",
                channel_type="VSM",
                unit="kW",
            )
        ]
    )

    result = resolve_profile(dataset, profile)

    assert not result.is_valid
    assert [(item.definition.semantic_name, item.expected_unit, item.actual_unit) for item in result.unit_mismatches] == [
        ("battery_power", "kW", "W")
    ]


def test_duplicate_source_name_is_ambiguous() -> None:
    dataset = _dataset(
        [
            _channel("speed_a__col_001", "Chassis_Speed", "kph"),
            _channel("speed_b__col_002", "Chassis_Speed", "kph", column=2),
        ]
    )
    profile = _profile(
        raw_channels=[
            RawChannelDefinition(
                semantic_name="chassis_speed",
                source_name="Chassis_Speed",
                report_name="Speed",
                channel_type="VSM",
                unit="kph",
            )
        ]
    )

    result = resolve_profile(dataset, profile)

    assert not result.is_valid
    assert len(result.ambiguous) == 1
    assert {channel.channel_id for channel in result.ambiguous[0].candidates} == {"speed_a__col_001", "speed_b__col_002"}


def test_profile_resolution_is_independent_from_runtime_column_positions() -> None:
    profile = _profile(
        raw_channels=[
            RawChannelDefinition(
                semantic_name="chassis_speed",
                source_name="Chassis_Speed",
                report_name="Speed",
                channel_type="VSM",
                unit="kph",
            )
        ]
    )

    first = resolve_profile(_dataset([_channel("chassis_speed__col_001", "Chassis_Speed", "kph")]), profile)
    reordered = resolve_profile(_dataset([_channel("chassis_speed__col_039", "Chassis_Speed", "kph", column=39)]), profile)

    assert first.resolved_channel_ids["chassis_speed"] == "chassis_speed__col_001"
    assert reordered.resolved_channel_ids["chassis_speed"] == "chassis_speed__col_039"


def test_electric_robosprayer_profile_resolution_against_reference_csv() -> None:
    dataset = load_data_file(_robosprayer_csv(), ImportOptions())
    profile = load_reporting_profile(ELECTRIC_PROFILE)

    result = resolve_profile(dataset, profile)

    assert result.is_valid
    assert len(result.resolved) == 288
    assert len(result.profile.raw_channels) == 289
    assert not result.missing_required
    assert [item.definition.semantic_name for item in result.missing_optional] == ["track_height"]
    assert result.resolved_channel_ids["chassis_speed"] == "chassis_speed__col_039"
    assert result.resolved_channel_ids["electricsystem_battery_soc"] == "electricsystem_battery_soc__col_091"
    assert result.resolved_channel_ids["agrochemical_discharge_force"] == "hitchrear_force_z_vehiclecoordinates__col_167"
    assert result.resolved["agrochemical_discharge_force"].definition.source_name == "HitchRear_Force_Z_VehicleCoordinates"
    assert result.resolved["agrochemical_discharge_force"].definition.unit == "N"
    assert result.resolved["agrochemical_discharge_force"].is_all_zero


def test_latest_electric_profile_resolves_optional_track_height() -> None:
    dataset = load_data_file(_latest_electric_csv(), ImportOptions())
    profile = load_reporting_profile(ELECTRIC_PROFILE)

    result = resolve_profile(dataset, profile)

    assert result.is_valid
    assert not result.missing_required
    assert "track_height" in result.resolved_channel_ids
    assert result.resolved["track_height"].definition.source_name == "Track_Height"
    assert result.resolved["track_height"].definition.report_name == "Road Height"
    assert result.resolved["track_height"].definition.required is False


def test_agrochemical_force_mapping_is_independent_from_runtime_column_position() -> None:
    definition = load_reporting_profile(ELECTRIC_PROFILE).raw_by_semantic_name()["agrochemical_discharge_force"]
    profile = _profile(raw_channels=[definition])

    first = resolve_profile(
        _dataset([_channel("hitchrear_force_z_vehiclecoordinates__col_167", "HitchRear_Force_Z_VehicleCoordinates", "N", column=167)]),
        profile,
    )
    reordered = resolve_profile(
        _dataset([_channel("hitchrear_force_z_vehiclecoordinates__col_601", "HitchRear_Force_Z_VehicleCoordinates", "N", column=601)]),
        profile,
    )

    assert first.resolved_channel_ids["agrochemical_discharge_force"] == "hitchrear_force_z_vehiclecoordinates__col_167"
    assert reordered.resolved_channel_ids["agrochemical_discharge_force"] == "hitchrear_force_z_vehiclecoordinates__col_601"


def test_hybrid_profile_extends_electric_profile() -> None:
    electric = load_reporting_profile(ELECTRIC_PROFILE)
    hybrid = load_reporting_profile(HYBRID_PROFILE)

    assert len(electric.raw_channels) == 289
    assert len(electric.math_channels) == 29
    assert len(electric.plots) == 14
    assert len(hybrid.raw_channels) == 295
    assert len(hybrid.math_channels) == 32
    assert len(hybrid.plots) == 20
    assert {channel.semantic_name for channel in hybrid.raw_channels} >= {
        "engine_fuel_consumption",
        "engine_fuelconsumption_specific",
        "fuel_flow",
        "engine_speed",
        "engine_torque",
        "generator_torque_1",
    }


def test_electric_profile_marks_torque_placeholders_as_raw_fallbacks_and_plots_wheel_loads() -> None:
    profile = load_reporting_profile(ELECTRIC_PROFILE)
    math_by_name = profile.math_by_semantic_name()
    for semantic_name in (
        "driveshaft_torque_fl",
        "driveshaft_torque_fr",
        "driveshaft_torque_rl",
        "driveshaft_torque_rr",
    ):
        definition = math_by_name[semantic_name]
        assert definition.fallback_when_raw_missing is True
        assert definition.expression == "0"

    wheel_loads = profile.plots_by_id()["wheel_loads"]
    assert wheel_loads.title == "Wheel Loads"
    assert [series.semantic_name for series in wheel_loads.series] == [
        "wheel_load_dynamic_fl",
        "wheel_load_dynamic_fr",
        "wheel_load_dynamic_rl",
        "wheel_load_dynamic_rr",
    ]
    assert all(profile.raw_by_semantic_name()[series.semantic_name].for_plot for series in wheel_loads.series)


def test_hybrid_generator_torque_uses_sergio_corrected_source_mapping() -> None:
    profile = load_reporting_profile(HYBRID_PROFILE)
    generator_torque = profile.raw_by_semantic_name()["generator_torque_1"]

    assert generator_torque.source_name == "Engine_AuxiliaryTorque_1"
    assert generator_torque.report_name == "ICE Generator Torque"
    assert generator_torque.channel_type == "AVL"
    assert generator_torque.for_plot
    assert generator_torque.source_name != "Generator Torque_1"


def test_hybrid_generator_torque_resolution_is_independent_from_runtime_column_position() -> None:
    definition = load_reporting_profile(HYBRID_PROFILE).raw_by_semantic_name()["generator_torque_1"]
    profile = _profile(raw_channels=[definition])

    first = resolve_profile(
        _dataset([_channel("engine_auxiliarytorque_1__col_133", "Engine_AuxiliaryTorque_1", "Nm", column=133)]),
        profile,
    )
    reordered = resolve_profile(
        _dataset([_channel("engine_auxiliarytorque_1__col_599", "Engine_AuxiliaryTorque_1", "Nm", column=599)]),
        profile,
    )

    assert first.resolved_channel_ids["generator_torque_1"] == "engine_auxiliarytorque_1__col_133"
    assert reordered.resolved_channel_ids["generator_torque_1"] == "engine_auxiliarytorque_1__col_599"


def test_hybrid_profile_against_electric_csv_resolves_inactive_hybrid_channels() -> None:
    dataset = load_data_file(_robosprayer_csv(), ImportOptions())
    result = resolve_profile(dataset, load_reporting_profile(HYBRID_PROFILE))

    assert result.is_valid
    assert len(result.resolved) == 294
    assert len(result.profile.raw_channels) == 295
    assert not result.missing_required
    assert [item.definition.semantic_name for item in result.missing_optional] == ["track_height"]
    assert result.resolved["engine_speed"].channel.channel_id == "engine_speed__col_149"
    assert result.resolved["engine_speed"].is_all_zero
    assert result.resolved["engine_torque"].is_all_zero
    assert result.resolved["engine_fuel_consumption"].channel.channel_id == "engine_fuelconsumption_absolut__col_138"
    assert result.resolved["fuel_flow"].channel.channel_id == "engine_fuelconsumption_volumeflow__col_140"
    assert result.resolved["generator_torque_1"].channel.channel_id == "engine_auxiliarytorque_1__col_133"
    assert result.resolved["generator_torque_1"].channel.source_name == "Engine_AuxiliaryTorque_1"
    assert result.resolved["generator_torque_1"].is_all_zero


def test_semantic_math_dependency_resolution() -> None:
    dataset = _dataset(
        [
            _channel("track_time__col_001", "Track_Time", "s"),
            _channel("distance__col_002", "Track_Distance", "m", column=2),
            _channel("em1_speed__col_003", "ElectricSystem_EM1_Speed", "rpm", column=3),
            _channel("em1_torque__col_004", "ElectricSystem_EM1_Torque", "Nm", column=4),
        ]
    )
    profile = _profile(
        raw_channels=[
            RawChannelDefinition("track_time", "Track_Time", "Time (s)", "VSM", unit="s"),
            RawChannelDefinition("track_distance", "Track_Distance", "Distance (m)", "VSM", unit="m"),
            RawChannelDefinition("em1_speed", "ElectricSystem_EM1_Speed", "EDU Speed FL", "VSM", unit="rpm"),
            RawChannelDefinition("em1_torque", "ElectricSystem_EM1_Torque", "EDU Torque FL", "VSM", unit="Nm"),
        ],
        math_channels=[
            MathChannelDefinition(
                semantic_name="time_minutes",
                source_name="Time",
                report_name="Time (min)",
                unit="min",
                dependencies=("track_time",),
                expression="track_time / 60",
            ),
            MathChannelDefinition(
                semantic_name="edu_mech_power_fl",
                source_name="EDU Mech Power FL",
                report_name="EDU Mech Power FL",
                unit="kW",
                dependencies=("em1_speed", "em1_torque"),
                expression="em1_speed * em1_torque / rpm_nm_to_kw_divisor",
            ),
            MathChannelDefinition(
                semantic_name="total_edu_mech_power",
                source_name="Total EDU Mech Power",
                report_name="Total EDU Mech Power",
                unit="kW",
                dependencies=("edu_mech_power_fl",),
                expression="edu_mech_power_fl",
            ),
        ],
    )

    result = resolve_profile(dataset, profile)
    by_name = {item.definition.semantic_name: item for item in result.math_dependencies}

    assert by_name["time_minutes"].resolved_dependencies == {"track_time": "track_time__col_001"}
    assert by_name["edu_mech_power_fl"].resolved_dependencies == {
        "em1_speed": "em1_speed__col_003",
        "em1_torque": "em1_torque__col_004",
    }
    assert by_name["total_edu_mech_power"].math_dependencies == ("edu_mech_power_fl",)
    assert by_name["total_edu_mech_power"].is_resolved


def _profile(
    raw_channels: list[RawChannelDefinition],
    math_channels: list[MathChannelDefinition] | None = None,
) -> ReportingProfile:
    return ReportingProfile(
        version=1,
        metadata=ProfileMetadata(profile_id="test_profile", name="Test Profile"),
        raw_channels=tuple(raw_channels),
        math_channels=tuple(math_channels or []),
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


def _dataset(channels: list[ChannelInfo]) -> ImportedDataset:
    columns = len(channels)
    values = np.zeros((3, columns), dtype=np.float64)
    for index in range(columns):
        values[:, index] = np.array([0.0, float(index + 1), float(index + 2)])
    quality = DataQualityReport(
        source_file="test.csv",
        source_sha256="sha",
        file_type="csv",
        sheet_name=None,
        header_row=1,
        unit_row=2,
        data_start_row=3,
        data_end_row=5,
        sample_count=3,
        channel_count=columns,
        raw_channel_count=columns,
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
        values=values,
    )
