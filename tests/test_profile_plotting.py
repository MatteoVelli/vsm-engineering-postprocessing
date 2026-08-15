from __future__ import annotations

from pathlib import Path

import numpy as np
import pytest

from vsm_postprocessing.importer import ImportOptions, load_data_file
from vsm_postprocessing.models import ChannelInfo, DataQualityReport, ImportedDataset
from vsm_postprocessing.plotting_engine import PlotDefaults, PlotStyle, render_plots
from vsm_postprocessing.profile_plotting import render_profile_plots
from vsm_postprocessing.report_profile import (
    MathChannelDefinition,
    ProfileMetadata,
    ProfilePlotDefinition,
    ProfilePlotSeriesDefinition,
    RawChannelDefinition,
    ReportingProfile,
    load_reporting_profile,
)


REFERENCE_CSV = Path(
    "reference_files/RoboSprayer_3500Kg_Electric_12kph_Batt_50kW_Motor_63RPM_Susp_Cool_Rough_Crop_Field_05.csv"
)
ELECTRIC_PROFILE = Path("config/report_profiles/robosprayer_electric.yaml")
HYBRID_PROFILE = Path("config/report_profiles/robosprayer_hybrid.yaml")
CAIMAN_SOURCE_WORKBOOK = Path(
    "reference_files/Sprayer_Caiman_SP_9300Kg_Hybrid_Gen80kW_30kph_74Ht_4000KgAQ_57-4pcSOC_5-80_1C2G_02.xlsx"
)
CAIMAN_PLOTTING_CONFIG = Path("config/plotting_example.yaml")
CAIMAN_MATH_CONFIG = Path("config/math_channels_example.yaml")


def test_profile_plotting_resolves_semantic_raw_series(tmp_path: Path) -> None:
    profile = _profile(
        raw_channels=[
            RawChannelDefinition("time", "Time", "Time", "VSM", unit="s"),
            RawChannelDefinition("speed", "Speed", "Speed", "VSM", unit="kph"),
        ],
        plots=[
            ProfilePlotDefinition(
                "speed_plot",
                "Speed",
                "time",
                (ProfilePlotSeriesDefinition("speed"),),
            )
        ],
    )

    result = render_profile_plots(_dataset([_channel("t__col_010", "Time", "s"), _channel("s__col_020", "Speed", "kph", 20)]), profile, tmp_path)

    assert result.rendered_plot_count == 1
    assert result.rendered_plots[0].x_channel_id == "time"
    assert result.rendered_plots[0].primary_series_ids == ("speed",)
    assert Path(result.rendered_plots[0].png_file).exists()


def test_profile_plotting_resolves_semantic_math_series_and_secondary_axis(tmp_path: Path) -> None:
    profile = _profile(
        raw_channels=[
            RawChannelDefinition("time", "Time", "Time", "VSM", unit="s"),
            RawChannelDefinition("power", "Power", "Power", "VSM", unit="kW"),
        ],
        math_channels=[
            MathChannelDefinition("double_power", "Double Power", "Double Power", "kW", ("power",), "power * 2")
        ],
        plots=[
            ProfilePlotDefinition(
                "power_plot",
                "Power",
                "time",
                (
                    ProfilePlotSeriesDefinition("power", axis="primary"),
                    ProfilePlotSeriesDefinition("double_power", axis="secondary"),
                ),
                primary_y_label="Power [kW]",
                secondary_y_label="Double Power [kW]",
            )
        ],
    )

    result = render_profile_plots(
        _dataset(
            [_channel("t__col_001", "Time", "s"), _channel("p__col_002", "Power", "kW", 2)],
            [[0.0, 1.0], [1.0, 3.0], [2.0, 2.0]],
        ),
        profile,
        tmp_path,
    )

    item = result.rendered_plots[0]
    assert item.primary_series_ids == ("power",)
    assert item.secondary_series_ids == ("double_power",)
    assert item.axes_count == 2


def test_profile_plotting_handles_multiple_series_optional_missing_required_missing_and_zero_series(tmp_path: Path) -> None:
    profile = _profile(
        raw_channels=[
            RawChannelDefinition("time", "Time", "Time", "VSM", unit="s"),
            RawChannelDefinition("speed", "Speed", "Speed", "VSM", unit="kph"),
            RawChannelDefinition("zero_power", "Zero Power", "Zero Power", "VSM", unit="kW"),
            RawChannelDefinition("optional_missing", "Missing Optional", "Missing Optional", "VSM", unit="kW", required=False),
            RawChannelDefinition("required_missing", "Missing Required", "Missing Required", "VSM", unit="kW"),
        ],
        plots=[
            ProfilePlotDefinition(
                "optional_omitted",
                "Optional Omitted",
                "time",
                (
                    ProfilePlotSeriesDefinition("speed"),
                    ProfilePlotSeriesDefinition("optional_missing", required=False),
                    ProfilePlotSeriesDefinition("zero_power"),
                ),
            ),
            ProfilePlotDefinition(
                "required_unavailable",
                "Required Unavailable",
                "time",
                (ProfilePlotSeriesDefinition("required_missing"),),
            ),
        ],
    )

    result = render_profile_plots(
        _dataset(
            [
                _channel("time__col_100", "Time", "s", 100),
                _channel("speed__col_200", "Speed", "kph", 200),
                _channel("zero__col_300", "Zero Power", "kW", 300),
            ],
            [[0.0, 1.0, 0.0], [1.0, 2.0, 0.0], [2.0, 3.0, 0.0]],
        ),
        profile,
        tmp_path,
    )

    assert result.rendered_plots[0].primary_series_ids == ("speed", "zero_power")
    assert result.unavailable_plots[0].definition.plot_id == "required_unavailable"
    assert result.unavailable_plots[0].missing_semantic_names == ("required_missing",)
    zero_summary = next(item for item in result.series_summaries["optional_omitted"] if item.semantic_name == "zero_power")
    assert zero_summary.is_constant
    assert zero_summary.is_all_zero


def test_electric_profile_plot_loading_count_rendering_and_representative_mappings(tmp_path: Path) -> None:
    dataset = load_data_file(REFERENCE_CSV, ImportOptions())
    profile = load_reporting_profile(ELECTRIC_PROFILE)

    result = render_profile_plots(dataset, profile, tmp_path, defaults=_fast_defaults())
    by_id = {plot.plot_id: plot for plot in profile.plots}

    assert len(profile.plots) == 12
    assert result.configured_plot_count == 12
    assert result.rendered_plot_count == 10
    assert result.unavailable_plot_count == 2
    assert {item.definition.plot_id for item in result.unavailable_plots} == {
        "agrochemical_discharge_vs_distance",
        "agrochemical_discharge_and_charge_vs_time",
    }
    assert by_id["battery_soc"].x == "distance_km"
    assert [item.semantic_name for item in by_id["battery_soc"].series] == [
        "chassis_speed",
        "electricsystem_battery_soc",
    ]
    assert by_id["battery_energy_time_based"].status == "RECONSTRUCTED"
    assert all(Path(item.png_file).exists() for item in result.rendered_plots)
    assert all(result.values_by_semantic_name[item.x].size == dataset.quality.sample_count for item in profile.plots)


def test_hybrid_profile_plot_inheritance_generator_mapping_and_zero_inactive_rendering(tmp_path: Path) -> None:
    dataset = load_data_file(REFERENCE_CSV, ImportOptions())
    profile = load_reporting_profile(HYBRID_PROFILE)

    result = render_profile_plots(dataset, profile, tmp_path, defaults=_fast_defaults())
    by_id = {plot.plot_id: plot for plot in profile.plots}

    assert len(profile.plots) == 18
    assert result.rendered_plot_count == 16
    assert result.unavailable_plot_count == 2
    assert by_id["generator_power"].series[0].semantic_name == "generator_power_1"
    np.testing.assert_allclose(result.values_by_semantic_name["generator_power_1"], 0.0)
    generator_summary = next(item for item in result.series_summaries["generator_power"] if item.semantic_name == "generator_power_1")
    assert generator_summary.is_all_zero
    assert generator_summary.source_name == "Generator Power_1"


@pytest.mark.skipif(not CAIMAN_SOURCE_WORKBOOK.exists(), reason="Client source workbook is not present")
def test_legacy_caiman_plotting_path_remains_unchanged(tmp_path: Path) -> None:
    result = render_plots(
        CAIMAN_SOURCE_WORKBOOK,
        CAIMAN_PLOTTING_CONFIG,
        tmp_path / "caiman",
        ImportOptions(strict=True),
        math_config_file=CAIMAN_MATH_CONFIG,
    )

    assert result.plot_count == 24
    assert result.series_count == 45


def _fast_defaults() -> PlotDefaults:
    return PlotDefaults(
        width_inches=4,
        height_inches=2.4,
        dpi=70,
        line_width=0.8,
        style=PlotStyle(output_formats=("png", "svg"), legend_fontsize=6, tick_fontsize=6),
    )


def _profile(
    raw_channels: list[RawChannelDefinition],
    plots: list[ProfilePlotDefinition],
    math_channels: list[MathChannelDefinition] | None = None,
) -> ReportingProfile:
    return ReportingProfile(
        version=1,
        metadata=ProfileMetadata(profile_id="test_profile", name="Test Profile"),
        raw_channels=tuple(raw_channels),
        math_channels=tuple(math_channels or []),
        plots=tuple(plots),
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


def _dataset(channels: list[ChannelInfo], values: list[list[float]] | None = None) -> ImportedDataset:
    if values is None:
        values = [[0.0 for _ in channels], [1.0 for _ in channels], [2.0 for _ in channels]]
    array = np.asarray(values, dtype=np.float64)
    return ImportedDataset(
        source_path=Path("test.csv"),
        channels=channels,
        quality=DataQualityReport(
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
        ),
        values=array,
    )
