from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import load_workbook

from vsm_postprocessing.excel_report_engine import generate_profile_excel_report
from vsm_postprocessing.importer import ImportOptions


REFERENCE_CSV = Path(
    "reference_files/RoboSprayer_3500Kg_Electric_12kph_Batt_50kW_Motor_63RPM_Susp_Cool_Rough_Crop_Field_05.csv"
)
ELECTRIC_PROFILE = Path("config/report_profiles/robosprayer_electric.yaml")
HYBRID_PROFILE = Path("config/report_profiles/robosprayer_hybrid.yaml")
DATA_START_ROW = 8


@pytest.fixture(scope="module")
def electric_report(tmp_path_factory: pytest.TempPathFactory):
    return generate_profile_excel_report(
        REFERENCE_CSV,
        ELECTRIC_PROFILE,
        tmp_path_factory.mktemp("profile_excel_electric"),
        ImportOptions(),
    )


@pytest.fixture(scope="module")
def electric_workbook(electric_report):
    return load_workbook(electric_report.report_path, data_only=True)


def test_profile_excel_report_generates_reopenable_electric_workbook(electric_report, electric_workbook) -> None:
    assert electric_report.report_path.exists()
    assert electric_report.sample_count == 3853
    assert electric_report.source_raw_channel_count == 607
    assert electric_report.report_channel_count == 317
    assert electric_report.vsm_count == 178
    assert electric_report.avl_count == 110
    assert electric_report.math_count == 29
    assert electric_report.statistic_count == 27
    assert electric_report.kpi_count == 9
    assert electric_report.plot_count == 12
    assert electric_workbook.sheetnames == [
        "RoboSprayer Electric",
        "Statistics KPIs",
        "Plots",
        "Metadata",
        "Template Comparison",
    ]


def test_profile_excel_report_exports_semantic_raw_and_math_channels_with_dynamic_rows(
    electric_report,
    electric_workbook,
) -> None:
    sheet = electric_workbook["RoboSprayer Electric"]
    channel_types = [sheet.cell(4, col).value for col in range(1, electric_report.report_channel_count + 1)]

    assert channel_types.count("VSM") == 178
    assert channel_types.count("AVL") == 110
    assert channel_types.count("MATH") == 29
    assert not any("__col_" in str(sheet.cell(6, col).value) for col in range(1, electric_report.report_channel_count + 1))
    assert sheet.max_column == 317
    assert sheet.cell(DATA_START_ROW + 3853, 1).value is None

    by_semantic = {channel.channel_id: index + 1 for index, channel in enumerate(electric_report.report_channels)}
    data_end_row = DATA_START_ROW + electric_report.sample_count - 1
    assert sheet.cell(DATA_START_ROW, by_semantic["track_time"]).value == pytest.approx(0.0)
    assert sheet.cell(data_end_row, by_semantic["time_minutes"]).value == pytest.approx(64.2)
    assert sheet.cell(data_end_row, by_semantic["distance_km"]).value == pytest.approx(11.9996)


def test_profile_excel_report_statistics_kpis_and_correct_rms_values(electric_report, electric_workbook) -> None:
    sheet = electric_workbook["Statistics KPIs"]
    statistics = {sheet.cell(row, 1).value: sheet.cell(row, 5).value for row in range(2, 29)}
    kpi_start = 31
    kpis = {sheet.cell(row, 1).value: sheet.cell(row, 3).value for row in range(kpi_start + 1, kpi_start + 10)}

    assert statistics["battery_power_rms"] == pytest.approx(28.716636645770492)
    assert statistics["battery_power_rms"] != pytest.approx(13.506611297860566)
    assert statistics["battery_heatflow_rms"] == pytest.approx(2.840107066457429)
    assert statistics["battery_heatflow_rms"] != pytest.approx(1.3358187681981721)
    assert statistics["agrochemical_discharge_max"] == pytest.approx(0.0)
    assert kpis["battery_capacity_used"] == pytest.approx(33.65367)
    assert kpis["battery_energy_consumption_wh_per_km"] == pytest.approx(2804.565985532851)
    assert kpis["range_85_battery_km"] == pytest.approx(12.76114517141972)


def test_profile_excel_report_includes_plots_metadata_and_template_comparison(
    electric_report,
    electric_workbook,
) -> None:
    plots_sheet = electric_workbook["Plots"]
    metadata_sheet = electric_workbook["Metadata"]
    comparison_sheet = electric_workbook["Template Comparison"]

    assert len(plots_sheet._images) == 12
    assert [plots_sheet.cell(row, 1).value for row in range(2, 14)] == [
        plot.plot_id for plot in electric_report.plotting_result.rendered_plots
    ]

    metadata = {metadata_sheet.cell(row, 1).value: metadata_sheet.cell(row, 2).value for row in range(2, 36)}
    assert metadata["Source sample count"] == 3853
    assert metadata["Workbook data start row"] == DATA_START_ROW
    assert metadata["Workbook data end row"] == DATA_START_ROW + 3853 - 1
    assert metadata["Exported report channels"] == 317
    assert metadata["Rendered plots"] == 12

    statuses = [comparison_sheet.cell(row, 4).value for row in range(2, comparison_sheet.max_row + 1)]
    assert "UNAVAILABLE" not in statuses
    assert "INTENTIONAL CORRECTION" in statuses


def test_profile_excel_report_hybrid_dry_run_remains_profile_generic(tmp_path: Path) -> None:
    result = generate_profile_excel_report(
        REFERENCE_CSV,
        HYBRID_PROFILE,
        tmp_path / "profile_excel_hybrid",
        ImportOptions(),
    )
    workbook = load_workbook(result.report_path, data_only=True)

    assert workbook.sheetnames[0] == "RoboSprayer Hybrid"
    assert result.sample_count == 3853
    assert result.source_raw_channel_count == 607
    assert result.report_channel_count == 326
    assert result.vsm_count == 183
    assert result.avl_count == 111
    assert result.math_count == 32
    assert result.statistic_count == 36
    assert result.kpi_count == 9
    assert result.plot_count == 18
    assert len(workbook["Plots"]._images) == 18
