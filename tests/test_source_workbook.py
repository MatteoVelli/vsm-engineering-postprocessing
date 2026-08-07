from pathlib import Path

import pytest
from openpyxl import Workbook

from vsm_postprocessing.errors import DataValidationError
from vsm_postprocessing.importer import ImportOptions, inspect_data_file


SOURCE_WORKBOOK = Path(__file__).resolve().parents[1] / "reference_files" / (
    "Sprayer_Caiman_SP_9300Kg_Hybrid_Gen80kW_30kph_74Ht_4000KgAQ_57-4pcSOC_5-80_1C2G_02.xlsx"
)


@pytest.mark.skipif(not SOURCE_WORKBOOK.exists(), reason="Client reference workbook is not present")
def test_supplied_source_workbook_acceptance_criteria() -> None:
    result = inspect_data_file(SOURCE_WORKBOOK)
    quality = result.quality

    assert quality.is_valid
    assert quality.sheet_name == "Sprayer_Caiman_SP_9300Kg_Hybrid"
    assert quality.header_row == 3
    assert quality.unit_row == 4
    assert quality.data_start_row == 5
    assert quality.data_end_row == 1870
    assert quality.sample_count == 1866
    assert quality.channel_count == 70
    assert quality.raw_channel_count == 45
    assert quality.math_channel_count == 25
    assert quality.time_channel_name == "Track_Time"
    assert quality.time_unit == "s"
    assert quality.time_start == 0.0
    assert quality.time_end == 1865.0
    assert quality.nominal_time_step == 1.0
    assert quality.time_is_strictly_increasing
    assert quality.duplicate_timestamp_count == 0
    assert quality.missing_cell_count == 0
    assert quality.invalid_numeric_cell_count == 0
    assert quality.non_finite_cell_count == 0

    channel_ids = [channel.channel_id for channel in result.channels]
    assert len(channel_ids) == len(set(channel_ids))


def test_nonmonotonic_time_fails_strict_validation(tmp_path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    sheet.append(["Track_Time", "Speed"])
    sheet.append(["s", "kph"])
    sheet.append([0, 0])
    sheet.append([1, 10])
    sheet.append([1, 20])
    path = tmp_path / "nonmonotonic.xlsx"
    workbook.save(path)

    with pytest.raises(DataValidationError, match="not strictly increasing"):
        inspect_data_file(path, ImportOptions(header_row=1, unit_row=2, data_start_row=3))


def test_duplicate_display_names_still_receive_unique_ids(tmp_path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    sheet.append(["Track_Time", "Power", "Power"])
    sheet.append(["s", "kW", "kW"])
    sheet.append([0, 1, 2])
    sheet.append([1, 3, 4])
    path = tmp_path / "duplicates.xlsx"
    workbook.save(path)

    result = inspect_data_file(path, ImportOptions(header_row=1, unit_row=2, data_start_row=3))
    ids = [channel.channel_id for channel in result.channels]

    assert len(ids) == len(set(ids))
    assert result.channels[1].source_name == result.channels[2].source_name == "Power"


def test_csv_import_with_units(tmp_path: Path) -> None:
    path = tmp_path / "data.csv"
    path.write_text("Track_Time;Speed\ns;kph\n0;0\n1;12.5\n2;20\n", encoding="utf-8")

    result = inspect_data_file(path)

    assert result.quality.is_valid
    assert result.quality.sample_count == 3
    assert result.quality.channel_count == 2
    assert result.quality.raw_channel_count == 2
    assert result.quality.math_channel_count == 0
    assert result.quality.nominal_time_step == 1.0
