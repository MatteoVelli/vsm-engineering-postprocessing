from __future__ import annotations

import csv
import json
from pathlib import Path

import pytest
from openpyxl import Workbook

from vsm_postprocessing.channel_manager import (
    export_channel_selection,
    load_selection_config,
    select_channels,
)
from vsm_postprocessing.errors import ChannelSelectionError, ConfigurationError
from vsm_postprocessing.importer import ImportOptions


def _write_small_workbook(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    sheet.append(["Track_Time", "Speed", "Power", "Power"])
    sheet.append(["s", "kph", "kW", "W"])
    sheet.append([0, 0, 10, 10000])
    sheet.append([1, 12.5, 20, 20000])
    sheet.append([2, 20, 30, 30000])
    workbook.save(path)


def _write_config(path: Path, export_channels: list[str], include_time: bool = True) -> None:
    lines = [
        "version: 1",
        "selection:",
        f"  include_time: {'true' if include_time else 'false'}",
        "  export_channels:",
        *[f"    - {channel_id}" for channel_id in export_channels],
        "output:",
        "  data_filename: selected.csv",
        "  include_units_row: true",
        "  float_precision: 12",
    ]
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def test_selects_by_channel_id_and_exports_in_configured_order(tmp_path: Path) -> None:
    workbook_path = tmp_path / "data.xlsx"
    config_path = tmp_path / "selection.yaml"
    output_dir = tmp_path / "outputs"
    _write_small_workbook(workbook_path)
    _write_config(
        config_path,
        ["power__col_004", "speed__col_002"],
        include_time=True,
    )

    result = select_channels(
        workbook_path,
        config_path,
        ImportOptions(header_row=1, unit_row=2, data_start_row=3),
    )
    outputs = export_channel_selection(result, output_dir)

    assert [channel.channel_id for channel in result.selected_channels] == [
        "track_time__col_001",
        "power__col_004",
        "speed__col_002",
    ]
    assert result.selected_values.tolist() == [
        [0.0, 10000.0, 0.0],
        [1.0, 20000.0, 12.5],
        [2.0, 30000.0, 20.0],
    ]

    with outputs["selected_data"].open(encoding="utf-8", newline="") as handle:
        rows = list(csv.reader(handle))
    assert rows[0] == ["track_time__col_001", "power__col_004", "speed__col_002"]
    assert rows[1] == ["s", "W", "kph"]
    assert rows[2] == ["0", "10000", "0"]
    assert rows[-1] == ["2", "30000", "20"]

    manifest = json.loads(outputs["selection_manifest"].read_text(encoding="utf-8"))
    assert manifest["sample_count"] == 3
    assert manifest["selected_channel_count"] == 3


def test_include_time_does_not_duplicate_explicit_time_channel(tmp_path: Path) -> None:
    workbook_path = tmp_path / "data.xlsx"
    config_path = tmp_path / "selection.yaml"
    _write_small_workbook(workbook_path)
    _write_config(config_path, ["speed__col_002", "track_time__col_001"], include_time=True)

    result = select_channels(
        workbook_path,
        config_path,
        ImportOptions(header_row=1, unit_row=2, data_start_row=3),
    )

    assert [channel.channel_id for channel in result.selected_channels] == [
        "track_time__col_001",
        "speed__col_002",
    ]


def test_missing_channel_name_explains_that_ids_are_required(tmp_path: Path) -> None:
    workbook_path = tmp_path / "data.xlsx"
    config_path = tmp_path / "selection.yaml"
    _write_small_workbook(workbook_path)
    _write_config(config_path, ["Power"], include_time=False)

    with pytest.raises(ChannelSelectionError, match="display names are not unique keys") as exc_info:
        select_channels(
            workbook_path,
            config_path,
            ImportOptions(header_row=1, unit_row=2, data_start_row=3),
        )

    message = str(exc_info.value)
    assert "power__col_003" in message
    assert "power__col_004" in message


def test_duplicate_channel_ids_are_rejected_in_configuration(tmp_path: Path) -> None:
    config_path = tmp_path / "selection.yaml"
    _write_config(config_path, ["speed__col_002", "speed__col_002"])

    with pytest.raises(ConfigurationError, match="duplicate channel IDs"):
        load_selection_config(config_path)


def test_output_filename_cannot_escape_output_directory(tmp_path: Path) -> None:
    config_path = tmp_path / "selection.yaml"
    config_path.write_text(
        """version: 1
selection:
  include_time: true
  export_channels:
    - speed__col_002
output:
  data_filename: ../outside.csv
""",
        encoding="utf-8",
    )

    with pytest.raises(ConfigurationError, match="without directories"):
        load_selection_config(config_path)
