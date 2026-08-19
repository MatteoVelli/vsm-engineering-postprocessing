from __future__ import annotations

from pathlib import Path

import pytest

from vsm_postprocessing.errors import ConfigurationError
from vsm_postprocessing.excel_report_engine import load_excel_report_config


def _write_config(path: Path, *, channels: str = "  - time__col_001\n", output: str = "report.xlsx", bottom: str = "    - max\n") -> None:
    path.write_text(
        f"""version: 1
report:
  title: Test Report
  report_sheet: Report
  metadata_sheet: Metadata
channels:
{channels}statistics:
  top_rms: []
  kpis: []
  bottom_operations:
{bottom}plots:
  include: []
  columns: 2
  width_px: 600
  height_px: 300
output:
  filename: {output}
  keep_plot_assets: true
""",
        encoding="utf-8",
    )


def test_excel_report_config_loads(tmp_path: Path) -> None:
    path = tmp_path / "report.yaml"
    _write_config(path)
    config = load_excel_report_config(path)
    assert config.version == 1
    assert config.report_sheet == "Report"
    assert config.metadata_sheet == "Metadata"
    assert config.channel_ids == ("time__col_001",)
    assert config.bottom_operations == ("max",)
    assert config.output_filename == "report.xlsx"


def test_duplicate_report_channels_are_rejected(tmp_path: Path) -> None:
    path = tmp_path / "report.yaml"
    _write_config(path, channels="  - time__col_001\n  - time__col_001\n")
    with pytest.raises(ConfigurationError, match="duplicate channel IDs"):
        load_excel_report_config(path)


def test_invalid_bottom_operation_is_rejected(tmp_path: Path) -> None:
    path = tmp_path / "report.yaml"
    _write_config(path, bottom="    - median\n")
    with pytest.raises(ConfigurationError, match="unsupported operations"):
        load_excel_report_config(path)


def test_output_must_be_plain_xlsx_filename(tmp_path: Path) -> None:
    path = tmp_path / "report.yaml"
    _write_config(path, output="nested/report.xlsx")
    with pytest.raises(ConfigurationError, match="plain .xlsx filename"):
        load_excel_report_config(path)


def test_sergio_reference_layout_config_loads(tmp_path: Path) -> None:
    path = tmp_path / "report.yaml"
    path.write_text(
        """version: 1
report:
  title: Test Report
  report_sheet: Report
  metadata_sheet: Metadata
channels:
  - time__col_001
statistics:
  top_rms: []
  kpis: []
  bottom_operations:
    - max
  bottom_summary: []
layout:
  profile: sergio_reference
  plot_placement: kpi_panel
  blank_separator_columns: 1
  channel_width: 12
  header_row_height: 106
  unit_row_height: 16
plots:
  include: []
  columns: 2
  width_px: 600
  height_px: 300
output:
  filename: report.xlsx
  keep_plot_assets: true
""",
        encoding="utf-8",
    )
    config = load_excel_report_config(path)
    assert config.layout_profile == "sergio_reference"
    assert config.plot_placement == "kpi_panel"
    assert config.header_row_height == 106
    assert config.unit_row_height == 16


def test_invalid_layout_profile_is_rejected(tmp_path: Path) -> None:
    path = tmp_path / "report.yaml"
    path.write_text(
        """version: 1
report:
  title: Test Report
channels:
  - time__col_001
statistics:
  top_rms: []
  kpis: []
  bottom_operations:
    - max
plots:
  include: []
layout:
  profile: copied_by_hand
output:
  filename: report.xlsx
""",
        encoding="utf-8",
    )
    with pytest.raises(ConfigurationError, match="layout.profile"):
        load_excel_report_config(path)
