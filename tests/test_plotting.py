from __future__ import annotations

import csv
import json
import struct
from pathlib import Path

import matplotlib.pyplot as plt
import pytest

from vsm_postprocessing.errors import ConfigurationError, PlottingError
from vsm_postprocessing.importer import ImportOptions
from vsm_postprocessing.plotting_engine import load_plotting_config, render_plots
from conftest import (
    CAIMAN_PROFILE_REFERENCE_DESCRIPTION,
    CAIMAN_PROFILE_REFERENCE_XLSX,
    CAIMAN_REFERENCE_DESCRIPTION,
    CAIMAN_REFERENCE_XLSX,
    require_private_reference_file,
)


def _write_csv(path: Path) -> None:
    path.write_text(
        "Time,Speed,Power\n"
        "s,kph,kW\n"
        "0,0,10\n"
        "1,20,15\n"
        "2,30,5\n",
        encoding="utf-8",
    )


def _write_plot_config(path: Path, plots: str) -> None:
    path.write_text(
        f"""version: 1
defaults:
  width_inches: 6
  height_inches: 3
  dpi: 80
  grid: true
  legend: true
  line_width: 1.0
plots:
{plots}
""",
        encoding="utf-8",
    )


def test_basic_plot_is_rendered_with_metadata(tmp_path: Path) -> None:
    data_path = tmp_path / "data.csv"
    config_path = tmp_path / "plots.yaml"
    output_dir = tmp_path / "out"
    _write_csv(data_path)
    _write_plot_config(
        config_path,
        """  - plot_id: speed_plot
    title: Speed
    x_channel_id: time__col_001
    output_filename: speed.png
    series:
      - channel_id: speed__col_002
        axis: primary
""",
    )

    result = render_plots(data_path, config_path, output_dir)

    assert result.sample_count == 3
    assert result.plot_count == 1
    assert result.series_count == 1
    image_path = Path(result.rendered_plots[0].output_file)
    assert image_path.exists()
    assert image_path.parent == output_dir / "png"
    assert image_path.read_bytes()[:8] == b"\x89PNG\r\n\x1a\n"
    assert (output_dir / "plot_catalogue.csv").exists()
    assert (output_dir / "plot_manifest.json").exists()
    assert (output_dir / "plotting_summary.txt").exists()


def _png_size(path: Path) -> tuple[int, int]:
    with path.open("rb") as handle:
        header = handle.read(24)
    return struct.unpack(">II", header[16:24])


def test_engineering_style_generates_expected_dimensions_svg_and_clean_labels(tmp_path: Path) -> None:
    data_path = tmp_path / "data.csv"
    config_path = tmp_path / "plots.yaml"
    _write_csv(data_path)
    config_path.write_text(
        """version: 1
defaults:
  width_inches: 7
  height_inches: 4
  dpi: 100
  grid: true
  legend: true
  line_width: 1.4
style:
  title_fontsize: 14
  axis_label_fontsize: 10
  tick_fontsize: 8
  legend_fontsize: 8
  output_formats: [png, svg]
plots:
  - plot_id: clean
    title: Vehicle Speed vs Time
    x_channel_id: time__col_001
    x_label: Time [s]
    primary_y_label: Speed [kph]
    secondary_y_label: Power [kW]
    output_filename: clean.png
    series:
      - channel_id: speed__col_002
        axis: primary
        label: Speed
      - channel_id: power__col_003
        axis: secondary
        label: power__col_003
""",
        encoding="utf-8",
    )

    before = set(plt.get_fignums())
    result = render_plots(data_path, config_path, tmp_path / "out")
    after = set(plt.get_fignums())
    item = result.rendered_plots[0]

    assert before == after
    assert item.figure_width_inches == 7
    assert item.figure_height_inches == 4
    assert item.dpi == 100
    assert item.axes_count == 2
    assert _png_size(Path(item.png_file)) == (700, 400)
    assert item.svg_file is not None
    assert Path(item.svg_file).exists()
    assert item.legend_labels == ("Speed", "power")
    assert "__col_" not in item.title
    assert all("__col_" not in label for label in item.legend_labels)


def test_optional_phase_boundaries_are_rendered_from_provenance(tmp_path: Path) -> None:
    data_path = tmp_path / "data.csv"
    config_path = tmp_path / "plots.yaml"
    provenance_path = tmp_path / "duty_cycle_provenance.csv"
    _write_csv(data_path)
    provenance_path.write_text(
        "output_sample_index,phase_id,phase_type\n"
        "0,P01,field\n"
        "1,P01,field\n"
        "2,P02,road\n",
        encoding="utf-8",
    )
    _write_plot_config(
        config_path,
        """  - plot_id: phase_plot
    title: Speed
    x_channel_id: time__col_001
    output_filename: speed.png
    show_phase_boundaries: true
    show_phase_labels: true
    series:
      - channel_id: speed__col_002
        axis: primary
""",
    )

    result = render_plots(data_path, config_path, tmp_path / "out", phase_provenance_file=provenance_path)

    assert len(result.phase_boundaries) == 2
    assert result.phase_aware_plot_count == 1
    assert result.rendered_plots[0].phase_boundary_count == 1


def test_repeated_generation_does_not_accumulate_open_figures(tmp_path: Path) -> None:
    data_path = tmp_path / "data.csv"
    config_path = tmp_path / "plots.yaml"
    _write_csv(data_path)
    _write_plot_config(
        config_path,
        """  - plot_id: speed_plot
    title: Speed
    x_channel_id: time__col_001
    output_filename: speed.png
    series:
      - channel_id: speed__col_002
        axis: primary
""",
    )

    before = set(plt.get_fignums())
    render_plots(data_path, config_path, tmp_path / "out1")
    render_plots(data_path, config_path, tmp_path / "out2")
    assert set(plt.get_fignums()) == before


def test_plot_manifest_records_engineering_quality_metadata(tmp_path: Path) -> None:
    data_path = tmp_path / "data.csv"
    config_path = tmp_path / "plots.yaml"
    output_dir = tmp_path / "out"
    _write_csv(data_path)
    _write_plot_config(
        config_path,
        """  - plot_id: speed_plot
    title: Speed
    x_channel_id: time__col_001
    output_filename: speed.png
    series:
      - channel_id: speed__col_002
        axis: primary
""",
    )

    render_plots(data_path, config_path, output_dir)
    manifest = json.loads((output_dir / "plot_manifest.json").read_text(encoding="utf-8"))

    assert manifest["defaults"]["style"]["title_fontsize"] == 15.0
    assert manifest["plots"][0]["png_file"].endswith("speed.png")
    assert manifest["plots"][0]["axes_count"] == 1


def test_secondary_axis_series_is_recorded(tmp_path: Path) -> None:
    data_path = tmp_path / "data.csv"
    config_path = tmp_path / "plots.yaml"
    _write_csv(data_path)
    _write_plot_config(
        config_path,
        """  - plot_id: combined
    title: Combined
    x_channel_id: time__col_001
    primary_y_label: Speed [kph]
    secondary_y_label: Power [kW]
    series:
      - channel_id: speed__col_002
        axis: primary
      - channel_id: power__col_003
        axis: secondary
""",
    )

    result = render_plots(data_path, config_path, tmp_path / "out")
    item = result.rendered_plots[0]
    assert item.primary_series_ids == ("speed__col_002",)
    assert item.secondary_series_ids == ("power__col_003",)


def test_missing_channel_reports_suggestions(tmp_path: Path) -> None:
    data_path = tmp_path / "data.csv"
    config_path = tmp_path / "plots.yaml"
    _write_csv(data_path)
    _write_plot_config(
        config_path,
        """  - plot_id: bad
    title: Bad
    x_channel_id: time__col_001
    series:
      - channel_id: speeed__col_002
""",
    )

    with pytest.raises(PlottingError, match="Configured plotting channel IDs were not found") as exc_info:
        render_plots(data_path, config_path, tmp_path / "out")
    assert "speed__col_002" in str(exc_info.value)


def test_duplicate_plot_ids_are_rejected(tmp_path: Path) -> None:
    config_path = tmp_path / "plots.yaml"
    _write_plot_config(
        config_path,
        """  - plot_id: duplicate
    title: One
    x_channel_id: time__col_001
    series:
      - channel_id: speed__col_002
  - plot_id: duplicate
    title: Two
    x_channel_id: time__col_001
    series:
      - channel_id: power__col_003
""",
    )

    with pytest.raises(ConfigurationError, match="duplicate plot IDs"):
        load_plotting_config(config_path)


def test_duplicate_output_filenames_are_rejected(tmp_path: Path) -> None:
    config_path = tmp_path / "plots.yaml"
    _write_plot_config(
        config_path,
        """  - plot_id: one
    title: One
    x_channel_id: time__col_001
    output_filename: same.png
    series:
      - channel_id: speed__col_002
  - plot_id: two
    title: Two
    x_channel_id: time__col_001
    output_filename: same.png
    series:
      - channel_id: power__col_003
""",
    )

    with pytest.raises(ConfigurationError, match="duplicate output filenames"):
        load_plotting_config(config_path)


def test_invalid_axis_is_rejected(tmp_path: Path) -> None:
    config_path = tmp_path / "plots.yaml"
    _write_plot_config(
        config_path,
        """  - plot_id: invalid_axis
    title: Invalid
    x_channel_id: time__col_001
    series:
      - channel_id: speed__col_002
        axis: tertiary
""",
    )

    with pytest.raises(ConfigurationError, match="axis must be 'primary' or 'secondary'"):
        load_plotting_config(config_path)


def test_output_filename_must_be_plain_png(tmp_path: Path) -> None:
    config_path = tmp_path / "plots.yaml"
    _write_plot_config(
        config_path,
        """  - plot_id: bad_file
    title: Invalid
    x_channel_id: time__col_001
    output_filename: nested/bad.png
    series:
      - channel_id: speed__col_002
""",
    )

    with pytest.raises(ConfigurationError, match="plain .png filename"):
        load_plotting_config(config_path)


def test_unknown_configuration_keys_are_rejected(tmp_path: Path) -> None:
    config_path = tmp_path / "plots.yaml"
    _write_plot_config(
        config_path,
        """  - plot_id: bad_key
    title: Invalid
    x_channel_id: time__col_001
    typo_field: true
    series:
      - channel_id: speed__col_002
""",
    )

    with pytest.raises(ConfigurationError, match="Unknown key"):
        load_plotting_config(config_path)


PROJECT_ROOT = Path(__file__).resolve().parents[1]
MATH_CONFIG = PROJECT_ROOT / "config" / "math_channels_example.yaml"
PLOTTING_CONFIG = PROJECT_ROOT / "config" / "plotting_example.yaml"
REPORT_PLOTTING_CONFIG = PROJECT_ROOT / "config" / "plotting_reference_report.yaml"
CHART_INVENTORY = PROJECT_ROOT / "docs" / "phase_1" / "excel_chart_inventory.csv"


def test_reference_plot_config_maps_all_meaningful_excel_charts() -> None:
    config = load_plotting_config(REPORT_PLOTTING_CONFIG)
    with CHART_INVENTORY.open(newline="", encoding="utf-8-sig") as handle:
        meaningful = [row for row in csv.DictReader(handle) if row["status"] == "meaningful"]

    assert len(config.plots) == 18
    assert {item.reference_chart_number for item in config.plots} == {
        int(row["chart_number"]) for row in meaningful
    }
    by_number = {item.reference_chart_number: item for item in config.plots}
    for row in meaningful:
        definition = by_number[int(row["chart_number"])]
        assert definition.title.strip() == row["title"].strip()
        assert len(definition.series) == int(row["series_count"])


def test_supplied_source_workbook_plotting_acceptance(tmp_path: Path) -> None:
    source_workbook = require_private_reference_file(CAIMAN_REFERENCE_XLSX, CAIMAN_REFERENCE_DESCRIPTION)
    result = render_plots(
        source_workbook,
        PLOTTING_CONFIG,
        tmp_path / "plots",
        ImportOptions(strict=True),
        math_config_file=MATH_CONFIG,
    )

    assert result.sample_count == 1866
    assert len(result.channels_by_id) == 83
    assert result.plot_count == 24
    assert result.series_count == 45
    assert all(Path(item.output_file).exists() for item in result.rendered_plots)


def test_supplied_report_plotting_config_channels_are_resolvable(tmp_path: Path) -> None:
    report_workbook = require_private_reference_file(
        CAIMAN_PROFILE_REFERENCE_XLSX,
        CAIMAN_PROFILE_REFERENCE_DESCRIPTION,
    )
    # Render just the full configured reference set to prove every mapped source range is resolvable.
    result = render_plots(
        report_workbook,
        REPORT_PLOTTING_CONFIG,
        tmp_path / "reference_plots",
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
    assert result.plot_count == 18
    assert result.series_count == 34
    assert {item.reference_chart_number for item in result.rendered_plots} == {
        1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 16, 17, 18, 19, 20
    }
