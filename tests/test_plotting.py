from __future__ import annotations

import json
import struct
from pathlib import Path

import matplotlib.pyplot as plt
import pytest

from vsm_postprocessing.errors import ConfigurationError, PlottingError
from vsm_postprocessing.plotting_engine import load_plotting_config, render_plots


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
