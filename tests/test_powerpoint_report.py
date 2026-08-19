from __future__ import annotations

from pathlib import Path

import pytest

from vsm_postprocessing.errors import ConfigurationError
from vsm_postprocessing.powerpoint_report_engine import load_powerpoint_report_config

PROJECT_ROOT = Path(__file__).resolve().parents[1]
POWERPOINT_CONFIG = PROJECT_ROOT / "config" / "powerpoint_report_example.yaml"


def test_powerpoint_config_loads() -> None:
    config = load_powerpoint_report_config(POWERPOINT_CONFIG)
    assert config.version == 1
    assert config.output_filename == "vsm_engineering_report.pptx"
    assert len(config.slides) == 4
    assert config.slides[1].slide_type == "plot_pair"
    assert config.slides[1].plot_ids == ("speed_vs_time", "battery_soc_vs_distance")


def test_powerpoint_config_rejects_invalid_slide_type(tmp_path: Path) -> None:
    path = tmp_path / "bad.yaml"
    path.write_text(
        """version: 1
presentation:
  title: Test
slides:
  - slide_id: bad
    type: magic
    title: Bad
    statistics: []
    plots: []
""",
        encoding="utf-8",
    )
    with pytest.raises(ConfigurationError, match="type must be one of"):
        load_powerpoint_report_config(path)


def test_powerpoint_config_rejects_three_plots(tmp_path: Path) -> None:
    path = tmp_path / "bad.yaml"
    path.write_text(
        """version: 1
presentation:
  title: Test
slides:
  - slide_id: plots
    type: plot_pair
    title: Too many
    statistics: []
    plots: [a, b, c]
""",
        encoding="utf-8",
    )
    with pytest.raises(ConfigurationError, match="one or two"):
        load_powerpoint_report_config(path)

