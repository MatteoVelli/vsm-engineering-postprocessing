from __future__ import annotations

import copy
import json
import zipfile
from dataclasses import replace
from pathlib import Path

import pytest
from pptx import Presentation

from vsm_postprocessing.importer import ImportOptions
from vsm_postprocessing.profile_powerpoint_report_engine import (
    _hybrid_subsystem_active,
    _profile_powerpoint_config,
    inspect_reference_powerpoint_layout,
)
from vsm_postprocessing.ui_config import generate_reporting_profile_engineering_report


PROJECT_ROOT = Path(__file__).resolve().parents[1]
REFERENCE_DECK = PROJECT_ROOT / "reference_files" / "Hybrid_SP_Caiman_Sprayer_Report_FINAL (1).pptx"
ROBOSPRAYER_CSV = PROJECT_ROOT / "reference_files" / "RoboSprayer_3500Kg_Electric_12kph_Batt_50kW_Motor_63RPM_Susp_Cool_Rough_Crop_Field_05.csv"
ELECTRIC_PROFILE = PROJECT_ROOT / "config" / "report_profiles" / "robosprayer_electric.yaml"
HYBRID_PROFILE = PROJECT_ROOT / "config" / "report_profiles" / "robosprayer_hybrid.yaml"


@pytest.fixture(scope="module")
def electric_report(tmp_path_factory: pytest.TempPathFactory):
    return generate_reporting_profile_engineering_report(
        ROBOSPRAYER_CSV,
        ELECTRIC_PROFILE,
        tmp_path_factory.mktemp("profile_ppt_electric"),
        ImportOptions(strict=True),
    )


@pytest.fixture(scope="module")
def hybrid_report(tmp_path_factory: pytest.TempPathFactory):
    return generate_reporting_profile_engineering_report(
        ROBOSPRAYER_CSV,
        HYBRID_PROFILE,
        tmp_path_factory.mktemp("profile_ppt_hybrid"),
        ImportOptions(strict=True),
    )


@pytest.mark.skipif(not REFERENCE_DECK.exists(), reason="Reference PowerPoint deck is not present")
def test_reference_powerpoint_layout_inspection_matches_sergio_ready_deck() -> None:
    spec = inspect_reference_powerpoint_layout(REFERENCE_DECK)

    assert spec.slide_count == 10
    assert spec.slide_width_in == pytest.approx(13.333)
    assert spec.slide_height_in == pytest.approx(7.5)
    assert spec.titles[:3] == (
        "VSM ENGINEERING   \xb7   FEASIBILITY STUDY",
        "System and Mission Overview",
        "Duty-Cycle Executive Results",
    )
    assert spec.slides[0].picture_count >= 1
    assert spec.slides[6].title == "Range Extender and Generator"


@pytest.mark.skipif(not ROBOSPRAYER_CSV.exists(), reason="RoboSprayer CSV is not present")
def test_profile_powerpoint_generates_reopenable_electric_and_hybrid_decks(electric_report, hybrid_report) -> None:
    assert electric_report.presentation_path.name == "RoboSprayer_Electric_Engineering_Report.pptx"
    assert hybrid_report.presentation_path.name == "RoboSprayer_Hybrid_Engineering_Report.pptx"

    for result in (electric_report, hybrid_report):
        assert result.report_path.exists()
        assert result.presentation_path.exists()
        assert result.slide_count == 10
        with zipfile.ZipFile(result.presentation_path) as package:
            assert package.testzip() is None
            media = [name for name in package.namelist() if name.startswith("ppt/media/")]
        assert len(media) >= 10
        prs = Presentation(result.presentation_path)
        assert len(prs.slides) == 10


@pytest.mark.skipif(not ROBOSPRAYER_CSV.exists(), reason="RoboSprayer CSV is not present")
def test_profile_powerpoint_slide_titles_are_profile_conditional(electric_report, hybrid_report) -> None:
    assert _slide_titles(electric_report.presentation_path) == [
        "RoboSprayer Electric",
        "System and Simulation Overview",
        "Executive Results",
        "Vehicle Operation",
        "Battery and Electrical Energy System",
        "Battery Power and Energy Recovery",
        "Energy Consumption and Estimated Range",
        "Agrochemical and Battery Behaviour",
        "Traction, EDU and Auxiliary Energy Demand",
        "Simulation Summary",
    ]
    assert _slide_titles(hybrid_report.presentation_path) == [
        "RoboSprayer Hybrid",
        "System and Simulation Overview",
        "Executive Results",
        "Vehicle Operation",
        "Battery and Electrical Energy System",
        "Battery Power and Energy Recovery",
        "Range Extender and Generator",
        "Agrochemical and Battery Behaviour",
        "Traction, EDU and Auxiliary Energy Demand",
        "Simulation Summary",
    ]


@pytest.mark.skipif(not ROBOSPRAYER_CSV.exists(), reason="RoboSprayer CSV is not present")
def test_profile_powerpoint_uses_profile_kpis_and_omits_caiman_leakage(electric_report) -> None:
    text = _visible_text(electric_report.presentation_path)
    values = {
        item.definition.statistic_id: item.value
        for item in electric_report.excel_result.statistics_result.statistics
    }
    values.update(
        {
            item.definition.kpi_id: item.value
            for item in electric_report.excel_result.statistics_result.kpis
        }
    )

    assert f"{values['battery_soc_last']:.2f} %" in text
    assert f"{values['battery_power_rms']:.2f} kW" in text
    assert f"{values['distance_km_last']:.2f} km" in text
    assert f"{values['chassis_speed_max']:.2f} kph" in text
    assert "Generator" not in text
    assert "Fuel" not in text
    assert "Caiman" not in text
    assert "17,418" not in text
    assert "114.00" not in text
    assert "290.28" not in text
    assert "39.84" not in text
    assert "23.94" not in text
    assert "C:\\Users\\" not in text
    assert "Desktop\\Agro Project" not in text


@pytest.mark.skipif(not ROBOSPRAYER_CSV.exists(), reason="RoboSprayer CSV is not present")
def test_profile_powerpoint_hybrid_inactive_subsystem_handling(hybrid_report) -> None:
    text = _visible_text(hybrid_report.presentation_path)

    assert _hybrid_subsystem_active(hybrid_report.excel_result) is False
    assert "Range Extender and Generator" in text
    assert "ICE/generator channels resolved; no operating activity detected in this simulation." in text
    assert "Resolved ICE/generator channels are inactive/all-zero for this run." in text


@pytest.mark.skipif(not ROBOSPRAYER_CSV.exists(), reason="RoboSprayer CSV is not present")
def test_profile_powerpoint_future_active_hybrid_uses_active_slide_copy(hybrid_report) -> None:
    excel_result = copy.copy(hybrid_report.excel_result)
    statistics_result = copy.copy(hybrid_report.excel_result.statistics_result)
    statistics_result.statistics = [
        replace(item, value=1200.0)
        if item.definition.statistic_id == "engine_speed_max"
        else item
        for item in statistics_result.statistics
    ]
    excel_result.statistics_result = statistics_result

    assert _hybrid_subsystem_active(excel_result) is True
    config = _profile_powerpoint_config(excel_result, output_filename="active_hybrid.pptx")
    slide_7 = config["slides"][6]
    assert slide_7["title"] == "Range Extender and Generator"
    assert slide_7["subtitle"] == "Engine demand, fuel consumption and generator behavior"
    assert slide_7["body"] == [
        "Range-extender activity is derived from non-zero resolved engine/generator statistics."
    ]


@pytest.mark.skipif(not ROBOSPRAYER_CSV.exists(), reason="RoboSprayer CSV is not present")
def test_profile_powerpoint_plot_ids_and_manifest_are_deterministic(electric_report, hybrid_report) -> None:
    electric_manifest = json.loads(electric_report.manifest_path.read_text(encoding="utf-8"))
    hybrid_manifest = json.loads(hybrid_report.manifest_path.read_text(encoding="utf-8"))

    assert electric_manifest["slides"][3]["plots"] == ["speed_vs_distance"]
    assert electric_manifest["slides"][6]["plots"] == [
        "battery_energy_distance_based",
        "auxiliaries_energy_consumption",
    ]
    assert hybrid_manifest["slides"][6]["plots"] == [
        "generator_power",
        "engine_power_and_fuel_vs_time",
    ]
    assert electric_manifest["slides"][6]["title"] == "Energy Consumption and Estimated Range"
    assert hybrid_manifest["slides"][6]["title"] == "Range Extender and Generator"


def test_streamlit_profile_workflow_exposes_powerpoint_without_custom_analysis_regression() -> None:
    source = (PROJECT_ROOT / "src" / "vsm_postprocessing" / "ui_app.py").read_text(encoding="utf-8")

    assert "Custom Analysis" in source
    assert "generate_reporting_profile_engineering_report" in source
    assert "Download PowerPoint Engineering Report" in source
    assert "PowerPoint output is not available yet" not in source


def _slide_titles(path: Path) -> list[str]:
    prs = Presentation(path)
    titles = []
    for slide in prs.slides:
        titles.append(next(getattr(shape, "text", "").strip() for shape in slide.shapes if getattr(shape, "text", "").strip()))
    return titles


def _visible_text(path: Path) -> str:
    prs = Presentation(path)
    return "\n".join(
        getattr(shape, "text", "")
        for slide in prs.slides
        for shape in slide.shapes
        if getattr(shape, "text", "").strip()
    )
