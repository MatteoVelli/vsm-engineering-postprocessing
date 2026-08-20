from __future__ import annotations

from pathlib import Path

from vsm_postprocessing.report_metadata import resolve_report_metadata
from vsm_postprocessing.report_profile import load_reporting_profile


ELECTRIC_PROFILE = Path("config/report_profiles/robosprayer_electric.yaml")
HYBRID_PROFILE = Path("config/report_profiles/robosprayer_hybrid.yaml")


def test_report_metadata_resolves_robosprayer_electric_filename() -> None:
    profile = load_reporting_profile(ELECTRIC_PROFILE)
    metadata = resolve_report_metadata(
        "RoboSprayer_3500Kg_Electric_12kph_Batt_50kW_Mot_63RPM_Susp_Cool_Rough_Grad_Discharge.csv",
        profile,
    )

    assert metadata.machine_name == "RoboSprayer"
    assert metadata.powertrain_name == "Electric"
    assert metadata.report_title == "RoboSprayer Electric"
    assert metadata.safe_output_stem == "RoboSprayer_Electric"


def test_report_metadata_resolves_caiman_hybrid_filename() -> None:
    profile = load_reporting_profile(HYBRID_PROFILE)
    metadata = resolve_report_metadata(
        "Sprayer_Caiman_SP_9300Kg_Hybrid_12x1Km_4000Kg_Chem.csv",
        profile,
    )

    assert metadata.machine_name == "Caiman SP"
    assert metadata.powertrain_name == "Hybrid"
    assert metadata.report_title == "Caiman SP Hybrid"
    assert metadata.safe_output_stem == "Caiman_SP_Hybrid"


def test_report_metadata_user_override_wins() -> None:
    profile = load_reporting_profile(HYBRID_PROFILE)
    metadata = resolve_report_metadata(
        "Sprayer_Caiman_SP_9300Kg_Hybrid_12x1Km_4000Kg_Chem.csv",
        profile,
        machine_name_override="Caiman SP Prototype B",
    )

    assert metadata.machine_name == "Caiman SP Prototype B"
    assert metadata.report_title == "Caiman SP Prototype B Hybrid"
    assert metadata.detection_source == "user_override"


def test_report_metadata_source_metadata_precedes_filename() -> None:
    profile = load_reporting_profile(HYBRID_PROFILE)
    metadata = resolve_report_metadata(
        "Sprayer_Caiman_SP_9300Kg_Hybrid_12x1Km_4000Kg_Chem.csv",
        profile,
        source_metadata={"vehicle_name": "Caiman SP Lab Unit"},
    )

    assert metadata.machine_name == "Caiman SP Lab Unit"
    assert metadata.detection_source == "source_metadata"


def test_report_metadata_unknown_future_machine_is_readable() -> None:
    profile = load_reporting_profile(ELECTRIC_PROFILE)
    metadata = resolve_report_metadata(
        "OrchardSeeder_X9_Prototype_6100Kg_Electric_Test_Run_07.csv",
        profile,
    )

    assert metadata.machine_name == "OrchardSeeder X9 Prototype"
    assert metadata.report_title == "OrchardSeeder X9 Prototype Electric"
    assert metadata.safe_output_stem == "OrchardSeeder_X9_Prototype_Electric"
