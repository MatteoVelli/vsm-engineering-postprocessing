from __future__ import annotations

import os
from pathlib import Path

import pytest


PROJECT_ROOT = Path(__file__).resolve().parents[1]
REFERENCE_FILES_DIR = Path(os.environ.get("VSM_TEST_REFERENCE_FILES_DIR", PROJECT_ROOT / "reference_files"))
if not REFERENCE_FILES_DIR.is_absolute():
    REFERENCE_FILES_DIR = PROJECT_ROOT / REFERENCE_FILES_DIR

ROBOSPRAYER_REFERENCE_CSV = (
    REFERENCE_FILES_DIR
    / "RoboSprayer_3500Kg_Electric_12kph_Batt_50kW_Motor_63RPM_Susp_Cool_Rough_Crop_Field_05.csv"
)
CAIMAN_REFERENCE_XLSX = (
    REFERENCE_FILES_DIR
    / "Sprayer_Caiman_SP_9300Kg_Hybrid_Gen80kW_30kph_74Ht_4000KgAQ_57-4pcSOC_5-80_1C2G_02.xlsx"
)
CAIMAN_PROFILE_REFERENCE_XLSX = REFERENCE_FILES_DIR / "Sprayer_Caiman_SP_9300Kg_Electrification_03.xlsx"

ROBOSPRAYER_REFERENCE_DESCRIPTION = "RoboSprayer raw reference CSV"
CAIMAN_REFERENCE_DESCRIPTION = "private Caiman reference workbook"
CAIMAN_PROFILE_REFERENCE_DESCRIPTION = "private Caiman profile reference workbook"


def pytest_configure(config: pytest.Config) -> None:
    config.addinivalue_line(
        "markers",
        "private_reference: test requires an optional private client reference dataset",
    )


def require_private_reference_file(path: Path, description: str) -> Path:
    candidate = Path(path)
    if not candidate.exists():
        pytest.skip(f"private client reference dataset not available: {description} ({candidate})")
    return candidate


def private_reference_mark(path: Path, description: str) -> pytest.MarkDecorator:
    return pytest.mark.skipif(
        not Path(path).exists(),
        reason=f"private client reference dataset not available: {description}",
    )


def require_private_reference_files(*items: tuple[Path, str]) -> tuple[Path, ...]:
    return tuple(require_private_reference_file(path, description) for path, description in items)
