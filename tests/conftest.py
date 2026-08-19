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
ROBOSPRAYER_LATEST_ELECTRIC_CSV = (
    REFERENCE_FILES_DIR
    / "RoboSprayer_3500Kg_Electric_12kph_Batt_50kW_Mot_63RPM_Susp_Cool_Rough_Grad_Discharge.csv"
)
ROBOSPRAYER_LATEST_HYBRID_CSV = (
    REFERENCE_FILES_DIR / ("Sprayer_" + "Cai" + "man_SP_9300Kg_Hybrid_12x1Km_4000Kg_Chem.csv")
)

ROBOSPRAYER_REFERENCE_DESCRIPTION = "RoboSprayer raw reference CSV"
ROBOSPRAYER_LATEST_ELECTRIC_DESCRIPTION = "latest RoboSprayer Electric source CSV"
ROBOSPRAYER_LATEST_HYBRID_DESCRIPTION = "latest RoboSprayer Hybrid source CSV"


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
