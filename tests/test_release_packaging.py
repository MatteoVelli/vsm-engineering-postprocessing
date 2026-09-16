from __future__ import annotations

import hashlib
import tomllib
from pathlib import Path
from zipfile import ZipFile

import vsm_postprocessing
from vsm_postprocessing.release_builder import _CLIENT_DIRECTORIES, build_client_release


PROJECT_ROOT = Path(__file__).resolve().parents[1]


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    digest.update(path.read_bytes())
    return digest.hexdigest()


def test_release_version_metadata_is_consistent() -> None:
    with (PROJECT_ROOT / "pyproject.toml").open("rb") as handle:
        pyproject = tomllib.load(handle)
    assert pyproject["project"]["version"] == vsm_postprocessing.__version__
    assert vsm_postprocessing.__version__ == "1.3.0"


def test_client_release_excludes_private_and_development_artifacts(tmp_path: Path) -> None:
    result = build_client_release(PROJECT_ROOT, tmp_path / "dist")
    assert result.archive_path.exists()
    assert result.checksum_path.exists()

    with ZipFile(result.archive_path) as archive:
        names = archive.namelist()

    assert any(name.endswith("/RELEASE_MANIFEST.json") for name in names)
    assert any(name.endswith("/SETUP_VSM_TOOL.bat") for name in names)
    assert any(name.endswith("/START_VSM_TOOL.bat") for name in names)
    for asset_name in (
        "astauto-light-text_web.jpg",
        "RoboSprayer_Electric_Report_Astauto_Colours.pptx",
        "Caiman_SP_Hybrid_Report_Astauto_Colours.pptx",
        "RoboSprayer_Electric_Report_Astauto_v7.pptx",
        "Robo_Sprayer_Electrification_Tamplate_Electric_03.xlsx",
        "Robo_Sprayer_Electrification_Tamplate_Hybrid_04.xlsx",
    ):
        assert any(name.endswith(f"/reference_files/{asset_name}") for name in names)
    retired_asset_fragment = "/assets/scenarios/" + "cai" + "man" + "_sp_hybrid/"
    assert not any(retired_asset_fragment in name for name in names)
    assert not any("duty" + "_cycle" in name for name in names)
    assert not any("/tests/" in name for name in names)
    assert not any("Sprayer_" + "Cai" + "man" in name for name in names)
    assert not any("Road_Profile_Reference" in name for name in names)
    assert not any("Wheel_Steering_Angles_Reference" in name for name in names)
    assert not any("outputs/end_to_end" in name for name in names)
    assert "assets" not in _CLIENT_DIRECTORIES


def test_client_release_build_is_deterministic(tmp_path: Path) -> None:
    first = build_client_release(PROJECT_ROOT, tmp_path / "first")
    second = build_client_release(PROJECT_ROOT, tmp_path / "second")
    assert _sha256(first.archive_path) == _sha256(second.archive_path)


def test_client_release_does_not_require_removed_assets_directory(tmp_path: Path) -> None:
    project = tmp_path / "project"
    for directory in (".streamlit", "config", "scripts", "src", "docs", "reference_files", "outputs"):
        (project / directory).mkdir(parents=True)
    for filename in ("SETUP_VSM_TOOL.bat", "START_VSM_TOOL.bat", "README.md", "CHANGELOG.md", "pyproject.toml", ".python-version"):
        (project / filename).write_text(f"{filename}\n", encoding="utf-8")
    for relative in (
        "docs/CLIENT_QUICK_START.md",
        "docs/TROUBLESHOOTING.md",
        "docs/CONFIGURATION_GUIDE.md",
        "docs/KNOWN_LIMITATIONS.md",
        "docs/FINAL_RELEASE_NOTES.md",
        "docs/GIT_RELEASE.md",
        "reference_files/README.md",
        "outputs/.gitkeep",
        ".streamlit/config.toml",
        "config/example.yaml",
        "scripts/start.ps1",
        "src/package.py",
        "reference_files/astauto-light-text_web.jpg",
        "reference_files/RoboSprayer_Electric_Report_Astauto_Colours.pptx",
        "reference_files/Caiman_SP_Hybrid_Report_Astauto_Colours.pptx",
        "reference_files/RoboSprayer_Electric_Report_Astauto_v7.pptx",
        "reference_files/Robo_Sprayer_Electrification_Tamplate_Electric_03.xlsx",
        "reference_files/Robo_Sprayer_Electrification_Tamplate_Hybrid_04.xlsx",
    ):
        path = project / relative
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(f"{relative}\n", encoding="utf-8")

    result = build_client_release(project, tmp_path / "dist")

    assert result.archive_path.exists()
    assert not (project / "assets").exists()
