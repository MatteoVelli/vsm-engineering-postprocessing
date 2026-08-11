from __future__ import annotations

import hashlib
import tomllib
from pathlib import Path
from zipfile import ZipFile

import vsm_postprocessing
from vsm_postprocessing.release_builder import build_client_release


PROJECT_ROOT = Path(__file__).resolve().parents[1]


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    digest.update(path.read_bytes())
    return digest.hexdigest()


def test_release_version_metadata_is_consistent() -> None:
    with (PROJECT_ROOT / "pyproject.toml").open("rb") as handle:
        pyproject = tomllib.load(handle)
    assert pyproject["project"]["version"] == vsm_postprocessing.__version__
    assert vsm_postprocessing.__version__ == "1.2.1"


def test_client_release_excludes_private_and_development_artifacts(tmp_path: Path) -> None:
    result = build_client_release(PROJECT_ROOT, tmp_path / "dist")
    assert result.archive_path.exists()
    assert result.checksum_path.exists()

    with ZipFile(result.archive_path) as archive:
        names = archive.namelist()

    assert any(name.endswith("/RELEASE_MANIFEST.json") for name in names)
    assert any(name.endswith("/START_VSM_TOOL.bat") for name in names)
    assert not any("/tests/" in name for name in names)
    assert not any("Sprayer_Caiman" in name for name in names)
    assert not any(name.lower().endswith((".xlsx", ".pptx")) for name in names)
    assert not any("outputs/end_to_end" in name for name in names)


def test_client_release_build_is_deterministic(tmp_path: Path) -> None:
    first = build_client_release(PROJECT_ROOT, tmp_path / "first")
    second = build_client_release(PROJECT_ROOT, tmp_path / "second")
    assert _sha256(first.archive_path) == _sha256(second.archive_path)
