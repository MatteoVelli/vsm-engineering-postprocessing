"""Deterministic client-release packaging for the VSM application."""

from __future__ import annotations

import hashlib
import json
from dataclasses import dataclass
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

from .version import __version__


_FIXED_ZIP_TIMESTAMP = (2026, 1, 1, 0, 0, 0)


@dataclass(frozen=True)
class ReleaseBuildResult:
    archive_path: Path
    checksum_path: Path
    archive_sha256: str
    file_count: int


_CLIENT_ROOT_FILES = (
    "START_VSM_TOOL.bat",
    "README.md",
    "CHANGELOG.md",
    "pyproject.toml",
    ".python-version",
)

_CLIENT_DIRECTORIES = (
    ".streamlit",
    "config",
    "scripts",
    "src",
)

_CLIENT_DOCS = (
    "docs/CLIENT_QUICK_START.md",
    "docs/TROUBLESHOOTING.md",
    "docs/CONFIGURATION_GUIDE.md",
    "docs/KNOWN_LIMITATIONS.md",
    "docs/FINAL_RELEASE_NOTES.md",
    "docs/GIT_RELEASE.md",
)

_CLIENT_PLACEHOLDERS = (
    "reference_files/README.md",
    "outputs/.gitkeep",
)

_EXCLUDED_NAMES = {
    "__pycache__",
    ".pytest_cache",
    ".mypy_cache",
    ".ruff_cache",
    ".venv",
    ".git",
}

_EXCLUDED_SUFFIXES = {".pyc", ".pyo"}


def _sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _is_allowed(path: Path, project_root: Path) -> bool:
    relative = path.relative_to(project_root)
    if any(part in _EXCLUDED_NAMES for part in relative.parts):
        return False
    if path.suffix.lower() in _EXCLUDED_SUFFIXES:
        return False
    return path.is_file()


def _collect_client_files(project_root: Path) -> list[Path]:
    candidates: list[Path] = []

    for relative in _CLIENT_ROOT_FILES + _CLIENT_DOCS + _CLIENT_PLACEHOLDERS:
        path = project_root / relative
        if not path.exists():
            raise FileNotFoundError(f"Required release file is missing: {relative}")
        candidates.append(path)

    for directory in _CLIENT_DIRECTORIES:
        root = project_root / directory
        if not root.exists():
            raise FileNotFoundError(f"Required release directory is missing: {directory}")
        candidates.extend(path for path in root.rglob("*") if _is_allowed(path, project_root))

    unique = {path.resolve(): path for path in candidates if _is_allowed(path, project_root)}
    return sorted(unique.values(), key=lambda path: path.relative_to(project_root).as_posix().lower())


def _zip_write_bytes(archive: ZipFile, archive_name: str, data: bytes) -> None:
    info = ZipInfo(archive_name, date_time=_FIXED_ZIP_TIMESTAMP)
    info.compress_type = ZIP_DEFLATED
    info.external_attr = 0o644 << 16
    archive.writestr(info, data)


def build_client_release(project_root: str | Path, output_dir: str | Path) -> ReleaseBuildResult:
    """Build a deterministic client ZIP and SHA-256 checksum file."""

    root = Path(project_root).expanduser().resolve()
    destination = Path(output_dir).expanduser().resolve()
    destination.mkdir(parents=True, exist_ok=True)

    files = _collect_client_files(root)
    package_root = f"VSM_Engineering_PostProcessing_v{__version__}"
    archive_path = destination / f"VSM_Engineering_PostProcessing_v{__version__}_Client.zip"
    checksum_path = destination / f"VSM_Engineering_PostProcessing_v{__version__}_Client.sha256"

    manifest_files: list[dict[str, object]] = []
    for path in files:
        relative = path.relative_to(root).as_posix()
        data = path.read_bytes()
        manifest_files.append(
            {
                "path": relative,
                "size_bytes": len(data),
                "sha256": _sha256_bytes(data),
            }
        )

    release_manifest = {
        "product": "VSM Engineering Data Post-Processing",
        "version": __version__,
        "release_type": "client_source_distribution",
        "deterministic_engineering_calculations": True,
        "client_reference_data_included": False,
        "file_count": len(manifest_files),
        "files": manifest_files,
    }
    manifest_bytes = (json.dumps(release_manifest, indent=2, sort_keys=True) + "\n").encode("utf-8")

    with ZipFile(archive_path, "w") as archive:
        for path in files:
            relative = path.relative_to(root).as_posix()
            _zip_write_bytes(archive, f"{package_root}/{relative}", path.read_bytes())
        _zip_write_bytes(archive, f"{package_root}/RELEASE_MANIFEST.json", manifest_bytes)

    archive_sha256 = _sha256_file(archive_path)
    checksum_path.write_text(f"{archive_sha256}  {archive_path.name}\n", encoding="utf-8")

    return ReleaseBuildResult(
        archive_path=archive_path,
        checksum_path=checksum_path,
        archive_sha256=archive_sha256,
        file_count=len(manifest_files) + 1,
    )
