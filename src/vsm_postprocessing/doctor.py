from __future__ import annotations

import importlib.metadata
import os
import platform
import shutil
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from .channel_manager import load_selection_config
from .errors import ConfigurationError, VSMPostProcessingError
from .excel_report_engine import load_excel_report_config
from .math_engine import load_math_config
from .pipeline_engine import load_pipeline_config
from .plotting_engine import load_plotting_config
from .powerpoint_report_engine import load_powerpoint_report_config
from .statistics_engine import load_statistics_config


@dataclass(frozen=True)
class DoctorCheck:
    name: str
    status: str
    detail: str


@dataclass(frozen=True)
class DoctorReport:
    project_root: Path
    checks: tuple[DoctorCheck, ...]

    @property
    def error_count(self) -> int:
        return sum(check.status == "FAIL" for check in self.checks)

    @property
    def warning_count(self) -> int:
        return sum(check.status == "WARN" for check in self.checks)

    @property
    def status(self) -> str:
        return "FAIL" if self.error_count else "PASS"


_REQUIRED_PACKAGES = (
    "numpy",
    "openpyxl",
    "PyYAML",
    "matplotlib",
    "streamlit",
    "python-pptx",
)


def run_doctor(project_root: str | Path, pipeline_config: str | Path | None = None) -> DoctorReport:
    root = Path(project_root).expanduser().resolve()
    checks: list[DoctorCheck] = []

    checks.append(_check_python_version())
    checks.append(_check_uv())
    checks.extend(_check_packages())
    checks.append(_check_project_structure(root))
    checks.append(_check_output_writable(root / "outputs"))
    checks.append(_check_disk_space(root))

    config_path = Path(pipeline_config).expanduser().resolve() if pipeline_config else root / "config" / "end_to_end_example.yaml"
    checks.extend(_check_pipeline_bundle(config_path))

    return DoctorReport(project_root=root, checks=tuple(checks))


def _check_python_version() -> DoctorCheck:
    version = sys.version_info
    detail = f"Python {version.major}.{version.minor}.{version.micro} ({sys.executable})"
    if version < (3, 11):
        return DoctorCheck("Python version", "FAIL", detail + "; Python >= 3.11 is required")
    return DoctorCheck("Python version", "PASS", detail)


def _check_uv() -> DoctorCheck:
    location = shutil.which("uv")
    if location:
        return DoctorCheck("uv launcher", "PASS", location)
    return DoctorCheck(
        "uv launcher",
        "WARN",
        "uv is not on PATH. The already-created .venv can still run, but one-click setup/update will not work.",
    )


def _check_packages() -> list[DoctorCheck]:
    checks: list[DoctorCheck] = []
    for package in _REQUIRED_PACKAGES:
        try:
            version = importlib.metadata.version(package)
        except importlib.metadata.PackageNotFoundError:
            checks.append(DoctorCheck(f"Package {package}", "FAIL", "not installed"))
        else:
            checks.append(DoctorCheck(f"Package {package}", "PASS", version))
    try:
        app_version = importlib.metadata.version("vsm-postprocessing")
    except importlib.metadata.PackageNotFoundError:
        checks.append(DoctorCheck("VSM package", "FAIL", "vsm-postprocessing is not installed"))
    else:
        checks.append(DoctorCheck("VSM package", "PASS", app_version))
    return checks


def _check_project_structure(root: Path) -> DoctorCheck:
    required = ["config", "scripts", "src", "outputs"]
    missing = [name for name in required if not (root / name).exists()]
    if missing:
        return DoctorCheck("Project structure", "FAIL", "missing: " + ", ".join(missing))
    return DoctorCheck("Project structure", "PASS", str(root))


def _check_output_writable(output_dir: Path) -> DoctorCheck:
    try:
        output_dir.mkdir(parents=True, exist_ok=True)
        fd, name = tempfile.mkstemp(prefix=".vsm_write_test_", dir=output_dir)
        os.close(fd)
        Path(name).unlink(missing_ok=True)
    except OSError as exc:
        return DoctorCheck("Output directory", "FAIL", f"not writable: {exc}")
    return DoctorCheck("Output directory", "PASS", str(output_dir))


def _check_disk_space(root: Path) -> DoctorCheck:
    try:
        free = shutil.disk_usage(root).free
    except OSError as exc:
        return DoctorCheck("Free disk space", "WARN", f"could not determine: {exc}")
    gib = free / (1024 ** 3)
    detail = f"{gib:.2f} GiB free"
    if gib < 0.5:
        return DoctorCheck("Free disk space", "FAIL", detail + "; at least 0.5 GiB is recommended")
    if gib < 2.0:
        return DoctorCheck("Free disk space", "WARN", detail + "; 2 GiB or more is recommended")
    return DoctorCheck("Free disk space", "PASS", detail)


def _check_pipeline_bundle(config_path: Path) -> list[DoctorCheck]:
    if not config_path.exists():
        return [DoctorCheck("Default pipeline configuration", "WARN", f"not found: {config_path}")]

    try:
        config = load_pipeline_config(config_path)
    except ConfigurationError as exc:
        # A release can intentionally omit Sergio's source workbook. If that is the only issue,
        # the UI remains fully usable because it creates a runtime pipeline config for uploaded data.
        text = str(exc)
        if "input.file does not exist" in text:
            return [
                DoctorCheck(
                    "Default pipeline configuration",
                    "WARN",
                    "configuration is present but its example source file is not available; UI uploads are still supported",
                )
            ]
        return [DoctorCheck("Default pipeline configuration", "FAIL", text)]

    validators: tuple[tuple[str, Path, Callable[[Path], object]], ...] = (
        ("Channel-selection configuration", config.channel_selection_config, load_selection_config),
        ("Math configuration", config.math_config, load_math_config),
        ("Statistics configuration", config.statistics_config, load_statistics_config),
        ("Plotting configuration", config.plotting_config, load_plotting_config),
        ("Excel statistics configuration", config.excel_statistics_config, load_statistics_config),
        ("Excel report configuration", config.excel_report_config, load_excel_report_config),
    )
    entries = list(validators)
    if config.powerpoint_report_config is not None:
        entries.append(("PowerPoint report configuration", config.powerpoint_report_config, load_powerpoint_report_config))

    checks = [DoctorCheck("Default pipeline configuration", "PASS", str(config_path))]
    for name, path, loader in entries:
        try:
            loader(path)
        except VSMPostProcessingError as exc:
            checks.append(DoctorCheck(name, "FAIL", str(exc)))
        else:
            checks.append(DoctorCheck(name, "PASS", str(path)))
    return checks
