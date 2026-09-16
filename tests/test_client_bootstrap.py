from __future__ import annotations

import os
import shutil
import subprocess
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
POWERSHELL = Path(os.environ.get("SystemRoot", r"C:\Windows")) / "System32" / "WindowsPowerShell" / "v1.0" / "powershell.exe"


def _read(relative: str) -> str:
    return (PROJECT_ROOT / relative).read_text(encoding="utf-8")


def _write_fake_python(path: Path, version: str, architecture: str = "AMD64") -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        f"@echo off\r\necho {version}^|{architecture}^|%~f0\r\n",
        encoding="utf-8",
    )
    return path


def _probe_candidate(tmp_path: Path, candidate: Path) -> str:
    if not POWERSHELL.exists():
        return "SKIPPED"
    project = tmp_path / "Tool"
    project.mkdir(parents=True, exist_ok=True)
    completed = subprocess.run(
        [
            str(POWERSHELL),
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(PROJECT_ROOT / "scripts" / "client_setup.ps1"),
            "-ProjectRoot",
            str(project),
            "-CandidateProbeExecutable",
            str(candidate),
            "-NoPause",
        ],
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        check=False,
    )
    assert completed.returncode == 0
    return completed.stdout


def test_client_launchers_delegate_to_package_local_scripts() -> None:
    setup = _read("SETUP_VSM_TOOL.bat")
    start = _read("START_VSM_TOOL.bat")

    assert "client_setup.ps1" in setup
    assert "client_start.ps1" in start
    assert "%~dp0" in setup
    assert "%~dp0" in start
    assert 'set "TOOL_ROOT=%PACKAGE_ROOT%Tool"' in setup
    assert 'set "TOOL_ROOT=%PACKAGE_ROOT%Tool"' in start
    assert "%TOOL_ROOT%\\scripts\\client_setup.ps1" in setup
    assert "%TOOL_ROOT%\\scripts\\client_start.ps1" in start
    assert r"C:\Users\user\Desktop\Agro Project" not in setup
    assert r"C:\Users\user\Desktop\Agro Project" not in start


def test_client_start_uses_package_local_runtime_only() -> None:
    script = _read("scripts/client_start.ps1")

    assert '.venv\\Scripts\\python.exe"' in script
    assert "-m streamlit run" in script
    assert "Please double-click SETUP_VSM_TOOL.bat first." in script
    assert "Get-Command uv" not in script
    assert "Get-Command streamlit" not in script
    assert r"C:\Users\user\Desktop\Agro Project" not in script


def test_client_setup_bootstrap_contract_is_client_safe() -> None:
    script = _read("scripts/client_setup.ps1")

    assert "https://www.python.org/ftp/python/" in script
    assert "https://www.nuget.org/api/v2/package/python/" in script
    assert 'BootstrapPythonVersion = "3.11.9"' in script
    assert 'ValidatedPythonRange = ">=3.11,<3.12"' in script
    assert 'ValidatedPythonMinors = @("3.11")' in script
    assert "python-$BootstrapPythonVersion-amd64.exe" in script
    assert "python-$BootstrapPythonVersion-amd64.nupkg" in script
    assert "Configured Python bootstrap $Description was not found at the expected official URL" in script
    assert "Downloaded $Description is unexpectedly small" in script
    assert "does not have the expected file signature" in script
    assert "could not write the Python runtime download" in script
    assert "Falling back to official CPython NuGet runtime package" in script
    assert "InstallAllUsers=0" in script
    assert "Include_pip=1" in script
    assert "PrependPath=0" in script
    assert "WindowsApps\\python*.exe" in script
    assert "Read-Host \"Press Y to continue or N to cancel\"" in script
    assert '"-m", "venv", $VenvDir' in script
    assert '"-m", "pip", "install", "-e", $ProjectRoot' in script
    assert "vsm_postprocessing.doctor_cli" in script
    assert "logs" in script
    assert "Get-Command uv" not in script
    assert r"C:\Users\user\Desktop\Agro Project" not in script


def test_sergio_package_builder_validates_real_bootstrap_installer_url() -> None:
    builder = _read("outputs/client_delivery/build_sergio_bootstrap_package.py")

    assert 'BOOTSTRAP_PYTHON_VERSION = "3.11.9"' in builder
    assert "https://www.python.org/ftp/python/" in builder
    assert "https://www.nuget.org/api/v2/package/python/" in builder
    assert "python-{BOOTSTRAP_PYTHON_VERSION}-amd64.exe" in builder
    assert "validate_bootstrap_installer_url()" in builder
    assert "validate_bootstrap_nuget_url()" in builder
    assert "Configured Python bootstrap {label} was not found" in builder
    assert 'expected_prefix=b"MZ"' in builder
    assert 'expected_prefix=b"PK"' in builder


def test_client_readme_agrees_with_bootstrap_target() -> None:
    builder = _read("outputs/client_delivery/build_sergio_bootstrap_package.py")

    assert "Python 3.11.x runtime" in builder
    assert "CPython 3.11.9" in builder
    assert "64-bit runtime" in builder
    assert "Python 3.11.15" not in builder


def test_runtime_dependency_bounds_avoid_unvalidated_streamlit_packages() -> None:
    pyproject = _read("pyproject.toml")

    assert '"streamlit>=1.37,<1.50"' in pyproject
    assert '"streamlit>=1.37,<2"' not in pyproject


def test_client_setup_accepts_python_311_candidate(tmp_path: Path) -> None:
    output = _probe_candidate(tmp_path, _write_fake_python(tmp_path / "Python311" / "python.cmd", "3.11.9"))
    if output == "SKIPPED":
        return
    assert "VALIDATED_PYTHON 3.11.9" in output


def test_client_setup_rejects_python_310_candidate(tmp_path: Path) -> None:
    output = _probe_candidate(tmp_path, _write_fake_python(tmp_path / "Python310" / "python.cmd", "3.10.20"))
    if output == "SKIPPED":
        return
    assert "REJECTED_PYTHON" in output
    assert "VALIDATED_PYTHON" not in output


def test_client_setup_rejects_future_python_candidate(tmp_path: Path) -> None:
    output = _probe_candidate(tmp_path, _write_fake_python(tmp_path / "Python314" / "python.cmd", "3.14.7"))
    if output == "SKIPPED":
        return
    assert "REJECTED_PYTHON" in output
    assert "VALIDATED_PYTHON" not in output


def test_client_setup_ignores_windows_store_alias_path(tmp_path: Path) -> None:
    alias = tmp_path / "WindowsApps" / "python.exe"
    alias.parent.mkdir(parents=True)
    alias.write_text("not a real interpreter", encoding="utf-8")
    output = _probe_candidate(tmp_path, alias)
    if output == "SKIPPED":
        return
    assert "REJECTED_PYTHON" in output
    assert "VALIDATED_PYTHON" not in output


def test_client_setup_accepts_package_local_python_311_runtime(tmp_path: Path) -> None:
    if not POWERSHELL.exists():
        return
    project = tmp_path / "Tool With Spaces"
    scripts = project / "scripts"
    scripts.mkdir(parents=True)
    shutil.copy2(PROJECT_ROOT / "scripts" / "client_setup.ps1", scripts / "client_setup.ps1")
    subprocess.run(
        [sys.executable, "-m", "venv", str(project / ".venv")],
        check=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
    )
    (project / "pyproject.toml").write_text("[project]\nname='probe'\nversion='0'\n", encoding="utf-8")

    completed = subprocess.run(
        [
            str(POWERSHELL),
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(scripts / "client_setup.ps1"),
            "-ProjectRoot",
            str(project),
            "-ProbePythonOnly",
            "-NoPause",
        ],
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        check=False,
    )
    assert completed.returncode == 0
    assert "FOUND_PYTHON 3.11" in completed.stdout
    assert str(project / ".venv" / "Scripts" / "python.exe") in completed.stdout


def test_client_setup_probe_modes_do_not_modify_host_python(tmp_path: Path) -> None:
    if not POWERSHELL.exists():
        return

    project = tmp_path / "Tool With Spaces"
    scripts = project / "scripts"
    scripts.mkdir(parents=True)
    shutil.copy2(PROJECT_ROOT / "scripts" / "client_setup.ps1", scripts / "client_setup.ps1")
    (project / "pyproject.toml").write_text("[project]\nname='probe'\nversion='0'\n", encoding="utf-8")

    missing = subprocess.run(
        [
            str(POWERSHELL),
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(scripts / "client_setup.ps1"),
            "-ProjectRoot",
            str(project),
            "-ProbePythonOnly",
            "-SimulateMissingPython",
            "-NoPause",
        ],
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        check=False,
    )
    assert missing.returncode == 0
    assert "A compatible VSM runtime was not found." in missing.stdout
    assert "Your existing Python installations will not be modified." in missing.stdout
    assert "MISSING_PYTHON_BOOTSTRAP_AVAILABLE" in missing.stdout
    assert "Python was not found" not in missing.stdout
    assert "Microsoft Store" not in missing.stdout

    existing_env = os.environ.copy()
    existing_env["VSM_PYTHON_EXE"] = sys.executable
    existing = subprocess.run(
        [
            str(POWERSHELL),
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(scripts / "client_setup.ps1"),
            "-ProjectRoot",
            str(project),
            "-ProbePythonOnly",
            "-NoPause",
        ],
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        env=existing_env,
        check=False,
    )
    assert existing.returncode == 0
    assert "FOUND_PYTHON" in existing.stdout
