"""Exercise runner isolation and destructive-operation boundaries with tiny probes."""
from __future__ import annotations

import importlib.util
import json
import os
from pathlib import Path
import subprocess
import sys
import xml.etree.ElementTree as ET

import pytest


spec = importlib.util.spec_from_file_location(
    "vsm_maintenance_runner", Path(__file__).resolve().parents[1] / "maintenance/run_tests.py")
runner = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = runner
spec.loader.exec_module(runner)


@pytest.fixture
def project(tmp_path, monkeypatch):
    root = tmp_path / "tiny project"
    (root / "tests").mkdir(parents=True)
    (root / "pyproject.toml").write_text('[tool.pytest.ini_options]\naddopts = "-q"\n')
    (root / "tests/test_probe.py").write_text('''import os
from pathlib import Path
def test_probe(tmp_path):
    assert Path(os.environ['TEMP']).parent == tmp_path.parents[1]
    assert os.environ['PYTHONDONTWRITEBYTECODE'] == '1'
    assert os.environ['PYTEST_ADDOPTS'] == ''
    assert 'PYTEST_PLUGINS' not in os.environ
    (tmp_path / 'diagnostic.txt').write_text('probe evidence')
    assert os.environ.get('VSM_PROBE_FAIL') != '1', 'expected synthetic failure'
''')
    parent = tmp_path / "external scratch"
    parent.mkdir()
    historical = parent / "vsm-test-pre-existing"
    historical.mkdir()
    (historical / "sentinel.txt").write_text("preserve")
    monkeypatch.setattr(runner, "PROJECT_ROOT", root)
    monkeypatch.delenv("VSM_PROBE_FAIL", raising=False)
    return root, parent


def invoke(project, *options):
    _, parent = project
    return runner.main(["--temp-root", str(parent), *options, "tests/test_probe.py"])


def owned_runs(parent):
    return [p for p in parent.iterdir() if p.name != "vsm-test-pre-existing"]


def test_success_cleans_only_new_workspace(project, monkeypatch):
    root, parent = project
    monkeypatch.setenv("PYTEST_ADDOPTS", "--basetemp=must-not-be-used")
    monkeypatch.setenv("PYTEST_PLUGINS", "must_not_be_imported")
    before = {str(p.relative_to(root)) for p in root.rglob("*")}
    assert invoke(project) == 0
    assert not owned_runs(parent)
    assert (parent / "vsm-test-pre-existing/sentinel.txt").read_text() == "preserve"
    assert {str(p.relative_to(root)) for p in root.rglob("*")} == before


def test_expected_failure_retains_log_junit_and_artifacts(project, monkeypatch):
    _, parent = project
    monkeypatch.setenv("VSM_PROBE_FAIL", "1")
    assert invoke(project) == 1
    workspace, = owned_runs(parent)
    assert "expected synthetic failure" in (workspace / "pytest.log").read_text()
    assert int(ET.parse(workspace / "junit.xml").getroot()[0].get("failures")) == 1
    assert list((workspace / "pytest").rglob("diagnostic.txt"))
    assert json.loads((workspace / "run.json").read_text())["cwd"] == str(project[0])
    assert (parent / "vsm-test-pre-existing/sentinel.txt").read_text() == "preserve"


def test_keep_success_retains_review_evidence(project):
    assert invoke(project, "--keep-success") == 0
    workspace, = owned_runs(project[1])
    assert (workspace / "junit.xml").is_file()
    assert "1 passed" in (workspace / "pytest.log").read_text()


def test_dry_run_creates_no_workspace_or_process(project, monkeypatch, capsys):
    def unexpected(*args, **kwargs):
        pytest.fail("Dry run started a process")
    monkeypatch.setattr(runner.subprocess, "Popen", unexpected)
    assert invoke(project, "--dry-run") == 0
    plan = json.loads(capsys.readouterr().out)
    assert plan["temporary_parent"] == str(project[1])
    assert not owned_runs(project[1])


def test_environment_configures_temporary_parent(project, monkeypatch):
    monkeypatch.setenv("VSM_TEST_TEMP_ROOT", str(project[1]))
    assert runner.main(["--dry-run", "tests/test_probe.py"]) == 0
    assert not owned_runs(project[1])


@pytest.mark.parametrize("selection", ["pyproject.toml", "../outside.py"])
def test_invalid_selection_never_creates_workspace(project, selection):
    with pytest.raises(SystemExit) as exc:
        runner.main(["--temp-root", str(project[1]), selection])
    assert exc.value.code == 2
    assert not owned_runs(project[1])


def test_repository_temporary_parent_is_refused(project):
    with pytest.raises(SystemExit) as exc:
        runner.main(["--temp-root", str(project[0]), "tests/test_probe.py"])
    assert exc.value.code == 2
    assert not owned_runs(project[1])


def test_user_cannot_override_basetemp_or_start_default_full_suite(project):
    for arguments in ([], ["--basetemp", str(project[1]), "tests/test_probe.py"]):
        with pytest.raises(SystemExit) as exc:
            runner.main(arguments)
        assert exc.value.code == 2
    assert not owned_runs(project[1])


def test_cleanup_refuses_changed_owner(project):
    workspace = runner.create_workspace(project[1])
    (workspace.path / runner.MARKER).write_text("different owner")
    with pytest.raises(ValueError, match="ownership"):
        runner.remove_owned_workspace(workspace)
    assert workspace.path.is_dir()


def test_cleanup_refuses_replaced_directory_even_with_matching_marker(project):
    workspace = runner.create_workspace(project[1])
    retained = workspace.path.with_name("original-workspace")
    workspace.path.rename(retained)
    workspace.path.mkdir()
    (workspace.path / runner.MARKER).write_text(workspace.token)
    with pytest.raises(ValueError, match="identity"):
        runner.remove_owned_workspace(workspace)
    assert workspace.path.is_dir() and retained.is_dir()


def test_cleanup_refuses_changed_boundary(project):
    workspace = runner.create_workspace(project[1])
    forged = runner.Workspace(workspace.path, project[0], workspace.identity, workspace.token)
    with pytest.raises(ValueError, match="boundary"):
        runner.remove_owned_workspace(forged)
    assert workspace.path.is_dir()


def test_cleanup_refuses_junction_or_symlink_and_preserves_target(project):
    workspace = runner.create_workspace(project[1])
    target = project[1] / "vsm-test-pre-existing"
    link = workspace.path / "external-link"
    if os.name == "nt":
        subprocess.run(["cmd", "/c", "mklink", "/J", str(link), str(target)],
                       check=True, capture_output=True)
    else:
        link.symlink_to(target, target_is_directory=True)
    try:
        with pytest.raises(ValueError, match="reparse|Link"):
            runner.remove_owned_workspace(workspace)
        assert (target / "sentinel.txt").read_text() == "preserve"
    finally:
        if os.name == "nt":
            link.rmdir()
        else:
            link.unlink()


@pytest.mark.parametrize("error,expected", [(OSError("synthetic start error"), 2), (KeyboardInterrupt(), 130)])
def test_start_error_or_interruption_preserves_workspace(project, monkeypatch, error, expected):
    def interrupted(*args):
        raise error
    monkeypatch.setattr(runner, "run_pytest", interrupted)
    assert invoke(project) == expected
    workspace, = owned_runs(project[1])
    assert (workspace / runner.MARKER).is_file()
