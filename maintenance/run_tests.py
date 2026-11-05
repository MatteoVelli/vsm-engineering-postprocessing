"""Opt-in pytest runner; invoke with an already validated Python environment."""
from __future__ import annotations

import argparse
from dataclasses import dataclass
import json
import os
from pathlib import Path
import shutil
import stat
import subprocess
import sys
import tempfile
import uuid


PROJECT_ROOT = Path(__file__).resolve().parents[1]
MARKER = ".vsm-test-owner"


@dataclass(frozen=True)
class Workspace:
    path: Path
    parent: Path
    identity: tuple[int, int]
    token: str


def external_parent(path: str | Path) -> Path:
    parent = Path(path).expanduser().resolve(strict=True)
    if not parent.is_dir() or parent.is_relative_to(PROJECT_ROOT):
        raise ValueError("Temporary parent must be an existing directory outside the repository")
    return parent


def create_workspace(parent: Path) -> Workspace:
    parent = external_parent(parent)
    path = Path(tempfile.mkdtemp(prefix="vsm-test-", dir=parent))
    info = path.stat()
    token = uuid.uuid4().hex
    (path / MARKER).write_text(token, encoding="ascii")
    return Workspace(path, parent, (info.st_dev, info.st_ino), token)


def remove_owned_workspace(workspace: Workspace) -> None:
    """Remove only this invocation's unchanged workspace; fail closed on links."""
    path = workspace.path
    if (path.resolve(strict=True) != path or path.parent != workspace.parent
            or path.is_relative_to(PROJECT_ROOT) or not path.name.startswith("vsm-test-")):
        raise ValueError("Workspace boundary changed; cleanup refused")
    info = path.stat()
    if (info.st_dev, info.st_ino) != workspace.identity:
        raise ValueError("Workspace identity changed; cleanup refused")
    # Inspect without following links, including Windows junctions/reparse points.
    pending = [path]
    while pending:
        entry = pending.pop()
        info = entry.lstat()
        if (stat.S_ISLNK(info.st_mode)
                or getattr(info, "st_file_attributes", 0) & stat.FILE_ATTRIBUTE_REPARSE_POINT):
            raise ValueError(f"Link or reparse point found; cleanup refused: {entry}")
        if stat.S_ISDIR(info.st_mode):
            pending.extend(entry.iterdir())
    if (path / MARKER).read_text(encoding="ascii") != workspace.token:
        raise ValueError("Workspace ownership changed; cleanup refused")
    # No permission-changing retry handler and no cleanup of the parent or siblings.
    shutil.rmtree(path)


def selected_tests(selections: list[str]) -> list[str]:
    result = []
    for selection in selections:
        filename, separator, node = selection.partition("::")
        path = (PROJECT_ROOT / filename).resolve(strict=True)
        if not path.is_relative_to(PROJECT_ROOT / "tests") or path.suffix != ".py" or not path.is_file():
            raise ValueError("Select explicit Python test files or nodes inside tests/")
        result.append(str(path) + (separator + node if separator else ""))
    return result


def test_command(workspace: Path, selections: list[str]) -> list[str]:
    return [sys.executable, "-B", "-m", "pytest", "-c", str(PROJECT_ROOT / "pyproject.toml"),
            "-o", "addopts=", "-q", "-ra", "-p", "no:cacheprovider",
            "--basetemp", str(workspace / "pytest"),
            "--junitxml", str(workspace / "junit.xml"), *selections]


def run_pytest(workspace: Path, selections: list[str]) -> int:
    env = os.environ.copy()
    for name in ("tmp", "mpl", "cache"):
        (workspace / name).mkdir()
    env.update(PYTHONDONTWRITEBYTECODE="1", PYTEST_DISABLE_PLUGIN_AUTOLOAD="1",
               PYTEST_ADDOPTS="", PYTHONIOENCODING="utf-8", MPLBACKEND="Agg",
               MPLCONFIGDIR=str(workspace / "mpl"), XDG_CACHE_HOME=str(workspace / "cache"),
               TEMP=str(workspace / "tmp"), TMP=str(workspace / "tmp"), TMPDIR=str(workspace / "tmp"))
    env.pop("PYTEST_PLUGINS", None)
    env["PYTHONPATH"] = os.pathsep.join(filter(None, [str(PROJECT_ROOT / "src"), env.get("PYTHONPATH")]))
    command = test_command(workspace, selections)
    (workspace / "run.json").write_text(json.dumps({"command": command, "cwd": str(PROJECT_ROOT)}, indent=2), encoding="utf-8")
    with (workspace / "pytest.log").open("w", encoding="utf-8") as log:
        with subprocess.Popen(command, cwd=PROJECT_ROOT, env=env, stdout=subprocess.PIPE,
                              stderr=subprocess.STDOUT, text=True, encoding="utf-8", errors="replace") as process:
            for line in process.stdout:
                log.write(line)
                log.flush()
                print(line, end="", flush=True)
            return process.wait()


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--temp-root", default=os.environ.get("VSM_TEST_TEMP_ROOT"),
                        help="Existing external directory; defaults to the system temp directory")
    parser.add_argument("--dry-run", action="store_true", help="Show the plan without creating a workspace")
    parser.add_argument("--keep-success", action="store_true", help="Retain a successful run for review")
    parser.add_argument("tests", nargs="+", help="Explicit tests/*.py files or file::test nodes")
    args = parser.parse_args(argv)
    try:
        parent = external_parent(args.temp_root or tempfile.gettempdir())
        selections = selected_tests(args.tests)
    except (OSError, ValueError) as exc:
        parser.error(str(exc))
    if args.dry_run:
        print(json.dumps({"temporary_parent": str(parent), "command": test_command(parent / "vsm-test-<unique>", selections),
                          "cleanup": "Only the fresh owned workspace after pytest exits 0; failures retained"}, indent=2))
        return 0
    workspace = create_workspace(parent)
    print(f"Test workspace: {workspace.path}", flush=True)
    try:
        code = run_pytest(workspace.path, selections)
    except (OSError, KeyboardInterrupt) as exc:
        print(f"Run interrupted or could not start: {exc}. Evidence retained: {workspace.path}", file=sys.stderr)
        return 130 if isinstance(exc, KeyboardInterrupt) else 2
    if code != 0 or args.keep_success:
        print(f"pytest exit code {code}; workspace retained: {workspace.path}", flush=True)
        return code
    try:
        remove_owned_workspace(workspace)
    except (OSError, ValueError) as exc:
        print(f"Tests passed but cleanup failed: {exc}. Inspect: {workspace.path}", file=sys.stderr)
        return 1
    print(f"Tests passed; removed owned workspace: {workspace.path}", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
