from __future__ import annotations

import argparse
from pathlib import Path

from .doctor import run_doctor


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="vsm-doctor",
        description="Check the local VSM post-processing installation before client use.",
    )
    parser.add_argument(
        "--project-root",
        default=".",
        help="Project root containing config/, scripts/, src/ and outputs/",
    )
    parser.add_argument(
        "--pipeline-config",
        default=None,
        help="Optional end-to-end YAML configuration to validate",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    report = run_doctor(Path(args.project_root), args.pipeline_config)
    print("VSM client-readiness check")
    print(f"  status: {report.status}")
    for check in report.checks:
        print(f"  [{check.status}] {check.name}: {check.detail}")
    print(f"  warnings: {report.warning_count}")
    print(f"  failures: {report.error_count}")
    return 0 if report.status == "PASS" else 2


if __name__ == "__main__":
    raise SystemExit(main())
