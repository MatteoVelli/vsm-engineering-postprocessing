from __future__ import annotations

import argparse
import sys

from .errors import VSMPostProcessingError
from .pipeline_engine import run_pipeline


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="vsm-run",
        description="Run the complete deterministic VSM post-processing pipeline from source data to Excel and PowerPoint reports.",
    )
    parser.add_argument(
        "config_file",
        nargs="?",
        default="config/end_to_end_example.yaml",
        help="Path to the end-to-end pipeline YAML configuration",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        result = run_pipeline(args.config_file)
    except VSMPostProcessingError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2

    print("VSM end-to-end pipeline completed")
    print(f"  status: {result.status}")
    print(f"  stages passed: {result.completed_stage_count}/{len(result.stages)}")
    for stage in result.stages:
        metrics = ", ".join(f"{key}={value}" for key, value in stage.metrics.items())
        print(f"  {stage.name}: {stage.status}" + (f" | {metrics}" if metrics else ""))
    if result.report_path is not None:
        print(f"  final Excel report: {result.report_path}")
    if result.powerpoint_path is not None:
        print(f"  final PowerPoint report: {result.powerpoint_path}")
    print(f"  manifest: {result.manifest_path}")
    print(f"  summary: {result.summary_path}")
    print(f"  log: {result.log_path}")
    print(f"  duration: {result.duration_seconds:.3f} s")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
