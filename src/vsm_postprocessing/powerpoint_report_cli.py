from __future__ import annotations

import argparse
import sys

from .errors import VSMPostProcessingError
from .importer import ImportOptions
from .powerpoint_report_engine import generate_powerpoint_report


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="vsm-pptx",
        description="Generate a deterministic PowerPoint engineering report from selected VSM statistics and plots.",
    )
    parser.add_argument("input_file", help="Path to the source .xlsx or .csv file")
    parser.add_argument("--config", required=True, help="PowerPoint-report YAML configuration")
    parser.add_argument("--math-config", help="Optional math-channel YAML configuration")
    parser.add_argument("--statistics-config", required=True, help="Statistics YAML configuration")
    parser.add_argument("--plotting-config", required=True, help="Plotting YAML configuration")
    parser.add_argument("--output-dir", default="outputs/powerpoint_report", help="Output directory")
    parser.add_argument("--sheet", dest="sheet_name", help="Excel source sheet override")
    parser.add_argument("--header-row", type=int, help="1-based header row override")
    parser.add_argument("--unit-row", type=int, help="1-based unit row override")
    parser.add_argument("--data-start-row", type=int, help="1-based first data row override")
    parser.add_argument("--data-end-row", type=int, help="1-based final data row override")
    parser.add_argument("--last-channel-column", type=int, help="1-based final source-channel column override")
    parser.add_argument("--time-channel", help="Exact source name of the time channel")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    options = ImportOptions(
        sheet_name=args.sheet_name,
        header_row=args.header_row,
        unit_row=args.unit_row,
        data_start_row=args.data_start_row,
        data_end_row=args.data_end_row,
        last_channel_column=args.last_channel_column,
        time_channel=args.time_channel,
        strict=True,
    )
    try:
        result = generate_powerpoint_report(
            args.input_file,
            args.config,
            args.statistics_config,
            args.plotting_config,
            args.output_dir,
            options,
            math_config_file=args.math_config,
        )
    except VSMPostProcessingError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2

    print("VSM PowerPoint report completed")
    print("  status: PASS")
    print(f"  samples: {result.sample_count}")
    print(f"  slides: {result.slide_count}")
    print(f"  statistics used: {result.statistic_count}")
    print(f"  plots used: {result.plot_count}")
    print(f"  presentation: {result.presentation_path}")
    print(f"  manifest: {result.manifest_path}")
    print(f"  summary: {result.summary_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
