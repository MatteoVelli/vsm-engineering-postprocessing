from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .errors import VSMPostProcessingError
from .excel_report_engine import generate_excel_report
from .importer import ImportOptions


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="vsm-excel",
        description="Generate a deterministic engineering Excel report from configured VSM processing results.",
    )
    parser.add_argument("input_file", help="Path to the source .xlsx or .csv file")
    parser.add_argument("--config", required=True, help="Excel-report YAML configuration")
    parser.add_argument("--math-config", help="Optional math-channel YAML configuration")
    parser.add_argument("--statistics-config", required=True, help="Statistics YAML configuration")
    parser.add_argument("--plotting-config", required=True, help="Plotting YAML configuration")
    parser.add_argument("--output-dir", default="outputs/excel_report", help="Output directory")
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
        result = generate_excel_report(
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

    print("VSM Excel report completed")
    print("  status: PASS")
    print(f"  samples: {result.sample_count}")
    print(f"  report channels: {result.channel_count}")
    print(f"  statistics available: {result.statistic_count}")
    print(f"  plots embedded: {result.plot_count}")
    print(f"  layout profile: {result.config.layout_profile}")
    print(f"  plot placement: {result.config.plot_placement}")
    print(f"  workbook: {result.report_path}")
    print(f"  manifest: {result.manifest_path}")
    print(f"  summary: {result.summary_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
