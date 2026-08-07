from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .errors import VSMPostProcessingError
from .importer import ImportOptions, export_inspection, inspect_data_file


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="vsm-inspect",
        description="Inspect VSM-style XLSX/CSV data and export a deterministic channel catalogue.",
    )
    parser.add_argument("input_file", help="Path to the source .xlsx or .csv file")
    parser.add_argument("--output-dir", default="outputs", help="Directory for generated inspection files")
    parser.add_argument("--sheet", dest="sheet_name", help="Excel sheet name override")
    parser.add_argument("--header-row", type=int, help="1-based header row override")
    parser.add_argument("--unit-row", type=int, help="1-based unit row override")
    parser.add_argument("--data-start-row", type=int, help="1-based first data row override")
    parser.add_argument("--time-channel", help="Exact source name of the time channel")
    parser.add_argument(
        "--allow-invalid",
        action="store_true",
        help="Export findings even when validation errors are present",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    options = ImportOptions(
        sheet_name=args.sheet_name,
        header_row=args.header_row,
        unit_row=args.unit_row,
        data_start_row=args.data_start_row,
        time_channel=args.time_channel,
        strict=not args.allow_invalid,
    )

    try:
        result = inspect_data_file(args.input_file, options)
        outputs = export_inspection(result, args.output_dir)
    except VSMPostProcessingError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2

    report = result.quality
    print("VSM inspection completed")
    print(f"  status: {'PASS' if report.is_valid else 'FAIL'}")
    print(f"  samples: {report.sample_count}")
    print(f"  channels: {report.channel_count} ({report.raw_channel_count} raw, {report.math_channel_count} math)")
    print(f"  time: {report.time_channel_name}, {report.time_start} to {report.time_end} {report.time_unit or ''}")
    print(f"  nominal step: {report.nominal_time_step}")
    print("  outputs:")
    for name, path in outputs.items():
        print(f"    {name}: {Path(path)}")
    return 0 if report.is_valid else 1


if __name__ == "__main__":
    raise SystemExit(main())
