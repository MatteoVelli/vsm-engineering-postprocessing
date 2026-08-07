from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .errors import VSMPostProcessingError
from .importer import ImportOptions
from .math_engine import calculate_math_channels, export_math_channels


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="vsm-math",
        description="Calculate deterministic, configured VSM math channels and export verification evidence.",
    )
    parser.add_argument("input_file", help="Path to the source .xlsx or .csv file")
    parser.add_argument("--config", required=True, help="Path to the versioned YAML math-channel file")
    parser.add_argument("--output-dir", default="outputs/math_channels", help="Output directory")
    parser.add_argument("--sheet", dest="sheet_name", help="Excel sheet name override")
    parser.add_argument("--header-row", type=int, help="1-based header row override")
    parser.add_argument("--unit-row", type=int, help="1-based unit row override")
    parser.add_argument("--data-start-row", type=int, help="1-based first data row override")
    parser.add_argument("--time-channel", help="Exact source name of the time channel")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    options = ImportOptions(
        sheet_name=args.sheet_name,
        header_row=args.header_row,
        unit_row=args.unit_row,
        data_start_row=args.data_start_row,
        time_channel=args.time_channel,
        strict=True,
    )

    try:
        result = calculate_math_channels(args.input_file, args.config, options)
        outputs = export_math_channels(result, args.output_dir)
    except VSMPostProcessingError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2

    print("VSM math-channel calculation completed")
    print("  status: PASS")
    print(f"  samples: {result.sample_count}")
    print(f"  source channels exported: {result.source_channel_count}")
    print(f"  math channels calculated: {result.math_channel_count}")
    print(f"  output channels: {result.output_channel_count}")
    print("  calculation order:")
    for index, channel_id in enumerate(result.calculation_order, start=1):
        print(f"    {index:02d}. {channel_id}")
    print("  comparisons:")
    if result.comparisons:
        for comparison in result.comparisons:
            print(
                f"    {'PASS' if comparison.passed else 'FAIL'} | "
                f"{comparison.math_channel_id} vs {comparison.reference_channel_id} | "
                f"max abs error {comparison.max_absolute_error:.12g}"
            )
    else:
        print("    none configured")
    print("  outputs:")
    for name, path in outputs.items():
        print(f"    {name}: {Path(path)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
