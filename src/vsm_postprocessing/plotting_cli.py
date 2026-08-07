from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .errors import VSMPostProcessingError
from .importer import ImportOptions
from .plotting_engine import render_plots


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="vsm-plot",
        description="Render deterministic configured PNG plots from raw and math VSM channels.",
    )
    parser.add_argument("input_file", help="Path to the source .xlsx or .csv file")
    parser.add_argument("--config", required=True, help="Path to the versioned YAML plotting file")
    parser.add_argument("--math-config", help="Optional math-channel YAML file used before plotting")
    parser.add_argument("--output-dir", default="outputs/plots", help="Output directory")
    parser.add_argument("--sheet", dest="sheet_name", help="Excel sheet name override")
    parser.add_argument("--header-row", type=int, help="1-based header row override")
    parser.add_argument("--unit-row", type=int, help="1-based unit row override")
    parser.add_argument("--data-start-row", type=int, help="1-based first data row override")
    parser.add_argument("--data-end-row", type=int, help="1-based final data row override")
    parser.add_argument("--last-channel-column", type=int, help="1-based final channel column override")
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
        result = render_plots(
            args.input_file,
            args.config,
            args.output_dir,
            options,
            math_config_file=args.math_config,
        )
    except VSMPostProcessingError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2

    print("VSM plotting completed")
    print("  status: PASS")
    print(f"  samples: {result.sample_count}")
    print(f"  available channels: {len(result.channels_by_id)}")
    print(f"  plots rendered: {result.plot_count}")
    print(f"  series rendered: {result.series_count}")
    print("  plots:")
    for index, item in enumerate(result.rendered_plots, start=1):
        reference = f" | reference chart {item.reference_chart_number}" if item.reference_chart_number else ""
        print(f"    {index:02d}. {item.plot_id} | {Path(item.output_file).name}{reference}")
    print(f"  output directory: {Path(args.output_dir).expanduser().resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
