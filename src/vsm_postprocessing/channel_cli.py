from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .channel_manager import export_channel_selection, select_channels
from .errors import VSMPostProcessingError
from .importer import ImportOptions


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="vsm-select",
        description="Select VSM channels by stable channel_id and export a traceable reduced dataset.",
    )
    parser.add_argument("input_file", help="Path to the source .xlsx or .csv file")
    parser.add_argument("--config", required=True, help="Path to the versioned YAML channel-selection file")
    parser.add_argument("--output-dir", default="outputs/channel_selection", help="Output directory")
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
        result = select_channels(args.input_file, args.config, options)
        outputs = export_channel_selection(result, args.output_dir)
    except VSMPostProcessingError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2

    print("VSM channel selection completed")
    print("  status: PASS")
    print(f"  samples: {result.sample_count}")
    print(f"  selected channels: {result.channel_count}")
    print("  order:")
    for index, channel in enumerate(result.selected_channels, start=1):
        print(f"    {index:02d}. {channel.channel_id} | {channel.display_name} [{channel.unit or '-'}]")
    print("  outputs:")
    for name, path in outputs.items():
        print(f"    {name}: {Path(path)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
