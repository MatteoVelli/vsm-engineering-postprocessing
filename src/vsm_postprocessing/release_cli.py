"""CLI for deterministic client-release packaging."""

from __future__ import annotations

import argparse
from pathlib import Path

from .release_builder import build_client_release


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Build the deterministic VSM client release ZIP.")
    parser.add_argument("--project-root", default=".", help="Project root directory.")
    parser.add_argument("--output-dir", default="dist", help="Release output directory.")
    return parser


def main() -> int:
    args = build_parser().parse_args()
    result = build_client_release(Path(args.project_root), Path(args.output_dir))
    print("VSM client release completed")
    print("  status: PASS")
    print(f"  files packaged: {result.file_count}")
    print(f"  archive: {result.archive_path}")
    print(f"  sha256: {result.archive_sha256}")
    print(f"  checksum: {result.checksum_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
