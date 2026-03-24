#!/usr/bin/env python3

from __future__ import annotations

import argparse
import re
import shutil
import sys
import unicodedata
from pathlib import Path


FILENAME_PATTERN = re.compile(
    r"^(?P<region>.+)_(?P<ownership>ГОС|НЕГОС)_(?P<format>.+)\.xlsx$"
)

OWNERSHIP_MAP = {
    "гос": "governmental",
    "негос": "commercial",
}

FORMAT_MAP = {
    "очная": "full_time",
    "заочная": "part_time",
    "очнозаочная": "hybrid",
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Restructure workbook files into format/ownership/region directories."
        )
    )
    parser.add_argument(
        "--source-dir",
        default="input",
        help="Directory with source .xlsx files. Default: %(default)s",
    )
    parser.add_argument(
        "--target-dir",
        default="001_output",
        help="Directory for the refactored tree. Default: %(default)s",
    )
    parser.add_argument(
        "--move",
        action="store_true",
        help="Move files instead of copying them.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Show planned destination paths without writing files.",
    )
    return parser.parse_args()


def normalize_token(value: str) -> str:
    normalized = unicodedata.normalize("NFKC", value).strip().lower()
    return re.sub(r"[\s_\-–—]+", "", normalized)


def normalize_region(value: str) -> str:
    return unicodedata.normalize("NFC", value).strip().lower()


def should_skip(path: Path) -> bool:
    return path.name.startswith("~$") or path.name.startswith(".")


def iter_source_files(source_dir: Path) -> list[Path]:
    return sorted(
        path
        for path in source_dir.rglob("*.xlsx")
        if path.is_file() and not should_skip(path)
    )


def build_destination(path: Path, target_dir: Path) -> Path:
    match = FILENAME_PATTERN.match(path.name)
    if not match:
        raise ValueError(f"Unrecognized filename: {path}")

    region = normalize_region(match.group("region"))
    ownership_key = normalize_token(match.group("ownership"))
    format_key = normalize_token(match.group("format"))

    try:
        ownership_dir = OWNERSHIP_MAP[ownership_key]
    except KeyError as exc:
        raise ValueError(f"Unknown ownership token in {path.name}") from exc

    try:
        format_dir = FORMAT_MAP[format_key]
    except KeyError as exc:
        raise ValueError(f"Unknown format token in {path.name}") from exc

    return target_dir / format_dir / ownership_dir / region / path.name


def main() -> int:
    args = parse_args()
    source_dir = Path(args.source_dir).expanduser().resolve()
    target_dir = Path(args.target_dir).expanduser().resolve()

    if not source_dir.is_dir():
        print(f"Source directory does not exist: {source_dir}", file=sys.stderr)
        return 1

    source_files = iter_source_files(source_dir)
    if not source_files:
        print(f"No .xlsx files found in {source_dir}", file=sys.stderr)
        return 1

    planned_moves: list[tuple[Path, Path]] = []
    for source_path in source_files:
        destination_path = build_destination(source_path, target_dir)
        planned_moves.append((source_path, destination_path))

    seen_destinations: set[Path] = set()
    for source_path, destination_path in planned_moves:
        if destination_path in seen_destinations:
            print(
                f"Duplicate destination generated for {source_path}: {destination_path}",
                file=sys.stderr,
            )
            return 1
        seen_destinations.add(destination_path)

        if destination_path.exists():
            print(
                f"Destination already exists: {destination_path}",
                file=sys.stderr,
            )
            return 1

    if args.dry_run:
        for source_path, destination_path in planned_moves:
            print(f"{source_path} -> {destination_path}")
        print(f"Planned {len(planned_moves)} file(s)")
        return 0

    for source_path, destination_path in planned_moves:
        destination_path.parent.mkdir(parents=True, exist_ok=True)
        if args.move:
            shutil.move(str(source_path), str(destination_path))
        else:
            shutil.copy2(source_path, destination_path)

    action = "Moved" if args.move else "Copied"
    print(f"{action} {len(planned_moves)} file(s) into {target_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
