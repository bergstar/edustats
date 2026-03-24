#!/usr/bin/env python3

from __future__ import annotations

import os
import shutil
import sys
import time
from concurrent.futures import ProcessPoolExecutor, as_completed
from pathlib import Path

from openpyxl import load_workbook


SOURCE_DIR_NAME = "001_output"
TARGET_DIR_NAME = "002_output"
MAX_WORKERS = 10
BAR_WIDTH = 32


def should_skip(path: Path) -> bool:
    return path.name.startswith("~$") or path.name.startswith(".")


def iter_workbooks(source_dir: Path) -> list[Path]:
    return sorted(
        path
        for path in source_dir.rglob("*.xlsx")
        if path.is_file() and not should_skip(path)
    )


def recreate_target_dir(target_dir: Path) -> None:
    if target_dir.exists():
        shutil.rmtree(target_dir)
    target_dir.mkdir(parents=True, exist_ok=True)


def collect_sheet_names(workbook_path: Path) -> list[str]:
    workbook = load_workbook(workbook_path, read_only=True, data_only=False)
    try:
        return list(workbook.sheetnames)
    finally:
        workbook.close()


def export_workbook(workbook_path_str: str, destination_dir_str: str) -> int:
    workbook_path = Path(workbook_path_str)
    destination_dir = Path(destination_dir_str)
    sheet_names = collect_sheet_names(workbook_path)
    exported_sheets = 0

    for sheet_name in sheet_names:
        workbook = load_workbook(workbook_path)
        try:
            for worksheet in list(workbook.worksheets):
                if worksheet.title != sheet_name:
                    workbook.remove(worksheet)

            workbook.active = 0
            destination_path = destination_dir / f"{sheet_name}.xlsx"
            destination_path.parent.mkdir(parents=True, exist_ok=True)
            workbook.save(destination_path)
            exported_sheets += 1
        finally:
            workbook.close()

    return exported_sheets


def format_duration(seconds: float) -> str:
    total_seconds = int(seconds)
    minutes, seconds = divmod(total_seconds, 60)
    hours, minutes = divmod(minutes, 60)
    return f"{hours:02d}:{minutes:02d}:{seconds:02d}"


def print_progress(
    completed_books: int,
    total_books: int,
    exported_sheets: int,
    started_at: float,
) -> None:
    ratio = completed_books / total_books if total_books else 1.0
    filled = int(BAR_WIDTH * ratio)
    bar = "#" * filled + "-" * (BAR_WIDTH - filled)
    elapsed = time.monotonic() - started_at
    speed = exported_sheets / elapsed if elapsed > 0 else 0.0

    print(
        (
            f"\r[{bar}] {completed_books}/{total_books} books"
            f" | {exported_sheets} sheets"
            f" | {speed:6.1f} sheets/s"
            f" | {format_duration(elapsed)}"
        ),
        end="",
        flush=True,
    )


def main() -> int:
    base_dir = Path(__file__).resolve().parent
    source_dir = (base_dir / SOURCE_DIR_NAME).resolve()
    target_dir = (base_dir / TARGET_DIR_NAME).resolve()

    if not source_dir.is_dir():
        print(f"Source directory does not exist: {source_dir}", file=sys.stderr)
        return 1

    workbook_paths = iter_workbooks(source_dir)
    if not workbook_paths:
        print(f"No .xlsx files found in {source_dir}", file=sys.stderr)
        return 1

    recreate_target_dir(target_dir)

    worker_count = min(MAX_WORKERS, len(workbook_paths), os.cpu_count() or 1)
    print(
        f"Exporting {len(workbook_paths)} workbooks with {worker_count} workers...",
        flush=True,
    )

    started_at = time.monotonic()
    completed_books = 0
    exported_sheets = 0
    print_progress(completed_books, len(workbook_paths), exported_sheets, started_at)

    with ProcessPoolExecutor(max_workers=worker_count) as executor:
        future_to_workbook = {}
        for workbook_path in workbook_paths:
            relative_region_dir = workbook_path.parent.relative_to(source_dir)
            destination_dir = target_dir / relative_region_dir
            future = executor.submit(
                export_workbook,
                str(workbook_path),
                str(destination_dir),
            )
            future_to_workbook[future] = workbook_path

        for future in as_completed(future_to_workbook):
            workbook_path = future_to_workbook[future]
            try:
                sheet_count = future.result()
            except Exception as exc:
                print()
                print(f"Failed while exporting {workbook_path}: {exc}", file=sys.stderr)
                return 1

            completed_books += 1
            exported_sheets += sheet_count
            print_progress(
                completed_books,
                len(workbook_paths),
                exported_sheets,
                started_at,
            )

    print()
    print(
        f"Exported {exported_sheets} single-sheet file(s) from "
        f"{len(workbook_paths)} workbook(s) into {target_dir}",
        flush=True,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
