#!/usr/bin/env python3

from __future__ import annotations

import shutil
import sys
from pathlib import Path

from openpyxl import load_workbook

SOURCE_DIR_NAME = "001_output"
TARGET_DIR_NAME = "002_output"


def should_skip(path: Path) -> bool:
    return path.name.startswith("~$") or path.name.startswith(".")


def iter_workbooks(source_dir: Path) -> list[Path]:
    return sorted(
        path
        for path in source_dir.rglob("*.xlsx")
        if path.is_file() and not should_skip(path)
    )


def collect_sheet_plans(
    source_dir: Path, target_dir: Path, workbook_paths: list[Path]
) -> list[tuple[Path, str, Path]]:
    plans: list[tuple[Path, str, Path]] = []

    for workbook_index, workbook_path in enumerate(workbook_paths, start=1):
        relative_region_dir = workbook_path.parent.relative_to(source_dir)
        print(
            f"[DEBUG] Reading workbook {workbook_index}/{len(workbook_paths)}: "
            f"{workbook_path}",
            flush=True,
        )

        workbook = load_workbook(workbook_path, read_only=True, data_only=False)
        try:
            sheet_names = list(workbook.sheetnames)
        finally:
            workbook.close()

        print(
            f"[DEBUG] Found {len(sheet_names)} sheet(s) in {workbook_path.name}",
            flush=True,
        )

        for sheet_name in sheet_names:
            destination_path = target_dir / relative_region_dir / f"{sheet_name}.xlsx"
            plans.append((workbook_path, sheet_name, destination_path))

    return plans


def recreate_target_dir(target_dir: Path) -> None:
    if target_dir.exists():
        print(f"[DEBUG] Removing existing output directory: {target_dir}", flush=True)
        shutil.rmtree(target_dir)

    print(f"[DEBUG] Creating output directory: {target_dir}", flush=True)
    target_dir.mkdir(parents=True, exist_ok=True)


def export_single_sheet(
    workbook_path: Path, sheet_name: str, destination_path: Path
) -> None:
    print(
        f"[DEBUG] Exporting sheet '{sheet_name}' from {workbook_path.name} "
        f"to {destination_path}",
        flush=True,
    )

    workbook = load_workbook(workbook_path)
    try:
        for worksheet in list(workbook.worksheets):
            if worksheet.title != sheet_name:
                workbook.remove(worksheet)

        workbook.active = 0
        destination_path.parent.mkdir(parents=True, exist_ok=True)
        workbook.save(destination_path)
    finally:
        workbook.close()


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

    print(
        f"[DEBUG] Found {len(workbook_paths)} workbook(s) in {source_dir}", flush=True
    )
    plans = collect_sheet_plans(source_dir, target_dir, workbook_paths)
    if not plans:
        print(f"No sheets found in workbooks under {source_dir}", file=sys.stderr)
        return 1

    recreate_target_dir(target_dir)

    for export_index, (workbook_path, sheet_name, destination_path) in enumerate(
        plans, start=1
    ):
        print(
            f"[DEBUG] Export {export_index}/{len(plans)}",
            flush=True,
        )
        export_single_sheet(workbook_path, sheet_name, destination_path)

    print(
        f"Exported {len(plans)} single-sheet file(s) from "
        f"{len(workbook_paths)} workbook(s) into {target_dir}",
        flush=True,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
