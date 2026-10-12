
import os
import shutil
import traceback
from pathlib import Path

import win32com.client as win32


# ==========================================
# Settings
# ==========================================

import sys
from pathlib import Path

if getattr(sys, "frozen", False):
    APP_DIR = Path(sys.executable).resolve().parent
else:
    APP_DIR = Path(__file__).resolve().parent

SOURCE_FILE = APP_DIR / "در راه.xlsm"
OUTPUT_DIR = APP_DIR / "Branches"

SOURCE_FILE = APP_DIR / "در راه.xlsm"
OUTPUT_DIR = APP_DIR / "Branches"

SHEET_NAMES = [
    "کاله آمل",
    "کاله تهران",
    "سس",
    "نوشیدنی",
]

BRANCH_COLUMN = 5  # Column E
HEADER_ROW = 1

XL_UP = -4162


def get_unique_branches(workbook):
    """Extract unique branch names from column E of all four sheets."""

    branches = {}

    for sheet_name in SHEET_NAMES:
        sheet = workbook.Worksheets(sheet_name)

        last_row = sheet.Cells(
            sheet.Rows.Count,
            BRANCH_COLUMN
        ).End(XL_UP).Row

        if last_row <= HEADER_ROW:
            continue

        values = sheet.Range(
            sheet.Cells(HEADER_ROW + 1, BRANCH_COLUMN),
            sheet.Cells(last_row, BRANCH_COLUMN),
        ).Value2

        # Excel returns a scalar for a single-cell range.
        if last_row == HEADER_ROW + 1:
            values = ((values,),)

        for row in values:
            value = row[0] if isinstance(row, tuple) else row

            if value is None:
                continue

            branch_name = str(value).strip()

            if branch_name:
                branches.setdefault(
                    branch_name.casefold(),
                    branch_name
                )

    return list(branches.values())


def safe_filename(name):
    """Replace invalid Windows filename characters."""

    invalid_chars = '<>:"/\\|?*'

    cleaned = "".join(
        "_" if char in invalid_chars or ord(char) < 32 else char
        for char in name
    )

    return cleaned.strip().rstrip(". ")


def keep_only_branch(sheet, branch_name):
    """Delete rows belonging to other branches, from bottom to top."""

    last_row = sheet.Cells(
        sheet.Rows.Count,
        BRANCH_COLUMN
    ).End(XL_UP).Row

    for row_index in range(last_row, HEADER_ROW, -1):
        value = sheet.Cells(
            row_index,
            BRANCH_COLUMN
        ).Value2

        current_branch = (
            "" if value is None else str(value).strip()
        )

        if current_branch.casefold() != branch_name.casefold():
            sheet.Rows(row_index).Delete()


def main():
    if not SOURCE_FILE.is_file():
        print(f"ERROR: Source file not found:\n{SOURCE_FILE}")
        return 1

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    excel = None
    source_workbook = None
    created_count = 0
    failed_count = 0

    try:
        # Start a separate Excel instance.
        excel = win32.DispatchEx("Excel.Application")

        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        excel.AutomationSecurity = 3

        # Open the original workbook read-only.
        source_workbook = excel.Workbooks.Open(
            Filename=str(SOURCE_FILE.resolve()),
            ReadOnly=True,
            UpdateLinks=0,
        )

        branches = get_unique_branches(source_workbook)

        if not branches:
            print("No branch names found in column E.")
            return 1

        print(f"Branches found: {len(branches)}")
        print(f"Output folder: {OUTPUT_DIR}\n")

        source_stem = SOURCE_FILE.stem
        extension = SOURCE_FILE.suffix

        for branch_name in branches:
            output_path = OUTPUT_DIR / (
                f"{source_stem} {safe_filename(branch_name)}{extension}"
            )

            branch_workbook = None

            try:
                # Do not overwrite an existing output file.
                if output_path.exists():
                    output_path.unlink()

                # Copy the original, preserving its Excel file format.
                shutil.copy2(SOURCE_FILE, output_path)

                # Open the copy in Excel.
                branch_workbook = excel.Workbooks.Open(
                    Filename=str(output_path.resolve()),
                    UpdateLinks=0,
                )

                # Keep this branch's rows in each original sheet.
                for sheet_name in SHEET_NAMES:
                    sheet = branch_workbook.Worksheets(sheet_name)
                    keep_only_branch(sheet, branch_name)

                branch_workbook.Save()
                created_count += 1

                print(f"[OK] {output_path.name}")

            except Exception as error:
                failed_count += 1
                print(f"[ERROR] {branch_name}: {error}")

            finally:
                if branch_workbook is not None:
                    try:
                        branch_workbook.Close(SaveChanges=False)
                    except Exception:
                        pass

        print("\n========== SUMMARY ==========")
        print(f"Created: {created_count}")
        print(f"Failed:  {failed_count}")
        print(f"Folder:  {OUTPUT_DIR}")

        return 1 if failed_count else 0

    finally:
        if source_workbook is not None:
            try:
                source_workbook.Close(SaveChanges=False)
            except Exception:
                pass

        if excel is not None:
            try:
                excel.ScreenUpdating = True
                excel.EnableEvents = True
                excel.Quit()
            except Exception:
                pass


if __name__ == "__main__":
    try:
        exit_code = main()
    except Exception:
        traceback.print_exc()
        exit_code = 1

    input("\nPress Enter to close...")
    raise SystemExit(exit_code)