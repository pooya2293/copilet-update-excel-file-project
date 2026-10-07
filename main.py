"""Process demand rows, inventory matches, and in-transit orders."""

import shutil
import sys
import traceback
from datetime import datetime
from pathlib import Path

import win32com.client as win32

from tasks.date import increment_b1
from tasks.inv import copy_inventory_matches
from tasks.open_order import import_in_transit_orders
from tasks.remove import remove_demand_ranges
from tasks.sales import update_sales_trend

APP_DIR = (
    Path(sys.executable).resolve().parent
    if getattr(sys, "frozen", False)
    else Path(__file__).resolve().parent
)
DEFAULT_WORKBOOK = APP_DIR / "1.xlsb"
IN_TRANSIT_WORKBOOK = APP_DIR / "در راه.xlsm"
INVENTORY_WORKBOOK = APP_DIR / "inv.XLSX"
SOURCE_SALES_WORKBOOK = APP_DIR / "083.XLSX"
TREND_WORKBOOK = APP_DIR / "trend.xlsx"
ERROR_LOG = APP_DIR / "main-error.log"


def report_failure(message: str, exit_code: int) -> int:
    print(message, file=sys.stderr)
    try:
        with ERROR_LOG.open("a", encoding="utf-8") as error_log:
            error_log.write(
                f"\n[{datetime.now().astimezone().isoformat()}]\n"
                f"{message}\n"
            )
    except OSError as error:
        print(f"Could not write error log {ERROR_LOG}: {error}", file=sys.stderr)
    return exit_code


def main(argv: list[str] | None = None) -> int:
    arguments = sys.argv[1:] if argv is None else argv
    if len(arguments) > 1:
        return report_failure("Usage: main.exe [workbook path]", 2)

    workbook = Path(arguments[0]) if arguments else DEFAULT_WORKBOOK
    if not workbook.is_file():
        return report_failure(f"Workbook not found: {workbook}", 3)
    if not IN_TRANSIT_WORKBOOK.is_file():
        return report_failure(f"Workbook not found: {IN_TRANSIT_WORKBOOK}", 3)
    if not INVENTORY_WORKBOOK.is_file():
        return report_failure(f"Workbook not found: {INVENTORY_WORKBOOK}", 3)
    if not SOURCE_SALES_WORKBOOK.is_file():
        return report_failure(f"Workbook not found: {SOURCE_SALES_WORKBOOK}", 3)
    if not TREND_WORKBOOK.is_file():
        return report_failure(f"Workbook not found: {TREND_WORKBOOK}", 3)

    backup_path = str(workbook) + ".bak"
    shutil.copy2(workbook, backup_path)
    print(f"Backup created: {backup_path}")

    print(f"Processing workbook: {workbook}")
    excel = None
    destination_wb = None
    in_transit_wb = None
    inventory_wb = None
    source_sales_wb = None
    trend_wb = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        excel.AutomationSecurity = 3

        destination_wb = excel.Workbooks.Open(Filename=str(workbook))
        increment_b1(destination_wb)
        in_transit_wb = excel.Workbooks.Open(
            Filename=str(IN_TRANSIT_WORKBOOK),
            ReadOnly=True,
            UpdateLinks=0,
        )
        inventory_wb = excel.Workbooks.Open(
            Filename=str(INVENTORY_WORKBOOK),
            ReadOnly=True,
            UpdateLinks=0,
        )
        source_sales_wb = excel.Workbooks.Open(
            Filename=str(SOURCE_SALES_WORKBOOK)
        )
        trend_wb = excel.Workbooks.Open(Filename=str(TREND_WORKBOOK))

        remove_demand_ranges(destination_wb)
        import_in_transit_orders(destination_wb, in_transit_wb)
        lookup_value = copy_inventory_matches(destination_wb, inventory_wb)
        update_sales_trend(
            source_sales_wb,
            trend_wb,
            destination_wb,
            excel,
            lookup_value,
        )

        source_sales_wb.Save()
        trend_wb.Save()
        destination_wb.Save()
        print(f"Saved workbook: {workbook}")
    finally:
        try:
            if trend_wb is not None:
                trend_wb.Close(SaveChanges=False)
        finally:
            try:
                if source_sales_wb is not None:
                    source_sales_wb.Close(SaveChanges=False)
            finally:
                try:
                    if inventory_wb is not None:
                        inventory_wb.Close(SaveChanges=False)
                finally:
                    try:
                        if in_transit_wb is not None:
                            in_transit_wb.Close(SaveChanges=False)
                    finally:
                        try:
                            if destination_wb is not None:
                                destination_wb.Close(SaveChanges=False)
                        finally:
                            if excel is not None:
                                excel.Quit()
    final_excel = None
    try:
        final_excel = win32.DispatchEx("Excel.Application")
        final_excel.Visible = False
        final_excel.DisplayAlerts = False
        final_excel.AutomationSecurity = 3
        final_workbook = final_excel.Workbooks.Open(
            Filename=str(workbook.resolve()),
            UpdateLinks=0,
        )
        final_workbook.Activate()
        final_excel.ScreenUpdating = True
        final_excel.WindowState = -4143
        final_excel.Visible = True
    except Exception as error:
        if final_excel is not None:
            final_excel.Quit()
        raise RuntimeError(
            "Processing finished, but Excel could not display the final "
            f"workbook {workbook}: {error}"
        ) from error
    print(f"Opened final workbook: {workbook}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception:
        details = traceback.format_exc()
        try:
            with ERROR_LOG.open("a", encoding="utf-8") as error_log:
                error_log.write(
                    f"\n[{datetime.now().astimezone().isoformat()}]\n"
                    f"{details}"
                )
        except OSError as error:
            print(f"Could not write error log {ERROR_LOG}: {error}", file=sys.stderr)
        print(details, file=sys.stderr, end="")
        print(f"Detailed error log: {ERROR_LOG}", file=sys.stderr)
        raise SystemExit(1)
