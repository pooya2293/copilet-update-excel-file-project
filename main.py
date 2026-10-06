"""Process demand rows, inventory matches, and in-transit orders."""

import os
import shutil
import sys
from pathlib import Path

import win32com.client as win32

from tasks.date import increment_b1
from tasks.inv import copy_inventory_matches
from tasks.open_order import import_in_transit_orders
from tasks.remove import remove_demand_ranges
from tasks.sales import update_sales_trend

DEFAULT_WORKBOOK = Path(__file__).with_name("main.xlsb")
IN_TRANSIT_WORKBOOK = Path(__file__).with_name("در راه.xlsm")
INVENTORY_WORKBOOK = Path(__file__).with_name("inv.XLSX")
SOURCE_SALES_WORKBOOK = Path(__file__).with_name("083.XLSX")
TREND_WORKBOOK = Path(__file__).with_name("trend.xlsx")


def main(argv: list[str] | None = None) -> int:
    arguments = sys.argv[1:] if argv is None else argv
    if len(arguments) > 1:
        print("Usage: python main.py [workbook path]")
        return 2

    workbook = Path(arguments[0]) if arguments else DEFAULT_WORKBOOK
    if not workbook.is_file():
        print(f"Workbook not found: {workbook}")
        return 3
    if not IN_TRANSIT_WORKBOOK.is_file():
        print(f"Workbook not found: {IN_TRANSIT_WORKBOOK}")
        return 3
    if not INVENTORY_WORKBOOK.is_file():
        print(f"Workbook not found: {INVENTORY_WORKBOOK}")
        return 3
    if not SOURCE_SALES_WORKBOOK.is_file():
        print(f"Workbook not found: {SOURCE_SALES_WORKBOOK}")
        return 3
    if not TREND_WORKBOOK.is_file():
        print(f"Workbook not found: {TREND_WORKBOOK}")
        return 3

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
    try:
        os.startfile(str(workbook))
    except OSError as error:
        raise RuntimeError(
            f"Processing finished, but could not open the final workbook: "
            f"{workbook}"
        ) from error
    print(f"Opened final workbook: {workbook}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
