"""Process demand rows, inventory matches, and in-transit orders."""

import gc
import shutil
import sys
import traceback
from datetime import datetime
from pathlib import Path
from time import perf_counter
from typing import Any

import pythoncom
import pywintypes
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
EXCEL_CALC_MANUAL = -4135


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


def report_warning(message: str) -> None:
    print(message, file=sys.stderr)
    try:
        with ERROR_LOG.open("a", encoding="utf-8") as error_log:
            error_log.write(
                f"\n[{datetime.now().astimezone().isoformat()}]\n"
                f"WARNING: {message}\n"
            )
    except OSError as error:
        print(f"Could not write error log {ERROR_LOG}: {error}", file=sys.stderr)


def _release_com_object(obj: Any, *, label: str = "COM object") -> None:
    """Release a COM object and force Python to drop stale references."""
    if obj is None:
        return
    try:
        if hasattr(obj, "Close"):
            try:
                obj.Close(SaveChanges=False)
            except Exception:
                pass
    finally:
        try:
            if hasattr(obj, "_oleobj_"):
                pythoncom.CoDisconnectObject(obj._oleobj_, 0)
        except Exception:
            pass
        try:
            pythoncom.CoReleaseUnusedLibraries()
        except Exception:
            pass
        gc.collect()
        try:
            del obj
        except Exception:
            pass


def _close_workbook(
    workbook: Any,
    *,
    save_changes: bool = False,
    label: str = "workbook",
) -> None:
    """Close one workbook and release lingering COM references."""
    if workbook is None:
        return
    try:
        print(f"Closing {label}...")
        workbook.Close(SaveChanges=save_changes)
        print(f"Closed {label}.")
    except Exception as error:
        print(f"Warning while closing {label}: {error}", file=sys.stderr)
    finally:
        _release_com_object(workbook, label=label)


def _run_timed_stage(label: str, action: Any) -> Any:
    """Execute a workbook stage and emit precise elapsed-time diagnostics."""
    started = perf_counter()
    print(f"[START] {label}...")
    try:
        result = action()
    finally:
        elapsed = perf_counter() - started
        print(f"[DONE] {label} - {elapsed:.2f}s")
    return result


def update_one_destination(
    workbook: Path,
    excel: Any,
    in_transit_wb: Any,
    inventory_wb: Any,
    source_sales_wb: Any,
    trend_wb: Any,
) -> None:
    """Run the existing update sequence for one destination workbook."""
    backup_path = str(workbook) + ".bak"
    backup_started = perf_counter()
    shutil.copy2(workbook, backup_path)
    print(f"Backup created: {backup_path} (in {perf_counter() - backup_started:.2f}s)")

    print(f"Processing workbook: {workbook}")
    destination_wb = None
    try:
        start = perf_counter()
        print(f"Opening {workbook.name}...")
        destination_wb = excel.Workbooks.Open(Filename=str(workbook),UpdateLinks= 0)
        print(
            f"Opened {workbook.name} in {perf_counter() - start:.2f}s."
        )
        _run_timed_stage("Preparing workbook", lambda: increment_b1(destination_wb))
        _run_timed_stage(
            "1000 update",
            lambda: remove_demand_ranges(destination_wb),
        )
        
        _run_timed_stage(
            "Open Order update",
            lambda: import_in_transit_orders(destination_wb, in_transit_wb),
        )
        lookup_value = _run_timed_stage(
            "zdsd037 update",
            lambda: copy_inventory_matches(destination_wb, inventory_wb),
        )
        _run_timed_stage(
            "sales-2 update",
            lambda: update_sales_trend(
                source_sales_wb,
                trend_wb,
                destination_wb,
                excel,
                lookup_value,
            ),
        )

        print(f"Saving {workbook.name}...")
        save_started = perf_counter()
        source_sales_wb.Save()
        trend_wb.Save()
        destination_wb.Save()
        print(f"Saved workbook: {workbook} (in {perf_counter() - save_started:.2f}s)")
    finally:
        _close_workbook(destination_wb, label=workbook.name)
        destination_wb = None


def _excel_app_settings(excel: Any) -> dict[str, Any]:
    """Capture current Excel app settings so they can be restored safely."""
    return {
        "Calculation": excel.Calculation,
        "DisplayAlerts": excel.DisplayAlerts,
        "ScreenUpdating": excel.ScreenUpdating,
        "EnableEvents": excel.EnableEvents,
        "AutomationSecurity": excel.AutomationSecurity,
        "Visible": excel.Visible,
    }


def _preloaded_workbook_names(excel: Any) -> list[str]:
    """List workbooks Excel loaded before this application opened its inputs."""
    workbooks = excel.Workbooks
    return [
        str(workbooks.Item(index).Name)
        for index in range(1, workbooks.Count + 1)
    ]


def main(argv: list[str] | None = None) -> int:
    arguments = sys.argv[1:] if argv is None else argv
    if len(arguments) > 1:
        return report_failure("Usage: main.exe [workbook path]", 2)

    if arguments:
        destination_files = [Path(arguments[0])]
        if not destination_files[0].is_file():
            return report_failure(
                f"Workbook not found: {destination_files[0]}", 3
            )
        skipped_files: list[Path] = []
    else:
        destination_files = [
            APP_DIR / f"{number}.xlsb"
            for number in range(1, 6)
            if (APP_DIR / f"{number}.xlsb").is_file()
        ]
        skipped_files = [
            APP_DIR / f"{number}.xlsb"
            for number in range(1, 6)
            if not (APP_DIR / f"{number}.xlsb").is_file()
        ]
    if not destination_files:
        print("No destination workbooks found.")
        print("Successfully updated: none")
        print("Failed: none")
        print(
            "Not found / skipped: "
            + ", ".join(path.name for path in skipped_files)
        )
        return 3
    print(
        "Found destination files: "
        + ", ".join(path.name for path in destination_files)
    )

    if not IN_TRANSIT_WORKBOOK.is_file():
        return report_failure(f"Workbook not found: {IN_TRANSIT_WORKBOOK}", 3)
    if not INVENTORY_WORKBOOK.is_file():
        return report_failure(f"Workbook not found: {INVENTORY_WORKBOOK}", 3)
    if not SOURCE_SALES_WORKBOOK.is_file():
        return report_failure(f"Workbook not found: {SOURCE_SALES_WORKBOOK}", 3)
    if not TREND_WORKBOOK.is_file():
        return report_failure(f"Workbook not found: {TREND_WORKBOOK}", 3)

    excel = None
    in_transit_wb = None
    inventory_wb = None
    source_sales_wb = None
    trend_wb = None
    successful_files: list[Path] = []
    failed_files: list[Path] = []
    original_settings: dict[str, Any] | None = None
    calculation_changed = False
    try:
        excel = win32.DispatchEx("Excel.Application")
        original_settings = _excel_app_settings(excel)
        try:
            excel.Calculation = EXCEL_CALC_MANUAL
            calculation_changed = True
            print("Excel calculation mode set to Manual.")
        except pywintypes.com_error as error:
            try:
                preloaded_workbooks = _preloaded_workbook_names(excel)
            except pywintypes.com_error as diagnostic_error:
                preloaded_workbooks = [
                    f"<could not inspect: {diagnostic_error}>"
                ]
            report_warning(
                "Excel refused Manual calculation before the application "
                "opened its source workbooks. Continuing with the existing "
                f"calculation mode ({original_settings['Calculation']}); "
                f"preloaded workbooks: {preloaded_workbooks or 'none'}. "
                f"Excel reported: {error}"
            )
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        excel.AutomationSecurity = 3

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
            Filename=str(SOURCE_SALES_WORKBOOK),
            UpdateLinks=0
        )
        trend_wb = excel.Workbooks.Open(Filename=str(TREND_WORKBOOK),
            UpdateLinks=0)

        excel.Calculation = EXCEL_CALC_MANUAL

        for workbook in destination_files:
            print(f"\n========== Processing {workbook.name} ==========")
            try:
                update_one_destination(
                    workbook,
                    excel,
                    in_transit_wb,
                    inventory_wb,
                    source_sales_wb,
                    trend_wb,
                )
            except Exception as error:
                failed_files.append(workbook)
                print(f"{workbook.name} failed: {error}", file=sys.stderr)
                report_failure(
                    f"Failed to update {workbook}:\n{traceback.format_exc()}",
                    1,
                )
            else:
                successful_files.append(workbook)
                print(f"========== {workbook.name} completed ==========")
    finally:
        try:
            _close_workbook(trend_wb, label="trend.xlsx")
        finally:
            try:
                _close_workbook(source_sales_wb, label="083.XLSX")
            finally:
                try:
                    _close_workbook(inventory_wb, label="inv.XLSX")
                finally:
                    try:
                        _close_workbook(in_transit_wb, label="در راه.xlsm")
                    finally:
                        if excel is not None:
                            if original_settings is not None:
                                settings_to_restore = {
                                    "DisplayAlerts": original_settings["DisplayAlerts"],
                                    "ScreenUpdating": original_settings["ScreenUpdating"],
                                    "EnableEvents": original_settings["EnableEvents"],
                                    "AutomationSecurity": original_settings[
                                        "AutomationSecurity"
                                    ],
                                    "Visible": original_settings["Visible"],
                                }
                                if calculation_changed:
                                    settings_to_restore = {
                                        "Calculation": original_settings[
                                            "Calculation"
                                        ],
                                        **settings_to_restore,
                                    }
                                for setting, value in settings_to_restore.items():
                                    try:
                                        setattr(excel, setting, value)
                                    except pywintypes.com_error as error:
                                        report_warning(
                                            "Could not restore Excel "
                                            f"{setting} setting to {value!r}: "
                                            f"{error}"
                                        )
                            print("Quitting Excel application...")
                            try:
                                excel.Quit()
                            except Exception as error:
                                print(
                                    f"Warning while quitting Excel: {error}",
                                    file=sys.stderr,
                                )
                            finally:
                                _release_com_object(excel, label="Excel")

    print(
        "Successfully updated: "
        + (
            ", ".join(path.name for path in successful_files)
            if successful_files
            else "none"
        )
    )
    print(
        "Failed: "
        + (
            ", ".join(path.name for path in failed_files)
            if failed_files
            else "none"
        )
    )
    print(
        "Not found / skipped: "
        + (
            ", ".join(path.name for path in skipped_files)
            if skipped_files
            else "none"
        )
    )

    if successful_files:
        final_excel = None
        try:
            final_excel = win32.DispatchEx("Excel.Application")
            final_excel.Visible = False
            final_excel.DisplayAlerts = False
            final_excel.AutomationSecurity = 3
            for workbook in successful_files:
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
            return report_failure(
                "Processing finished, but Excel could not display the "
                f"updated workbooks: {error}",
                1,
            )
        print(
            "Opened final workbook(s): "
            + ", ".join(path.name for path in successful_files)
        )

    return 1 if failed_files else 0


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
