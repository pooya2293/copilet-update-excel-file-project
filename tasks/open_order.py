"""Import in-transit order rows into the Mashhad workbook."""

from typing import Any, Optional


def import_in_transit_orders(
    destination_workbook: Any,
    source_workbook: Any,
    destination_sheet_name: str = "Open Order",
    source_sheet_names: Optional[list[str]] = None,
) -> None:
    """Copy F:R rows from open source sheets into an open destination workbook."""
    sheet_names = list(
        ["کاله آمل", "کاله تهران", "سس ,مربا و برنج ", "نوشیدنی"]
        if source_sheet_names is None
        else source_sheet_names
    )
    if not sheet_names:
        raise ValueError("At least one source worksheet name must be provided.")

    destination_ws = destination_workbook.Worksheets(destination_sheet_name)
    destination_ws.Range("F2:R1000").ClearContents()
    destination_row = 2
    copied_rows = 0

    for sheet_name in sheet_names:
        source_ws = source_workbook.Worksheets(sheet_name)
        last_row = source_ws.Cells(source_ws.Rows.Count, 6).End(-4162).Row
        if last_row < 2:
            continue

        row_count = last_row - 1
        source_range = source_ws.Range(f"A2:M{last_row}")
        destination_range = destination_ws.Range(
            f"F{destination_row}:R{destination_row + row_count - 1}"
        )
        destination_range.Value = source_range.Value
        destination_row += row_count + 7
        copied_rows += row_count

    print(
        f"Copied {copied_rows} row(s) to "
        f"{destination_sheet_name}!F:R"
    )
