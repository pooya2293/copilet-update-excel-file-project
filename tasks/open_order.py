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

    lookup_value = destination_workbook.Worksheets("1000").Range("C4").Value2
    if lookup_value is None or not str(lookup_value).strip():
        raise ValueError("Lookup cell 1000!C4 is blank.")

    destination_ws = destination_workbook.Worksheets(destination_sheet_name)
    matching_rows_by_sheet = []

    for sheet_name in sheet_names:
        source_ws = source_workbook.Worksheets(sheet_name)
        last_row = source_ws.Cells(source_ws.Rows.Count, 6).End(-4162).Row
        if last_row < 2:
            continue

        source_values = source_ws.Range(f"A2:M{last_row}").Value2
        matching_rows = [
            row
            for row in source_values
            if row[5] is not None
            and str(row[5]).casefold() == str(lookup_value).casefold()
        ]
        if matching_rows:
            matching_rows_by_sheet.append(matching_rows)

    copied_rows = sum(len(rows) for rows in matching_rows_by_sheet)
    if not copied_rows:
        raise ValueError(
            "No in-transit rows found with column F equal to "
            f"1000!C4 value {lookup_value!r}. Destination was not changed."
        )

    destination_ws.Range("F2:R1000").ClearContents()
    destination_row = 2
    for matching_rows in matching_rows_by_sheet:
        row_count = len(matching_rows)
        destination_range = destination_ws.Range(
            f"F{destination_row}:R{destination_row + row_count - 1}"
        )
        destination_range.Value = tuple(matching_rows)
        destination_row += row_count + 7

    print(
        f"Copied {copied_rows} row(s) to "
        f"{destination_sheet_name}!F:R"
    )
