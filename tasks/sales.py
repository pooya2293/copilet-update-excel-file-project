"""Update the trend PivotTable and append its filtered results to Mashhad."""

from typing import Any


def update_sales_trend(
    source_workbook: Any,
    trend_workbook: Any,
    destination_workbook: Any,
    excel: Any,
    lookup_value: Any,
) -> None:
    """Refresh sales data and append the selected-plant PivotTable results."""
    if lookup_value is None or not str(lookup_value).strip():
        raise ValueError("Plant lookup value from 1000!D5 is blank.")

    # Convert each specified source column independently using Text to Columns.
    try:
        source_sheet = source_workbook.Worksheets("Sheet1")
        source_used_range = source_sheet.UsedRange
        last_source_row = (
            source_used_range.Row + source_used_range.Rows.Count - 1
        )
        last_source_column = (
            source_used_range.Column + source_used_range.Columns.Count - 1
        )
        for column_number in (2, 8, 9):
            column_range = source_sheet.Range(
                source_sheet.Cells(1, column_number),
                source_sheet.Cells(last_source_row, column_number),
            )
            column_range.TextToColumns(
                Destination=column_range.Cells(1, 1),
                DataType=1,
                TextQualifier=1,
                ConsecutiveDelimiter=False,
                Tab=False,
                Semicolon=False,
                Comma=False,
                Space=False,
                Other=False,
            )

        source_data = source_sheet.Range(
            source_sheet.Cells(1, 1),
            source_sheet.Cells(last_source_row, last_source_column),
        ).Value2
    except Exception as error:
        raise RuntimeError(
            "Step 1 failed while processing 083.XLSX columns B, H, and I "
            f"with Text to Columns: {error}"
        ) from error

    # Replace DB cell contents while retaining its existing formatting.
    try:
        db_sheet = trend_workbook.Worksheets("DB")
        db_sheet.UsedRange.ClearContents()
        source_range = source_sheet.Range(
            source_sheet.Cells(1, 1),
            source_sheet.Cells(last_source_row, last_source_column),
        )
        destination_range = db_sheet.Range(
            db_sheet.Cells(1, 1),
            db_sheet.Cells(last_source_row, last_source_column),
        )
        source_range.Copy()
        trend_workbook.Activate()
        db_sheet.Activate()
        destination_range.PasteSpecial(Paste=-4163)

        pasted_data = destination_range.Value2
        if pasted_data != source_data:
            raise ValueError(
                "Pasted DB data does not match the source data from 083.XLSX."
            )
    except Exception as error:
        raise RuntimeError(
            f"Step 2 failed while replacing the trend.xlsx DB data: {error}"
        ) from error

    # Refresh all workbook connections and the PivotTable from the replaced DB.
    try:
        sales_sheet = trend_workbook.Worksheets("sales")
        pivot_table = sales_sheet.Range("A2").PivotTable
        trend_workbook.RefreshAll()
        excel.CalculateUntilAsyncQueriesDone()
        refreshed_record_count = pivot_table.PivotCache().RecordCount
        if refreshed_record_count != last_source_row:
            raise RuntimeError(
                "Refresh All completed, but the PivotCache contains "
                f"{refreshed_record_count} records; expected "
                f"{last_source_row}."
            )
    except Exception as error:
        raise RuntimeError(
            "Step 3 failed while refreshing the sales PivotTable "
            f"from DB: {error}"
        ) from error

    # Find the Plant slicer, select exactly the lookup plant, and verify it.
    try:
        slicer_cache = _find_slicer_cache(trend_workbook, "Plant")
        items = slicer_cache.SlicerItems
        matching_items = []
        for index in range(1, items.Count + 1):
            item = items.Item(index)
            if str(item.Caption).strip().casefold() == str(
                lookup_value
            ).strip().casefold() or str(item.Name).strip().casefold() == str(
                lookup_value
            ).strip().casefold():
                matching_items.append(item)

        if len(matching_items) != 1:
            raise ValueError(
                f"Plant slicer must contain exactly one item matching "
                f"{lookup_value!r}; found {len(matching_items)}."
            )

        target_item = matching_items[0]
        slicer_cache.ClearManualFilter()
        for index in range(1, items.Count + 1):
            item = items.Item(index)
            if str(item.Name) != str(target_item.Name):
                item.Selected = False
        target_item.Selected = True

        selected_items = [
            items.Item(index)
            for index in range(1, items.Count + 1)
            if items.Item(index).Selected
        ]
        if (
            len(selected_items) != 1
            or str(selected_items[0].Name) != str(target_item.Name)
        ):
            raise RuntimeError(
                f"Plant slicer did not apply the requested filter "
                f"{lookup_value!r}."
            )
        excel.CalculateUntilAsyncQueriesDone()
    except Exception as error:
        raise RuntimeError(
            f"Step 4 failed while filtering the Plant slicer: {error}"
        ) from error

    # Read the PivotTable's actual bounds after its slicer filter is applied.
    try:
        pivot_range = pivot_table.TableRange1
        if pivot_range.Row != 1 or pivot_range.Column != 1:
            raise ValueError(
                "The sales PivotTable is not anchored at cell A1."
            )
        last_pivot_row = pivot_range.Row + pivot_range.Rows.Count - 1
        if last_pivot_row < 2:
            raise ValueError("The filtered PivotTable contains no data rows.")

        pivot_values = sales_sheet.Range(
            f"A2:C{last_pivot_row}"
        ).Value2
        if last_pivot_row == 2:
            pivot_values = (pivot_values,)
        pivot_rows = tuple(
            tuple(row)
            for row in pivot_values
            if any(value is not None and value != "" for value in row)
        )
        if not pivot_rows:
            raise ValueError("The filtered PivotTable contains no populated rows.")
    except Exception as error:
        raise RuntimeError(
            "Step 5 failed while reading the filtered PivotTable data: "
            f"{error}"
        ) from error

    # Append after the existing column A data without overwriting any cells.
    try:
        destination_sheet = destination_workbook.Worksheets("sales-2")
        last_destination_row = destination_sheet.Cells(
            destination_sheet.Rows.Count, 1
        ).End(-4162).Row
        last_cell_value = destination_sheet.Cells(
            last_destination_row, 1
        ).Value2
        first_empty_row = (
            1
            if last_destination_row == 1 and last_cell_value is None
            else last_destination_row + 1
        )
        append_range = destination_sheet.Range(
            f"A{first_empty_row}:C{first_empty_row + len(pivot_rows) - 1}"
        )
        existing_values = append_range.Value2
        if any(
            value is not None and value != ""
            for row in existing_values
            for value in row
        ):
            raise ValueError(
                f"Append range {append_range.Address()} already contains data."
            )
        append_range.Value2 = pivot_rows
    except Exception as error:
        raise RuntimeError(
            "Step 6 failed while appending PivotTable rows to sales-2: "
            f"{error}"
        ) from error

    print(
        f"Appended {len(pivot_rows)} filtered sales row(s) "
        f"to sales-2 starting at row {first_empty_row}."
    )

    # Add the next sales period and its trend lookup to the summary sheet.
    try:
        summary_sheet = destination_workbook.Worksheets("فروش تفکیکی")
        last_summary_column = summary_sheet.Cells(
            2, summary_sheet.Columns.Count
        ).End(-4159).Column
        previous_header_cell = summary_sheet.Cells(2, last_summary_column)
        if previous_header_cell.Value2 is None or previous_header_cell.Value2 == "":
            raise ValueError("Row 2 does not contain a last populated cell.")

        new_column = last_summary_column + 1
        new_header_cell = summary_sheet.Cells(2, new_column)
        new_header_cell.FormulaR1C1 = "=RC[-1]+1"

        formula_cell = summary_sheet.Cells(3, new_column)
        formula_cell.Formula = (
            '=IFERROR(VLOOKUP(A3,[trend.xlsx]trend!$A:$E,5,0),"")'
        )
        last_summary_row = summary_sheet.Cells(
            summary_sheet.Rows.Count, 1
        ).End(-4162).Row
        if last_summary_row > 3:
            summary_sheet.Range(
                formula_cell,
                summary_sheet.Cells(last_summary_row, new_column),
            ).FillDown()
    except Exception as error:
        raise RuntimeError(
            "Step 7 failed while adding the sales trend formula column "
            f"to فروش تفکیکی: {error}"
        ) from error


def _find_slicer_cache(workbook: Any, caption: str) -> Any:
    """Return the cache for the slicer with the requested visible caption."""
    caches = workbook.SlicerCaches
    for cache_index in range(1, caches.Count + 1):
        cache = caches.Item(cache_index)
        for slicer_index in range(1, cache.Slicers.Count + 1):
            slicer = cache.Slicers.Item(slicer_index)
            if str(slicer.Caption).strip().casefold() == caption.casefold():
                return cache
    raise ValueError(f"Slicer {caption!r} was not found.")
