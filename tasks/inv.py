"""Copy inventory rows matching the workbook's 1000!D5 lookup value."""

from typing import Any, Optional


def copy_inventory_matches(
    destination_workbook: Any,
    inventory_workbook: Optional[Any] = None,
) -> Any:
    """Replace zdsd037 columns B:D and return the existing plant lookup value."""
    if inventory_workbook is None:
        raise ValueError("An open inventory workbook must be provided.")

    inventory_sheet = inventory_workbook.Worksheets("Sheet1")
    lookup_sheet = destination_workbook.Worksheets("1000")
    destination_sheet = destination_workbook.Worksheets("zdsd037")
    lookup_value = lookup_sheet.Range("D5").Value2
    if lookup_value is None or not str(lookup_value).strip():
        raise ValueError("Lookup cell 1000!D5 is blank.")

    last_inventory_row = inventory_sheet.Cells(
        inventory_sheet.Rows.Count, 7
    ).End(-4162).Row
    inventory_values = inventory_sheet.Range(
        f"A1:G{last_inventory_row}"
    ).Value2
    if last_inventory_row == 1:
        inventory_values = (inventory_values,)

    matching_rows = [
        (row[0], row[1], row[2])
        for row in inventory_values
        if row[6] is not None
        and str(row[6]).casefold() == str(lookup_value).casefold()
    ]
    if not matching_rows:
        raise ValueError(
            "No inventory rows found with column G equal to "
            f"1000!D5 value {lookup_value!r}. Destination was not changed."
        )

    used_range = destination_sheet.UsedRange
    last_destination_row = used_range.Row + used_range.Rows.Count - 1
    destination_sheet.Range(f"B1:D{last_destination_row}").ClearContents()
    destination_sheet.Range(
        f"B1:D{len(matching_rows)}"
    ).Value2 = tuple(matching_rows)

    print(
        f"Copied {len(matching_rows)} matching row(s) to zdsd037!B:D"
    )
    return lookup_value
