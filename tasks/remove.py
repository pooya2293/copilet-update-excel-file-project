"""Clear G:R in rows whose column D contains the demand keyword."""

from typing import Any, Iterable, Optional


def remove_demand_ranges(
    workbook: Any,
    sheets: Optional[Iterable[str]] = None,
    demand_keyword: str = "demand",
) -> None:
    """Clear matching cells in an already-open Excel workbook."""
    sheet_names = list(
        ["1000", "1m03", "1400", "1800"] if sheets is None else sheets
    )
    if not sheet_names:
        raise ValueError("At least one worksheet name must be provided.")
    if not demand_keyword:
        raise ValueError("The demand keyword cannot be empty.")

    for sheet_name in sheet_names:
        sheet = workbook.Worksheets(sheet_name)
        used_range = sheet.UsedRange
        last_row = used_range.Row + used_range.Rows.Count - 1
        values = sheet.Range(f"D1:D{last_row}").Value2
        if last_row == 1:
            values = ((values,),)

        ranges: list[str] = []
        first_match = 0
        last_match = 0
        cleared_rows = 0

        for row in range(1, last_row + 1):
            value = values[row - 1][0]
            if value is not None and demand_keyword.casefold() in str(value).casefold():
                cleared_rows += 1
                if first_match == 0:
                    first_match = row
                elif row != last_match + 1:
                    ranges.append(f"G{first_match}:R{last_match}")
                    first_match = row
                last_match = row

        if first_match:
            ranges.append(f"G{first_match}:R{last_match}")

        batch: list[str] = []
        batch_length = 0
        for range_address in ranges:
            next_length = len(range_address) + (1 if batch else 0)
            if batch and batch_length + next_length > 180:
                sheet.Range(",".join(batch)).ClearContents()
                batch.clear()
                batch_length = 0
                next_length = len(range_address)
            batch.append(range_address)
            batch_length += next_length

        if batch:
            sheet.Range(",".join(batch)).ClearContents()

        print(f"{sheet_name}: cleared {cleared_rows} row(s)")
