
"""Small operations on already-open workbook objects."""

from datetime import datetime, timedelta
from typing import Any


def increment_b1(workbook: Any) -> None:
    """Increment 1000!B1 and skip Friday."""
    worksheet = workbook.Worksheets("1000")
    cell = worksheet.Range("B1")
    current_value = cell.Value2

    if not isinstance(current_value, (int, float)):
        raise ValueError(
            f"Cannot increment 1000!B1: expected a date serial, "
            f"got {current_value!r}."
        )

    # Convert Excel serial date to a Python datetime.
    excel_epoch = datetime(1899, 12, 30)
    current_date = excel_epoch + timedelta(days=current_value)

    # Increment by one day.
    next_date = current_date + timedelta(days=1)

    # If the new date is Friday, skip to Saturday.
    if next_date.weekday() == 4:
        next_date += timedelta(days=1)

    # Write the updated date back to Excel.
    cell.Value2 = (next_date - excel_epoch).total_seconds() / 86400
