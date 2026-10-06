"""Small operations on already-open workbook objects."""

from typing import Any


def increment_b1(workbook: Any) -> None:
    """Increment cell 1000!B1 in an already-open workbook."""
    worksheet = workbook.Worksheets("1000")
    cell = worksheet.Range("B1")
    current_value = cell.Value2
    if not isinstance(current_value, (int, float)):
        raise ValueError(
            f"Cannot increment 1000!B1: expected a number, got {current_value!r}."
        )

    cell.Value2 = current_value + 1