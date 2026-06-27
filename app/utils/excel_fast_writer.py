from __future__ import annotations

from contextlib import contextmanager
from datetime import date, datetime, time
from decimal import Decimal
from typing import Any, Callable, Iterable, Mapping, Sequence

XL_CALCULATION_MANUAL = -4135
DEFAULT_CHUNK_SIZE = 5000


def _to_excel_com_value(value: Any) -> Any:
    """Convert Python values to values that are safe for bulk COM Range.Value write."""
    if value is None:
        return ""
    if isinstance(value, Decimal):
        return float(value)
    if isinstance(value, date) and not isinstance(value, datetime):
        return datetime.combine(value, time.min)
    return value


def write_excel_matrix(ws: Any, matrix: Sequence[Sequence[Any]], *, start_row: int = 1, start_col: int = 1) -> None:
    """Write a 2D matrix to Excel in one COM Range.Value assignment.

    Cell-by-cell COM writes are very slow on large exports. Assigning a whole
    rectangular Range.Value keeps the same workbook/worksheet formatting logic,
    but reduces thousands of COM calls to one call per chunk.
    """
    if not matrix:
        return

    max_cols = max((len(row) for row in matrix), default=0)
    if max_cols <= 0:
        return

    prepared = []
    for row in matrix:
        values = [_to_excel_com_value(value) for value in row]
        if len(values) < max_cols:
            values.extend([""] * (max_cols - len(values)))
        prepared.append(tuple(values))

    row_count = len(prepared)
    first_cell = ws.Cells(start_row, start_col)
    last_cell = ws.Cells(start_row + row_count - 1, start_col + max_cols - 1)
    target_range = ws.Range(first_cell, last_cell)

    if row_count == 1 and max_cols == 1:
        target_range.Value = prepared[0][0]
    else:
        target_range.Value = tuple(prepared)


def _build_excel_row(
    row: Any,
    headers: Sequence[Any],
    value_getter: Callable[[Any, Any, int], Any] | None,
) -> list[Any]:
    row_values: list[Any] = []
    for column_index, header in enumerate(headers):
        if value_getter is not None:
            value = value_getter(row, header, column_index)
        elif isinstance(row, Mapping):
            value = row.get(header, "")
        elif column_index < len(row):  # type: ignore[arg-type]
            value = row[column_index]  # type: ignore[index]
        else:
            value = ""
        row_values.append(value)
    return row_values


def write_excel_table(
    ws: Any,
    headers: Sequence[Any],
    rows: Iterable[Any],
    *,
    header_getter: Callable[[Any], Any] | None = None,
    value_getter: Callable[[Any, Any, int], Any] | None = None,
    start_row: int = 1,
    start_col: int = 1,
    chunk_size: int = DEFAULT_CHUNK_SIZE,
) -> None:
    """Write headers and rows to Excel with fast chunked Range.Value assignments.

    value_getter receives (row, header, zero_based_column_index). If no getter is
    passed, dict rows are read by header name, sequence rows by index.
    """
    header_list = list(headers)
    if not header_list:
        return

    get_header = header_getter or (lambda header: header)
    write_excel_matrix(ws, [[get_header(header) for header in header_list]], start_row=start_row, start_col=start_col)

    excel_row = start_row + 1
    buffer: list[list[Any]] = []
    safe_chunk_size = max(int(chunk_size or DEFAULT_CHUNK_SIZE), 1)

    for row in rows:
        buffer.append(_build_excel_row(row, header_list, value_getter))
        if len(buffer) >= safe_chunk_size:
            write_excel_matrix(ws, buffer, start_row=excel_row, start_col=start_col)
            excel_row += len(buffer)
            buffer.clear()

    if buffer:
        write_excel_matrix(ws, buffer, start_row=excel_row, start_col=start_col)


@contextmanager
def excel_fast_mode(excel: Any):
    """Temporarily reduce Excel UI/recalculation overhead during export."""
    previous: dict[str, Any] = {}
    for attr, value in (
        ("ScreenUpdating", False),
        ("EnableEvents", False),
        ("Calculation", XL_CALCULATION_MANUAL),
    ):
        try:
            previous[attr] = getattr(excel, attr)
            setattr(excel, attr, value)
        except Exception:
            pass

    try:
        yield
    finally:
        for attr, value in previous.items():
            try:
                setattr(excel, attr, value)
            except Exception:
                pass
