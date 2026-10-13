"""Reconcile committed reference rows without reloading the entire result set.

Call only after a successful database commit, when rows have stable database IDs.
The caller is responsible for fetching the affected rows in one query and for
filtering them according to the currently active page filters.
"""
from __future__ import annotations

from collections.abc import Callable, Mapping
from typing import Any

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QTableWidget

from app.utils.table_view_context import capture_table_context, restore_table_context


def reconcile_committed_rows(
    table: QTableWidget,
    changed_ids: set[int],
    visible_rows_by_id: Mapping[int, Any],
    paint_row: Callable[[int, Any], None],
) -> None:
    """Replace only affected UI rows, preserving selection/sort/scroll by DB ID.

    ``visible_rows_by_id`` contains changed records still matching the filters.
    Missing IDs (deleted/filtered away) are removed. Uses only QTableWidget
    item data; no PostgreSQL call is made by this function.
    """
    if not changed_ids:
        return
    state = capture_table_context(table)
    table.setSortingEnabled(False)  # prevent row movement while inserting cells
    try:
        positions: dict[int, int] = {}
        for row in range(table.rowCount() - 1, -1, -1):
            item = table.item(row, 0)
            if item is None:
                continue
            raw_id = item.data(Qt.ItemDataRole.UserRole)
            if raw_id is None:
                raw_id = item.text()
            try:
                record_id = int(raw_id)
            except (ValueError, TypeError):
                continue
            if record_id in changed_ids:
                positions[record_id] = row
                table.removeRow(row)
        # For unsorted tables, retain the approximate original position.
        for record_id, data in sorted(visible_rows_by_id.items(), key=lambda entry: positions.get(entry[0], table.rowCount())):
            at = min(positions.get(record_id, table.rowCount()), table.rowCount())
            table.insertRow(at)
            paint_row(at, data)
    finally:
        table.setSortingEnabled(state.sorting_enabled)
        restore_table_context(table, state)
