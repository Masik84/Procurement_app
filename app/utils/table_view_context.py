"""Preserve QTableWidget navigation state across a deliberate data reload.

Only visual state belongs here; database IDs remain the page's source of truth.
Call capture before clearing and restore after sorting has been restored.
"""
from __future__ import annotations

from dataclasses import dataclass
from functools import wraps
from typing import Any, Callable

from PySide6.QtCore import Qt, QItemSelection, QItemSelectionModel
from PySide6.QtWidgets import QTableWidget

# Dedicated per-row identity. UserRole and UserRole + 1 already carry
# column names and row IDs for page-specific edit handlers.
ROW_ID_ROLE = Qt.ItemDataRole.UserRole + 3


@dataclass(frozen=True)
class TableViewContext:
    selected_ids: frozenset[Any]
    current_id: Any
    current_column: int
    vertical_scroll: int
    horizontal_scroll: int
    sorting_enabled: bool
    sort_column: int
    sort_order: Qt.SortOrder


def _key(table: QTableWidget, row: int, key_for_row: Callable[[int], Any] | None = None):
    if key_for_row is not None:
        if not 0 <= row < table.rowCount():
            return None
        try:
            return key_for_row(row)
        except (IndexError, KeyError):
            # A just-cleared batch may have fewer row keys than stale Qt rows.
            return None
    if row < 0 or row >= table.rowCount():
        return None
    item = table.item(row, 0)
    if item is None:
        return None
    # The first cell of several staging tables uses UserRole for a field
    # name. Never interpret that field name as a row key.
    for role in (ROW_ID_ROLE, Qt.ItemDataRole.UserRole + 1, Qt.ItemDataRole.UserRole):
        value = item.data(role)
        if value is None or isinstance(value, bool):
            continue
        try:
            return int(value)
        except (ValueError, TypeError):
            continue
    try:
        return int(item.text())
    except (ValueError, TypeError):
        return None


def capture_table_context(
    table: QTableWidget, *, key_for_row: Callable[[int], Any] | None = None
) -> TableViewContext:
    header = table.horizontalHeader()
    return TableViewContext(
        selected_ids=frozenset(k for row in {i.row() for i in table.selectedIndexes()}
                               if (k := _key(table, row, key_for_row)) is not None),
        current_id=_key(table, table.currentRow(), key_for_row),
        current_column=table.currentColumn(),
        vertical_scroll=table.verticalScrollBar().value(),
        horizontal_scroll=table.horizontalScrollBar().value(),
        sorting_enabled=table.isSortingEnabled(),
        sort_column=header.sortIndicatorSection(),
        sort_order=header.sortIndicatorOrder(),
    )


def restore_table_context(
    table: QTableWidget, state: TableViewContext, *,
    key_for_row: Callable[[int], Any] | None = None
) -> None:
    # The caller has finished populating. Sort exactly once, then look up the IDs.
    if state.sorting_enabled and 0 <= state.sort_column < table.columnCount():
        table.sortItems(state.sort_column, state.sort_order)
    table.clearSelection()
    current_row = None
    for row in range(table.rowCount()):
        key = _key(table, row, key_for_row)
        if key in state.selected_ids:
            # SelectRows allows selection through the first column's range.
            table.selectRow(row) if len(state.selected_ids) == 1 else None
        if key == state.current_id and key is not None:
            current_row = row
    if len(state.selected_ids) > 1 and table.columnCount():
        selection = QItemSelection()
        for row in range(table.rowCount()):
            if _key(table, row, key_for_row) in state.selected_ids:
                selection.select(table.model().index(row, 0),
                                 table.model().index(row, max(0, table.columnCount() - 1)))
        if table.columnCount() and not selection.isEmpty():
            table.selectionModel().select(selection, QItemSelectionModel.SelectionFlag.Select)
    if current_row is not None and table.columnCount():
        table.setCurrentCell(current_row, max(0, min(state.current_column, table.columnCount() - 1)),
                             QItemSelectionModel.SelectionFlag.NoUpdate)
    table.verticalScrollBar().setValue(state.vertical_scroll)
    table.horizontalScrollBar().setValue(state.horizontal_scroll)


def tag_table_row_ids(table: QTableWidget, row_ids) -> None:
    """Attach stable database IDs to QTableWidgetItem instances, not row indexes.

    Must be called after populating rows with sorting temporarily disabled.
    """
    for row_index, record_id in enumerate(row_ids):
        if row_index >= table.rowCount():
            break
        item = table.item(row_index, 0)
        if item is None:
            from PySide6.QtWidgets import QTableWidgetItem
            item = QTableWidgetItem()
            table.setItem(row_index, 0, item)
        item.setData(ROW_ID_ROLE, int(record_id))


def table_row_id(table: QTableWidget, row: int) -> int | None:
    """Return the committed ID belonging to the current *visual* row."""
    if row < 0 or row >= table.rowCount():
        return None
    item = table.item(row, 0)
    if item is None:
        return None
    for role in (ROW_ID_ROLE, Qt.ItemDataRole.UserRole + 1, Qt.ItemDataRole.UserRole):
        value = item.data(role)
        if value is None or isinstance(value, bool):
            continue
        try:
            return int(value)
        except (ValueError, TypeError):
            continue
    return None


def find_table_row_id(table: QTableWidget, record_id: int) -> int:
    for row in range(table.rowCount()):
        if table_row_id(table, row) == record_id:
            return row
    return -1


def stable_row_editor(method):
    """Resolve editor's visual row from its DB ID and paint with sorting paused.

    For page handlers shaped (self, visual_row, record_id, ...). An active
    QComboBox may outlive a sort. Never trust the captured visual_row. The
    database write and the existing business logic remain in the handler.
    """
    @wraps(method)
    def wrapped(self, row: int, record_id: int, *args, **kwargs):
        current_row = find_table_row_id(self.table, record_id)
        if current_row < 0:
            return None  # the edited record was removed/reloaded meanwhile
        was_sorting = self.table.isSortingEnabled()
        self.table.setSortingEnabled(False)
        try:
            return method(self, current_row, record_id, *args, **kwargs)
        finally:
            self.table.setSortingEnabled(was_sorting)
    return wrapped
