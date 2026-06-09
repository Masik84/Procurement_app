from __future__ import annotations

from typing import Any

from PySide6.QtCore import QPoint, Qt
from PySide6.QtGui import QGuiApplication
from PySide6.QtWidgets import QAbstractItemView, QCheckBox, QComboBox, QMenu, QTableWidget, QTableWidgetItem

DELETE_MESSAGE = "Полное удаление строк будет сделано при сохранении"


def _message(owner: Any, text: str) -> None:
    if hasattr(owner, "show_message"):
        owner.show_message(text)
        return
    label = getattr(getattr(owner, "ui", None), "label_msg", None)
    if label is not None:
        label.setText(text)
        label.setVisible(True)


def configure_table_selection(table: QTableWidget) -> None:
    table.setSelectionMode(QAbstractItemView.ExtendedSelection)
    table.setSelectionBehavior(QAbstractItemView.SelectRows)


def _selected_rows(table: QTableWidget, fallback_row: int | None = None) -> list[int]:
    rows = sorted({idx.row() for idx in table.selectedIndexes()})
    if not rows and fallback_row is not None and fallback_row >= 0:
        rows = [fallback_row]
    return rows


def _cell_text(table: QTableWidget, row: int, col: int) -> str:
    widget = table.cellWidget(row, col)
    if isinstance(widget, QComboBox):
        return widget.currentText()
    if isinstance(widget, QCheckBox):
        return "да" if widget.isChecked() else "нет"
    item = table.item(row, col)
    return item.text() if item else ""


def _copy_field(table: QTableWidget, row: int, col: int) -> None:
    if row < 0 or col < 0:
        return
    QGuiApplication.clipboard().setText(_cell_text(table, row, col))


def _copy_rows(table: QTableWidget, rows: list[int]) -> None:
    if not rows:
        return
    lines = []
    for row in rows:
        lines.append("\t".join(_cell_text(table, row, col) for col in range(table.columnCount())))
    QGuiApplication.clipboard().setText("\n".join(lines))


def _row_key(owner: Any, table: QTableWidget, row: int) -> Any:
    row_ids = getattr(owner, "_table_row_ids", None)
    if isinstance(row_ids, list) and 0 <= row < len(row_ids):
        return row_ids[row]

    first_item = table.item(row, 0)
    if first_item is not None:
        # Prefer Qt.UserRole because reference pages store the real row key there
        # (for example: "history::<id>" / "current::<product_id>::<supplier_id>").
        # Qt.UserRole + 1 is often used for sorting values, so using it first can
        # put a date/number into _pending_deletes instead of the DB row key.
        for role in (Qt.UserRole, Qt.UserRole + 1):
            value = first_item.data(role)
            if value is not None:
                return value

    text = _cell_text(table, row, 0).strip()
    if not text:
        return None
    try:
        return int(text)
    except ValueError:
        return text


def _qt_enum_int(value: Any) -> int:
    """Return a stable int for Qt enum values across PySide/PyQt versions."""
    if value is None:
        return 0
    raw = getattr(value, "value", value)
    try:
        return int(raw)
    except (TypeError, ValueError):
        return int(getattr(raw, "value", 0))


def _snapshot_row(owner: Any, table: QTableWidget, row: int) -> dict[str, Any]:
    cells: list[dict[str, Any]] = []
    for col in range(table.columnCount()):
        item = table.item(row, col)
        widget = table.cellWidget(row, col)
        cell: dict[str, Any] = {
            "text": item.text() if item else "",
            "flags": item.flags() if item else (Qt.ItemIsEnabled | Qt.ItemIsSelectable),
            "check_state": _qt_enum_int(item.checkState()) if item else None,
            "alignment": item.textAlignment() if item else (Qt.AlignLeft | Qt.AlignVCenter),
            "user_roles": {},
            "widget": None,
        }
        if item is not None:
            for role in (Qt.UserRole, Qt.UserRole + 1, Qt.UserRole + 2, Qt.UserRole + 3):
                value = item.data(role)
                if value is not None:
                    cell["user_roles"][_qt_enum_int(role)] = value
        if isinstance(widget, QComboBox):
            cell["widget"] = {
                "type": "combo",
                "items": [widget.itemText(i) for i in range(widget.count())],
                "current": widget.currentText(),
                "enabled": widget.isEnabled(),
                "properties": {
                    "row_id": widget.property("row_id"),
                    "edit_row_id": widget.property("edit_row_id"),
                    "combo_role": widget.property("combo_role"),
                },
            }
        elif isinstance(widget, QCheckBox):
            cell["widget"] = {"type": "check", "checked": widget.isChecked(), "enabled": widget.isEnabled()}
        cells.append(cell)
    return {"row": row, "key": _row_key(owner, table, row), "cells": cells}


def _restore_snapshot(owner: Any, table: QTableWidget, snapshot: dict[str, Any]) -> None:
    row = min(int(snapshot["row"]), table.rowCount())
    key = snapshot.get("key")
    row_ids = getattr(owner, "_table_row_ids", None)
    if isinstance(row_ids, list):
        row_ids.insert(row, key)

    table.insertRow(row)
    for col, cell in enumerate(snapshot["cells"]):
        item = QTableWidgetItem(cell.get("text", ""))
        item.setFlags(cell.get("flags", Qt.ItemIsEnabled | Qt.ItemIsSelectable))
        if cell.get("check_state") is not None:
            item.setCheckState(Qt.CheckState(cell["check_state"]))
        item.setTextAlignment(cell.get("alignment", Qt.AlignLeft | Qt.AlignVCenter))
        for role_int, value in cell.get("user_roles", {}).items():
            item.setData(Qt.ItemDataRole(_qt_enum_int(role_int)), value)
        table.setItem(row, col, item)

        widget_data = cell.get("widget")
        if widget_data and widget_data.get("type") == "combo":
            combo = QComboBox()
            combo.addItems(widget_data.get("items", []))
            idx = combo.findText(widget_data.get("current", ""))
            if idx >= 0:
                combo.setCurrentIndex(idx)
            combo.setEnabled(bool(widget_data.get("enabled", True)))
            for prop_name, prop_value in widget_data.get("properties", {}).items():
                if prop_value is not None:
                    combo.setProperty(prop_name, prop_value)
            table.setCellWidget(row, col, combo)
        elif widget_data and widget_data.get("type") == "check":
            check = QCheckBox()
            check.setChecked(bool(widget_data.get("checked", False)))
            check.setEnabled(bool(widget_data.get("enabled", True)))
            table.setCellWidget(row, col, check)


def delete_selected_rows_visual(owner: Any, table: QTableWidget, rows: list[int]) -> None:
    if not rows:
        return
    if not hasattr(owner, "_pending_deletes"):
        owner._pending_deletes = set()
    if not hasattr(owner, "_deleted_row_snapshots"):
        owner._deleted_row_snapshots = []

    rows = [row for row in rows if 0 <= row < table.rowCount()]
    snapshots = [_snapshot_row(owner, table, row) for row in rows]

    before_hook = getattr(owner, "before_standard_table_rows_deleted", None)
    if callable(before_hook):
        before_hook(table, rows, snapshots)

    owner._deleted_row_snapshots.extend(snapshots)

    new_rows = getattr(owner, "_new_rows", set())
    pending_changes = getattr(owner, "_pending_changes", None)
    for snapshot in snapshots:
        key = snapshot.get("key")
        if key is None:
            continue
        if isinstance(new_rows, set) and key in new_rows:
            new_rows.discard(key)
        else:
            owner._pending_deletes.add(key)
        if isinstance(pending_changes, dict):
            pending_changes.pop(key, None)

    row_ids = getattr(owner, "_table_row_ids", None)
    for row in sorted(rows, reverse=True):
        if isinstance(row_ids, list) and 0 <= row < len(row_ids):
            row_ids.pop(row)
        table.removeRow(row)

    after_hook = getattr(owner, "after_standard_table_rows_deleted", None)
    if callable(after_hook):
        after_hook(table, rows, snapshots)

    _message(owner, getattr(owner, "TABLE_DELETE_MESSAGE", DELETE_MESSAGE))


def undo_visual_delete(owner: Any, table: QTableWidget) -> None:
    snapshots = getattr(owner, "_deleted_row_snapshots", [])
    if not snapshots:
        _message(owner, "Нет строк для восстановления")
        return
    restored_snapshots = []
    while snapshots:
        snapshot = snapshots.pop()
        restored_snapshots.append(snapshot)
        key = snapshot.get("key")
        pending_deletes = getattr(owner, "_pending_deletes", None)
        if isinstance(pending_deletes, set) and key in pending_deletes:
            pending_deletes.discard(key)
        _restore_snapshot(owner, table, snapshot)

    after_hook = getattr(owner, "after_standard_table_rows_restored", None)
    if callable(after_hook):
        after_hook(table, restored_snapshots)

    _message(owner, "Удаление строк отменено")


def _viewport_position(table: QTableWidget, position: QPoint) -> QPoint:
    if table.indexAt(position).isValid():
        return position
    mapped = table.viewport().mapFrom(table, position)
    if table.indexAt(mapped).isValid():
        return mapped
    return position


def _apply_menu_style(menu: QMenu) -> None:
    menu.setStyleSheet(
        """
        QMenu {
            background-color: #fffaf4;
            color: #262626;
            border: 1px solid #f28223;
            padding: 4px;
            font: 9pt "Tahoma";
        }
        QMenu::item {
            color: #262626;
            background-color: transparent;
            padding: 6px 28px 6px 22px;
        }
        QMenu::item:selected {
            background-color: #f28223;
            color: #ffffff;
        }
        QMenu::item:disabled {
            color: #8a8a8a;
        }
        QMenu::separator {
            height: 1px;
            background: #d7b18c;
            margin: 4px 8px;
        }
        """
    )


def show_standard_table_context_menu(owner: Any, table: QTableWidget, position) -> None:
    configure_table_selection(table)
    viewport_pos = _viewport_position(table, position)
    index = table.indexAt(viewport_pos)

    if index.isValid() and index.row() not in _selected_rows(table):
        table.selectRow(index.row())
    rows = _selected_rows(table, index.row() if index.isValid() else None)

    menu = QMenu(table)
    _apply_menu_style(menu)

    copy_field_action = menu.addAction("Копировать поле")
    copy_rows_action = menu.addAction("Копировать строку/строки")
    menu.addSeparator()
    delete_action = menu.addAction("Удалить строку/строки")
    undo_delete_action = menu.addAction("Отменить удаление строки/строк")

    copy_field_action.setEnabled(index.isValid())
    copy_rows_action.setEnabled(bool(rows))
    delete_action.setEnabled(bool(rows))
    undo_delete_action.setEnabled(bool(getattr(owner, "_deleted_row_snapshots", [])))

    copy_field_action.triggered.connect(lambda: _copy_field(table, index.row(), index.column()))
    copy_rows_action.triggered.connect(lambda: _copy_rows(table, rows))
    delete_action.triggered.connect(lambda: delete_selected_rows_visual(owner, table, rows))
    undo_delete_action.triggered.connect(lambda: undo_visual_delete(owner, table))

    menu.exec(table.viewport().mapToGlobal(viewport_pos))


def install_standard_table_context_menu(owner: Any, table: QTableWidget) -> None:
    table.setContextMenuPolicy(Qt.CustomContextMenu)
    configure_table_selection(table)

    # Disconnect only the handler installed by this helper.
    # Calling disconnect() without a previously connected slot can emit a
    # RuntimeWarning in PyQt/PySide, even when wrapped in try/except.
    old_handler = getattr(table, "_standard_context_menu_handler", None)
    if old_handler is not None:
        try:
            table.customContextMenuRequested.disconnect(old_handler)
        except (TypeError, RuntimeError):
            pass

    handler = lambda position: show_standard_table_context_menu(owner, table, position)
    table._standard_context_menu_handler = handler
    table.customContextMenuRequested.connect(handler)


def apply_pending_table_deletes_to_db(session: Any, owner: Any) -> None:
    """Apply pending visual deletes for temp/import pages that save via services.

    Regular reference pages already delete using their own apply_pending_changes().
    This helper is intentionally limited to pages that call it explicitly before
    calculation/saving.
    """
    pending_deletes = set(getattr(owner, "_pending_deletes", set()) or set())
    if not pending_deletes:
        return

    cls_name = owner.__class__.__name__
    if cls_name == "CustomerCostsPage":
        from app.db.models import TempCustomerCostImport, TempCustomerCostOption
        session.query(TempCustomerCostOption).filter(
            TempCustomerCostOption.temp_import_id.in_(pending_deletes)
        ).delete(synchronize_session=False)
        session.query(TempCustomerCostImport).filter(
            TempCustomerCostImport.id.in_(pending_deletes)
        ).delete(synchronize_session=False)
    elif cls_name == "TargetPricesPage":
        from app.db.models import TempTargetPriceImport, TempTargetPriceOption
        session.query(TempTargetPriceOption).filter(
            TempTargetPriceOption.temp_import_id.in_(pending_deletes)
        ).delete(synchronize_session=False)
        session.query(TempTargetPriceImport).filter(
            TempTargetPriceImport.id.in_(pending_deletes)
        ).delete(synchronize_session=False)
    elif cls_name == "ProductStockPage":
        model_map = getattr(owner, "MODEL_BY_MODE", None)
        model = None
        if isinstance(model_map, dict):
            model = model_map.get(getattr(owner, "_mode", None))
        if model is None:
            from app.db.models import TempIsImport, TempStockImport, TempSupplierOrdersImport
            mode = getattr(owner, "_mode", None)
            if mode == "stock":
                model = TempStockImport
            elif mode == "orders":
                model = TempSupplierOrdersImport
            else:
                model = TempIsImport
        session.query(model).filter(model.id.in_(pending_deletes)).delete(synchronize_session=False)
    else:
        # For pages not known here, deletion is handled in apply_pending_changes().
        return

    pending_deletes.clear()
    snapshots = getattr(owner, "_deleted_row_snapshots", None)
    if isinstance(snapshots, list):
        snapshots.clear()
