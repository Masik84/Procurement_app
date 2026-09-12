from __future__ import annotations

import logging
from datetime import date, datetime, time
from pathlib import Path
from typing import Any, Sequence

from PySide6.QtCore import QFile, QDate, Qt, QTimer, QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtUiTools import QUiLoader
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QComboBox,
    QFileDialog,
    QMessageBox,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from app.db.db import SessionLocal
from app.db.models import Product, ProductUc3History
from app.exports.product_uc3_exporter import ProductUc3Exporter
from app.imports.product_uc3_importer import ProductUc3Importer
from app.services.product_uc3_service import ProductUc3Service
from app.ui.table_style import resize_columns_for_multiline_headers, setup_data_table
from app.utils.checked_filter_dialog import CheckedFilterDialog, FilterOption
from app.utils.gui_table_actions import commit_active_table_item_editors
from app.utils.text import clean_multi_spaces, normalize_product_name
from app.workers.excel_export_worker import start_excel_export


logger = logging.getLogger(__name__)

BASE_DIR = Path(__file__).resolve().parents[2]
UI_PATH = BASE_DIR / "app" / "ui" / "windows" / "product_uc3.ui"

COL_PRODUCT = 0
COL_TARGET = 1
COL_WA = 2
COL_DATE = 3


class SortableTableWidgetItem(QTableWidgetItem):
    def __lt__(self, other):
        if isinstance(other, QTableWidgetItem):
            left = self.data(Qt.UserRole + 1)
            right = other.data(Qt.UserRole + 1)
            if left is not None and right is not None:
                try:
                    return left < right
                except Exception:
                    pass
        return super().__lt__(other)


def load_ui(ui_path: Path):
    loader = QUiLoader()
    ui_file = QFile(str(ui_path))
    if not ui_file.open(QFile.ReadOnly):
        raise RuntimeError(f"Не удалось открыть UI: {ui_path}")
    try:
        widget = loader.load(ui_file)
    finally:
        ui_file.close()
    if widget is None:
        raise RuntimeError(f"Не удалось загрузить UI: {ui_path}")
    return widget


class ProductUc3Page(QWidget):
    HEADERS = ["Product Name", "Target uC3", "Walk-Away uC3", "Change date"]
    TABLE_DELETE_MESSAGE = "Полное удаление строк будет сделано при сохранении"

    def __init__(self):
        super().__init__()
        self.ui = load_ui(UI_PATH)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.ui)

        self.table = self.ui.table
        self._updating_table = False
        self._updating_filters = False
        self._pending_changes: dict[str, dict[str, Any]] = {}
        self._pending_deletes: set[str] = set()
        self._deleted_row_snapshots: list[dict[str, Any]] = []
        self._new_rows: set[str] = set()
        self._temp_row_id = -1
        self._selected_brand_values: set[str] | None = None
        self._selected_family_values: set[str] | None = None
        self._selected_product_ids: set[int] | None = None
        self._date_filter_changed = False
        self._excel_export_thread = None
        self._excel_export_worker = None

        self.setup_ui()
        self.setup_connections()
        self.load_find_brands()
        self._refresh_filter_buttons(prune=True)
        self.find_rows()

    def get_session(self):
        return SessionLocal()

    def setup_ui(self):
        setup_data_table(self.table, sorting=True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.setColumnCount(len(self.HEADERS))
        self.table.setHorizontalHeaderLabels(self.HEADERS)
        self.table.horizontalHeader().setVisible(True)

        for date_edit in (self.ui.line_Start_date, self.ui.line_End_date):
            date_edit.setCalendarPopup(True)
            date_edit.setDisplayFormat("dd.MM.yyyy")
            date_edit.setSpecialValueText("")
        today = QDate.currentDate()
        self.ui.line_Start_date.setDate(today)
        self.ui.line_End_date.setDate(today)
        self.ui.chb_CalcToday.setChecked(True)

        from app.utils.gui_table_actions import install_standard_table_context_menu
        install_standard_table_context_menu(self, self.table)
        self.clear_message()

    def setup_connections(self):
        self.table.cellDoubleClicked.connect(self.start_cell_edit)
        self.table.itemChanged.connect(self.on_item_changed)

        self.ui.cbo_FindBrand.currentTextChanged.connect(self.refresh_current_product_combo)
        self.ui.line_FindProduct.textChanged.connect(self.refresh_current_product_combo)

        self.ui.btn_FilterBrand.clicked.connect(self.open_brand_filter)
        self.ui.btn_FilterProductFamily.clicked.connect(self.open_family_filter)
        self.ui.btn_FilterProduct.clicked.connect(self.open_product_filter)
        self.ui.line_NameSearch.textChanged.connect(self.on_name_search_changed)
        self.ui.line_Start_date.dateChanged.connect(self.on_date_changed)
        self.ui.line_End_date.dateChanged.connect(self.on_date_changed)
        self.ui.chb_CalcToday.stateChanged.connect(self.on_current_changed)

        self.ui.btn_Search.clicked.connect(self.find_rows)
        self.ui.btn_AddLine.clicked.connect(self.add_line)
        self.ui.btn_DownFile.clicked.connect(self.download_template)
        self.ui.btn_Import.clicked.connect(self.import_excel)
        self.ui.btn_SaveExcel.clicked.connect(self.save_excel)
        self.ui.btn_Save.clicked.connect(self.apply_pending_changes)

    def show_message(self, text: str):
        self.ui.label_msg.setText(text or "Сообщений нет")
        self.ui.label_msg.setVisible(True)

    def clear_message(self):
        self.ui.label_msg.setText("Сообщений нет")
        self.ui.label_msg.setVisible(True)

    def show_error_message(self, text: str):
        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Warning)
        msg.setWindowTitle("Ошибка")
        msg.setTextFormat(Qt.PlainText)
        msg.setText(str(text or "Неизвестная ошибка"))
        copy_btn = msg.addButton("Copy", QMessageBox.ActionRole)
        msg.addButton(QMessageBox.Ok)
        msg.exec()
        if msg.clickedButton() == copy_btn:
            QApplication.clipboard().setText(str(text or ""))

    # ------------------------------------------------------------------
    # Product editor search in the TOP panel. This intentionally does not
    # filter the table; it only filters the product combo opened in a cell.
    # ------------------------------------------------------------------
    def load_find_brands(self):
        current = clean_multi_spaces(self.ui.cbo_FindBrand.currentText())
        with self.get_session() as session:
            brands = [
                row[0]
                for row in (
                    session.query(Product.brand)
                    .filter(Product.brand.isnot(None), Product.brand != "")
                    .distinct()
                    .order_by(Product.brand.asc())
                    .all()
                )
                if row[0]
            ]
        self.ui.cbo_FindBrand.blockSignals(True)
        self.ui.cbo_FindBrand.clear()
        self.ui.cbo_FindBrand.addItem("-")
        self.ui.cbo_FindBrand.addItems(brands)
        if current and self.ui.cbo_FindBrand.findText(current) >= 0:
            self.ui.cbo_FindBrand.setCurrentText(current)
        self.ui.cbo_FindBrand.blockSignals(False)

    def _get_filtered_products_for_editor(self) -> list[Product]:
        brand_filter = clean_multi_spaces(self.ui.cbo_FindBrand.currentText())
        text_filter = clean_multi_spaces(self.ui.line_FindProduct.text())
        with self.get_session() as session:
            query = session.query(Product).filter(Product.name.isnot(None), Product.name != "")
            if brand_filter and brand_filter != "-":
                query = query.filter(Product.brand == brand_filter)
            if text_filter:
                query = query.filter(Product.name.ilike(f"%{text_filter}%"))
            return query.order_by(Product.name.asc()).all()

    def refresh_current_product_combo(self):
        row = self.table.currentRow()
        if row < 0 or row >= self.table.rowCount():
            return
        combo = self.table.cellWidget(row, COL_PRODUCT)
        if isinstance(combo, QComboBox) and combo.property("combo_role") == "product_combo":
            self.populate_product_combo(combo, keep_current=True)

    def _row_key_at(self, row: int) -> str | None:
        if row < 0 or row >= self.table.rowCount():
            return None
        for col in range(self.table.columnCount()):
            widget = self.table.cellWidget(row, col)
            if widget is not None:
                value = widget.property("row_key")
                if value is not None:
                    return str(value)
            item = self.table.item(row, col)
            if item is not None:
                value = item.data(Qt.UserRole)
                if value is not None:
                    return str(value)
        return None

    def _find_table_row_by_key(self, row_key: str) -> int:
        for row in range(self.table.rowCount()):
            if self._row_key_at(row) == row_key:
                return row
        return -1

    def _row_product_id(self, row_key: str) -> int | None:
        pending = self._pending_changes.get(row_key, {})
        if "product_id" in pending:
            return pending["product_id"]
        if row_key.startswith("db::"):
            row_id = int(row_key.split("::", 1)[1])
            with self.get_session() as session:
                obj = session.query(ProductUc3History).filter(ProductUc3History.id == row_id).first()
                return int(obj.product_id) if obj else None
        return None

    def populate_product_combo(self, combo: QComboBox, *, keep_current: bool, selected_product_id: int | None = None):
        current_id = combo.currentData() if keep_current else selected_product_id
        try:
            current_id = int(current_id) if current_id is not None else None
        except Exception:
            current_id = None
        products = self._get_filtered_products_for_editor()
        combo.blockSignals(True)
        combo.clear()
        combo.addItem("", None)
        for product in products:
            combo.addItem(product.name, int(product.id))
        if current_id is not None and combo.findData(current_id) < 0:
            with self.get_session() as session:
                product = session.query(Product).filter(Product.id == current_id).first()
                if product:
                    combo.addItem(product.name, int(product.id))
        idx = combo.findData(current_id)
        combo.setCurrentIndex(idx if idx >= 0 else 0)
        combo.blockSignals(False)

    def start_cell_edit(self, row: int, column: int):
        if self._updating_table or row < 0 or row >= self.table.rowCount():
            return
        if column != COL_PRODUCT:
            return
        row_key = self._row_key_at(row)
        if row_key is None:
            return
        combo = QComboBox()
        combo.setProperty("combo_role", "product_combo")
        combo.setProperty("row_key", row_key)
        self.populate_product_combo(combo, keep_current=False, selected_product_id=self._row_product_id(row_key))
        combo.activated.connect(lambda _i, r=row, key=row_key, c=combo: self.finish_product_edit(r, key, c))
        self.table.setCellWidget(row, column, combo)
        combo.setFocus()
        QTimer.singleShot(0, combo.showPopup)

    def finish_product_edit(self, row: int, row_key: str, combo: QComboBox):
        product_id = combo.currentData()
        try:
            product_id = int(product_id) if product_id is not None else None
        except Exception:
            product_id = None
        product_name = clean_multi_spaces(combo.currentText())
        self._pending_changes.setdefault(row_key, {})["product_id"] = product_id
        self._pending_changes[row_key]["product_name"] = product_name
        self._updating_table = True
        try:
            self.table.removeCellWidget(row, COL_PRODUCT)
            self.table.setItem(row, COL_PRODUCT, self._build_item(product_name, row_key, align_left=True, editable=False))
        finally:
            self._updating_table = False

    # ------------------------------------------------------------------
    # LEFT panel filters for the displayed table.
    # ------------------------------------------------------------------
    def _open_checked_filter(self, *, title: str, options: Sequence[FilterOption], selected):
        dialog = CheckedFilterDialog(self, title=title, options=options, selected_keys=selected)
        return dialog.exec_and_get_selection()

    def _base_product_query(self, session, *, ignore: str | None = None):
        query = session.query(Product).join(ProductUc3History, ProductUc3History.product_id == Product.id)
        if ignore != "brand" and self._selected_brand_values is not None:
            query = query.filter(Product.brand.in_(sorted(self._selected_brand_values)))
        if ignore != "family" and self._selected_family_values is not None:
            query = query.filter(Product.family.in_(sorted(self._selected_family_values)))
        if ignore != "product" and self._selected_product_ids is not None:
            query = query.filter(Product.id.in_(sorted(self._selected_product_ids)))
        name_search = clean_multi_spaces(self.ui.line_NameSearch.text())
        if name_search:
            query = query.filter(Product.name.ilike(f"%{name_search}%"))
        return query.distinct()

    def _brand_options(self) -> list[FilterOption]:
        with self.get_session() as session:
            rows = (
                self._base_product_query(session, ignore="brand")
                .with_entities(Product.brand)
                .filter(Product.brand.isnot(None), Product.brand != "")
                .distinct()
                .order_by(Product.brand.asc())
                .all()
            )
        return [FilterOption(key=value, label=value, search_text=value) for (value,) in rows if value]

    def _family_options(self) -> list[FilterOption]:
        with self.get_session() as session:
            rows = (
                self._base_product_query(session, ignore="family")
                .with_entities(Product.family)
                .filter(Product.family.isnot(None), Product.family != "")
                .distinct()
                .order_by(Product.family.asc())
                .all()
            )
        return [FilterOption(key=value, label=value, search_text=value) for (value,) in rows if value]

    def _product_options(self) -> list[FilterOption]:
        with self.get_session() as session:
            rows = (
                self._base_product_query(session, ignore="product")
                .with_entities(Product.id, Product.name, Product.brand, Product.family, Product.pack)
                .order_by(Product.name.asc())
                .all()
            )
        result = []
        for product_id, name, brand, family, pack in rows:
            if product_id is None or not name:
                continue
            search = " ".join(str(value or "") for value in (product_id, name, brand, family, pack))
            result.append(FilterOption(key=int(product_id), label=str(name), search_text=search))
        return result

    def open_brand_filter(self):
        accepted, selected = self._open_checked_filter(
            title="Фильтр по брендам", options=self._brand_options(), selected=self._selected_brand_values
        )
        if accepted:
            self._selected_brand_values = None if selected is None else {str(value) for value in selected}
            self._refresh_filter_buttons(prune=True)

    def open_family_filter(self):
        accepted, selected = self._open_checked_filter(
            title="Фильтр по Product Family", options=self._family_options(), selected=self._selected_family_values
        )
        if accepted:
            self._selected_family_values = None if selected is None else {str(value) for value in selected}
            self._refresh_filter_buttons(prune=True)

    def open_product_filter(self):
        accepted, selected = self._open_checked_filter(
            title="Фильтр по продуктам", options=self._product_options(), selected=self._selected_product_ids
        )
        if accepted:
            self._selected_product_ids = None if selected is None else {int(value) for value in selected}
            self._refresh_filter_buttons(prune=True)

    def _refresh_filter_buttons(self, *, prune: bool = False):
        if prune:
            try:
                if self._selected_brand_values is not None:
                    available = {str(o.key) for o in self._brand_options()}
                    self._selected_brand_values &= available
                if self._selected_family_values is not None:
                    available = {str(o.key) for o in self._family_options()}
                    self._selected_family_values &= available
                if self._selected_product_ids is not None:
                    available = {int(o.key) for o in self._product_options()}
                    self._selected_product_ids &= available
            except Exception:
                logger.exception("Не удалось обновить фильтры Target uC3")
        self.ui.btn_FilterBrand.setText(
            "все Бренды" if self._selected_brand_values is None else f"все Бренды ({len(self._selected_brand_values)})"
        )
        self.ui.btn_FilterProductFamily.setText(
            "все Product Family" if self._selected_family_values is None else f"все Product Family ({len(self._selected_family_values)})"
        )
        self.ui.btn_FilterProduct.setText(
            "все Продукты" if self._selected_product_ids is None else f"все Продукты ({len(self._selected_product_ids)})"
        )

    def on_name_search_changed(self):
        if not self._updating_filters:
            self._refresh_filter_buttons(prune=True)

    def on_date_changed(self):
        if not self._updating_filters:
            self._date_filter_changed = True

    def on_current_changed(self):
        if self._updating_filters:
            return
        self.find_rows()

    @staticmethod
    def _qdate_to_date(value: QDate) -> date:
        return date(value.year(), value.month(), value.day())

    def _history_period(self) -> tuple[datetime, datetime] | None:
        if self.ui.chb_CalcToday.isChecked():
            return None
        start_q = self.ui.line_Start_date.date()
        end_q = self.ui.line_End_date.date()
        today = QDate.currentDate()
        if not self._date_filter_changed and start_q == today and end_q == today:
            return None
        start = self._qdate_to_date(start_q)
        end = self._qdate_to_date(end_q)
        if start > end:
            raise ValueError("Дата начала периода не может быть больше даты окончания периода")
        return datetime.combine(start, time.min), datetime.combine(end, time.max)

    def _query_rows(self) -> list[tuple[ProductUc3History, Product]]:
        current_only = self.ui.chb_CalcToday.isChecked()
        period = self._history_period()
        with self.get_session() as session:
            query = (
                session.query(ProductUc3History, Product)
                .join(Product, Product.id == ProductUc3History.product_id)
            )
            if self._selected_brand_values is not None:
                query = query.filter(Product.brand.in_(sorted(self._selected_brand_values)))
            if self._selected_family_values is not None:
                query = query.filter(Product.family.in_(sorted(self._selected_family_values)))
            if self._selected_product_ids is not None:
                query = query.filter(Product.id.in_(sorted(self._selected_product_ids)))
            name_search = clean_multi_spaces(self.ui.line_NameSearch.text())
            if name_search:
                query = query.filter(Product.name.ilike(f"%{name_search}%"))
            if period is not None:
                start_dt, end_dt = period
                query = query.filter(
                    ProductUc3History.change_date >= start_dt,
                    ProductUc3History.change_date <= end_dt,
                )
            rows = query.order_by(
                Product.name.asc(),
                ProductUc3History.change_date.desc(),
                ProductUc3History.id.desc(),
            ).all()

        if not current_only:
            return rows
        result = []
        seen: set[int] = set()
        for history, product in rows:
            pid = int(history.product_id)
            if pid in seen:
                continue
            seen.add(pid)
            result.append((history, product))
        return result

    def find_rows(self):
        try:
            rows = self._query_rows()
            data = [
                {
                    "row_key": f"db::{history.id}",
                    "product_id": int(history.product_id),
                    "product_name": product.name or "",
                    "target_uc3": history.target_uc3,
                    "walk_away_uc3": history.walk_away_uc3,
                    "change_date": history.change_date,
                }
                for history, product in rows
            ]
            self._pending_changes.clear()
            self._pending_deletes.clear()
            self._deleted_row_snapshots.clear()
            self._new_rows.clear()
            self._populate_table(data)
            self.show_message(f"Найдено строк: {len(data)}")
        except Exception as exc:
            self.show_error_message(str(exc))

    def _build_item(self, value, row_key: str, *, editable: bool, align_left: bool = False, sort_value=None):
        if isinstance(value, datetime):
            text = value.strftime("%d.%m.%Y")
        elif value is None:
            text = ""
        else:
            text = str(value)
        item = SortableTableWidgetItem(text)
        flags = Qt.ItemIsEnabled | Qt.ItemIsSelectable
        if editable:
            flags |= Qt.ItemIsEditable
        item.setFlags(flags)
        item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter if align_left else Qt.AlignCenter)
        item.setData(Qt.UserRole, row_key)
        item.setData(Qt.UserRole + 1, value if sort_value is None else sort_value)
        return item

    def _populate_table(self, rows: list[dict]):
        self._updating_table = True
        self.table.setSortingEnabled(False)
        try:
            self.table.clearContents()
            self.table.setRowCount(len(rows))
            self.table.setColumnCount(len(self.HEADERS))
            self.table.setHorizontalHeaderLabels(self.HEADERS)
            for row_index, data in enumerate(rows):
                row_key = str(data["row_key"])
                self.table.setItem(row_index, COL_PRODUCT, self._build_item(data.get("product_name", ""), row_key, editable=False, align_left=True))
                self.table.setItem(row_index, COL_TARGET, self._build_item(data.get("target_uc3"), row_key, editable=True))
                self.table.setItem(row_index, COL_WA, self._build_item(data.get("walk_away_uc3"), row_key, editable=True))
                self.table.setItem(row_index, COL_DATE, self._build_item(data.get("change_date"), row_key, editable=False))
            self.table.resizeColumnsToContents()
            resize_columns_for_multiline_headers(self.table)
            if self.table.columnCount() >= 4:
                self.table.setColumnWidth(COL_PRODUCT, max(self.table.columnWidth(COL_PRODUCT), 300))
                self.table.setColumnWidth(COL_TARGET, max(self.table.columnWidth(COL_TARGET), 100))
                self.table.setColumnWidth(COL_WA, max(self.table.columnWidth(COL_WA), 120))
                self.table.setColumnWidth(COL_DATE, max(self.table.columnWidth(COL_DATE), 100))
        finally:
            self.table.setSortingEnabled(True)
            self._updating_table = False

    def on_item_changed(self, item: QTableWidgetItem):
        if self._updating_table:
            return
        col = item.column()
        if col not in {COL_TARGET, COL_WA}:
            return
        key = item.data(Qt.UserRole)
        if key is None:
            return
        key = str(key)
        field = "target_uc3" if col == COL_TARGET else "walk_away_uc3"
        self._pending_changes.setdefault(key, {})[field] = item.text().strip()

    def add_line(self):
        key = f"new::{self._temp_row_id}"
        self._temp_row_id -= 1
        self._new_rows.add(key)
        self._pending_changes[key] = {
            "product_id": None,
            "product_name": "",
            "target_uc3": None,
            "walk_away_uc3": None,
        }
        self.table.setSortingEnabled(False)
        row = self.table.rowCount()
        self.table.insertRow(row)
        self._updating_table = True
        try:
            self.table.setItem(row, COL_PRODUCT, self._build_item("", key, editable=False, align_left=True))
            self.table.setItem(row, COL_TARGET, self._build_item(None, key, editable=True))
            self.table.setItem(row, COL_WA, self._build_item(None, key, editable=True))
            self.table.setItem(row, COL_DATE, self._build_item(None, key, editable=False))
        finally:
            self._updating_table = False
            self.table.setSortingEnabled(True)
        row = self._find_table_row_by_key(key)
        if row >= 0:
            self.table.setCurrentCell(row, COL_PRODUCT)
            self.start_cell_edit(row, COL_PRODUCT)
        self.show_message("Добавлена строка")

    def _row_display_values(self, table_row: int, row_key: str) -> dict:
        pending = dict(self._pending_changes.get(row_key, {}))
        if "product_id" not in pending:
            pending["product_id"] = self._row_product_id(row_key)
        if "product_name" not in pending:
            item = self.table.item(table_row, COL_PRODUCT)
            pending["product_name"] = item.text().strip() if item else ""
        if "target_uc3" not in pending:
            item = self.table.item(table_row, COL_TARGET)
            pending["target_uc3"] = item.text().strip() if item else None
        if "walk_away_uc3" not in pending:
            item = self.table.item(table_row, COL_WA)
            pending["walk_away_uc3"] = item.text().strip() if item else None
        return pending

    def apply_pending_changes(self):
        commit_active_table_item_editors(self.table)
        try:
            with self.get_session() as session:
                service = ProductUc3Service(session)
                # Delete explicitly requested historical rows first.
                for key in list(self._pending_deletes):
                    if isinstance(key, str) and key.startswith("db::"):
                        row_id = int(key.split("::", 1)[1])
                        session.query(ProductUc3History).filter(ProductUc3History.id == row_id).delete(synchronize_session=False)
                session.flush()

                created = 0
                unchanged = 0
                reassigned = 0
                for table_row in range(self.table.rowCount()):
                    key = self._row_key_at(table_row)
                    if key is None or key in self._pending_deletes:
                        continue
                    if key not in self._new_rows and key not in self._pending_changes:
                        continue
                    values = self._row_display_values(table_row, key)
                    product_id = values.get("product_id")
                    if not product_id:
                        raise ValueError(f"Строка {table_row + 1}: выберите Product Name")
                    product_id = int(product_id)

                    changes = self._pending_changes.get(key, {})
                    if key.startswith("db::"):
                        # Product editing is a correction of the historical row itself,
                        # matching the Price History page. It must not create a new
                        # dated uC3 value by itself.
                        row_id = int(key.split("::", 1)[1])
                        history_row = (
                            session.query(ProductUc3History)
                            .filter(ProductUc3History.id == row_id)
                            .first()
                        )
                        if history_row is None:
                            continue
                        if int(history_row.product_id) != product_id:
                            history_row.product_id = product_id
                            session.flush()
                            reassigned += 1

                        # Target/WA edits are immutable-history changes: compare
                        # against the product's current values and append only when
                        # the truncated values really changed.
                        if not ({"target_uc3", "walk_away_uc3"} & set(changes)):
                            continue

                    _obj, was_created = service.save_values(
                        product_id=product_id,
                        target_uc3=values.get("target_uc3"),
                        walk_away_uc3=values.get("walk_away_uc3"),
                    )
                    if was_created:
                        created += 1
                    else:
                        unchanged += 1
                session.commit()

            self._pending_changes.clear()
            self._pending_deletes.clear()
            self._deleted_row_snapshots.clear()
            self._new_rows.clear()
            self.find_rows()
            parts = [f"Сохранено новых значений: {created}", f"без изменений: {unchanged}"]
            if reassigned:
                parts.append(f"изменён Product: {reassigned}")
            self.show_message("; ".join(parts))
        except Exception as exc:
            self.show_error_message(str(exc))

    def download_template(self):
        try:
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "Сохранить шаблон",
                str(BASE_DIR / "Target_uC3_ImportTemplate.xlsx"),
                "Excel files (*.xlsx)",
            )
            if not file_path:
                return
            ProductUc3Exporter().export_template(file_path)
            self.show_message("Шаблон сохранен")
        except Exception as exc:
            self.show_error_message(str(exc))

    def _find_product_for_import(self, session, product_name: str) -> Product | None:
        name = clean_multi_spaces(product_name)
        if not name:
            return None
        product = session.query(Product).filter(Product.name == name).first()
        if product:
            return product
        normalized = normalize_product_name(name)
        if not normalized:
            return None
        for candidate in session.query(Product).filter(Product.name.isnot(None)).all():
            if normalize_product_name(candidate.name) == normalized:
                return candidate
        return None

    def import_excel(self):
        try:
            file_path, _ = QFileDialog.getOpenFileName(
                self,
                "Выберите файл Target uC3",
                "",
                "Excel files (*.xls *.xlsx)",
            )
            if not file_path:
                return
            rows = ProductUc3Importer().read_excel(file_path)
            if not rows:
                self.show_message("Нет строк для импорта")
                return

            preview = []
            missing = []
            self._pending_changes.clear()
            self._pending_deletes.clear()
            self._deleted_row_snapshots.clear()
            self._new_rows.clear()
            with self.get_session() as session:
                for source in rows:
                    key = f"new::{self._temp_row_id}"
                    self._temp_row_id -= 1
                    product = self._find_product_for_import(session, source.get("product_name", ""))
                    product_id = int(product.id) if product else None
                    product_name = product.name if product else clean_multi_spaces(source.get("product_name"))
                    if product is None:
                        missing.append(product_name)
                    values = {
                        "product_id": product_id,
                        "product_name": product_name,
                        "target_uc3": source.get("target_uc3"),
                        "walk_away_uc3": source.get("walk_away_uc3"),
                    }
                    self._new_rows.add(key)
                    self._pending_changes[key] = values
                    preview.append({"row_key": key, **values, "change_date": None})

            self._populate_table(preview)
            if missing:
                self.show_message(
                    f"Загружено строк: {len(preview)}. Не сопоставлено продуктов: {len(missing)}. "
                    "Выберите Product Name в таблице и нажмите Сохранить."
                )
            else:
                self.show_message(f"Загружено строк: {len(preview)}. Нажмите Сохранить для записи в БД.")
        except Exception as exc:
            self.show_error_message(str(exc))

    def _rows_for_excel(self) -> list[dict]:
        rows = []
        for table_row in range(self.table.rowCount()):
            key = self._row_key_at(table_row)
            if key is None or key in self._pending_deletes:
                continue
            values = self._row_display_values(table_row, key)
            date_item = self.table.item(table_row, COL_DATE)
            change_date = date_item.text().strip() if date_item else ""
            rows.append(
                {
                    "Product Name": values.get("product_name", ""),
                    "Target uC3": ProductUc3Service.truncate_integer(values.get("target_uc3")),
                    "Walk-Away uC3": ProductUc3Service.truncate_integer(values.get("walk_away_uc3")),
                    "Change date": change_date,
                }
            )
        return rows

    def save_excel(self):
        try:
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "Сохранить Target uC3",
                str(BASE_DIR / "Target_uC3.xlsx"),
                "Excel files (*.xlsx)",
            )
            if not file_path:
                return
            rows = self._rows_for_excel()

            def do_export():
                return ProductUc3Exporter().export_rows(rows, file_path)

            def done(output_path):
                QDesktopServices.openUrl(QUrl.fromLocalFile(str(output_path)))
                self.show_message("Файл сохранен")

            if not start_excel_export(self, do_export, on_finished=done, on_error=lambda text: self.show_error_message(str(text))):
                self.show_message("Excel файл уже формируется. Можно продолжать работать в программе.")
        except Exception as exc:
            self.show_error_message(str(exc))
