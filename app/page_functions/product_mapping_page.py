from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any, Sequence

import pandas as pd
from PySide6.QtCore import QFile, Qt, QTimer
from PySide6.QtGui import QDesktopServices
from PySide6.QtUiTools import QUiLoader
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QHBoxLayout,
    QFileDialog,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from app.db.db import SessionLocal
from app.db.models import Product
from app.imports.product_mapping_portfolio_importer import ProductMappingPortfolioImporter
from app.services.product_mapping_service import ProductMappingService
from app.services.product_matching_service import MissingPackTypeError
from app.ui.table_style import format_table_field_value, resize_columns_for_multiline_headers, setup_data_table
from app.utils.checked_filter_dialog import CheckedFilterDialog, FilterOption
from app.utils.message_dialogs import show_error
from app.utils.pack_type_prompt import ask_pack_type
from app.utils.gui_table_actions import install_standard_table_context_menu
from app.utils.parsers import parse_loose_number
from app.utils.text import clean_multi_spaces


BASE_DIR = Path(__file__).resolve().parents[2]
UI_PATH = BASE_DIR / "app" / "ui" / "windows" / "product_mapping.ui"


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


class ProductMappingPage(QWidget):
    TABLE_DELETE_MESSAGE = "Удаление сопоставлений будет выполнено после Сохранить"

    COLUMNS = [
        "sales_code",
        "sales_article",
        "sales_product_name",
        "sales_pack",
        "sales_qty_in_box",
        "sales_brand",
        "sales_is_excise",
        "status",
        "product_name",
        "product_qty_in_box",
        "new_product_name",
        "new_brand",
        "new_pack",
        "new_qty_in_box",
        "new_is_excise",
    ]
    HEADERS = [
        "Код",
        "Артикул",
        "Product Name from Portfolio",
        "Упаковка",
        "Кол-во в упак",
        "Brand from Portfolio",
        "Акциз",
        "Статус",
        "Our Product Name",
        "Qty in Box",
        "Product name (for new)",
        "Brand (for new)",
        "Pack (for new)",
        "Qty in Box (for new)",
        "Excise duty (for new)",
    ]
    EDITABLE_FIELDS = {"new_product_name", "new_brand", "new_pack", "new_qty_in_box"}
    NUMERIC_FIELDS = {"sales_pack", "sales_qty_in_box", "product_qty_in_box", "new_pack", "new_qty_in_box"}
    COL_PRODUCT = COLUMNS.index("product_name")
    COL_NEW_EXCISE = COLUMNS.index("new_is_excise")

    def __init__(self):
        super().__init__()
        self.ui = load_ui(UI_PATH)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.ui)

        self.table = self.ui.table
        self._all_rows: list[dict] = []
        self._rows: list[dict] = []
        self._mode = "search"
        self._updating_table = False
        self._selected_brand_values: set[str] | None = None
        self._selected_family_values: set[str] | None = None
        self._selected_product_ids: set[int] | None = None

        # Left-side filter controls are only applied after Search, like on the
        # other reference pages.  These values are a snapshot of the filters
        # that were actually applied to the table.
        self._applied_brand_values: set[str] | None = None
        self._applied_family_values: set[str] | None = None
        self._applied_product_ids: set[int] | None = None
        self._applied_name_search: str = ""
        self._product_meta: dict[int, tuple[str, str, str]] = {}
        self._pending_deletes: set[str] = set()
        self._deleted_row_snapshots: list[dict] = []
        self._visually_deleted_codes: set[str] = set()
        self._show_ignored = False
        self._portfolio_df: pd.DataFrame | None = None

        self.setup_ui()
        self.setup_connections()
        self.load_find_brands()
        self._load_product_meta()
        self._refresh_filter_buttons()
        self.show_message("Выберите файл Портфель или нажмите Search")

    def get_session(self):
        return SessionLocal()

    def service(self, session) -> ProductMappingService:
        return ProductMappingService(session)

    def setup_ui(self):
        setup_data_table(self.table, sorting=True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.setColumnCount(len(self.HEADERS))
        self.table.setHorizontalHeaderLabels(self.HEADERS)
        self.table.horizontalHeader().setVisible(True)
        install_standard_table_context_menu(self, self.table)
        self.clear_message()

    def setup_connections(self):
        self.ui.btn_Search.clicked.connect(self.search_saved)
        self.ui.btn_Import.clicked.connect(self.import_portfolio)
        self.ui.btn_ResetAll.clicked.connect(self.reset_all)
        self.ui.btn_Save.clicked.connect(self.save)
        self.ui.btn_SaveExcel.clicked.connect(self.save_excel)

        self.ui.btn_FilterBrand.clicked.connect(self.open_brand_filter)
        self.ui.btn_FilterProductFamily.clicked.connect(self.open_family_filter)
        self.ui.btn_FilterProduct.clicked.connect(self.open_product_filter)

        self.table.cellDoubleClicked.connect(self.start_product_edit)
        self.table.itemChanged.connect(self.on_item_changed)
        self.ui.cbo_FindBrand.currentTextChanged.connect(self.refresh_current_product_combo)
        self.ui.line_FindProduct.textChanged.connect(self.refresh_current_product_combo)

    def show_message(self, text: str):
        self.ui.label_msg.setText(text or "Сообщений нет")
        self.ui.label_msg.setVisible(True)

    def clear_message(self):
        self.ui.label_msg.setText("Сообщений нет")
        self.ui.label_msg.setVisible(True)

    def show_error_message(self, text: str):
        show_error(self, text)

    # ------------------------------------------------------------------
    # Product selector used only to change Our Product Name in a GUI row.
    # ------------------------------------------------------------------
    def load_find_brands(self):
        current = clean_multi_spaces(self.ui.cbo_FindBrand.currentText())
        with self.get_session() as session:
            brands = self.service(session).get_brand_values()
        self.ui.cbo_FindBrand.blockSignals(True)
        self.ui.cbo_FindBrand.clear()
        self.ui.cbo_FindBrand.addItem("-")
        self.ui.cbo_FindBrand.addItems(brands)
        idx = self.ui.cbo_FindBrand.findText(current)
        self.ui.cbo_FindBrand.setCurrentIndex(idx if idx >= 0 else 0)
        self.ui.cbo_FindBrand.blockSignals(False)

    def _filtered_products_for_editor(self) -> list[Product]:
        brand = clean_multi_spaces(self.ui.cbo_FindBrand.currentText())
        text_filter = clean_multi_spaces(self.ui.line_FindProduct.text())
        with self.get_session() as session:
            return self.service(session).get_products_for_combo(brand, text_filter)

    def _row_by_code(self, code: str) -> dict | None:
        for row in self._all_rows:
            if clean_multi_spaces(row.get("sales_code")) == code:
                return row
        return None

    def _code_at_table_row(self, table_row: int) -> str:
        for col in range(self.table.columnCount()):
            item = self.table.item(table_row, col)
            if item is not None:
                code = item.data(Qt.UserRole)
                if code:
                    return str(code)
            widget = self.table.cellWidget(table_row, col)
            if widget is not None:
                code = widget.property("sales_code")
                if code:
                    return str(code)
        return ""

    def start_product_edit(self, table_row: int, column: int):
        if self._updating_table or column != self.COL_PRODUCT:
            return
        code = self._code_at_table_row(table_row)
        row = self._row_by_code(code)
        if row is None:
            return
        combo = QComboBox()
        combo.setEditable(True)
        combo.setInsertPolicy(QComboBox.NoInsert)
        combo.setProperty("sales_code", code)
        combo.addItem("", None)
        for product in self._filtered_products_for_editor():
            combo.addItem(product.name, int(product.id))
        current_id = row.get("product_id")
        if current_id and combo.findData(int(current_id)) < 0:
            with self.get_session() as session:
                product = session.query(Product).filter(Product.id == int(current_id)).first()
                if product:
                    combo.addItem(product.name, int(product.id))
        idx = combo.findData(int(current_id)) if current_id else -1
        if idx >= 0:
            combo.setCurrentIndex(idx)
        else:
            combo.setCurrentText(clean_multi_spaces(row.get("product_name")))
        combo.activated.connect(lambda _=None, tr=table_row, c=combo: self.finish_product_edit(tr, c))
        if combo.lineEdit() is not None:
            combo.lineEdit().returnPressed.connect(lambda tr=table_row, c=combo: self.finish_product_edit(tr, c))
        self.table.setCellWidget(table_row, column, combo)
        combo.setFocus()
        QTimer.singleShot(0, combo.showPopup)

    def refresh_current_product_combo(self):
        table_row = self.table.currentRow()
        if table_row < 0:
            return
        combo = self.table.cellWidget(table_row, self.COL_PRODUCT)
        if not isinstance(combo, QComboBox):
            return
        code = self._code_at_table_row(table_row)
        row = self._row_by_code(code)
        if row is None:
            return
        current_id = combo.currentData()
        current_text = combo.currentText()
        combo.blockSignals(True)
        combo.clear()
        combo.addItem("", None)
        for product in self._filtered_products_for_editor():
            combo.addItem(product.name, int(product.id))
        if current_id and combo.findData(current_id) < 0:
            with self.get_session() as session:
                product = session.query(Product).filter(Product.id == int(current_id)).first()
                if product:
                    combo.addItem(product.name, int(product.id))
        idx = combo.findData(current_id)
        if idx >= 0:
            combo.setCurrentIndex(idx)
        else:
            combo.setCurrentText(current_text)
        combo.blockSignals(False)

    def finish_product_edit(self, table_row: int, combo: QComboBox):
        code = str(combo.property("sales_code") or self._code_at_table_row(table_row))
        row = self._row_by_code(code)
        if row is None:
            return
        product_id = combo.currentData()
        text_value = clean_multi_spaces(combo.currentText())

        matched_id = None
        if text_value:
            for idx in range(combo.count()):
                if clean_multi_spaces(combo.itemText(idx)).upper() == text_value.upper():
                    matched_id = combo.itemData(idx)
                    text_value = combo.itemText(idx)
                    break

        if matched_id:
            product_id = int(matched_id)
            with self.get_session() as session:
                product = session.query(Product).filter(Product.id == product_id).first()
                if product:
                    row["product_id"] = product_id
                    row["product_name"] = product.name
                    row["product_qty_in_box"] = product.qty_in_box
                    row["new_product_name"] = ""
                    row["new_brand"] = ""
                    row["new_pack"] = None
                    row["new_qty_in_box"] = None
                    row["new_is_excise"] = None
        elif not text_value:
            row["product_id"] = None
            row["product_name"] = ""
            row["product_qty_in_box"] = None
            row["new_product_name"] = row.get("sales_product_name") or ""
            row["new_brand"] = row.get("sales_brand") or ""
            row["new_pack"] = row.get("sales_pack")
            row["new_qty_in_box"] = row.get("sales_qty_in_box")
            row["new_is_excise"] = bool(row.get("sales_is_excise"))
        else:
            # Typed text not present in Products means a proposed new product.
            row["product_id"] = None
            row["product_name"] = ""
            row["product_qty_in_box"] = None
            row["new_product_name"] = text_value.upper()
            row["new_brand"] = row.get("sales_brand") or ""
            row["new_pack"] = row.get("sales_pack")
            row["new_qty_in_box"] = row.get("sales_qty_in_box")
            row["new_is_excise"] = bool(row.get("sales_is_excise"))

        self.table.removeCellWidget(table_row, self.COL_PRODUCT)
        self.apply_filters()

    def _commit_open_product_editor(self):
        for table_row in range(self.table.rowCount()):
            combo = self.table.cellWidget(table_row, self.COL_PRODUCT)
            if isinstance(combo, QComboBox):
                self.finish_product_edit(table_row, combo)
                return

    # ------------------------------------------------------------------
    # Left filters affect only displayed rows.
    # ------------------------------------------------------------------
    def _load_product_meta(self):
        with self.get_session() as session:
            products = session.query(Product).all()
        self._product_meta = {
            int(p.id): (clean_multi_spaces(p.brand), clean_multi_spaces(p.family), clean_multi_spaces(p.name))
            for p in products
        }

    def _brand_for_row(self, row: dict) -> str:
        pid = row.get("product_id")
        if pid and int(pid) in self._product_meta:
            return self._product_meta[int(pid)][0]
        return clean_multi_spaces(row.get("sales_brand"))

    def _family_for_row(self, row: dict) -> str:
        pid = row.get("product_id")
        return self._product_meta.get(int(pid), ("", "", ""))[1] if pid else ""

    def _open_checked_filter(self, *, title: str, options: Sequence[FilterOption], selected):
        dialog = CheckedFilterDialog(self, title=title, options=options, selected_keys=selected)
        return dialog.exec_and_get_selection()

    def _brand_options(self) -> list[FilterOption]:
        values = sorted({self._brand_for_row(r) for r in self._all_rows if self._brand_for_row(r)})
        return [FilterOption(key=v, label=v, search_text=v) for v in values]

    def _family_options(self) -> list[FilterOption]:
        values = sorted({self._family_for_row(r) for r in self._all_rows if self._family_for_row(r)})
        return [FilterOption(key=v, label=v, search_text=v) for v in values]

    def _product_options(self) -> list[FilterOption]:
        ids = sorted({int(r["product_id"]) for r in self._all_rows if r.get("product_id")})
        result = []
        for pid in ids:
            brand, family, name = self._product_meta.get(pid, ("", "", ""))
            result.append(FilterOption(key=pid, label=name or str(pid), search_text=f"{pid} {name} {brand} {family}"))
        return result

    def open_brand_filter(self):
        accepted, selected = self._open_checked_filter(title="Фильтр по брендам", options=self._brand_options(), selected=self._selected_brand_values)
        if accepted:
            self._selected_brand_values = None if selected is None else {str(v) for v in selected}
            self._refresh_filter_buttons()

    def open_family_filter(self):
        accepted, selected = self._open_checked_filter(title="Фильтр по Product Family", options=self._family_options(), selected=self._selected_family_values)
        if accepted:
            self._selected_family_values = None if selected is None else {str(v) for v in selected}
            self._refresh_filter_buttons()

    def open_product_filter(self):
        accepted, selected = self._open_checked_filter(title="Фильтр по продуктам", options=self._product_options(), selected=self._selected_product_ids)
        if accepted:
            self._selected_product_ids = None if selected is None else {int(v) for v in selected}
            self._refresh_filter_buttons()

    def _refresh_filter_buttons(self):
        self.ui.btn_FilterBrand.setText("все Бренды" if self._selected_brand_values is None else f"все Бренды ({len(self._selected_brand_values)})")
        self.ui.btn_FilterProductFamily.setText("все Product Family" if self._selected_family_values is None else f"все Product Family ({len(self._selected_family_values)})")
        self.ui.btn_FilterProduct.setText("все Продукты" if self._selected_product_ids is None else f"все Продукты ({len(self._selected_product_ids)})")

    def _capture_left_filter_state(self):
        """Apply the current left-side controls only when Search is pressed."""
        self._applied_brand_values = (
            None if self._selected_brand_values is None else set(self._selected_brand_values)
        )
        self._applied_family_values = (
            None if self._selected_family_values is None else set(self._selected_family_values)
        )
        self._applied_product_ids = (
            None if self._selected_product_ids is None else set(self._selected_product_ids)
        )
        self._applied_name_search = clean_multi_spaces(
            self.ui.line_NameSearch.text()
        ).casefold()

    def _clear_applied_left_filters(self):
        self._applied_brand_values = None
        self._applied_family_values = None
        self._applied_product_ids = None
        self._applied_name_search = ""

    def apply_filters(self):
        text_filter = self._applied_name_search
        rows = []
        for row in self._all_rows:
            code = clean_multi_spaces(row.get("sales_code"))
            if code and code in self._visually_deleted_codes:
                continue
            if self._applied_brand_values is not None and self._brand_for_row(row) not in self._applied_brand_values:
                continue
            if self._applied_family_values is not None and self._family_for_row(row) not in self._applied_family_values:
                continue
            if self._applied_product_ids is not None:
                pid = row.get("product_id")
                if not pid or int(pid) not in self._applied_product_ids:
                    continue
            if text_filter:
                haystack = " ".join(clean_multi_spaces(row.get(key)) for key in (
                    "sales_code", "sales_article", "sales_product_name", "sales_brand",
                    "status", "product_name", "new_product_name",
                )).casefold()
                if text_filter not in haystack:
                    continue
            rows.append(row)
        self._rows = rows
        self._populate_table(rows)

    # ------------------------------------------------------------------
    # Shared context menu integration.
    # ------------------------------------------------------------------
    def after_standard_table_rows_deleted(self, table, rows, snapshots):
        for snapshot in snapshots:
            code = clean_multi_spaces(snapshot.get("key"))
            if code:
                self._visually_deleted_codes.add(code)

    def after_standard_table_rows_restored(self, table, snapshots):
        for snapshot in snapshots:
            code = clean_multi_spaces(snapshot.get("key"))
            if code:
                self._visually_deleted_codes.discard(code)

    def _selected_codes_from_rows(self, table_rows: list[int]) -> list[str]:
        result: list[str] = []
        for table_row in table_rows:
            code = self._code_at_table_row(table_row)
            if code and code not in result:
                result.append(code)
        return result

    def populate_standard_table_context_menu(self, menu, table, rows, index):
        """Append mapping-specific actions to the common table menu."""
        menu.addSeparator()
        codes = self._selected_codes_from_rows(rows)
        selected_data = [self._row_by_code(code) for code in codes]
        selected_data = [row for row in selected_data if row is not None]

        ignore_action = menu.addAction("Не использовать / больше не проверять")
        restore_action = menu.addAction("Вернуть в проверку")
        menu.addSeparator()
        toggle_text = "Скрыть игнорируемые" if self._show_ignored else "Показать игнорируемые"
        toggle_action = menu.addAction(toggle_text)

        ignore_action.setEnabled(any(not bool(row.get("is_ignored")) for row in selected_data))
        restore_action.setEnabled(any(bool(row.get("is_ignored")) for row in selected_data))

        ignore_action.triggered.connect(lambda: self._set_selected_ignored(codes, True))
        restore_action.triggered.connect(lambda: self._set_selected_ignored(codes, False))
        toggle_action.triggered.connect(self._toggle_show_ignored)

    def _set_selected_ignored(self, codes: list[str], ignored: bool):
        rows = [self._row_by_code(code) for code in codes]
        rows = [row for row in rows if row is not None]
        if not rows:
            return
        try:
            with self.get_session() as session:
                count = self.service(session).set_ignored(rows, ignored=ignored)
                session.commit()

            # Ignore is persistent immediately: it is not tied to the ordinary
            # Save button, because its purpose is to stop future system checks.
            self._pending_deletes.difference_update(codes)
            self._visually_deleted_codes.difference_update(codes)
            self._deleted_row_snapshots.clear()

            if self._mode == "search":
                self.search_saved()
            else:
                self._all_rows = [
                    row for row in self._all_rows
                    if clean_multi_spaces(row.get("sales_code")) not in set(codes)
                ]
                self.apply_filters()

            self.show_message(
                f"Помечено как неиспользуемые: {count}"
                if ignored
                else f"Возвращено в проверку: {count}"
            )
        except Exception as exc:
            self.show_error_message(str(exc))

    def _toggle_show_ignored(self):
        self._show_ignored = not self._show_ignored
        self.search_saved()

    def _reset_visual_delete_state(self):
        self._pending_deletes.clear()
        self._deleted_row_snapshots.clear()
        self._visually_deleted_codes.clear()

    def reset_all(self):
        self.ui.line_FindProduct.clear()
        self.ui.cbo_FindBrand.setCurrentIndex(0)
        self.ui.line_NameSearch.clear()
        self._selected_brand_values = None
        self._selected_family_values = None
        self._selected_product_ids = None
        self._clear_applied_left_filters()
        self._show_ignored = False
        self._reset_visual_delete_state()
        self._refresh_filter_buttons()
        self.apply_filters()
        self.show_message("Фильтры и поля сброшены")

    # ------------------------------------------------------------------
    # Search / Import / Save
    # ------------------------------------------------------------------
    def search_saved(self):
        try:
            self._capture_left_filter_state()

            # After Portfolio import Search is only the left-filter apply button.
            # It must not replace the current Portfolio result with historical
            # rows from sales_product_links.
            if self._mode == "check" and self._portfolio_df is not None:
                self.apply_filters()
                self.show_message(f"Найдено: {len(self._rows)}")
                return

            with self.get_session() as session:
                rows = self.service(session).search_links(include_ignored=self._show_ignored)
            self._mode = "search"
            self._reset_visual_delete_state()
            self._all_rows = [dict(r) for r in rows]
            self._load_product_meta()
            self.apply_filters()
            self.show_message(
                f"Сохранённых сопоставлений: {len(rows)}; найдено: {len(self._rows)}"
            )
        except Exception as exc:
            self.show_error_message(str(exc))

    def _apply_check_result(self, result, *, loaded_rows: int | None = None):
        self._mode = "check"
        self._reset_visual_delete_state()
        self._all_rows = [dict(r) for r in result.rows]
        self._load_product_meta()
        self.apply_filters()

        if loaded_rows is not None:
            unresolved_names = sum(
                1 for row in result.rows
                if not row.get("product_id")
            )
            self.show_message(
                f"Загружено строк: {loaded_rows}; "
                f"не определено названий: {unresolved_names}."
            )
            return

        if not result.rows:
            self.show_message("Изменений не найдено")
        else:
            self.show_message(f"Требуют обработки: {len(result.rows)}")

    def import_portfolio(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите файл Портфель",
            "",
            "Excel files (*.xlsx *.xlsm *.xls)",
        )
        if not file_path:
            return

        try:
            portfolio_df = ProductMappingPortfolioImporter().read_excel(file_path)
            with self.get_session() as session:
                service = self.service(session)
                deleted = service.cleanup_stale_new_links(portfolio_df)
                # The stale NEW links must disappear from the DB before the
                # matching pass starts, exactly as the Portfolio refresh rule
                # requires.
                session.commit()
                result = service.check_products(portfolio_df)

            self._portfolio_df = portfolio_df.copy()
            self._clear_applied_left_filters()

            self._apply_check_result(result, loaded_rows=len(portfolio_df))
        except Exception as exc:
            self.show_error_message(str(exc))


    def save(self):
        self._commit_open_product_editor()
        pending_deletes = {clean_multi_spaces(code) for code in self._pending_deletes if clean_multi_spaces(code)}
        base_rows = self._all_rows if self._mode == "check" else self._rows
        rows = [
            row for row in base_rows
            if clean_multi_spaces(row.get("sales_code")) not in pending_deletes
            and not bool(row.get("is_ignored"))
        ]
        if not rows and not pending_deletes:
            self.show_message("Нет строк для сохранения")
            return
        try:
            with self.get_session() as session:
                deleted = self.service(session).delete_links(pending_deletes)
                count = self.service(session).save_rows(rows) if rows else 0
                session.commit()
            # Do not jump to Search after Portfolio save: Search shows the full
            # historical sales_product_links table and can contain legacy links
            # that are not present in the current filtered Portfolio. Re-run the
            # current Portfolio comparison instead, so the GUI stays tied to the
            # imported source. Search remains available as a separate explicit
            # action.
            if self._mode == "check" and self._portfolio_df is not None:
                with self.get_session() as session:
                    result = self.service(session).check_products(self._portfolio_df)
                self._apply_check_result(result)
            else:
                self.search_saved()

            if deleted:
                self.show_message(f"Сопоставления сохранены: {count}; удалено: {deleted}")
            else:
                self.show_message(f"Сопоставления сохранены: {count}")
        except MissingPackTypeError as exc:
            selected_pack = ask_pack_type(self, exc)
            if selected_pack is None:
                return
            requested_pack = parse_loose_number(exc.requested_pack)
            for row in self._all_rows:
                if row.get("product_id"):
                    continue
                pack = row.get("new_pack") if row.get("new_pack") not in (None, "") else row.get("sales_pack")
                if parse_loose_number(pack) == requested_pack:
                    row["new_pack"] = selected_pack
                    row["sales_pack"] = selected_pack
            self.apply_filters()
            self.save()
        except Exception as exc:
            self.show_error_message(str(exc))

    def save_excel(self):
        if not self._rows:
            self.show_message("Нет строк для выгрузки")
            return
        default = f"Product_Mapping_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
        path, _ = QFileDialog.getSaveFileName(self, "Сохранить Excel", str(BASE_DIR / default), "Excel files (*.xlsx)")
        if not path:
            return
        try:
            data = []
            for row in self._rows:
                data.append({header: row.get(field) for field, header in zip(self.COLUMNS, self.HEADERS)})
            pd.DataFrame(data).to_excel(path, index=False)
            self.show_message("Excel файл сохранен")
            QDesktopServices.openUrl(Path(path).as_uri())
        except Exception as exc:
            self.show_error_message(str(exc))

    # ------------------------------------------------------------------
    # Table rendering / editable new-product fields.
    # ------------------------------------------------------------------
    @staticmethod
    def _display(field: str, value) -> str:
        if value is None:
            return ""
        if isinstance(value, bool):
            return "Да" if value else "Нет"

        # Numeric display rules are shared application-wide in table_style.py:
        # Pack is a decimal with comma; Qty in Box is an integer without ',0'.
        if field in ProductMappingPage.NUMERIC_FIELDS:
            return format_table_field_value(field, value)
        return str(value)

    def _item(self, field: str, value, code: str, *, editable: bool = False, left: bool = False):
        item = SortableTableWidgetItem(self._display(field, value))
        flags = Qt.ItemIsEnabled | Qt.ItemIsSelectable
        if editable:
            flags |= Qt.ItemIsEditable
        item.setFlags(flags)
        item.setData(Qt.UserRole, code)
        item.setData(Qt.UserRole + 1, value)
        item.setTextAlignment((Qt.AlignLeft if left else Qt.AlignCenter) | Qt.AlignVCenter)
        return item

    def _build_checkbox_widget(self, code: str, checked: bool, enabled: bool) -> QWidget:
        # Keep the checkbox visually identical to Supplier Price.
        checkbox = QCheckBox()
        checkbox.setChecked(checked)
        checkbox.setEnabled(enabled)
        checkbox.setStyleSheet(
            """
            QCheckBox {
                background: transparent;
            }
            QCheckBox::indicator {
                width: 14px;
                height: 14px;
            }
            """
        )
        checkbox.toggled.connect(lambda state, c=code: self._new_excise_changed(c, state))

        container = QWidget()
        container.setProperty("sales_code", code)
        layout = QHBoxLayout(container)
        layout.addWidget(checkbox)
        layout.setAlignment(Qt.AlignCenter)
        layout.setContentsMargins(0, 0, 0, 0)
        return container

    def _populate_table(self, rows: list[dict]):
        self._updating_table = True
        self.table.setSortingEnabled(False)
        try:
            self.table.clearContents()
            self.table.setRowCount(len(rows))
            self.table.setColumnCount(len(self.HEADERS))
            self.table.setHorizontalHeaderLabels(self.HEADERS)
            for tr, row in enumerate(rows):
                code = clean_multi_spaces(row.get("sales_code"))
                for col, field in enumerate(self.COLUMNS):
                    if field == "new_is_excise":
                        self.table.setCellWidget(
                            tr,
                            col,
                            self._build_checkbox_widget(
                                code,
                                bool(row.get(field)),
                                not bool(row.get("product_id")),
                            ),
                        )
                        continue
                    editable = field in self.EDITABLE_FIELDS and not bool(row.get("product_id"))
                    left = field in {"sales_article", "sales_product_name", "sales_brand", "status", "product_name", "new_product_name", "new_brand"}
                    self.table.setItem(
                        tr,
                        col,
                        self._item(field, row.get(field), code, editable=editable, left=left),
                    )
            self.table.resizeColumnsToContents()
            resize_columns_for_multiline_headers(self.table)
            for col, field in enumerate(self.COLUMNS):
                if field in {"sales_product_name", "product_name", "new_product_name"}:
                    self.table.setColumnWidth(col, max(self.table.columnWidth(col), 250))
                elif field in {"status", "sales_article"}:
                    self.table.setColumnWidth(col, max(self.table.columnWidth(col), 130))
        finally:
            self.table.setSortingEnabled(True)
            self._updating_table = False

    def on_item_changed(self, item: QTableWidgetItem):
        if self._updating_table:
            return
        col = item.column()
        if col < 0 or col >= len(self.COLUMNS):
            return
        field = self.COLUMNS[col]
        if field not in self.EDITABLE_FIELDS:
            return
        code = str(item.data(Qt.UserRole) or "")
        row = self._row_by_code(code)
        if row is None or row.get("product_id"):
            return
        row[field] = item.text().strip()

    def _new_excise_changed(self, code: str, checked: bool):
        if self._updating_table:
            return
        row = self._row_by_code(code)
        if row is not None and not row.get("product_id"):
            row["new_is_excise"] = bool(checked)
