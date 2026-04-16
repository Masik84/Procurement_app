from __future__ import annotations

import shutil
from datetime import datetime
from pathlib import Path

import pandas as pd
from PySide6.QtCore import Qt, QFile, QEvent, QPoint
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QHBoxLayout,
    QLineEdit,
    QMenu,
    QMessageBox,
    QTableWidgetItem,
    QToolTip,
    QVBoxLayout,
    QWidget,
)
from PySide6.QtUiTools import QUiLoader
from sqlalchemy.orm import joinedload

from app.db.db import SessionLocal
from app.db.models import Product, TempProductSearchImport
from app.imports.product_search_importer import ProductSearchImporter
from app.services.product_search_service import ProductSearchService
from app.utils.batch import get_current_username
from app.utils.parsers import parse_loose_number
from app.utils.text import clean_multi_spaces
from app.ui.table_style import setup_data_table


BASE_DIR = Path(__file__).resolve().parents[2]
PRODUCT_SEARCH_UI = BASE_DIR / "app" / "ui" / "windows" / "product_search.ui"
PRODUCT_SEARCH_TEMPLATE = BASE_DIR / "ProductSearchTemplate.xlsx"


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


class ProductSearchPage(QWidget):
    def __init__(self):
        super().__init__()

        self.ui = load_ui(PRODUCT_SEARCH_UI)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.ui)

        self.imported_by = get_current_username()
        self.batch_id = ""
        self.selected_file_path = ""

        self._updating_table = False
        self._pending_changes: dict[int, dict] = {}
        self._pending_deletes: set[int] = set()
        self._new_rows: set[int] = set()
        self._table_row_ids: list[int] = []

        self.columns = [
            "selected_product_id",
            "source_article",
            "source_product_name",
            "new_product_name",
            "new_brand",
            "new_pack",
            "new_is_excise",
        ]
        self.headers = [
            "Product name",
            "Article",
            "Supplier Product name",
            "Product name (for new)",
            "Brand (for new)",
            "Pack (for new)",
            "Excise duty (for new)",
        ]
        self.numeric_columns = {"new_pack"}

        self.setup_ui()
        self.setup_connections()
        self.load_initial_state()

    def setup_ui(self):
        self.table = self.ui.table
        setup_data_table(self.table, sorting=False)
        self.table.horizontalHeader().setSectionsMovable(False)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.ui.label_msg.setText("Сообщений нет")

        self.ui.line_FindProduct.setToolTip("Фильтр по названию продукта из базы")
        self.ui.line_FindProduct.installEventFilter(self)

    def setup_connections(self):
        self.table.itemChanged.connect(self.on_item_changed)
        self.table.customContextMenuRequested.connect(self.show_context_menu)

        self.ui.btn_DownFile.clicked.connect(self.download_template)
        self.ui.btn_Import.clicked.connect(self.import_file)
        self.ui.btn_AddLine.clicked.connect(self.add_line)
        self.ui.btn_Save.clicked.connect(self.apply_pending_changes)
        self.ui.btn_Reset.clicked.connect(self.reset_form)

        self.ui.cbo_FindBrand.currentTextChanged.connect(self.refresh_current_product_combo)
        self.ui.line_FindProduct.textChanged.connect(self.refresh_current_product_combo)

    def get_session(self):
        return SessionLocal()

    def eventFilter(self, watched, event):
        if watched is self.ui.line_FindProduct and event.type() in {QEvent.Enter, QEvent.FocusIn, QEvent.MouseButtonPress}:
            QToolTip.showText(
                watched.mapToGlobal(QPoint(0, watched.height())),
                watched.toolTip(),
                watched,
            )

        if isinstance(watched, QComboBox):
            row_id = watched.property("row_id")
            role = watched.property("combo_role")
            if row_id and role and event.type() in {QEvent.FocusIn, QEvent.MouseButtonPress}:
                if role == "product_combo":
                    self.populate_product_combo(watched, int(row_id), keep_current=True)
                elif role == "brand_combo":
                    self.populate_brand_combo(watched, keep_current=True)

        return super().eventFilter(watched, event)

    def load_initial_state(self):
        self.start_new_batch()
        self.load_find_brands()
        self.cleanup_old_temp_rows()
        self.load_table_rows()

    def start_new_batch(self):
        with self.get_session() as session:
            service = ProductSearchService(session)
            self.batch_id = service.start_batch()

    def cleanup_old_temp_rows(self):
        with self.get_session() as session:
            service = ProductSearchService(session)
            service.cleanup_old_temp_rows(imported_by=self.imported_by)
            session.commit()

    def show_message(self, text: str):
        self.ui.label_msg.setText(text)

    def show_error_message(self, text: str):
        self.ui.label_msg.setText(text)
        QMessageBox.warning(self, "Ошибка", text)

    def load_find_brands(self):
        current_text = self.ui.cbo_FindBrand.currentText().strip()
        brands = self.get_brand_names()

        self.ui.cbo_FindBrand.blockSignals(True)
        self.ui.cbo_FindBrand.clear()
        self.ui.cbo_FindBrand.addItem("-")
        if brands:
            self.ui.cbo_FindBrand.addItems(brands)
        self.ui.cbo_FindBrand.blockSignals(False)

        if current_text:
            self.set_combo_text(self.ui.cbo_FindBrand, current_text if current_text != "" else "-")

    def get_brand_names(self) -> list[str]:
        with self.get_session() as session:
            rows = (
                session.query(Product.brand)
                .filter(Product.brand.isnot(None), Product.brand != "")
                .distinct()
                .order_by(Product.brand.asc())
                .all()
            )
        return [row[0] for row in rows if row[0]]

    def set_combo_text(self, combo: QComboBox, value: str):
        index = combo.findText(value)
        if index >= 0:
            combo.setCurrentIndex(index)
        elif combo.isEditable():
            combo.setCurrentText(value)

    def download_template(self):
        if not PRODUCT_SEARCH_TEMPLATE.exists():
            df = pd.DataFrame(columns=["Article", "Product name"])
            df.to_excel(PRODUCT_SEARCH_TEMPLATE, index=False)

        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Сохранить шаблон",
            str(Path.home() / "ProductSearchTemplate.xlsx"),
            "Excel files (*.xlsx)",
        )
        if not file_path:
            return

        if not file_path.lower().endswith(".xlsx"):
            file_path += ".xlsx"

        shutil.copyfile(PRODUCT_SEARCH_TEMPLATE, file_path)
        QDesktopServices.openUrl(Path(file_path).as_uri())
        self.show_message("Шаблон сохранен")

    def import_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите файл для поиска продуктов",
            "",
            "Excel files (*.xls *.xlsx)",
        )
        if not file_path:
            return

        try:
            importer = ProductSearchImporter()
            rows = importer.read_excel(file_path)

            with self.get_session() as session:
                service = ProductSearchService(session)
                service.import_rows_to_temp(
                    batch_id=self.batch_id,
                    imported_by=self.imported_by,
                    rows=rows,
                    import_date=datetime.now(),
                    replace_existing_batch_rows=True,
                )
                service.automatch_temp_rows(self.batch_id, self.imported_by)
                session.commit()

            self.selected_file_path = file_path
            self.load_table_rows()
            self.show_message("Данные импортированы")
        except Exception as e:
            self.show_error_message(str(e))

    def load_table_rows(self):
        with self.get_session() as session:
            rows = (
                session.query(TempProductSearchImport)
                .options(joinedload(TempProductSearchImport.selected_product))
                .filter(
                    TempProductSearchImport.batch_id == self.batch_id,
                    TempProductSearchImport.imported_by == self.imported_by,
                )
                .order_by(TempProductSearchImport.import_row_no.asc(), TempProductSearchImport.id.asc())
                .all()
            )

        self.display_rows(rows)

    def display_rows(self, rows: list[TempProductSearchImport]):
        self._updating_table = True
        self._table_row_ids = [row.id for row in rows]
        self._pending_deletes.clear()
        self._new_rows.clear()

        self.table.clear()
        self.table.setColumnCount(len(self.headers))
        self.table.setRowCount(len(rows))
        self.table.setHorizontalHeaderLabels(self.headers)

        for row_index, row in enumerate(rows):
            self._pending_changes.setdefault(row.id, {})

            self.table.setCellWidget(
                row_index,
                0,
                self.build_product_combo(
                    row.id,
                    row.selected_product_id,
                    row.selected_product.name if row.selected_product else "",
                ),
            )
            self.table.setItem(row_index, 1, self.build_table_item("source_article", row.source_article or ""))
            self.table.setItem(row_index, 2, self.build_table_item("source_product_name", row.source_product_name or ""))
            self.table.setItem(row_index, 3, self.build_table_item("new_product_name", row.new_product_name or ""))
            self.table.setCellWidget(
                row_index,
                4,
                self.build_brand_combo(row.id, row.new_brand or ""),
            )
            self.table.setItem(row_index, 5, self.build_table_item("new_pack", self.value_to_text(row.new_pack)))
            self.table.setCellWidget(
                row_index,
                6,
                self.build_checkbox_widget(row.id, bool(row.new_is_excise)),
            )

        self.table.resizeColumnsToContents()
        self.table.setColumnWidth(0, 260)
        self.table.setColumnWidth(1, 140)
        self.table.setColumnWidth(2, 280)
        self.table.setColumnWidth(3, 280)
        self.table.setColumnWidth(4, 160)
        self.table.setColumnWidth(5, 100)
        self.table.setColumnWidth(6, 110)

        self._updating_table = False

    def value_to_text(self, value: object) -> str:
        if value is None:
            return ""
        number = parse_loose_number(value)
        if number is None:
            return str(value)
        return str(number).replace(".", ",")

    def build_table_item(self, column_name: str, value: str) -> QTableWidgetItem:
        item = QTableWidgetItem(value)
        item.setData(Qt.UserRole, column_name)
        item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable)
        if column_name in self.numeric_columns:
            item.setTextAlignment(Qt.AlignCenter)
        else:
            item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        return item

    def build_product_combo(self, row_id: int, selected_product_id: int | None, selected_name: str) -> QComboBox:
        combo = QComboBox()
        combo.setProperty("row_id", row_id)
        combo.setProperty("combo_role", "product_combo")
        combo.installEventFilter(self)
        combo.setToolTip("Выберите продукт из базы")
        self.populate_product_combo(combo, row_id, keep_current=False, selected_product_id=selected_product_id, selected_name=selected_name)
        combo.currentIndexChanged.connect(lambda _=None, rid=row_id, cb=combo: self.on_product_combo_changed(rid, cb))
        return combo

    def populate_product_combo(
        self,
        combo: QComboBox,
        row_id: int,
        keep_current: bool,
        selected_product_id: int | None = None,
        selected_name: str = "",
    ):
        current_id = combo.currentData() if keep_current else selected_product_id
        current_name = combo.currentText().strip() if keep_current else selected_name

        products = self.get_filtered_products()
        combo.blockSignals(True)
        combo.clear()
        combo.addItem("", None)

        added_ids: set[int] = set()
        for product in products:
            combo.addItem(product.name, product.id)
            added_ids.add(product.id)

        if current_id and current_id not in added_ids and current_name:
            combo.insertItem(1, current_name, current_id)

        index = combo.findData(current_id)
        if index >= 0:
            combo.setCurrentIndex(index)
        elif current_name and combo.findText(current_name) >= 0:
            combo.setCurrentText(current_name)
        else:
            combo.setCurrentIndex(0)
        combo.blockSignals(False)

    def build_brand_combo(self, row_id: int, brand_name: str) -> QComboBox:
        combo = QComboBox()
        combo.setEditable(True)
        combo.setInsertPolicy(QComboBox.NoInsert)
        combo.setProperty("row_id", row_id)
        combo.setProperty("combo_role", "brand_combo")
        combo.installEventFilter(self)
        self.populate_brand_combo(combo, keep_current=False, current_text=brand_name)
        combo.currentTextChanged.connect(lambda _=None, rid=row_id, cb=combo: self.on_brand_combo_changed(rid, cb))
        return combo

    def populate_brand_combo(self, combo: QComboBox, keep_current: bool, current_text: str = ""):
        brand_value = combo.currentText().strip() if keep_current else current_text
        brands = self.get_brand_names()

        combo.blockSignals(True)
        combo.clear()
        combo.addItem("")
        if brands:
            combo.addItems(brands)
        if brand_value and combo.findText(brand_value) < 0:
            combo.addItem(brand_value)
        combo.setCurrentText(brand_value)
        combo.blockSignals(False)

    def build_checkbox_widget(self, row_id: int, checked: bool) -> QWidget:
        checkbox = QCheckBox()
        checkbox.setChecked(checked)
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
        checkbox.toggled.connect(lambda state, rid=row_id: self.on_checkbox_changed(rid, state))

        container = QWidget()
        layout = QHBoxLayout(container)
        layout.addWidget(checkbox)
        layout.setAlignment(Qt.AlignCenter)
        layout.setContentsMargins(0, 0, 0, 0)
        return container

    def on_product_combo_changed(self, row_id: int, combo: QComboBox):
        if self._updating_table:
            return
        self._pending_changes.setdefault(row_id, {})
        self._pending_changes[row_id]["selected_product_id"] = combo.currentData()

    def on_brand_combo_changed(self, row_id: int, combo: QComboBox):
        if self._updating_table:
            return
        self._pending_changes.setdefault(row_id, {})
        self._pending_changes[row_id]["new_brand"] = clean_multi_spaces(combo.currentText()) or None

    def on_checkbox_changed(self, row_id: int, checked: bool):
        if self._updating_table:
            return
        self._pending_changes.setdefault(row_id, {})
        self._pending_changes[row_id]["new_is_excise"] = bool(checked)

    def on_item_changed(self, item: QTableWidgetItem):
        if self._updating_table:
            return

        row = item.row()
        if row < 0 or row >= len(self._table_row_ids):
            return

        row_id = self._table_row_ids[row]
        column_name = item.data(Qt.UserRole)
        if not column_name:
            return

        value = clean_multi_spaces(item.text())
        self._pending_changes.setdefault(row_id, {})
        self._pending_changes[row_id][column_name] = value or None

    def show_context_menu(self, position):
        menu = QMenu()
        copy_action = menu.addAction("Копировать")
        delete_action = menu.addAction("Удалить строку")
        apply_action = menu.addAction("Применить изменения")
        revert_action = menu.addAction("Обновить")

        copy_action.triggered.connect(self.copy_cell_content)
        delete_action.triggered.connect(self.delete_selected_row)
        apply_action.triggered.connect(self.apply_pending_changes)
        revert_action.triggered.connect(self.load_table_rows)

        menu.exec_(self.table.viewport().mapToGlobal(position))

    def copy_cell_content(self):
        selected_items = self.table.selectedItems()
        if not selected_items:
            return

        clipboard = QApplication.clipboard()
        if len(selected_items) == 1:
            clipboard.setText(selected_items[0].text())
        else:
            rows = {}
            for item in selected_items:
                rows.setdefault(item.row(), {})[item.column()] = item.text()
            text_rows = []
            for _, cols in sorted(rows.items()):
                text_rows.append("	".join(value for _, value in sorted(cols.items())))
            clipboard.setText("".join(text_rows).strip())

        self.show_message("Скопировано")

    def delete_selected_row(self):
        row = self.table.currentRow()
        if row < 0 or row >= len(self._table_row_ids):
            self.show_error_message("Не выбрана строка")
            return

        row_id = self._table_row_ids[row]
        self._pending_deletes.add(row_id)
        self.apply_pending_changes(save_to_db_only=True)

    def refresh_current_product_combo(self):
        row = self.table.currentRow()
        if row < 0:
            return
        combo = self.table.cellWidget(row, 0)
        if isinstance(combo, QComboBox):
            row_id = combo.property("row_id")
            if row_id:
                self.populate_product_combo(combo, int(row_id), keep_current=True)

    def get_filtered_products(self) -> list[Product]:
        brand_filter = clean_multi_spaces(self.ui.cbo_FindBrand.currentText())
        text_filter = clean_multi_spaces(self.ui.line_FindProduct.text())

        with self.get_session() as session:
            query = session.query(Product).filter(Product.name.isnot(None), Product.name != "")
            if brand_filter and brand_filter != "-":
                query = query.filter(Product.brand == brand_filter)
            if text_filter:
                query = query.filter(Product.name.ilike(f"%{text_filter}%"))
            products = query.order_by(Product.name.asc()).all()
        return products

    def add_line(self):
        try:
            with self.get_session() as session:
                service = ProductSearchService(session)
                service.create_empty_temp_row(
                    batch_id=self.batch_id,
                    imported_by=self.imported_by,
                    import_date=datetime.now(),
                )
                session.commit()

            self.load_table_rows()
            if self.table.rowCount() > 0:
                self.table.setCurrentCell(0, 0)
            self.show_message("Добавлена строка")
        except Exception as e:
            self.show_error_message(str(e))

    def apply_pending_changes(self, save_to_db_only: bool = False):
        if not self._pending_changes and not self._pending_deletes and not save_to_db_only:
            self.show_message("Нет изменений")
            return

        try:
            with self.get_session() as session:
                service = ProductSearchService(session)

                for row_id, changes in self._pending_changes.items():
                    row = session.query(TempProductSearchImport).filter(TempProductSearchImport.id == row_id).first()
                    if row is None:
                        continue

                    for key, value in changes.items():
                        if key == "new_pack":
                            parsed = parse_loose_number(value)
                            setattr(row, key, parsed if parsed is not None else None)
                        else:
                            setattr(row, key, value)

                if self._pending_deletes:
                    session.query(TempProductSearchImport).filter(TempProductSearchImport.id.in_(self._pending_deletes)).delete(
                        synchronize_session=False
                    )

                if not save_to_db_only:
                    service.validate_new_products_before_save(self.batch_id, self.imported_by)
                    service.create_products_from_temp(self.batch_id, self.imported_by)
                    service.create_or_update_product_articles(self.batch_id, self.imported_by)

                session.commit()

            if not save_to_db_only:
                save_path, _ = QFileDialog.getSaveFileName(
                    self,
                    "Сохранить итоговый файл",
                    str(Path.home() / f"ProductSearch_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"),
                    "Excel files (*.xlsx)",
                )
                if not save_path:
                    self.show_message("Данные сохранены в БД. Сохранение файла отменено")
                else:
                    with self.get_session() as session:
                        service = ProductSearchService(session)
                        service.export_to_excel(self.batch_id, self.imported_by, save_path)
                    self.show_message("Данные сохранены")

            self._pending_changes.clear()
            self._pending_deletes.clear()
            self.load_find_brands()
            self.load_table_rows()
            if save_to_db_only:
                self.show_message("Строка удалена")
        except Exception as e:
            self.show_error_message(str(e))

    def reset_form(self):
        try:
            with self.get_session() as session:
                service = ProductSearchService(session)
                service.reset_batch(self.batch_id, self.imported_by)
                session.commit()

            self._pending_changes.clear()
            self._pending_deletes.clear()
            self.start_new_batch()
            self.selected_file_path = ""
            self.ui.line_FindProduct.clear()
            self.set_combo_text(self.ui.cbo_FindBrand, "-")
            self.load_table_rows()
            self.show_message("Форма очищена")
        except Exception as e:
            self.show_error_message(str(e))
