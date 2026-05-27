
from __future__ import annotations

from datetime import datetime
from pathlib import Path

from PySide6.QtCore import Qt, QFile, QEvent, QPoint, QTimer
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
from app.exports.product_search_exporter import ProductSearchExporter
from app.imports.product_search_importer import ProductSearchImporter
from app.services.product_search_service import ProductSearchService
from app.ui.table_style import *
from app.utils.batch import get_current_username
from app.utils.parsers import parse_loose_number
from app.utils.text import clean_multi_spaces


BASE_DIR = Path(__file__).resolve().parents[2]
PRODUCT_SEARCH_UI = BASE_DIR / "app" / "ui" / "windows" / "product_search.ui"


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
    COL_PRODUCT = 0
    COL_ARTICLE = 1
    COL_SUPPLIER_PRODUCT_NAME = 2
    COL_NEW_PRODUCT_NAME = 3
    COL_BRAND = 4
    COL_PACK = 5
    COL_EXCISE = 6

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
        self.table.cellDoubleClicked.connect(self.start_cell_edit)
        self.table.itemChanged.connect(self.on_item_changed)
        self.table.customContextMenuRequested.connect(self.show_context_menu)

        self.ui.btn_DownFile.clicked.connect(self.download_template)
        self.ui.btn_Import.clicked.connect(self.import_file)
        self.ui.btn_AddLine.clicked.connect(self.add_line)
        self.ui.btn_Save.clicked.connect(self.apply_pending_changes)
        self.ui.btn_Reset.clicked.connect(self.reset_form)

    def get_session(self):
        return SessionLocal()

    def eventFilter(self, watched, event):
        if watched is self.ui.line_FindProduct and event.type() in {QEvent.Enter, QEvent.FocusIn, QEvent.MouseButtonPress}:
            QToolTip.showText(
                watched.mapToGlobal(QPoint(0, watched.height())),
                watched.toolTip(),
                watched,
            )
        return super().eventFilter(watched, event)

    def load_initial_state(self):
        self.start_new_batch()
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

    def show_popup_error(self, text: str):
        QMessageBox.warning(self, "Ошибка", text)

    def set_combo_text(self, combo: QComboBox, value: str):
        index = combo.findText(value)
        if index >= 0:
            combo.setCurrentIndex(index)
        elif combo.isEditable():
            combo.setCurrentText(value)

    def download_template(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Сохранить шаблон",
            str(BASE_DIR / "ProductSearchTemplate.xlsx"),
            "Excel files (*.xlsx)",
        )
        if not file_path:
            return

        if not file_path.lower().endswith(".xlsx"):
            file_path += ".xlsx"

        try:
            exporter = ProductSearchExporter()
            exporter.export_template(file_path)
            QDesktopServices.openUrl(Path(file_path).as_uri())
            self.show_message("Шаблон сохранен")
        except PermissionError:
            self.show_popup_error(
                "Не удалось сохранить файл.\n\n"
                "Скорее всего, файл уже открыт в Excel.\n"
                "Закрой файл и попробуй снова."
            )
        except Exception as e:
            self.show_popup_error(f"Ошибка при сохранении шаблона: {str(e)}")

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

        self.table.clear()
        self.table.setColumnCount(len(self.headers))
        self.table.setRowCount(len(rows))
        self.table.setHorizontalHeaderLabels(self.headers)

        for row_index, row in enumerate(rows):
            product_name = row.selected_product.name if row.selected_product else ""
            brand_name = row.new_brand or ""

            self.table.setItem(
                row_index,
                self.COL_PRODUCT,
                self.build_table_item("selected_product_id", product_name, editable=False, align_left=True),
            )
            self.table.setItem(
                row_index,
                self.COL_ARTICLE,
                self.build_table_item("source_article", row.source_article or "", editable=True, align_left=True),
            )
            self.table.setItem(
                row_index,
                self.COL_SUPPLIER_PRODUCT_NAME,
                self.build_table_item("source_product_name", row.source_product_name or "", editable=True, align_left=True),
            )
            self.table.setItem(
                row_index,
                self.COL_NEW_PRODUCT_NAME,
                self.build_table_item("new_product_name", row.new_product_name or "", editable=True, align_left=True),
            )
            self.table.setItem(
                row_index,
                self.COL_BRAND,
                self.build_table_item("new_brand", brand_name, editable=False, align_left=True),
            )
            self.table.setItem(
                row_index,
                self.COL_PACK,
                self.build_table_item("new_pack", self.value_to_text(row.new_pack), editable=True, align_left=False),
            )
            self.table.setCellWidget(
                row_index,
                self.COL_EXCISE,
                self.build_checkbox_widget(row.id, bool(row.new_is_excise)),
            )

        self.table.resizeColumnsToContents()
        self._updating_table = False

    def value_to_text(self, value: object) -> str:
        if value is None:
            return ""
        number = parse_loose_number(value)
        if number is None:
            return str(value)
        return str(number).replace(".", ",")

    def build_table_item(self, column_name: str, value: str, *, editable: bool = True, align_left: bool = False) -> QTableWidgetItem:
        item = QTableWidgetItem(value)
        item.setData(Qt.UserRole, column_name)
        flags = Qt.ItemIsEnabled | Qt.ItemIsSelectable
        if editable:
            flags |= Qt.ItemIsEditable
        item.setFlags(flags)
        if column_name in self.numeric_columns and editable:
            item.setTextAlignment(Qt.AlignCenter)
        else:
            item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter if align_left else Qt.AlignCenter)
        return item

    def start_cell_edit(self, row: int, column: int):
        if self._updating_table:
            return

        if row < 0 or row >= len(self._table_row_ids):
            return

        row_id = self._table_row_ids[row]

        if column == self.COL_PRODUCT:
            current_name = self._get_cell_text(row, column)
            current_product_id = self._get_row_selected_product_id(row_id)

            combo = self.build_product_combo(row_id, current_product_id, current_name)
            combo.activated.connect(
                lambda _=None, r=row, rid=row_id, cb=combo: self.finish_product_edit(r, rid, cb)
            )
            if combo.lineEdit() is not None:
                combo.lineEdit().returnPressed.connect(
                    lambda r=row, rid=row_id, cb=combo: self.finish_product_edit(r, rid, cb)
                )

            self.table.setCellWidget(row, column, combo)
            combo.setFocus()
            QTimer.singleShot(0, combo.showPopup)

        elif column == self.COL_BRAND:
            current_text = self._get_cell_text(row, column)

            combo = self.build_brand_combo(row_id, current_text)
            combo.activated.connect(
                lambda _=None, r=row, rid=row_id, cb=combo: self.finish_brand_edit(r, rid, cb)
            )
            if combo.lineEdit() is not None:
                combo.lineEdit().returnPressed.connect(
                    lambda r=row, rid=row_id, cb=combo: self.finish_brand_edit(r, rid, cb)
                )

            self.table.setCellWidget(row, column, combo)
            combo.setFocus()
            QTimer.singleShot(0, combo.showPopup)

    def _get_cell_text(self, row: int, column: int) -> str:
        item = self.table.item(row, column)
        return item.text().strip() if item else ""

    def _get_row_selected_product_id(self, row_id: int):
        changes = self._pending_changes.get(row_id, {})
        if "selected_product_id" in changes:
            return changes.get("selected_product_id")

        with self.get_session() as session:
            row = (
                session.query(TempProductSearchImport)
                .filter(TempProductSearchImport.id == row_id)
                .first()
            )
            return row.selected_product_id if row else None

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

    def build_product_combo(self, row_id: int, selected_product_id: int | None, selected_name: str) -> QComboBox:
        combo = QComboBox()
        combo.setEditable(True)
        combo.setInsertPolicy(QComboBox.NoInsert)
        combo.setProperty("row_id", row_id)
        combo.setProperty("combo_role", "product_combo")
        combo.setToolTip("Выберите продукт из базы или впишите вручную")

        combo.addItem("", None)
        added_ids: set[int] = set()

        for product in self.get_filtered_products():
            combo.addItem(product.name, product.id)
            added_ids.add(product.id)

        if selected_product_id and selected_product_id not in added_ids and selected_name:
            combo.addItem(selected_name, selected_product_id)

        index = combo.findData(selected_product_id)
        if index >= 0:
            combo.setCurrentIndex(index)
        else:
            combo.setCurrentText(selected_name or "")

        return combo

    def build_brand_combo(self, row_id: int, brand_name: str) -> QComboBox:
        combo = QComboBox()
        combo.setEditable(True)
        combo.setInsertPolicy(QComboBox.NoInsert)
        combo.setProperty("row_id", row_id)
        combo.setProperty("combo_role", "brand_combo")
        combo.setToolTip("Выберите бренд из базы или впишите вручную")

        combo.addItem("")
        for brand in self.get_brand_names():
            combo.addItem(brand)

        if brand_name and combo.findText(brand_name) < 0:
            combo.addItem(brand_name)

        combo.setCurrentText(brand_name)
        return combo

    def _resolve_combo_product(self, combo: QComboBox):
        text = clean_multi_spaces(combo.currentText())
        if not text:
            return None, ""

        for index in range(combo.count()):
            item_text = clean_multi_spaces(combo.itemText(index))
            if item_text.lower() == text.lower():
                return combo.itemData(index), combo.itemText(index)

        return None, text

    def finish_product_edit(self, row: int, row_id: int, combo: QComboBox):
        if self.table.cellWidget(row, self.COL_PRODUCT) is not combo:
            return

        product_id, product_name = self._resolve_combo_product(combo)

        self._pending_changes.setdefault(row_id, {})
        self._pending_changes[row_id]["selected_product_id"] = product_id

        self.table.removeCellWidget(row, self.COL_PRODUCT)
        self._updating_table = True
        self.table.setItem(
            row,
            self.COL_PRODUCT,
            self.build_table_item("selected_product_id", product_name, editable=False, align_left=True),
        )
        self._updating_table = False
        self.table.resizeColumnsToContents()

    def finish_brand_edit(self, row: int, row_id: int, combo: QComboBox):
        if self.table.cellWidget(row, self.COL_BRAND) is not combo:
            return

        brand_name = clean_multi_spaces(combo.currentText()).upper() or None

        self._pending_changes.setdefault(row_id, {})
        self._pending_changes[row_id]["new_brand"] = brand_name

        self.table.removeCellWidget(row, self.COL_BRAND)
        self._updating_table = True
        self.table.setItem(
            row,
            self.COL_BRAND,
            self.build_table_item("new_brand", brand_name or "", editable=False, align_left=True),
        )
        self._updating_table = False
        self.table.resizeColumnsToContents()

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
        if not column_name or column_name in {"selected_product_id", "new_brand"}:
            return

        value = clean_multi_spaces(item.text()).upper()
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
                text_rows.append("\t".join(value for _, value in sorted(cols.items())))
            clipboard.setText("\n".join(text_rows).strip())

        self.show_message("Скопировано")

    def delete_selected_row(self):
        row = self.table.currentRow()
        if row < 0 or row >= len(self._table_row_ids):
            self.show_error_message("Не выбрана строка")
            return

        row_id = self._table_row_ids[row]
        self._pending_deletes.add(row_id)
        self.apply_pending_changes(save_to_db_only=True)

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
                self.table.setCurrentCell(0, self.COL_PRODUCT)
            self.show_message("Добавлена строка")
        except Exception as e:
            self.show_error_message(str(e))

    def _build_export_rows(self) -> list[dict]:
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

            data: list[dict] = []
            for row in rows:
                product = row.selected_product
                data.append(
                    {
                        "source_article": row.source_article or "",
                        "source_product_name": row.source_product_name or "",
                        "product_name": product.name if product else (row.new_product_name or ""),
                        "brand": product.brand if product else (row.new_brand or ""),
                        "pack": product.pack if product else row.new_pack,
                        "is_excise": product.is_excise if product else row.new_is_excise,
                    }
                )
            return data


    def _commit_open_editors(self):
        for row in range(self.table.rowCount()):
            for column in (self.COL_PRODUCT, self.COL_BRAND):
                widget = self.table.cellWidget(row, column)
                if not isinstance(widget, QComboBox):
                    continue
                if row < 0 or row >= len(self._table_row_ids):
                    continue
                row_id = self._table_row_ids[row]
                if column == self.COL_PRODUCT:
                    self.finish_product_edit(row, row_id, widget)
                elif column == self.COL_BRAND:
                    self.finish_brand_edit(row, row_id, widget)

    def apply_pending_changes(self, save_to_db_only: bool = False):
        self._commit_open_editors()

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

                    has_new_product_data = any([
                        bool(clean_multi_spaces(row.new_product_name)),
                        bool(clean_multi_spaces(row.new_brand)),
                        row.new_pack is not None,
                    ])
                    if row.selected_product_id is None and has_new_product_data and row.new_is_excise is None:
                        row.new_is_excise = False

                    if row.selected_product_id is not None:
                        row.new_product_name = None
                        row.new_brand = None
                        row.new_pack = None
                        row.new_is_excise = None

                session.flush()

                if self._pending_deletes:
                    session.query(TempProductSearchImport).filter(
                        TempProductSearchImport.id.in_(self._pending_deletes)
                    ).delete(synchronize_session=False)

                if not save_to_db_only:
                    service.validate_new_products_before_save(self.batch_id, self.imported_by)
                    service.create_products_from_temp(self.batch_id, self.imported_by)
                    service.create_or_update_product_articles(self.batch_id, self.imported_by)

                session.commit()

            if not save_to_db_only:
                save_path, _ = QFileDialog.getSaveFileName(
                    self,
                    "Сохранить итоговый файл",
                    str(BASE_DIR / f"ProductSearch_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"),
                    "Excel files (*.xlsx)",
                )
                if not save_path:
                    self.show_message("Данные сохранены в БД. Сохранение файла отменено")
                else:
                    if not save_path.lower().endswith(".xlsx"):
                        save_path += ".xlsx"

                    try:
                        exporter = ProductSearchExporter()
                        exporter.export_result(save_path, self._build_export_rows())
                        self.show_message("Данные сохранены")
                    except PermissionError:
                        self.show_popup_error(
                            "Не удалось сохранить файл.\n\n"
                            "Скорее всего, файл уже открыт в Excel.\n"
                            "Закрой файл и попробуй снова."
                        )

            self._pending_changes.clear()
            self._pending_deletes.clear()
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
            self.load_table_rows()
            self.show_message("Форма очищена")
        except Exception as e:
            self.show_error_message(str(e))
