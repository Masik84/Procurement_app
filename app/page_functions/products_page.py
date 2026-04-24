from __future__ import annotations

from pathlib import Path
from decimal import Decimal, InvalidOperation

import re
import pythoncom
import win32com.client as win32
from sqlalchemy.exc import SQLAlchemyError
from PySide6.QtWidgets import (
    QMessageBox,
    QMenu,
    QTableWidgetItem,
    QWidget,
    QApplication,
    QVBoxLayout,
    QCheckBox,
    QHBoxLayout,
    QComboBox,
    QFileDialog,
)
from PySide6.QtCore import Qt, QFile, QEvent
from PySide6.QtGui import QDesktopServices
from PySide6.QtUiTools import QUiLoader

from app.db.models import Product
from app.db.db import SessionLocal
from app.imports.product_importer import ProductImporter
from app.ui.table_style import *
from app.utils.parsers import parse_loose_number
from app.utils.text import clean_multi_spaces
from app.exports.product_exporter import ProductExporter


BASE_DIR = Path(__file__).resolve().parents[2]
PRODUCTS_UI = BASE_DIR / "app" / "ui" / "windows" / "products.ui"


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


class ProductsPage(QWidget):
    def __init__(self):
        super().__init__()

        self.ui = load_ui(PRODUCTS_UI)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.ui)

        self._updating_table = False
        self._original_values: dict[int, dict] = {}
        self._pending_changes: dict[int, dict] = {}
        self._pending_deletes: set[int] = set()
        self._new_rows: set[int] = set()
        self._temp_row_id = -1

        self.columns = ["id", "name", "brand", "pack", "is_excise", "family"]
        self.headers = ["id", "Product name", "Brand", "Pack", "Excise duty", "Product Family"]
        self.header_to_column = dict(zip(self.headers, self.columns))
        self.text_columns = {"name", "brand", "family"}

        self.setup_ui()
        self.setup_connections()
        self.refresh_all_comboboxes()

    def setup_ui(self):
        self.table = self.ui.table
        setup_data_table(self.table, sorting=True)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_context_menu)
        self.table.cellDoubleClicked.connect(self.on_cell_double_clicked)

    def setup_connections(self):
        self.table.itemChanged.connect(self.on_item_changed)
        self.ui.line_Brand.currentTextChanged.connect(self.fill_in_prod_fam_list)
        self.ui.line_Prod_Fam.currentTextChanged.connect(self.fill_in_prod_name_list)

        self.ui.btn_Search.clicked.connect(self.find_Product)
        self.ui.btn_AddLine.clicked.connect(self.add_line)
        self.ui.btn_Save.clicked.connect(self.apply_pending_changes)
        self.ui.btn_SaveExcel.clicked.connect(self.save_to_excel)

        if hasattr(self.ui, "btn_DownFile"):
            self.ui.btn_DownFile.clicked.connect(self.download_template)
        if hasattr(self.ui, "btn_Import"):
            self.ui.btn_Import.clicked.connect(self.import_products)
        if hasattr(self.ui, "btn_Reset"):
            self.ui.btn_Reset.clicked.connect(self.reset_form)

    def get_session(self):
        return SessionLocal()

    def show_context_menu(self, position):
        menu = QMenu()
        copy_action = menu.addAction("Копировать")
        delete_action = menu.addAction("Удалить строку")
        apply_action = menu.addAction("Применить изменения")
        revert_action = menu.addAction("Отменить изменения")

        copy_action.triggered.connect(self.copy_cell_content)
        delete_action.triggered.connect(self.delete_selected_row)
        apply_action.triggered.connect(self.apply_pending_changes)
        revert_action.triggered.connect(self.revert_changes)

        menu.exec_(self.table.viewport().mapToGlobal(position))

    def on_item_changed(self, item):
        if self._updating_table:
            return

        try:
            row = item.row()
            column = item.column()
            header_item = self.table.horizontalHeaderItem(column)
            id_item = self.table.item(row, 0)

            if not header_item or not id_item:
                return

            row_id_text = id_item.text().strip()
            if not row_id_text:
                return

            row_id = int(row_id_text)
            header = header_item.text()
            column_name = self.header_to_column.get(header)

            if not column_name or column_name == "id":
                return

            new_value = item.text()

            if row_id not in self._pending_changes:
                self._pending_changes[row_id] = {}

            self._pending_changes[row_id][column_name] = new_value

        except Exception as e:
            self.show_error_message(f"Ошибка: {str(e)}")

    def on_checkbox_changed(self, row_id, checked):
        if self._updating_table:
            return

        try:
            if row_id not in self._pending_changes:
                self._pending_changes[row_id] = {}

            self._pending_changes[row_id]["is_excise"] = bool(checked)
        except Exception as e:
            self.show_error_message(f"Ошибка: {str(e)}")

    def copy_cell_content(self):
        selected_items = self.table.selectedItems()
        if not selected_items:
            return

        clipboard = QApplication.clipboard()

        if len(selected_items) == 1:
            item = selected_items[0]
            text = self._get_cell_display_text(item.row(), item.column())
        else:
            rows = {}
            for item in selected_items:
                row = item.row()
                col = item.column()
                rows.setdefault(row, {})
                rows[row][col] = self._get_cell_display_text(row, col)

            text_rows = []
            for _, cols in sorted(rows.items()):
                text_rows.append("\t".join(value for _, value in sorted(cols.items())))
            text = "\n".join(text_rows)

        clipboard.setText(text.strip())
        self.show_message("Скопировано")

    def _get_cell_display_text(self, row, col):
        if self.columns[col] == "is_excise":
            checkbox = self._get_checkbox_from_cell(row, col)
            return "Да" if checkbox and checkbox.isChecked() else "Нет"

        item = self.table.item(row, col)
        return item.text() if item else ""

    def _get_checkbox_from_cell(self, row, col):
        container = self.table.cellWidget(row, col)
        if not container:
            return None

        checkbox = container.findChild(QCheckBox)
        return checkbox

    def delete_selected_row(self):
        row = self.table.currentRow()
        if row < 0:
            self.show_error_message("Не выбрана строка для удаления")
            return

        id_item = self.table.item(row, 0)
        if not id_item:
            self.show_error_message("Не удалось определить ID строки")
            return

        row_id_text = id_item.text().strip()
        if not row_id_text:
            self.show_error_message("Не удалось определить ID строки")
            return

        row_id = int(row_id_text)

        if row_id in self._new_rows:
            self._new_rows.discard(row_id)
            self._pending_changes.pop(row_id, None)
        else:
            self._pending_deletes.add(row_id)

        self.table.removeRow(row)
        self.show_message("Строка помечена на удаление")

    def revert_changes(self):
        try:
            self._pending_changes.clear()
            self._pending_deletes.clear()
            self._new_rows.clear()

            if self.has_active_filters():
                self.find_Product()
            else:
                self.table.clearContents()
                self.table.setRowCount(0)

            self.show_message("Изменения отменены")

        except Exception as e:
            self.show_error_message(f"Ошибка отката: {str(e)}")

    def apply_pending_changes(self):
        if not self._pending_changes and not self._pending_deletes:
            self.show_message("Нет изменений для применения")
            return

        try:
            with self.get_session() as session:
                if self._pending_deletes:
                    session.query(Product).filter(Product.id.in_(self._pending_deletes)).delete(
                        synchronize_session=False
                    )

                for row_id, changes in self._pending_changes.items():
                    if row_id in self._new_rows:
                        self._insert_or_update_imported_product(session, row_id, changes)
                    else:
                        self._update_product(session, row_id, changes)

                session.commit()

            self._pending_changes.clear()
            self._pending_deletes.clear()
            self._new_rows.clear()

            self.refresh_all_comboboxes()

            if self.has_active_filters():
                self.find_Product()
            else:
                self.table.clearContents()
                self.table.setRowCount(0)

            self.show_message("Данные успешно сохранены")

        except SQLAlchemyError as e:
            self.show_error_message(f"Ошибка сохранения в базу данных: {str(e)}")
        except Exception as e:
            self.show_error_message(f"Ошибка применения изменений: {str(e)}")

    def _normalize_product_changes(self, changes: dict) -> dict:
        name = clean_multi_spaces(changes.get("name", "")).upper()
        brand = clean_multi_spaces(changes.get("brand", "")).upper()
        family = clean_multi_spaces(changes.get("family", "")).upper()
        pack = self._to_decimal(changes.get("pack", ""), "Pack")
        is_excise = bool(changes.get("is_excise", False))

        if not name:
            raise Exception("Для продукта поле Product name обязательно")
        if not brand:
            raise Exception(f"Для '{name}' поле Brand обязательно")
        if pack is None:
            raise Exception(f"Для '{name}' поле Pack обязательно")

        family_calc = self._build_family_from_name(name, pack)
        if family and family != family_calc:
            raise Exception(
                f"Для '{name}' неверно заполнено family. Ожидается '{family_calc}'."
            )

        return {
            "name": name,
            "brand": brand,
            "pack": pack,
            "is_excise": is_excise,
            "family": family_calc,
        }

    def _find_existing_product_for_import(self, session, *, name: str, brand: str, pack: Decimal):
        product = (
            session.query(Product)
            .filter(
                Product.name == name,
                Product.brand == brand,
                Product.pack == pack,
            )
            .first()
        )
        if product is not None:
            return product

        return (
            session.query(Product)
            .filter(Product.name == name)
            .order_by(Product.id.asc())
            .first()
        )

    def _insert_or_update_imported_product(self, session, row_id, changes):
        data = self._normalize_product_changes(changes)
        existing = self._find_existing_product_for_import(
            session,
            name=data["name"],
            brand=data["brand"],
            pack=data["pack"],
        )

        if existing:
            existing.name = data["name"]
            existing.brand = data["brand"]
            existing.pack = data["pack"]
            existing.is_excise = data["is_excise"]
            existing.family = data["family"]
            return

        product = Product(
            name=data["name"],
            brand=data["brand"],
            pack=data["pack"],
            is_excise=data["is_excise"],
            family=data["family"],
        )
        session.add(product)

    def _update_product(self, session, row_id, changes):
        product = session.query(Product).filter(Product.id == row_id).first()
        if not product:
            raise Exception(f"Не найден продукт id={row_id}")

        merged_changes = {
            "name": changes.get("name", product.name or ""),
            "brand": changes.get("brand", product.brand or ""),
            "pack": changes.get("pack", product.pack),
            "is_excise": changes.get("is_excise", bool(product.is_excise)),
            "family": changes.get("family", product.family or ""),
        }
        data = self._normalize_product_changes(merged_changes)

        duplicate = (
            session.query(Product)
            .filter(
                Product.id != row_id,
                Product.name == data["name"],
                Product.brand == data["brand"],
                Product.pack == data["pack"],
            )
            .first()
        )
        if duplicate:
            raise Exception(
                f"Продукт '{data['name']}' / '{data['brand']}' / '{self._format_decimal_display(data['pack'])}' уже существует"
            )

        product.name = data["name"]
        product.brand = data["brand"]
        product.pack = data["pack"]
        product.is_excise = data["is_excise"]
        product.family = data["family"]

    def start_brand_edit(self, row: int):
        id_item = self.table.item(row, 0)
        if not id_item:
            return

        row_id_text = id_item.text().strip()
        if not row_id_text:
            return

        row_id = int(row_id_text)
        brand_col = self.columns.index("brand")

        current_item = self.table.item(row, brand_col)
        current_value = current_item.text().strip() if current_item else ""

        combo = QComboBox(self.table)
        combo.setEditable(True)
        combo.setInsertPolicy(QComboBox.NoInsert)
        combo.setFrame(False)

        brands = self._get_brand_values()
        combo.addItems(brands)

        if current_value and combo.findText(current_value) < 0:
            combo.addItem(current_value)

        combo.setCurrentText(current_value)

        combo.setProperty("edit_row", row)
        combo.setProperty("edit_row_id", row_id)

        combo.activated.connect(lambda *_, c=combo: self.finish_brand_edit_from_combo(c))
        combo.lineEdit().returnPressed.connect(lambda c=combo: self.finish_brand_edit_from_combo(c))

        self._updating_table = True
        self.table.setCellWidget(row, brand_col, combo)
        self._updating_table = False

        combo.setFocus()
        combo.lineEdit().selectAll()

    def on_cell_double_clicked(self, row, column):
        if self._updating_table:
            return

        if column < 0 or column >= len(self.columns):
            return

        column_name = self.columns[column]

        if column_name == "brand":
            self.start_brand_edit(row)

    def finish_brand_edit_from_combo(self, combo: QComboBox):
        row = combo.property("edit_row")
        row_id = combo.property("edit_row_id")

        if row is None or row_id is None:
            return

        brand_col = self.columns.index("brand")
        text = clean_multi_spaces(combo.currentText()).upper()

        if row_id not in self._pending_changes:
            self._pending_changes[row_id] = {}

        self._pending_changes[row_id]["brand"] = text

        self._updating_table = True
        self.table.removeCellWidget(row, brand_col)
        self.table.setItem(row, brand_col, self._build_table_item("brand", text))
        self._updating_table = False

        self.table.resizeColumnsToContents()

    def eventFilter(self, obj, event):
        if isinstance(obj, QComboBox) and obj.property("edit_row_id") is not None:
            if event.type() == QEvent.FocusOut:
                self.finish_brand_edit_from_combo(obj)
                return False

        return super().eventFilter(obj, event)

    def finish_brand_edit(self, row: int, row_id: int, combo: QComboBox):
        text = clean_multi_spaces(combo.currentText()).upper()

        if row_id not in self._pending_changes:
            self._pending_changes[row_id] = {}

        self._pending_changes[row_id]["brand"] = text

        self._updating_table = True
        self.table.removeCellWidget(row, self.columns.index("brand"))
        self.table.setItem(
            row,
            self.columns.index("brand"),
            self._build_table_item("brand", text),
        )
        self._updating_table = False

        self.table.resizeColumnsToContents()

    def _get_brand_values(self):
        with self.get_session() as session:
            rows = (
                session.query(Product.brand)
                .filter(Product.brand.isnot(None), Product.brand != "")
                .distinct()
                .order_by(Product.brand)
                .all()
            )
        return [row[0] for row in rows if row[0]]

    def refresh_all_comboboxes(self):
        self.fill_in_prod_brand_list()
        self.fill_in_prod_fam_list()
        self.fill_in_prod_name_list()

    def fill_in_prod_brand_list(self):
        try:
            with self.get_session() as session:
                brands = (
                    session.query(Product.brand)
                    .filter(Product.brand.isnot(None), Product.brand != "")
                    .distinct()
                    .order_by(Product.brand)
                    .all()
                )

            self._fill_combobox(self.ui.line_Brand, [row[0] for row in brands if row[0]])

        except Exception as e:
            self.show_error_message(f"Ошибка при получении брендов: {str(e)}")

    def fill_in_prod_fam_list(self):
        brand = self.ui.line_Brand.currentText()

        try:
            with self.get_session() as session:
                query = session.query(Product.family).filter(
                    Product.family.isnot(None),
                    Product.family != ""
                )
                if brand != "-":
                    query = query.filter(Product.brand == brand)

                families = query.distinct().order_by(Product.family).all()

            families = [row[0] for row in families if row[0]]
            current_value = self.ui.line_Prod_Fam.currentText()
            self._fill_combobox(self.ui.line_Prod_Fam, families)

            if current_value in families:
                self.ui.line_Prod_Fam.setCurrentText(current_value)

        except Exception as e:
            self.show_error_message(f"Ошибка при получении семейств: {str(e)}")
            self._fill_combobox(self.ui.line_Prod_Fam, [])

    def fill_in_prod_name_list(self):
        brand = self.ui.line_Brand.currentText()
        family = self.ui.line_Prod_Fam.currentText()

        try:
            with self.get_session() as session:
                query = session.query(Product).filter(Product.name.isnot(None), Product.name != "")

                if brand != "-":
                    query = query.filter(Product.brand == brand)
                if family != "-":
                    query = query.filter(Product.family == family)

                products = self._sort_products(query.all())
                product_names = [row.name for row in products if row.name]

            current_value = self.ui.line_Prod_name.currentText()
            self._fill_combobox(self.ui.line_Prod_name, product_names)

            if current_value in product_names:
                self.ui.line_Prod_name.setCurrentText(current_value)

        except Exception as e:
            self.show_error_message(f"Ошибка при получении продуктов: {str(e)}")
            self._fill_combobox(self.ui.line_Prod_name, [])

    def _fill_combobox(self, combobox, items):
        combobox.blockSignals(True)
        combobox.clear()
        combobox.addItem("-")
        if items:
            combobox.addItems(sorted(items))
        combobox.blockSignals(False)

    def _sort_products(self, products):
        return sorted(
            products,
            key=lambda x: (
                (x.family or "").lower(),
                -(float(x.pack) if str(x.pack).replace(".", "", 1).isdigit() else -999999)
            )
        )

    def get_Products_from_db(self):
        with self.get_session() as session:
            rows = self._sort_products(session.query(Product).all())

            data = []
            for row in rows:
                data.append({
                    "id": row.id,
                    "name": row.name,
                    "brand": row.brand,
                    "pack": row.pack,
                    "is_excise": bool(row.is_excise),
                    "family": row.family,
                })

        return data

    def _get_find_product_text(self):
        line_find = getattr(self.ui, "line_FindProduct", None)
        if line_find is None:
            return ""
        return line_find.text().strip().upper()

    def get_filtered_products(self):
        prod_data = self.get_Products_from_db()

        if not prod_data:
            return []

        brand = self.ui.line_Brand.currentText().strip()
        family = self.ui.line_Prod_Fam.currentText().strip()
        product_name = self.ui.line_Prod_name.currentText().strip()
        find_product_text = self._get_find_product_text()

        if brand != "-":
            prod_data = [row for row in prod_data if (row["brand"] or "") == brand]
        if family != "-":
            prod_data = [row for row in prod_data if (row["family"] or "") == family]
        if product_name != "-":
            prod_data = [row for row in prod_data if (row["name"] or "") == product_name]
        if find_product_text:
            prod_data = [
                row for row in prod_data
                if find_product_text in (row["name"] or "").upper()
            ]

        return prod_data

    def find_Product(self):
        self.table.clearContents()
        self.table.setRowCount(0)

        prod_data = self.get_filtered_products()

        if not prod_data:
            self._display_data([])
            self.show_message("Нет данных по заданным фильтрам")
            return

        self._display_data(prod_data)

    def download_template(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Сохранить шаблон",
            str(Path.home() / "ProductImportTemplate.xlsx"),
            "Excel files (*.xlsx)",
        )
        if not file_path:
            return

        if not file_path.lower().endswith(".xlsx"):
            file_path += ".xlsx"

        try:
            exporter = ProductExporter()
            exporter.export_template(file_path)
            QDesktopServices.openUrl(Path(file_path).as_uri())
            self.show_message("Шаблон сохранен")
        except PermissionError:
            self.show_error_message(
                "Не удалось сохранить файл Excel. Возможно, файл уже открыт."
            )
        except Exception as e:
            self.show_error_message(f"Ошибка при создании шаблона: {str(e)}")

    def import_products(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите файл продуктов",
            "",
            "Excel files (*.xls *.xlsx)",
        )
        if not file_path:
            return

        try:
            importer = ProductImporter()
            rows = importer.read_excel(file_path)
            self._load_imported_rows_to_table(rows)
            self.show_message(f"Импортировано строк: {len(rows)}")
        except Exception as e:
            self.show_error_message(str(e))

    def _load_imported_rows_to_table(self, rows: list[dict]):
        self._updating_table = True
        self.table.setSortingEnabled(False)

        self._pending_changes.clear()
        self._pending_deletes.clear()
        self._new_rows.clear()
        self._original_values.clear()
        self.table.clear()
        self.table.setColumnCount(len(self.headers))
        self.table.setRowCount(len(rows))
        self.table.setHorizontalHeaderLabels(self.headers)

        with self.get_session() as session:
            for row_index, imported in enumerate(rows):
                existing = self._find_existing_product_for_import(
                    session,
                    name=imported["name"],
                    brand=imported["brand"],
                    pack=Decimal(str(imported["pack"])),
                )

                if existing is not None:
                    row_id = int(existing.id)
                    self._pending_changes[row_id] = {
                        "name": imported["name"],
                        "brand": imported["brand"],
                        "pack": imported["pack"],
                        "is_excise": imported["is_excise"],
                        "family": imported["family"],
                    }
                else:
                    row_id = self._temp_row_id
                    self._temp_row_id -= 1
                    self._new_rows.add(row_id)
                    self._pending_changes[row_id] = {
                        "name": imported["name"],
                        "brand": imported["brand"],
                        "pack": imported["pack"],
                        "is_excise": imported["is_excise"],
                        "family": imported["family"],
                    }

                row_data = {
                    "id": row_id,
                    "name": imported["name"],
                    "brand": imported["brand"],
                    "pack": imported["pack"],
                    "is_excise": imported["is_excise"],
                    "family": imported["family"],
                }
                self._original_values[row_id] = row_data.copy()

                for col_index, col_name in enumerate(self.columns):
                    if col_name == "is_excise":
                        self.table.setCellWidget(
                            row_index,
                            col_index,
                            self._build_checkbox_widget(row_id, bool(row_data[col_name]))
                        )
                        continue

                    item = self._build_table_item(col_name, row_data[col_name])
                    self.table.setItem(row_index, col_index, item)

        self.table.resizeColumnsToContents()
        for i in range(self.table.columnCount()):
            if self.table.columnWidth(i) < 100:
                self.table.setColumnWidth(i, 100)

        self._updating_table = False

    def save_to_excel(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Сохранить Excel",
            "Products.xlsx",
            "Excel Files (*.xlsx)"
        )

        if not file_path:
            return

        if not file_path.lower().endswith(".xlsx"):
            file_path += ".xlsx"

        rows = self.get_filtered_products()
        if not rows:
            self.show_message("Нет данных для выгрузки")
            return

        try:
            self._export_products_to_excel(file_path, rows)
            self.show_message("Данные сохранены в Excel")
        except PermissionError:
            self.show_error_message(
                "Не удалось сохранить файл Excel. Возможно, файл уже открыт."
            )
        except Exception as e:
            self.show_error_message(f"Ошибка при сохранении Excel: {str(e)}")

    def _export_products_to_excel(self, file_path: str, rows: list[dict]):
        excel = None
        wb = None

        try:
            pythoncom.CoInitialize()

            excel = win32.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

            wb = excel.Workbooks.Add()
            ws = wb.Worksheets(1)
            ws.Name = "Sheet1"

            headers = [
                "ID",
                "Product name",
                "Brand",
                "Pack",
                "Excise duty",
                "Product Family",
            ]

            for col_index, header in enumerate(headers, start=1):
                ws.Cells(1, col_index).Value = header

            ws.Cells.Font.Name = "Aptos Narrow"
            ws.Cells.Font.Size = 11

            header_range = ws.Range("A1:F1")
            header_range.Font.Name = "Aptos Narrow"
            header_range.Font.Size = 11
            header_range.Font.Bold = True
            header_range.Interior.Color = 0xCDCDCD
            header_range.WrapText = True
            header_range.HorizontalAlignment = -4108
            header_range.VerticalAlignment = -4160

            ws.Rows(1).EntireRow.AutoFit()

            try:
                ws.Range("A1:F1").AutoFilter(1)
            except Exception:
                pass

            ws.Columns("A:A").ColumnWidth = 10
            ws.Columns("B:B").ColumnWidth = 30
            ws.Columns("C:C").ColumnWidth = 24
            ws.Columns("D:D").ColumnWidth = 12
            ws.Columns("E:E").ColumnWidth = 14
            ws.Columns("F:F").ColumnWidth = 24

            row_num = 2
            for row in rows:
                ws.Cells(row_num, 1).Value = row.get("id")
                ws.Cells(row_num, 2).Value = row.get("name", "") or ""
                ws.Cells(row_num, 3).Value = row.get("brand", "") or ""

                pack = row.get("pack")
                if pack not in (None, ""):
                    try:
                        ws.Cells(row_num, 4).Value = float(pack)
                    except Exception:
                        ws.Cells(row_num, 4).Value = str(pack)
                else:
                    ws.Cells(row_num, 4).Value = ""

                ws.Cells(row_num, 5).Value = "Да" if bool(row.get("is_excise")) else "Нет"
                ws.Cells(row_num, 6).Value = row.get("family", "") or ""
                row_num += 1

            # if row_num > 2:
            #     ws.Range(f"D2:D{row_num - 1}").NumberFormat = "General"

            wb.SaveAs(str(Path(file_path).resolve()))

        except PermissionError:
            raise
        except Exception as e:
            raise Exception(f"Ошибка при сохранении Excel: {e}")
        finally:
            try:
                if wb is not None:
                    wb.Close(SaveChanges=False)
            except Exception:
                pass

            try:
                if excel is not None:
                    excel.Quit()
            except Exception:
                pass

            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    def _display_data(self, data):
        self.table.clear()
        self.table.setColumnCount(0)
        self.table.setRowCount(0)

        if not data:
            self.show_message("Нет данных для отображения")
            return

        self._updating_table = True
        self._original_values.clear()

        self.table.setColumnCount(len(self.headers))
        self.table.setRowCount(len(data))
        self.table.setHorizontalHeaderLabels(self.headers)

        for row_index, row_data in enumerate(data):
            row_id = int(row_data["id"])
            self._original_values[row_id] = {}

            for col_index, col_name in enumerate(self.columns):
                if col_name == "is_excise":
                    checked = bool(row_data[col_name])
                    self._original_values[row_id][col_name] = checked
                    self.table.setCellWidget(
                        row_index,
                        col_index,
                        self._build_checkbox_widget(row_id, checked)
                    )
                    continue

                value = "" if row_data[col_name] is None else str(row_data[col_name])
                item = self._build_table_item(col_name, value)
                self._original_values[row_id][col_name] = value
                self.table.setItem(row_index, col_index, item)

        self.table.resizeColumnsToContents()

        for i in range(self.table.columnCount()):
            if self.table.columnWidth(i) < 100:
                self.table.setColumnWidth(i, 100)

        self._updating_table = False

    def _build_checkbox_widget(self, row_id, checked):
        checkbox = QCheckBox()
        checkbox.setChecked(checked)
        checkbox.setStyleSheet("""
            QCheckBox {
                background: transparent;
            }
            QCheckBox::indicator {
                width: 14px;
                height: 14px;
            }
        """)
        checkbox.toggled.connect(lambda state, rid=row_id: self.on_checkbox_changed(rid, state))

        container = QWidget()
        layout = QHBoxLayout(container)
        layout.addWidget(checkbox)
        layout.setAlignment(Qt.AlignCenter)
        layout.setContentsMargins(0, 0, 0, 0)
        container.setLayout(layout)

        return container

    def _build_table_item(self, col_name, value):
        display_value = "" if value is None else str(value)

        if col_name == "pack":
            parsed = parse_loose_number(value)
            if parsed is not None:
                display_value = self._format_decimal_display(parsed)
            item_text = format_table_value(display_value)
        else:
            item_text = display_value

        item = QTableWidgetItem(item_text)

        if col_name == "id":
            item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            item.setTextAlignment(Qt.AlignCenter)
            return item

        item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable)

        if col_name == "pack":
            item.setTextAlignment(Qt.AlignCenter)
        elif col_name in self.text_columns:
            item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        else:
            item.setTextAlignment(Qt.AlignCenter)

        return item

    def _to_decimal(self, value, field_name):
        if isinstance(value, Decimal):
            return value

        text = str(value).strip()
        if text == "":
            return None

        text = text.replace(",", ".")

        try:
            return Decimal(text)
        except (InvalidOperation, ValueError):
            raise Exception(f"Поле '{field_name}' должно быть числом")

    def _format_decimal_display(self, value) -> str:
        number = parse_loose_number(value)
        if number is None:
            return ""
        text = f"{float(number):.4f}".rstrip("0").rstrip(".")
        return text.replace(".", ",")

    def _build_family_from_name(self, name: str, pack: Decimal) -> str:
        product_name = clean_multi_spaces(name).upper()
        pack_value = parse_loose_number(pack)

        if not product_name:
            raise Exception("Не заполнено название продукта.")

        if pack_value is None:
            raise Exception("Поле 'Pack' должно быть числом.")

        matches = list(
            re.finditer(
                r"(?<!\d)([0-9]+(?:[.,][0-9]+)?)\s*(L|KG)\b",
                product_name,
                flags=re.IGNORECASE,
            )
        )

        for match in matches:
            found_num = parse_loose_number(match.group(1))
            if found_num is None:
                continue

            if float(found_num) == float(pack_value):
                family = product_name[:match.start()].strip()
                if not family:
                    raise Exception(
                        f"Для '{product_name}' не удалось определить family до упаковки."
                    )
                return family

        pack_display = self._format_decimal_display(pack)
        raise Exception(
            f"Для '{product_name}' проверь упаковку в названии. "
            f"Ожидается наличие '{pack_display}L' или '{pack_display}KG' внутри названия, "
            f"например: '... {pack_display}L BIB'."
        )

    def add_line(self):
        self.clear_message()
        self._updating_table = True

        self.table.setSortingEnabled(False)

        if self.table.columnCount() == 0:
            self.table.setColumnCount(len(self.headers))
            self.table.setHorizontalHeaderLabels(self.headers)

        self.table.insertRow(0)

        brand_value = self.ui.line_Brand.currentText()
        if brand_value == "-":
            brand_value = ""

        family_value = self.ui.line_Prod_Fam.currentText()
        if family_value == "-":
            family_value = ""

        row_id = self._temp_row_id
        self._temp_row_id -= 1
        self._new_rows.add(row_id)

        values = {
            "id": str(row_id),
            "name": "",
            "brand": brand_value,
            "pack": "",
            "is_excise": False,
            "family": family_value,
        }

        self._pending_changes[row_id] = {
            "name": "",
            "brand": brand_value,
            "pack": "",
            "is_excise": False,
            "family": family_value,
        }

        self._original_values[row_id] = values.copy()

        for col_index, col_name in enumerate(self.columns):
            if col_name == "is_excise":
                self.table.setCellWidget(0, col_index, self._build_checkbox_widget(row_id, False))
                continue

            item = self._build_table_item(col_name, values[col_name])
            self.table.setItem(0, col_index, item)

        for i in range(self.table.columnCount()):
            if self.table.columnWidth(i) < 100:
                self.table.setColumnWidth(i, 100)

        self._updating_table = False
        self.table.setCurrentCell(0, 1)
        self.show_message("Добавлена новая строка")

    def has_active_filters(self):
        return (
            self.ui.line_Brand.currentText() != "-"
            or self.ui.line_Prod_Fam.currentText() != "-"
            or self.ui.line_Prod_name.currentText() != "-"
            or bool(self._get_find_product_text())
        )

    def reset_form(self):
        try:
            self._pending_changes.clear()
            self._pending_deletes.clear()
            self._new_rows.clear()
            self._original_values.clear()
            self._temp_row_id = -1

            self.table.clearContents()
            self.table.setRowCount(0)

            self._fill_combobox(self.ui.line_Brand, [])
            self._fill_combobox(self.ui.line_Prod_Fam, [])
            self._fill_combobox(self.ui.line_Prod_name, [])
            line_find = getattr(self.ui, "line_FindProduct", None)
            if line_find is not None:
                line_find.clear()

            self.refresh_all_comboboxes()
            self.show_message("Форма очищена")
        except Exception as e:
            self.show_error_message(f"Ошибка очистки формы: {str(e)}")

    def show_message(self, text):
        self.ui.label_msg.setText(text)
        self.ui.label_msg.setProperty("active", True)
        self.ui.label_msg.style().unpolish(self.ui.label_msg)
        self.ui.label_msg.style().polish(self.ui.label_msg)
        self.ui.label_msg.setVisible(True)

    def clear_message(self):
        self.ui.label_msg.setText("")
        self.ui.label_msg.setProperty("active", False)
        self.ui.label_msg.style().unpolish(self.ui.label_msg)
        self.ui.label_msg.style().polish(self.ui.label_msg)
        self.ui.label_msg.setVisible(False)

    def show_error_message(self, text):
        msg = QMessageBox()
        msg.setWindowTitle("Ошибка")
        msg.setIcon(QMessageBox.Critical)
        msg.setMinimumSize(900, 600)

        if len(text) > 500:
            msg.setText("Произошла ошибка. Подробности ниже (используйте кнопку 'Show Details')")
            msg.setDetailedText(text)
        else:
            msg.setText(text)

        copy_button = msg.addButton("Copy", QMessageBox.ActionRole)
        msg.addButton(QMessageBox.Ok)

        def copy_text():
            QApplication.clipboard().setText(text)

        copy_button.clicked.connect(copy_text)
        msg.exec_()
