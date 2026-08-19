from __future__ import annotations

from pathlib import Path
from decimal import Decimal, InvalidOperation

import re
import pythoncom
import win32com.client as win32
from app.utils.excel_format_rules import save_workbook_xlsx
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
    QAbstractItemView,
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
from app.workers.excel_export_worker import start_excel_export
from app.exports.product_exporter import ProductExporter
from app.utils.output_headers import display_headers, standardize_output_header
from app.utils.excel_fast_writer import write_excel_table


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
        self._import_update_rows: set[int] = set()
        self._temp_row_id = -1
        self._excel_export_thread = None
        self._excel_export_worker = None

        self.columns = ["id", "name", "brand", "pack", "abc_category", "is_excise", "family"]
        self.headers = ["id", "Product name", "Brand", "Pack", "Категория ABC", "Excise duty", "Product Family"]
        self.header_to_column = dict(zip(self.headers, self.columns))
        self.text_columns = {"name", "brand", "family", "abc_category"}

        self.setup_ui()
        self.setup_connections()
        self.refresh_all_comboboxes()

    def setup_ui(self):
        self.table = self.ui.table
        setup_data_table(self.table, sorting=True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        from app.utils.gui_table_actions import install_standard_table_context_menu
        install_standard_table_context_menu(self, self.table)
        self.table.cellDoubleClicked.connect(self.on_cell_double_clicked)

    def setup_connections(self):
        self.table.itemChanged.connect(self.on_item_changed)
        self.ui.line_Brand.currentTextChanged.connect(self.on_brand_filter_changed)
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

    def on_brand_filter_changed(self):
        self.fill_in_prod_fam_list()
        self.fill_in_prod_name_list()

    def get_session(self):
        return SessionLocal()

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

    def revert_changes(self):
        try:
            self._pending_changes.clear()
            self._pending_deletes.clear()
            self._new_rows.clear()
            self._import_update_rows.clear()

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
                    session.flush()

                # Когда импорт пришел из Excel с заполненными ID, обновляем именно эти ID.
                # Если несколько продуктов меняют названия одновременно, PostgreSQL может увидеть
                # старое имя другого продукта как дубль. Поэтому сначала временно освобождаем
                # старые имена импортируемых строк, а потом ставим финальные значения.
                self._temporarily_free_import_update_names(session)

                for row_id, changes in self._pending_changes.items():
                    if row_id in self._new_rows:
                        self._insert_or_update_imported_product(session, row_id, changes)
                    else:
                        self._update_product(session, row_id, changes)

                session.commit()

            self._pending_changes.clear()
            self._pending_deletes.clear()
            self._new_rows.clear()
            self._import_update_rows.clear()

            self.refresh_all_comboboxes()
            self._reset_filter_controls_after_save()

            self.table.clearContents()
            self.table.setRowCount(0)

            self.show_message("Данные успешно сохранены")

        except SQLAlchemyError as e:
            self.show_error_message(f"Ошибка сохранения в базу данных: {str(e)}")
        except Exception as e:
            self.show_error_message(f"Ошибка применения изменений: {str(e)}")

    def _temporarily_free_import_update_names(self, session):
        """
        Excel-импорт с заполненными ID должен обновлять именно эти ID.

        В products.name есть уникальный индекс, поэтому при массовом переименовании
        возможны два конфликта:
        1) два импортируемых продукта меняются названиями местами;
        2) нужное новое название уже занято старым дублем, которого нет в Excel.

        Поэтому перед финальным UPDATE временно освобождаем:
        - старые названия всех строк, которые обновляются по ID из Excel;
        - названия сторонних дублей, если они совпадают с финальными именами импорта.
        """
        final_names_by_row_id: dict[int, str] = {}
        duplicate_names: dict[str, list[int]] = {}

        for row_id in list(self._import_update_rows):
            if row_id in self._pending_deletes or row_id in self._new_rows:
                continue

            changes = self._pending_changes.get(row_id)
            if not changes:
                continue

            new_name = clean_multi_spaces(changes.get("name", "")).upper()
            if not new_name:
                continue

            final_names_by_row_id[row_id] = new_name
            duplicate_names.setdefault(new_name, []).append(row_id)

        conflicts_inside_import = [
            f"{name}: {ids}"
            for name, ids in duplicate_names.items()
            if len(ids) > 1
        ]
        if conflicts_inside_import:
            raise Exception(
                "В Excel одно и то же Product name указано для нескольких разных ID:\n"
                + "\n".join(conflicts_inside_import)
            )

        if not final_names_by_row_id:
            return

        changed = False

        # 1. Освобождаем старые имена строк, которые точно обновляем по ID.
        for row_id, new_name in final_names_by_row_id.items():
            product = session.query(Product).filter(Product.id == row_id).first()
            if product is None:
                continue

            old_name = clean_multi_spaces(product.name or "").upper()
            if old_name and old_name != new_name:
                product.name = f"__TMP_IMPORT_RENAME_{row_id}__"
                changed = True

        if changed:
            session.flush()

        # 2. Если финальное имя занято другим продуктом не из импортируемых ID,
        # временно переименовываем этот старый дубль, чтобы обновление по ID прошло.
        target_ids = set(final_names_by_row_id.keys())
        for target_id, final_name in final_names_by_row_id.items():
            duplicate = (
                session.query(Product)
                .filter(Product.id != target_id, Product.name == final_name)
                .first()
            )
            if duplicate is None:
                continue

            if int(duplicate.id) in target_ids:
                # Этот продукт тоже будет обновлен из Excel, его имя уже освобождено выше.
                continue

            duplicate.name = f"__OLD_DUPLICATE_PRODUCT_{int(duplicate.id)}__"
            changed = True

        if changed:
            session.flush()

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
            )
            .first()
        )
        if duplicate:
            raise Exception(
                f"Нельзя сохранить продукт id={row_id} с названием '{data['name']}', "
                f"потому что такое название уже есть у продукта id={duplicate.id}."
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
                    "abc_category": row.abc_category or "-",
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
            str(BASE_DIR / "ProductImportTemplate.xlsx"),
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
        self._import_update_rows.clear()
        self._original_values.clear()
        self.table.clear()
        self.table.setColumnCount(len(self.headers))
        self.table.setRowCount(len(rows))
        self.table.setHorizontalHeaderLabels(display_headers(self.headers))

        with self.get_session() as session:
            for row_index, imported in enumerate(rows):
                imported_id = imported.get("id")
                existing = None

                # Если в Excel есть ID, обновляем именно этот продукт,
                # даже если в Excel изменили Product name / Brand / Pack.
                if imported_id is not None:
                    existing = (
                        session.query(Product)
                        .filter(Product.id == int(imported_id))
                        .first()
                    )

                # Если ID пустой или такого ID в БД нет, ищем как раньше:
                # сначала по name + brand + pack, потом по name.
                if existing is None:
                    existing = self._find_existing_product_for_import(
                        session,
                        name=imported["name"],
                        brand=imported["brand"],
                        pack=Decimal(str(imported["pack"])),
                    )

                imported_changes = {
                    "name": imported["name"],
                    "brand": imported["brand"],
                    "pack": imported["pack"],
                    "abc_category": (existing.abc_category or "-") if existing is not None else "-",
                    "is_excise": imported["is_excise"],
                    "family": imported["family"],
                }

                if existing is not None:
                    row_id = int(existing.id)
                    self._pending_changes[row_id] = imported_changes
                    if imported_id is not None and int(imported_id) == row_id:
                        self._import_update_rows.add(row_id)
                else:
                    row_id = self._temp_row_id
                    self._temp_row_id -= 1
                    self._new_rows.add(row_id)
                    self._pending_changes[row_id] = imported_changes

                row_data = {
                    "id": row_id,
                    "name": imported["name"],
                    "brand": imported["brand"],
                    "pack": imported["pack"],
                    "abc_category": (existing.abc_category or "-") if existing is not None else "-",
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
        self.table.setSortingEnabled(True)

    def save_to_excel(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Сохранить Excel",
            str(BASE_DIR / "Products.xlsx"),
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

        def do_export():
            self._export_products_to_excel(file_path, rows)
            return file_path

        def done(output_path):
            QDesktopServices.openUrl(Path(output_path).as_uri())
            self.show_message("Данные сохранены в Excel")

        def error(text):
            if "Permission" in str(text):
                self.show_error_message("Не удалось сохранить файл Excel. Возможно, файл уже открыт.")
            else:
                self.show_error_message(f"Ошибка при сохранении Excel: {text}")

        if not start_excel_export(self, do_export, on_finished=done, on_error=error):
            self.show_message("Excel файл уже формируется. Можно продолжать работать в программе.")
        else:
            self.show_message("Excel файл формируется в фоновом режиме. Можно продолжать работать в программе.")

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
                "Категория ABC",
                "Excise duty",
                "Product Family",
            ]

            def value_for_header(row, header, _col_index):
                if header == "ID":
                    return row.get("id")
                if header == "Product name":
                    return row.get("name", "") or ""
                if header == "Brand":
                    return row.get("brand", "") or ""
                if header == "Pack":
                    pack = row.get("pack")
                    if pack not in (None, ""):
                        try:
                            return float(pack)
                        except Exception:
                            return str(pack)
                    return ""
                if header == "Категория ABC":
                    return row.get("abc_category", "-") or "-"
                if header == "Excise duty":
                    return "Да" if bool(row.get("is_excise")) else "Нет"
                if header == "Product Family":
                    return row.get("family", "") or ""
                return ""

            write_excel_table(ws, headers, rows, header_getter=standardize_output_header, value_getter=value_for_header)

            ws.Cells.Font.Name = "Aptos Narrow"
            ws.Cells.Font.Size = 11

            header_range = ws.Range("A1:G1")
            header_range.Font.Name = "Aptos Narrow"
            header_range.Font.Size = 11
            header_range.Font.Bold = True
            header_range.Interior.Color = 0xCDCDCD
            header_range.WrapText = True
            header_range.HorizontalAlignment = -4108
            header_range.VerticalAlignment = -4160

            ws.Rows(1).EntireRow.AutoFit()

            try:
                ws.Range("A1:G1").AutoFilter(1)
            except Exception:
                pass

            ws.Columns("A:A").ColumnWidth = 10
            ws.Columns("B:B").ColumnWidth = 30
            ws.Columns("C:C").ColumnWidth = 24
            ws.Columns("D:D").ColumnWidth = 12
            ws.Columns("E:E").ColumnWidth = 12
            ws.Columns("F:F").ColumnWidth = 14
            ws.Columns("G:G").ColumnWidth = 24

            # if rows:
            #     ws.Range(f"D2:D{len(rows) + 1}").NumberFormat = "General"

            save_workbook_xlsx(wb, file_path)

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
        self.table.setSortingEnabled(False)
        self.table.clear()
        self.table.setColumnCount(0)
        self.table.setRowCount(0)

        if not data:
            self.table.setSortingEnabled(True)
            self.show_message("Нет данных для отображения")
            return

        self._updating_table = True
        self._original_values.clear()

        self.table.setColumnCount(len(self.headers))
        self.table.setRowCount(len(data))
        self.table.setHorizontalHeaderLabels(display_headers(self.headers))

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
        self.table.setSortingEnabled(True)

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

        if col_name in {"id", "abc_category"}:
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
            self.table.setHorizontalHeaderLabels(display_headers(self.headers))

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
            "abc_category": "-",
            "is_excise": False,
            "family": family_value,
        }

        self._pending_changes[row_id] = {
            "name": "",
            "brand": brand_value,
            "pack": "",
            "abc_category": "-",
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
        self.table.setSortingEnabled(True)
        self.table.setCurrentCell(0, 1)
        self.show_message("Добавлена новая строка")

    def _reset_filter_controls_after_save(self):
        for combo_name in ("line_Brand", "line_Prod_Fam", "line_Prod_name"):
            combo = getattr(self.ui, combo_name, None)
            if combo is None:
                continue

            combo.blockSignals(True)
            try:
                index = combo.findText("-")
                if index >= 0:
                    combo.setCurrentIndex(index)
                else:
                    combo.setCurrentText("-")
            finally:
                combo.blockSignals(False)

        line_find = getattr(self.ui, "line_FindProduct", None)
        if line_find is not None:
            line_find.clear()

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
            self._import_update_rows.clear()
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
