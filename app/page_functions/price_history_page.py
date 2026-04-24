from __future__ import annotations

from pathlib import Path
from datetime import datetime
from decimal import Decimal, InvalidOperation

from sqlalchemy.exc import SQLAlchemyError
from PySide6.QtWidgets import (
    QMessageBox,
    QMenu,
    QTableWidgetItem,
    QWidget,
    QApplication,
    QVBoxLayout,
    QComboBox,
    QFileDialog,
    QAbstractItemView,
)
from PySide6.QtCore import Qt, QFile
from PySide6.QtUiTools import QUiLoader

from app.db.models import Product, Supplier, PriceHistory, CurrentSupplierPrice
from app.db.db import SessionLocal
from app.ui.table_style import *
from app.imports.price_history_importer import PriceHistoryImporter
from app.exports.price_history_exporter import PriceHistoryExporter
from app.services.product_matching_service import ProductMatchingService
from app.utils.text import clean_multi_spaces, normalize_product_name
from app.utils.batch import get_current_username
from app.db.models import CurrentSupplierPrice, PriceHistory, Product, Supplier
from app.db.models import PriceHistory as PriceHistoryModel
from app.db.models import CurrentSupplierPrice as CurrentSupplierPriceModel


BASE_DIR = Path(__file__).resolve().parents[2]
PRICE_HISTORY_UI = BASE_DIR / "app" / "ui" / "windows" / "price_history.ui"


MODE_NONE = "-"
MODE_CURRENT = "Последние цены"
MODE_HISTORY = "История цен (вся)"


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


class PriceHistoryPage(QWidget):
    def __init__(self):
        super().__init__()

        self.ui = load_ui(PRICE_HISTORY_UI)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.ui)

        self.table = self.ui.table
        self.imported_by = get_current_username()

        self._updating_table = False
        self._pending_changes = {}
        self._pending_deletes = set()
        self._new_rows = set()
        self._temp_row_id = -1
        self._import_preview_active = False
        self._import_preview_rows = []

        self.columns = ["product_id", "supplier_id", "price_date", "price", "currency"]
        self.headers = ["Product name", "Supplier name", "Price date", "Price", "Currency"]
        self.header_to_column = dict(zip(self.headers, self.columns))

        self.setup_ui()
        self.setup_connections()
        self.refresh_all_comboboxes()

    def setup_ui(self):
        setup_data_table(self.table, sorting=True)
        self._init_date_filters()

        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)

        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.clear_message()

    def setup_connections(self):
        self.table.cellDoubleClicked.connect(self.start_cell_edit)
        self.table.itemChanged.connect(self.on_item_changed)
        self.table.customContextMenuRequested.connect(self.show_context_menu)

        self.ui.btn_Search.clicked.connect(self.find_rows)
        self.ui.btn_AddLine.clicked.connect(self.add_line)
        self.ui.btn_Save.clicked.connect(self.apply_pending_changes)
        self.ui.btn_SaveExcel.clicked.connect(self.save_excel)
        self.ui.btn_DownFile.clicked.connect(self.download_template)
        self.ui.btn_Import.clicked.connect(self.import_excel)

        self.ui.line_SupplName.currentTextChanged.connect(self.fill_in_prod_brand_list)
        self.ui.line_SupplName.currentTextChanged.connect(self.fill_in_prod_fam_list)
        self.ui.line_SupplName.currentTextChanged.connect(self.fill_in_prod_name_list)

        self.ui.line_Brand.currentTextChanged.connect(self.fill_in_prod_fam_list)
        self.ui.line_Brand.currentTextChanged.connect(self.fill_in_prod_name_list)

        self.ui.line_Prod_Fam.currentTextChanged.connect(self.fill_in_prod_name_list)

    def get_session(self):
        return SessionLocal()

    def get_mode(self):
        return self.ui.line_TableName.currentText().strip()

    def _fill_combobox(self, combobox, items):
        current_value = combobox.currentText()
        combobox.blockSignals(True)
        combobox.clear()
        combobox.addItem("-")
        if items:
            combobox.addItems(sorted(items))
        if current_value in items:
            combobox.setCurrentText(current_value)
        elif current_value == "-":
            combobox.setCurrentText("-")
        combobox.blockSignals(False)

    def _to_decimal(self, value, field_name):
        if isinstance(value, Decimal):
            return value

        text = str(value).strip().replace(",", ".")
        if text == "":
            return None

        try:
            return Decimal(text)
        except (InvalidOperation, ValueError):
            raise Exception(f"Поле '{field_name}' должно быть числом")

    def _to_datetime(self, value, field_name):
        if isinstance(value, datetime):
            return value

        text = str(value).strip()
        if text == "":
            return None

        for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d.%m.%Y %H:%M:%S", "%Y-%m-%d %H:%M:%S"):
            try:
                return datetime.strptime(text, fmt)
            except ValueError:
                pass

        raise Exception(f"Поле '{field_name}' должно быть датой в формате ДД.ММ.ГГГГ")

    def refresh_all_comboboxes(self):
        self.fill_in_table_list()
        self.fill_in_supplier_list()
        self.fill_in_prod_brand_list()
        self.fill_in_prod_fam_list()
        self.fill_in_prod_name_list()

    def _get_products_for_filters_query(self, session):
        supplier_name = self.ui.line_SupplName.currentText()

        query = session.query(Product).filter(
            Product.name.isnot(None),
            Product.name != ""
        )

        if supplier_name != "-":
            mode = self.get_mode()

            if mode == MODE_CURRENT:
                query = query.join(
                    CurrentSupplierPrice,
                    CurrentSupplierPrice.product_id == Product.id
                ).join(
                    Supplier,
                    CurrentSupplierPrice.supplier_id == Supplier.id
                ).filter(
                    Supplier.name == supplier_name
                )
            else:
                query = query.join(
                    PriceHistory,
                    PriceHistory.product_id == Product.id
                ).join(
                    Supplier,
                    PriceHistory.supplier_id == Supplier.id
                ).filter(
                    Supplier.name == supplier_name
                )

        return query.distinct()

    def start_cell_edit(self, row, column):
        if self._updating_table:
            return

        if column not in (0, 1):
            return

        date_item = self.table.item(row, 2)
        if not date_item:
            return

        row_key = date_item.data(Qt.UserRole)
        if not row_key:
            return

        if row_key.startswith("new::"):
            return

        if column == 0:
            products = self.get_filtered_products()
            current_product_id = self._get_row_product_id(row_key)
            combo = self._build_product_combo(row_key, current_product_id, products)
            combo.currentIndexChanged.connect(
                lambda _, r=row, rk=row_key, c=combo: self.finish_product_edit(r, rk, c)
            )
            self.table.setCellWidget(row, column, combo)
            combo.showPopup()

        elif column == 1:
            suppliers = self.get_filtered_suppliers()
            current_supplier_id = self._get_row_supplier_id(row_key)
            combo = self._build_supplier_combo(row_key, current_supplier_id, suppliers)
            combo.currentIndexChanged.connect(
                lambda _, r=row, rk=row_key, c=combo: self.finish_supplier_edit(r, rk, c)
            )
            self.table.setCellWidget(row, column, combo)
            combo.showPopup()

    def _get_row_product_id(self, row_key):
        if row_key.startswith("current::"):
            _, product_id, _ = row_key.split("::")
            return int(product_id)

        if row_key.startswith("history::"):
            with self.get_session() as session:
                _, history_id = row_key.split("::")
                row = session.query(PriceHistory).filter(PriceHistory.id == int(history_id)).first()
                return row.product_id if row else None

        return None

    def _get_row_supplier_id(self, row_key):
        if row_key.startswith("current::"):
            _, _, supplier_id = row_key.split("::")
            return int(supplier_id)

        if row_key.startswith("history::"):
            with self.get_session() as session:
                _, history_id = row_key.split("::")
                row = session.query(PriceHistory).filter(PriceHistory.id == int(history_id)).first()
                return row.supplier_id if row else None

        return None

    def finish_product_edit(self, row, row_key, combo):
        product_id = combo.currentData()
        product_name = combo.currentText().strip()

        if row_key not in self._pending_changes:
            self._pending_changes[row_key] = {}

        self._pending_changes[row_key]["product_id"] = product_id

        self.table.removeCellWidget(row, 0)
        self.table.setItem(row, 0, self._build_item(product_name, row_key, align_left=True))

    def finish_supplier_edit(self, row, row_key, combo):
        supplier_id = combo.currentData()
        supplier_name = combo.currentText().strip()

        if row_key not in self._pending_changes:
            self._pending_changes[row_key] = {}

        self._pending_changes[row_key]["supplier_id"] = supplier_id

        self.table.removeCellWidget(row, 1)
        self.table.setItem(row, 1, self._build_item(supplier_name, row_key, align_left=True))

    def fill_in_table_list(self):
        self._fill_combobox(self.ui.line_TableName, [MODE_CURRENT, MODE_HISTORY])

    def fill_in_supplier_list(self):
        try:
            with self.get_session() as session:
                suppliers = (
                    session.query(Supplier.name)
                    .filter(Supplier.name.isnot(None), Supplier.name != "")
                    .distinct()
                    .order_by(Supplier.name)
                    .all()
                )

            supplier_names = [row[0] for row in suppliers if row[0]]
            self._fill_combobox(self.ui.line_SupplName, supplier_names)

        except Exception as e:
            self.show_error_message(f"Ошибка при получении поставщиков: {str(e)}")

    def fill_in_prod_brand_list(self):
        try:
            with self.get_session() as session:
                query = self._get_products_for_filters_query(session)
                brands = (
                    query.with_entities(Product.brand)
                    .filter(Product.brand.isnot(None), Product.brand != "")
                    .distinct()
                    .order_by(Product.brand)
                    .all()
                )

            brand_names = [row[0] for row in brands if row[0]]
            self._fill_combobox(self.ui.line_Brand, brand_names)

        except Exception as e:
            self.show_error_message(f"Ошибка при получении брендов: {str(e)}")

    def fill_in_prod_fam_list(self):
        brand = self.ui.line_Brand.currentText()

        try:
            with self.get_session() as session:
                query = self._get_products_for_filters_query(session)

                if brand != "-":
                    query = query.filter(Product.brand == brand)

                families = (
                    query.with_entities(Product.family)
                    .filter(Product.family.isnot(None), Product.family != "")
                    .distinct()
                    .order_by(Product.family)
                    .all()
                )

            family_names = [row[0] for row in families if row[0]]
            current_value = self.ui.line_Prod_Fam.currentText()
            self._fill_combobox(self.ui.line_Prod_Fam, family_names)

            if current_value in family_names:
                self.ui.line_Prod_Fam.setCurrentText(current_value)

        except Exception as e:
            self.show_error_message(f"Ошибка при получении семейств: {str(e)}")
            self._fill_combobox(self.ui.line_Prod_Fam, [])

    def fill_in_prod_name_list(self):
        brand = self.ui.line_Brand.currentText()
        family = self.ui.line_Prod_Fam.currentText()

        try:
            with self.get_session() as session:
                query = self._get_products_for_filters_query(session)

                if brand != "-":
                    query = query.filter(Product.brand == brand)

                if family != "-":
                    query = query.filter(Product.family == family)

                products = query.order_by(Product.name).all()

            product_names = [row.name for row in products if row.name]
            current_value = self.ui.line_Prod_name.currentText()
            self._fill_combobox(self.ui.line_Prod_name, product_names)

            if current_value in product_names:
                self.ui.line_Prod_name.setCurrentText(current_value)

        except Exception as e:
            self.show_error_message(f"Ошибка при получении продуктов: {str(e)}")
            self._fill_combobox(self.ui.line_Prod_name, [])

    def get_filtered_products(self):
        brand = self.ui.line_Brand.currentText()
        family = self.ui.line_Prod_Fam.currentText()
        product_name = self.ui.line_Prod_name.currentText()

        with self.get_session() as session:
            query = session.query(Product).filter(Product.name.isnot(None), Product.name != "")

            if brand != "-":
                query = query.filter(Product.brand == brand)
            if family != "-":
                query = query.filter(Product.family == family)
            if product_name != "-":
                query = query.filter(Product.name == product_name)

            return query.order_by(Product.name).all()

    def get_filtered_suppliers(self):
        supplier_name = self.ui.line_SupplName.currentText()

        with self.get_session() as session:
            query = session.query(Supplier).filter(Supplier.name.isnot(None), Supplier.name != "")

            if supplier_name != "-":
                query = query.filter(Supplier.name == supplier_name)

            return query.order_by(Supplier.name).all()

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

    def copy_cell_content(self):
        selected_items = self.table.selectedItems()
        if not selected_items:
            return

        clipboard = QApplication.clipboard()

        if len(selected_items) == 1:
            text = selected_items[0].text()
        else:
            rows = {}
            for item in selected_items:
                rows.setdefault(item.row(), {})
                rows[item.row()][item.column()] = item.text()

            text = "\n".join(
                "\t".join(value for _, value in sorted(cols.items()))
                for _, cols in sorted(rows.items())
            )

        clipboard.setText(text.strip())
        self.show_message("Скопировано")

    def _init_date_filters(self):
        from PySide6.QtCore import QDate
        today = QDate.currentDate()
        self.ui.line_Start_date.setDate(today)
        self.ui.line_End_date.setDate(today)

    def _get_date_filters(self):
        from PySide6.QtCore import QDate
        start_qdate = self.ui.line_Start_date.date()
        end_qdate = self.ui.line_End_date.date()
        today = QDate.currentDate()

        if start_qdate == today and end_qdate == today:
            return None, None

        start_date = datetime(
            start_qdate.year(),
            start_qdate.month(),
            start_qdate.day(),
            0, 0, 0, 0,
        )

        end_date = datetime(
            end_qdate.year(),
            end_qdate.month(),
            end_qdate.day(),
            23, 59, 59, 999999,
        )

        return start_date, end_date

    def get_rows_from_db(self):
        mode = self.get_mode()
        start_date, end_date = self._get_date_filters()

        with self.get_session() as session:
            supplier_name = self.ui.line_SupplName.currentText()
            brand = self.ui.line_Brand.currentText()
            family = self.ui.line_Prod_Fam.currentText()
            product_name = self.ui.line_Prod_name.currentText()

            if mode == MODE_CURRENT:
                query = (
                    session.query(
                        CurrentSupplierPrice.product_id.label("product_id"),
                        CurrentSupplierPrice.supplier_id.label("supplier_id"),
                        CurrentSupplierPrice.last_update.label("price_date"),
                        CurrentSupplierPrice.price.label("price"),
                        CurrentSupplierPrice.currency.label("currency"),
                        Product.name.label("product_name"),
                        Supplier.name.label("supplier_name"),
                    )
                    .join(Product, CurrentSupplierPrice.product_id == Product.id)
                    .join(Supplier, CurrentSupplierPrice.supplier_id == Supplier.id)
                )

                if supplier_name != "-":
                    query = query.filter(Supplier.name == supplier_name)
                if brand != "-":
                    query = query.filter(Product.brand == brand)
                if family != "-":
                    query = query.filter(Product.family == family)
                if product_name != "-":
                    query = query.filter(Product.name == product_name)

                if start_date:
                    query = query.filter(CurrentSupplierPrice.last_update >= start_date)
                if end_date:
                    query = query.filter(CurrentSupplierPrice.last_update <= end_date)

                rows = query.order_by(
                    Product.name,
                    CurrentSupplierPrice.last_update.desc(),
                    Supplier.name,
                ).all()

                return [
                    {
                        "row_key": f"current::{row.product_id}::{row.supplier_id}",
                        "product_id": row.product_id,
                        "supplier_id": row.supplier_id,
                        "product_name": row.product_name or "",
                        "supplier_name": row.supplier_name or "",
                        "price_date": row.price_date,
                        "price": row.price,
                        "currency": row.currency,
                        "is_new": False,
                    }
                    for row in rows
                ]

            elif mode == MODE_HISTORY:
                query = (
                    session.query(
                        PriceHistory.id.label("history_id"),
                        PriceHistory.product_id.label("product_id"),
                        PriceHistory.supplier_id.label("supplier_id"),
                        PriceHistory.price_date.label("price_date"),
                        PriceHistory.price.label("price"),
                        PriceHistory.currency.label("currency"),
                        Product.name.label("product_name"),
                        Supplier.name.label("supplier_name"),
                    )
                    .join(Product, PriceHistory.product_id == Product.id)
                    .join(Supplier, PriceHistory.supplier_id == Supplier.id)
                )

                if supplier_name != "-":
                    query = query.filter(Supplier.name == supplier_name)
                if brand != "-":
                    query = query.filter(Product.brand == brand)
                if family != "-":
                    query = query.filter(Product.family == family)
                if product_name != "-":
                    query = query.filter(Product.name == product_name)

                if start_date:
                    query = query.filter(PriceHistory.price_date >= start_date)
                if end_date:
                    query = query.filter(PriceHistory.price_date <= end_date)

                rows = query.order_by(
                    Product.name,
                    PriceHistory.price_date.desc(),
                    Supplier.name,
                    PriceHistory.id.desc(),
                ).all()

                return [
                    {
                        "row_key": f"history::{row.history_id}",
                        "product_id": row.product_id,
                        "supplier_id": row.supplier_id,
                        "product_name": row.product_name or "",
                        "supplier_name": row.supplier_name or "",
                        "price_date": row.price_date,
                        "price": row.price,
                        "currency": row.currency,
                        "is_new": False,
                    }
                    for row in rows
                ]

        return []

    def find_rows(self):
        self._import_preview_active = False
        self._import_preview_rows = []
        mode = self.get_mode()

        if mode == MODE_NONE or not mode:
            self.show_message("Выбери таблицу")
            return

        data = self.get_rows_from_db()
        self._display_data(data)

        if not data:
            self.show_message("Нет данных по заданным фильтрам")
        
    def _build_item(self, value, row_key, align_left=False):
        if isinstance(value, datetime):
            text = value.strftime("%d.%m.%Y")
            sort_value = value
        elif isinstance(value, Decimal):
            text = f"{value:.2f}".replace(".", ",")
            sort_value = value
        elif isinstance(value, (int, float)):
            text = f"{float(value):.2f}".replace(".", ",")
            sort_value = float(value)
        else:
            text = "" if value is None else str(value)
            sort_value = text.casefold() if isinstance(text, str) else text

        item = SortableTableWidgetItem(text)
        item.setData(Qt.UserRole, row_key)
        item.setData(Qt.UserRole + 1, sort_value)
        item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable)
        item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter if align_left else Qt.AlignCenter)
        return item

    def _build_product_combo(self, row_key, selected_product_id, products):
        combo = QComboBox()
        combo.addItem("", None)

        for product in products:
            combo.addItem(product.name, product.id)

        if selected_product_id is not None and combo.findData(selected_product_id) < 0:
            with self.get_session() as session:
                product = session.query(Product).filter(Product.id == selected_product_id).first()
                if product:
                    combo.addItem(product.name, product.id)

        idx = combo.findData(selected_product_id)
        if idx >= 0:
            combo.setCurrentIndex(idx)

        combo.currentIndexChanged.connect(
            lambda _, rk=row_key, c=combo: self.on_combo_changed(rk, "product_id", c.currentData())
        )
        return combo

    def _build_supplier_combo(self, row_key, selected_supplier_id, suppliers):
        combo = QComboBox()
        combo.addItem("", None)

        for supplier in suppliers:
            combo.addItem(supplier.name, supplier.id)

        if selected_supplier_id is not None and combo.findData(selected_supplier_id) < 0:
            with self.get_session() as session:
                supplier = session.query(Supplier).filter(Supplier.id == selected_supplier_id).first()
                if supplier:
                    combo.addItem(supplier.name, supplier.id)

        idx = combo.findData(selected_supplier_id)
        if idx >= 0:
            combo.setCurrentIndex(idx)

        combo.currentIndexChanged.connect(
            lambda _, rk=row_key, c=combo: self.on_combo_changed(rk, "supplier_id", c.currentData())
        )
        return combo

    def _display_data(self, data):
        self.table.clear()
        self.table.setColumnCount(0)
        self.table.setRowCount(0)

        self._updating_table = True

        self.table.setColumnCount(len(self.headers))
        self.table.setRowCount(len(data))
        self.table.setHorizontalHeaderLabels(self.headers)

        products = self.get_filtered_products()
        suppliers = self.get_filtered_suppliers()

        for row_index, row_data in enumerate(data):
            row_key = row_data["row_key"]

            if row_key.startswith("new::"):
                self.table.setCellWidget(
                    row_index,
                    0,
                    self._build_product_combo(row_key, row_data["product_id"], products)
                )
                self.table.setCellWidget(
                    row_index,
                    1,
                    self._build_supplier_combo(row_key, row_data["supplier_id"], suppliers)
                )
            else:
                self.table.setItem(
                    row_index, 0,
                    self._build_item(row_data["product_name"], row_key, align_left=True)
                )
                self.table.setItem(
                    row_index, 1,
                    self._build_item(row_data["supplier_name"], row_key, align_left=True)
                )

            self.table.setItem(
                row_index, 2,
                self._build_item(row_data["price_date"], row_key, align_left=False)
            )
            self.table.setItem(
                row_index, 3,
                self._build_item(row_data["price"], row_key, align_left=False)
            )
            self.table.setItem(
                row_index, 4,
                self._build_item(row_data["currency"], row_key, align_left=True)
            )

        self.table.resizeColumnsToContents()

        self._updating_table = False

    def on_combo_changed(self, row_key, field_name, value):
        if self._updating_table:
            return

        if row_key not in self._pending_changes:
            self._pending_changes[row_key] = {}

        self._pending_changes[row_key][field_name] = value

    def on_item_changed(self, item):
        if self._updating_table:
            return

        row_key = item.data(Qt.UserRole)
        if row_key is None:
            return

        header = self.headers[item.column()]
        field_name = self.header_to_column.get(header)
        if not field_name:
            return

        value = item.text().strip()

        if field_name == "price":
            value = self._to_decimal(value, "Price")
        elif field_name == "price_date":
            value = self._to_datetime(value, "Price date")
        elif value == "":
            value = None

        if row_key not in self._pending_changes:
            self._pending_changes[row_key] = {}

        self._pending_changes[row_key][field_name] = value

    def add_line(self):
        mode = self.get_mode()
        if mode == MODE_NONE or not mode:
            self.show_message("Выбери таблицу")
            return

        row_key = f"new::{self._temp_row_id}"
        self._temp_row_id -= 1
        self._new_rows.add(row_key)

        if self._import_preview_active:
            data = list(self._import_preview_rows)
        else:
            data = self.get_rows_from_db()

        new_row = {
            "row_key": row_key,
            "product_id": None,
            "supplier_id": None,
            "product_name": "",
            "supplier_name": "",
            "price_date": None,
            "price": None,
            "currency": "",
            "is_new": True,
        }
        data.insert(0, new_row)

        if self._import_preview_active:
            self._import_preview_rows = data

        self._pending_changes[row_key] = {
            "product_id": None,
            "supplier_id": None,
            "price_date": None,
            "price": None,
            "currency": "",
        }

        self._display_data(data)
        self.table.setCurrentCell(0, 2)
        self.show_message("Добавлена новая строка")

    def delete_selected_row(self):
        selected_rows = sorted({index.row() for index in self.table.selectedIndexes()}, reverse=True)

        if not selected_rows:
            self.show_error_message("Не выбраны строки для удаления")
            return

        deleted_count = 0

        for row in selected_rows:
            row_key = None
            date_item = self.table.item(row, 2)
            if date_item:
                row_key = date_item.data(Qt.UserRole)

            if not row_key:
                continue

            if row_key in self._new_rows:
                self._new_rows.discard(row_key)
                self._pending_changes.pop(row_key, None)
                if self._import_preview_active:
                    self._import_preview_rows = [
                        r for r in self._import_preview_rows if r["row_key"] != row_key
                    ]
            else:
                self._pending_deletes.add(row_key)

            self.table.removeRow(row)
            deleted_count += 1

        if deleted_count:
            self.show_message(f"Строк помечено на удаление: {deleted_count}")
        else:
            self.show_error_message("Не удалось определить строки для удаления")

    def revert_changes(self):
        try:
            self._pending_changes.clear()
            self._pending_deletes.clear()
            self._new_rows.clear()

            if self._import_preview_active:
                self._display_data(self._import_preview_rows)
            elif self.get_mode() != MODE_NONE:
                self.find_rows()
            else:
                self.table.clearContents()
                self.table.setRowCount(0)

            self.show_message("Изменения отменены")

        except Exception as e:
            self.show_error_message(f"Ошибка отката: {str(e)}")

    def _merge_row_values(self, row_key, fallback):
        changes = self._pending_changes.get(row_key, {})
        result = fallback.copy()
        result.update(changes)
        return result

    def _save_price_to_history_and_current(self, session, values):
        currency = str(values["currency"]).strip().upper()
        price_date = values["price_date"]
        price_value = values["price"]
        product_id = values["product_id"]
        supplier_id = values["supplier_id"]

        row = PriceHistory(
            product_id=product_id,
            supplier_id=supplier_id,
            price_date=price_date,
            price=price_value,
            currency=currency,
        )
        session.add(row)

        current_row = session.query(CurrentSupplierPrice).filter(
            CurrentSupplierPrice.product_id == product_id,
            CurrentSupplierPrice.supplier_id == supplier_id,
        ).first()

        if current_row is None:
            current_row = CurrentSupplierPrice(
                product_id=product_id,
                supplier_id=supplier_id,
                last_update=price_date,
                price=price_value,
                currency=currency,
            )
            session.add(current_row)
        else:
            if current_row.last_update is None or current_row.last_update <= price_date:
                current_row.last_update = price_date
                current_row.price = price_value
                current_row.currency = currency

    def apply_pending_changes(self):
        mode = self.get_mode()
        if mode == MODE_NONE or not mode:
            self.show_message("Выбери таблицу")
            return

        if not self._pending_changes and not self._pending_deletes:
            self.show_message("Нет изменений для применения")
            return

        try:
            with self.get_session() as session:
                for row_key, changes in self._pending_changes.items():
                    if row_key.startswith("new::"):
                        values = self._merge_row_values(row_key, {
                            "product_id": None,
                            "supplier_id": None,
                            "price_date": None,
                            "price": None,
                            "currency": "",
                        })

                        if not values["product_id"]:
                            raise Exception("Поле 'Product name' обязательно")
                        if not values["supplier_id"]:
                            raise Exception("Поле 'Supplier name' обязательно")
                        if not values["price_date"]:
                            raise Exception("Поле 'Price date' обязательно")
                        if values["price"] is None:
                            raise Exception("Поле 'Price' обязательно")
                        if not str(values["currency"]).strip():
                            raise Exception("Поле 'Currency' обязательно")

                        if mode == MODE_CURRENT:
                            duplicate = (
                                session.query(CurrentSupplierPrice)
                                .filter(
                                    CurrentSupplierPrice.product_id == values["product_id"],
                                    CurrentSupplierPrice.supplier_id == values["supplier_id"],
                                )
                                .first()
                            )
                            if duplicate:
                                raise Exception("Такая строка уже существует в таблице последних цен")

                            row = CurrentSupplierPrice(
                                product_id=values["product_id"],
                                supplier_id=values["supplier_id"],
                                last_update=values["price_date"],
                                price=values["price"],
                                currency=str(values["currency"]).strip().upper(),
                            )
                            session.add(row)
                        else:
                            self._save_price_to_history_and_current(session, values)

                    else:
                        if mode == MODE_CURRENT:
                            _, product_id, supplier_id = row_key.split("::")
                            row = (
                                session.query(CurrentSupplierPrice)
                                .filter(
                                    CurrentSupplierPrice.product_id == int(product_id),
                                    CurrentSupplierPrice.supplier_id == int(supplier_id),
                                )
                                .first()
                            )
                            if not row:
                                continue

                            values = self._merge_row_values(row_key, {
                                "product_id": row.product_id,
                                "supplier_id": row.supplier_id,
                                "price_date": row.last_update,
                                "price": row.price,
                                "currency": row.currency,
                            })

                            if not values["product_id"]:
                                raise Exception("Поле 'Product name' обязательно")
                            if not values["supplier_id"]:
                                raise Exception("Поле 'Supplier name' обязательно")
                            if not values["price_date"]:
                                raise Exception("Поле 'Price date' обязательно")
                            if values["price"] is None:
                                raise Exception("Поле 'Price' обязательно")
                            if not str(values["currency"]).strip():
                                raise Exception("Поле 'Currency' обязательно")

                            row.product_id = values["product_id"]
                            row.supplier_id = values["supplier_id"]
                            row.last_update = values["price_date"]
                            row.price = values["price"]
                            row.currency = str(values["currency"]).strip().upper()

                        else:
                            _, history_id = row_key.split("::")
                            row = session.query(PriceHistory).filter(PriceHistory.id == int(history_id)).first()
                            if not row:
                                continue

                            values = self._merge_row_values(row_key, {
                                "product_id": row.product_id,
                                "supplier_id": row.supplier_id,
                                "price_date": row.price_date,
                                "price": row.price,
                                "currency": row.currency,
                            })

                            if not values["product_id"]:
                                raise Exception("Поле 'Product name' обязательно")
                            if not values["supplier_id"]:
                                raise Exception("Поле 'Supplier name' обязательно")
                            if not values["price_date"]:
                                raise Exception("Поле 'Price date' обязательно")
                            if values["price"] is None:
                                raise Exception("Поле 'Price' обязательно")
                            if not str(values["currency"]).strip():
                                raise Exception("Поле 'Currency' обязательно")

                            row.product_id = values["product_id"]
                            row.supplier_id = values["supplier_id"]
                            row.price_date = values["price_date"]
                            row.price = values["price"]
                            row.currency = str(values["currency"]).strip().upper()

                            current_row = session.query(CurrentSupplierPrice).filter(
                                CurrentSupplierPrice.product_id == values["product_id"],
                                CurrentSupplierPrice.supplier_id == values["supplier_id"],
                            ).first()
                            if current_row is None:
                                current_row = CurrentSupplierPrice(
                                    product_id=values["product_id"],
                                    supplier_id=values["supplier_id"],
                                    last_update=values["price_date"],
                                    price=values["price"],
                                    currency=str(values["currency"]).strip().upper(),
                                )
                                session.add(current_row)
                            elif current_row.last_update is None or current_row.last_update <= values["price_date"]:
                                current_row.last_update = values["price_date"]
                                current_row.price = values["price"]
                                current_row.currency = str(values["currency"]).strip().upper()

                for row_key in self._pending_deletes:
                    if row_key.startswith("current::"):
                        _, product_id, supplier_id = row_key.split("::")
                        session.query(CurrentSupplierPrice).filter(
                            CurrentSupplierPrice.product_id == int(product_id),
                            CurrentSupplierPrice.supplier_id == int(supplier_id),
                        ).delete(synchronize_session=False)

                    elif row_key.startswith("history::"):
                        _, history_id = row_key.split("::")
                        session.query(PriceHistory).filter(
                            PriceHistory.id == int(history_id)
                        ).delete(synchronize_session=False)

                session.commit()

            imported_saved = self._import_preview_active
            self._pending_changes.clear()
            self._pending_deletes.clear()
            self._new_rows.clear()
            self._import_preview_active = False
            self._import_preview_rows = []

            self.find_rows()
            if imported_saved:
                self.show_message("История цен успешно импортирована и сохранена")
            else:
                self.show_message("Данные успешно сохранены")

        except SQLAlchemyError as e:
            self.show_error_message(f"Ошибка сохранения в базу данных: {str(e)}")
        except Exception as e:
            self.show_error_message(f"Ошибка применения изменений: {str(e)}")

    def download_template(self):
        if self.get_mode() != MODE_HISTORY:
            self.show_message('Импорт возможен только в таблицу "История цен (вся)"')
            return

        try:
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "Сохранить шаблон",
                str(Path.home() / "PriceHistoryImportTemplate.xlsx"),
                "Excel files (*.xlsx)",
            )
            if not file_path:
                return

            if not file_path.lower().endswith(".xlsx"):
                file_path += ".xlsx"

            exporter = PriceHistoryExporter()
            exporter.export_template(file_path)
            self.show_message("Шаблон сохранен")
        except Exception as e:
            self.show_error_message(str(e))

    def _find_supplier_id_by_name(self, session, supplier_name: str):
        if not supplier_name:
            return None
        supplier = session.query(Supplier).filter(Supplier.name == supplier_name).first()
        return supplier.id if supplier else None

    def _find_product_for_import(self, session, row: dict):
        matching = ProductMatchingService(session)

        our_name = clean_multi_spaces(row.get("our_product_name"))
        article = clean_multi_spaces(row.get("supplier_article"))
        supplier_product_name = clean_multi_spaces(row.get("supplier_product_name"))

        if our_name:
            product = session.query(Product).filter(Product.name == our_name).first()
            if product:
                return product

            target = normalize_product_name(our_name)
            if target:
                products = session.query(Product).filter(Product.name.isnot(None)).all()
                for product in products:
                    if normalize_product_name(product.name) == target:
                        return product

        product = matching.find_price_import_product(
            supplier_article=article,
            supplier_product_name=supplier_product_name,
        )
        if product:
            return product

        return None

    def import_excel(self):
        if self.get_mode() != MODE_HISTORY:
            self.show_message('Импорт возможен только в таблицу "История цен (вся)"')
            return

        try:
            file_path, _ = QFileDialog.getOpenFileName(
                self,
                "Выберите файл истории цен",
                "",
                "Excel files (*.xls *.xlsx)",
            )
            if not file_path:
                return

            importer = PriceHistoryImporter()
            rows = importer.read_excel(file_path)
            if not rows:
                self.show_message("Нет строк для импорта")
                return

            preview_data = []
            self._pending_changes.clear()
            self._pending_deletes.clear()
            self._new_rows.clear()

            with self.get_session() as session:
                for src_row in rows:
                    row_key = f"new::{self._temp_row_id}"
                    self._temp_row_id -= 1
                    self._new_rows.add(row_key)

                    supplier_name = clean_multi_spaces(src_row.get("supplier_name"))
                    supplier_id = self._find_supplier_id_by_name(session, supplier_name)

                    product = self._find_product_for_import(session, src_row)
                    product_id = product.id if product else None
                    product_name = product.name if product else clean_multi_spaces(src_row.get("our_product_name"))

                    price_date = src_row.get("price_date")
                    if price_date is not None and hasattr(price_date, "to_pydatetime"):
                        price_date = price_date.to_pydatetime()

                    price_value = src_row.get("price")
                    currency = clean_multi_spaces(src_row.get("currency")).upper()

                    preview_row = {
                        "row_key": row_key,
                        "product_id": product_id,
                        "supplier_id": supplier_id,
                        "product_name": product_name or "",
                        "supplier_name": supplier_name or "",
                        "price_date": price_date,
                        "price": price_value,
                        "currency": currency,
                        "is_new": True,
                    }
                    preview_data.append(preview_row)
                    self._pending_changes[row_key] = {
                        "product_id": product_id,
                        "supplier_id": supplier_id,
                        "price_date": price_date,
                        "price": price_value,
                        "currency": currency,
                    }

            self._import_preview_active = True
            self._import_preview_rows = preview_data
            self._display_data(preview_data)
            self.show_message("Файл импортирован. Проверь предпросмотр и нажми Save")
        except Exception as e:
            self.show_error_message(str(e))

    def save_excel(self):
        mode = self.get_mode()
        if mode == MODE_NONE or not mode:
            self.show_message("Выбери таблицу")
            return

        try:
            supplier_name = clean_multi_spaces(self.ui.line_SupplName.currentText())
            timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

            if supplier_name and supplier_name != "-":
                safe_supplier = supplier_name
                for ch in ['\\', '/', ':', '*', '?', '"', '<', '>', '|']:
                    safe_supplier = safe_supplier.replace(ch, "_")
                default_name = f"Price history_{safe_supplier}_{timestamp}.xlsx"
            else:
                default_name = f"Price history_{timestamp}.xlsx"

            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "Сохранить Excel",
                default_name,
                "Excel files (*.xlsx)",
            )
            if not file_path:
                return

            data = self.get_rows_from_db()
            exporter = PriceHistoryExporter()

            report_type = "current" if mode == MODE_CURRENT else "history"
            exporter.export_rows(data, file_path, report_type=report_type)

            self.show_message("Excel файл сохранен")

        except Exception as e:
            self.show_error_message(str(e))

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
