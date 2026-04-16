from pathlib import Path
from datetime import datetime
from decimal import Decimal, InvalidOperation

from sqlalchemy.exc import SQLAlchemyError
from PySide6.QtWidgets import (
    QMessageBox,
    QHeaderView,
    QTableWidget,
    QMenu,
    QTableWidgetItem,
    QWidget,
    QApplication,
    QVBoxLayout,
    QComboBox,
)
from PySide6.QtCore import Qt, QFile
from PySide6.QtUiTools import QUiLoader

from app.db.models import Product, Supplier, PriceHistory, CurrentSupplierPrice
from app.db.db import SessionLocal
from app.ui.table_style import setup_data_table
from app.exports.price_history_export import PriceHistoryExport


BASE_DIR = Path(__file__).resolve().parents[2]
PRICE_HISTORY_UI = BASE_DIR / "app" / "ui" / "windows" / "price_history.ui"


MODE_NONE = "-"
MODE_CURRENT = "Последние цены"
MODE_HISTORY = "История цен (вся)"


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

        self._updating_table = False
        self._pending_changes = {}
        self._pending_deletes = set()
        self._new_rows = set()
        self._temp_row_id = -1

        self.columns = ["product_id", "supplier_id", "price_date", "price", "currency"]
        self.headers = ["Product name", "Supplier name", "Price date", "Price", "Currency"]
        self.header_to_column = dict(zip(self.headers, self.columns))

        self.setup_ui()
        self.setup_connections()
        self.refresh_all_comboboxes()

    def setup_ui(self):
        setup_data_table(self.table, sorting=True)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)

    def setup_connections(self):
        self.table.itemChanged.connect(self.on_item_changed)
        self.table.customContextMenuRequested.connect(self.show_context_menu)

        self.ui.btn_Search.clicked.connect(self.find_rows)
        self.ui.btn_AddLine.clicked.connect(self.add_line)
        self.ui.btn_Save.clicked.connect(self.apply_pending_changes)
        self.ui.btn_SaveExcel.clicked.connect(self.save_excel)

        self.ui.line_Brand.currentTextChanged.connect(self.fill_in_prod_fam_list)
        self.ui.line_Prod_Fam.currentTextChanged.connect(self.fill_in_prod_name_list)

    def get_session(self):
        return SessionLocal()

    def get_mode(self):
        return self.ui.line_TableName.currentText().strip()

    def is_current_mode(self):
        return self.get_mode() == MODE_CURRENT

    def is_history_mode(self):
        return self.get_mode() == MODE_HISTORY

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

        raise Exception(f"Поле '{field_name}' должно быть датой в формате dd.MM.yyyy")

    def refresh_all_comboboxes(self):
        self.fill_in_table_list()
        self.fill_in_supplier_list()
        self.fill_in_prod_brand_list()
        self.fill_in_prod_fam_list()
        self.fill_in_prod_name_list()

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
                brands = (
                    session.query(Product.brand)
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
                query = (
                    session.query(Product.family)
                    .filter(Product.family.isnot(None), Product.family != "")
                )

                if brand != "-":
                    query = query.filter(Product.brand == brand)

                families = query.distinct().order_by(Product.family).all()

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
                query = session.query(Product).filter(Product.name.isnot(None), Product.name != "")

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

    def get_rows_from_db(self):
        mode = self.get_mode()

        with self.get_session() as session:
            if mode == MODE_CURRENT:
                query = (
                    session.query(CurrentSupplierPrice)
                    .join(Product, CurrentSupplierPrice.product_id == Product.id)
                    .join(Supplier, CurrentSupplierPrice.supplier_id == Supplier.id)
                )

                supplier_name = self.ui.line_SupplName.currentText()
                brand = self.ui.line_Brand.currentText()
                family = self.ui.line_Prod_Fam.currentText()
                product_name = self.ui.line_Prod_name.currentText()

                if supplier_name != "-":
                    query = query.filter(Supplier.name == supplier_name)
                if brand != "-":
                    query = query.filter(Product.brand == brand)
                if family != "-":
                    query = query.filter(Product.family == family)
                if product_name != "-":
                    query = query.filter(Product.name == product_name)

                rows = (
                    query.order_by(Product.name, Supplier.name, CurrentSupplierPrice.last_update.desc())
                    .all()
                )

                data = []
                for row in rows:
                    product = session.query(Product).filter(Product.id == row.product_id).first()
                    supplier = session.query(Supplier).filter(Supplier.id == row.supplier_id).first()

                    data.append({
                        "row_key": f"current::{row.product_id}::{row.supplier_id}",
                        "product_id": row.product_id,
                        "supplier_id": row.supplier_id,
                        "product_name": product.name if product else "",
                        "supplier_name": supplier.name if supplier else "",
                        "price_date": row.last_update,
                        "price": row.price,
                        "currency": row.currency,
                        "is_new": False,
                    })

                return data

            elif mode == MODE_HISTORY:
                query = (
                    session.query(PriceHistory)
                    .join(Product, PriceHistory.product_id == Product.id)
                    .join(Supplier, PriceHistory.supplier_id == Supplier.id)
                )

                supplier_name = self.ui.line_SupplName.currentText()
                brand = self.ui.line_Brand.currentText()
                family = self.ui.line_Prod_Fam.currentText()
                product_name = self.ui.line_Prod_name.currentText()

                if supplier_name != "-":
                    query = query.filter(Supplier.name == supplier_name)
                if brand != "-":
                    query = query.filter(Product.brand == brand)
                if family != "-":
                    query = query.filter(Product.family == family)
                if product_name != "-":
                    query = query.filter(Product.name == product_name)

                rows = (
                    query.order_by(Product.name, Supplier.name, PriceHistory.price_date.desc(), PriceHistory.id.desc())
                    .all()
                )

                data = []
                for row in rows:
                    product = session.query(Product).filter(Product.id == row.product_id).first()
                    supplier = session.query(Supplier).filter(Supplier.id == row.supplier_id).first()

                    data.append({
                        "row_key": f"history::{row.id}",
                        "product_id": row.product_id,
                        "supplier_id": row.supplier_id,
                        "product_name": product.name if product else "",
                        "supplier_name": supplier.name if supplier else "",
                        "price_date": row.price_date,
                        "price": row.price,
                        "currency": row.currency,
                        "is_new": False,
                    })

                return data

        return []

    def find_rows(self):
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
            value = value.strftime("%d.%m.%Y")

        item = QTableWidgetItem("" if value is None else str(value))
        item.setData(Qt.UserRole, row_key)
        item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable)
        item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter if align_left else Qt.AlignCenter)
        return item

    def _build_product_combo(self, row_key, selected_product_id, products):
        combo = QComboBox()
        combo.addItem("", None)

        for product in products:
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

        for i in range(self.table.columnCount()):
            if self.table.columnWidth(i) < 120:
                self.table.setColumnWidth(i, 120)

        self._updating_table = False
        self.table.setSortingEnabled(True)

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

        data = self.get_rows_from_db()
        data.insert(0, {
            "row_key": row_key,
            "product_id": None,
            "supplier_id": None,
            "product_name": "",
            "supplier_name": "",
            "price_date": None,
            "price": None,
            "currency": "",
            "is_new": True,
        })

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
        row = self.table.currentRow()
        if row < 0:
            self.show_error_message("Не выбрана строка для удаления")
            return

        product_widget = self.table.cellWidget(row, 0)
        if not product_widget:
            self.show_error_message("Не удалось определить строку")
            return

        row_key = None
        date_item = self.table.item(row, 2)
        if date_item:
            row_key = date_item.data(Qt.UserRole)

        if not row_key:
            self.show_error_message("Не удалось определить строку")
            return

        if row_key in self._new_rows:
            self._new_rows.discard(row_key)
            self._pending_changes.pop(row_key, None)
        else:
            self._pending_deletes.add(row_key)

        self.table.removeRow(row)
        self.show_message("Строка помечена на удаление")

    def revert_changes(self):
        try:
            self._pending_changes.clear()
            self._pending_deletes.clear()
            self._new_rows.clear()

            if self.get_mode() != MODE_NONE:
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
                            row = PriceHistory(
                                product_id=values["product_id"],
                                supplier_id=values["supplier_id"],
                                price_date=values["price_date"],
                                price=values["price"],
                                currency=str(values["currency"]).strip().upper(),
                            )
                            session.add(row)

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

            self._pending_changes.clear()
            self._pending_deletes.clear()
            self._new_rows.clear()

            self.find_rows()
            self.show_message("Данные успешно сохранены")

        except SQLAlchemyError as e:
            self.show_error_message(f"Ошибка сохранения в базу данных: {str(e)}")
        except Exception as e:
            self.show_error_message(f"Ошибка применения изменений: {str(e)}")

    def save_excel(self):
        mode = self.get_mode()
        if mode == MODE_NONE or not mode:
            self.show_message("Выбери таблицу")
            return

        try:
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "Сохранить Excel",
                "price_history.xlsx",
                "Excel files (*.xlsx)",
            )
            if not file_path:
                return

            data = self.get_rows_from_db()
            exporter = PriceHistoryExport()
            exporter.export_rows(data, file_path)

            self.show_message("Excel файл сохранен")

        except Exception as e:
            self.show_error_message(str(e))

    def show_message(self, text):
        self.ui.label_msg.setText(text)
        self.ui.label_msg.setVisible(True)

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