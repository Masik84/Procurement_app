from pathlib import Path

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
    QCheckBox,
    QHBoxLayout,
)
from PySide6.QtCore import Qt, QFile
from PySide6.QtUiTools import QUiLoader

from app.db.models import Product
from app.db.db import SessionLocal


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
        self._original_values = {}
        self._pending_changes = {}
        self._pending_deletes = set()
        self._new_rows = set()
        self._temp_row_id = -1

        self.columns = ["id", "name", "brand", "pack", "is_excise", "family"]
        self.headers = ["id", "Product name", "Brand", "Pack", "Excise duty", "Product Family"]
        self.text_columns = {"name", "brand", "family"}

        self.setup_ui()
        self.setup_connections()
        self.refresh_all_comboboxes()

    def setup_ui(self):
        self.table = self.ui.table

        self.table.setSelectionBehavior(QTableWidget.SelectItems)
        self.table.setEditTriggers(QTableWidget.DoubleClicked | QTableWidget.EditKeyPressed)
        self.table.setAlternatingRowColors(True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.verticalHeader().setVisible(False)
        self.table.setSortingEnabled(True)
        self.table.setWordWrap(False)
        self.table.setTextElideMode(Qt.TextElideMode.ElideRight)

        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_context_menu)

        self.table.setStyleSheet("""
            QTableWidget {
                alternate-background-color: #f0f0f0;
                selection-background-color: #3daee9;
                selection-color: black;
            }
            QTableWidget::item {
                padding: 3px;
            }
            QTableWidget::item:editable {
                background-color: #fff5cc;
                border: 1px solid #ffcc66;
            }
            QTableWidget::item:focus {
                background-color: #f28223;
                border: 1px solid #ff9900;
                padding: 1px;
            }
            QTableWidget QLineEdit {
                background-color: #fff2cc;
                color: black;
                border: 1px solid #ff9900;
                padding: 1px;
            }
        """)

        self.table.setTabKeyNavigation(True)
        self.table.setCornerButtonEnabled(False)

    def setup_connections(self):
        self.table.itemChanged.connect(self.on_item_changed)
        self.ui.line_Brand.currentTextChanged.connect(self.fill_in_prod_fam_list)
        self.ui.line_Prod_Fam.currentTextChanged.connect(self.fill_in_prod_name_list)

        self.ui.btn_Search.clicked.connect(self.find_Product)
        self.ui.btn_AddLine.clicked.connect(self.add_line)
        self.ui.btn_Save.clicked.connect(self.apply_pending_changes)

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
            header = self.table.horizontalHeaderItem(column).text()
            id_item = self.table.item(row, 0)

            if not id_item:
                return

            row_id_text = id_item.text().strip()
            if not row_id_text:
                return

            row_id = int(row_id_text)

            if header == "id":
                return

            new_value = item.text()

            if row_id not in self._pending_changes:
                self._pending_changes[row_id] = {}

            self._pending_changes[row_id][header] = new_value

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
                        self._insert_product(session, changes)
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

    def _insert_product(self, session, changes):
        name = str(changes.get("name", "")).strip()
        brand = str(changes.get("brand", "")).strip()
        pack = str(changes.get("pack", "")).strip()
        family = str(changes.get("family", "")).strip()
        is_excise = bool(changes.get("is_excise", False))

        if not name:
            raise Exception("Для новой строки поле name обязательно")

        existing = session.query(Product).filter(Product.name == name).first()
        if existing:
            raise Exception(f"Продукт с name '{name}' уже существует")

        product = Product(
            name=name,
            brand=brand if brand else None,
            pack=pack if pack else None,
            is_excise=is_excise,
            family=family if family else None,
        )
        session.add(product)

    def _update_product(self, session, row_id, changes):
        product = session.query(Product).filter(Product.id == row_id).first()
        if not product:
            raise Exception(f"Не найден продукт id={row_id}")

        if "name" in changes:
            new_name = str(changes["name"]).strip()
            if not new_name:
                raise Exception("Поле name не может быть пустым")

            duplicate = (
                session.query(Product)
                .filter(Product.name == new_name, Product.id != row_id)
                .first()
            )
            if duplicate:
                raise Exception(f"Продукт с name '{new_name}' уже существует")

            product.name = new_name

        if "brand" in changes:
            value = str(changes["brand"]).strip()
            product.brand = value if value else None

        if "pack" in changes:
            value = str(changes["pack"]).strip()
            product.pack = value if value else None

        if "family" in changes:
            value = str(changes["family"]).strip()
            product.family = value if value else None

        if "is_excise" in changes:
            product.is_excise = bool(changes["is_excise"])

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

    def find_Product(self):
        self.table.clearContents()
        self.table.setRowCount(0)

        prod_data = self.get_Products_from_db()

        if not prod_data:
            self.show_message("Нет данных для отображения")
            return

        brand = self.ui.line_Brand.currentText()
        family = self.ui.line_Prod_Fam.currentText()
        product_name = self.ui.line_Prod_name.currentText()

        if brand != "-":
            prod_data = [row for row in prod_data if (row["brand"] or "") == brand]
        if family != "-":
            prod_data = [row for row in prod_data if (row["family"] or "") == family]
        if product_name != "-":
            prod_data = [row for row in prod_data if (row["name"] or "") == product_name]

        self._display_data(prod_data)

        if not prod_data:
            self.show_message("Нет данных по заданным фильтрам")

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
        item = QTableWidgetItem(value)

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

    def add_line(self):
        self._updating_table = True

        self.table.setSortingEnabled(False)
        self.table.insertRow(0)

        brand_value = self.ui.line_Brand.currentText()
        if brand_value == "-":
            brand_value = ""

        row_id = self._temp_row_id
        self._temp_row_id -= 1
        self._new_rows.add(row_id)

        values = {
            "id": str(row_id),
            "name": "",
            "brand": brand_value,
            "pack": "",
            "is_excise": False,
            "family": "",
        }

        self._pending_changes[row_id] = {
            "name": "",
            "brand": brand_value,
            "pack": "",
            "is_excise": False,
            "family": "",
        }

        for col_index, col_name in enumerate(self.columns):
            if col_name == "is_excise":
                self.table.setCellWidget(0, col_index, self._build_checkbox_widget(row_id, False))
                continue

            item = self._build_table_item(col_name, values[col_name])
            self.table.setItem(0, col_index, item)

        self._updating_table = False
        self.table.setCurrentCell(0, 1)
        self.show_message("Добавлена новая строка")

    def has_active_filters(self):
        return (
            self.ui.line_Brand.currentText() != "-"
            or self.ui.line_Prod_Fam.currentText() != "-"
            or self.ui.line_Prod_name.currentText() != "-"
        )

    def show_message(self, text):
        self.ui.label_msg.setText(text)
        self.ui.label_msg.setStyleSheet("""
            QLabel {
                background-color: #CCFF99;
                color: #12501A;
                border: 2px solid #12501A;
                border-radius: 5px;
                padding: 8px;
                font: 10pt "Tahoma";
                margin: 2px;
            }
        """)
        self.ui.label_msg.setVisible(True)

    def clear_message(self):
        self.ui.label_msg.setText("")
        self.ui.label_msg.setStyleSheet("")

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