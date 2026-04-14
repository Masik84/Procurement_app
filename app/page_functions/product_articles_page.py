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
)
from PySide6.QtCore import Qt, QFile
from PySide6.QtUiTools import QUiLoader

from app.db.models import Product, ProductArticle
from app.db.db import SessionLocal


BASE_DIR = Path(__file__).resolve().parents[2]
PRODUCT_ARTICLES_UI = BASE_DIR / "app" / "ui" / "windows" / "prod_articles.ui"


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


class ProductArticlesPage(QWidget):
    def __init__(self):
        super().__init__()

        self.ui = load_ui(PRODUCT_ARTICLES_UI)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.ui)

        self._updating_table = False
        self._original_values = {}
        self._pending_changes = {}
        self._pending_deletes = set()
        self._new_rows = set()
        self._temp_row_id = -1

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

        self._updating_table = False
        self.table.setTabKeyNavigation(True)
        self.table.setCornerButtonEnabled(False)

    def setup_connections(self):
        self.table.itemChanged.connect(self.on_item_changed)

        self.ui.line_Brand.currentTextChanged.connect(self.fill_in_prod_fam_list)
        self.ui.line_Prod_Fam.currentTextChanged.connect(self.fill_in_prod_name_list)

        self.ui.btn_Search.clicked.connect(self.find_product_articles)
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
        if not hasattr(self, "_updating_table") or self._updating_table:
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
                row = item.row()
                col = item.column()
                if row not in rows:
                    rows[row] = {}
                rows[row][col] = item.text()

            text_rows = []
            for _, cols in sorted(rows.items()):
                text_rows.append("\t".join(value for _, value in sorted(cols.items())))
            text = "\n".join(text_rows)

        clipboard.setText(text.strip())
        self.show_message("Скопировано")

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
                self.find_product_articles()
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
                    session.query(ProductArticle).filter(
                        ProductArticle.id.in_(self._pending_deletes)
                    ).delete(synchronize_session=False)

                for row_id, changes in self._pending_changes.items():
                    product_name = str(changes.get("Product name", "")).strip()
                    article = str(changes.get("Article", "")).strip()
                    variant_name = str(changes.get("Product name (variant)", "")).strip()

                    article_value = article if article else None
                    variant_value = variant_name if variant_name else None

                    if row_id in self._new_rows:
                        if not product_name:
                            raise Exception("Для новой строки поле Product name обязательно")

                        product = session.query(Product).filter(Product.name == product_name).first()
                        if not product:
                            raise Exception(f"Продукт с name '{product_name}' не найден")

                        product_article = ProductArticle(
                            product_id=product.id,
                            article=article_value,
                            name=variant_value,
                        )
                        session.add(product_article)

                    else:
                        product_article = (
                            session.query(ProductArticle)
                            .filter(ProductArticle.id == row_id)
                            .first()
                        )
                        if not product_article:
                            raise Exception(f"Не найден ProductArticle id={row_id}")

                        if "Product name" in changes:
                            if not product_name:
                                raise Exception("Поле Product name не может быть пустым")

                            product = session.query(Product).filter(Product.name == product_name).first()
                            if not product:
                                raise Exception(f"Продукт с name '{product_name}' не найден")

                            product_article.product_id = product.id

                        if "Article" in changes:
                            product_article.article = article_value

                        if "Product name (variant)" in changes:
                            product_article.name = variant_value

                session.commit()

            self._pending_changes.clear()
            self._pending_deletes.clear()
            self._new_rows.clear()

            self.refresh_all_comboboxes()

            if self.has_active_filters():
                self.find_product_articles()
            else:
                self.table.clearContents()
                self.table.setRowCount(0)

            self.show_message("Данные успешно сохранены")

        except SQLAlchemyError as e:
            self.show_error_message(f"Ошибка сохранения в базу данных: {str(e)}")
        except Exception as e:
            self.show_error_message(f"Ошибка применения изменений: {str(e)}")

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

            brands = [row[0] for row in brands if row[0]]
            self._fill_combobox(self.ui.line_Brand, brands)

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

                products = query.all()

            products.sort(
                key=lambda x: (
                    (x.family or "").lower(),
                    -(float(x.pack) if str(x.pack).replace(".", "", 1).isdigit() else -999999)
                )
            )

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

    def get_product_articles_from_db(self):
        with self.get_session() as session:
            rows = (
                session.query(ProductArticle, Product)
                .join(Product, ProductArticle.product_id == Product.id)
                .all()
            )

            data = []
            for article_row, product_row in rows:
                data.append({
                    "id": article_row.id,
                    "product_name": product_row.name,
                    "article": article_row.article,
                    "variant_name": article_row.name,
                    "family": product_row.family,
                    "pack": product_row.pack,
                    "brand": product_row.brand,
                })

            data.sort(
                key=lambda x: (
                    (x["family"] or "").lower(),
                    -(float(x["pack"]) if str(x["pack"]).replace(".", "", 1).isdigit() else -999999),
                    (x["product_name"] or "").lower(),
                    (x["article"] or "").lower(),
                    (x["variant_name"] or "").lower(),
                )
            )

        return data

    def find_product_articles(self):
        self.table.clearContents()
        self.table.setRowCount(0)

        article_data = self.get_product_articles_from_db()

        if article_data:
            brand = self.ui.line_Brand.currentText()
            family = self.ui.line_Prod_Fam.currentText()
            product_name = self.ui.line_Prod_name.currentText()

            if brand != "-":
                article_data = [row for row in article_data if (row["brand"] or "") == brand]
            if family != "-":
                article_data = [row for row in article_data if (row["family"] or "") == family]
            if product_name != "-":
                article_data = [row for row in article_data if (row["product_name"] or "") == product_name]

            self._display_data(article_data)

            if not article_data:
                self.show_message("Нет данных по заданным фильтрам")
        else:
            self.show_message("Нет данных для отображения")

    def _display_data(self, data):
        self.table.clear()
        self.table.setColumnCount(0)
        self.table.setRowCount(0)

        if not data:
            self.show_message("Нет данных для отображения")
            return

        self._updating_table = True
        self._original_values.clear()

        columns = ["id", "product_name", "article", "variant_name"]
        headers = ["id", "Product name", "Article", "Product name (variant)"]

        self.table.setColumnCount(len(headers))
        self.table.setRowCount(len(data))
        self.table.setHorizontalHeaderLabels(headers)

        for i, row_data in enumerate(data):
            row_id = int(row_data["id"])
            self._original_values[row_id] = {}

            for j, col in enumerate(columns):
                value = "" if row_data[col] is None else str(row_data[col])
                item = QTableWidgetItem(value)

                if col == "id":
                    item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                    item.setTextAlignment(Qt.AlignCenter)
                else:
                    item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable)
                    item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)

                self._original_values[row_id][col] = value
                self.table.setItem(i, j, item)

        self.table.resizeColumnsToContents()

        for i in range(self.table.columnCount()):
            if self.table.columnWidth(i) < 120:
                self.table.setColumnWidth(i, 120)

        self._updating_table = False

    def add_line(self):
        self._updating_table = True

        self.table.setSortingEnabled(False)
        self.table.insertRow(0)

        selected_product_name = self.ui.line_Prod_name.currentText()
        if selected_product_name == "-":
            selected_product_name = ""

        row_id = self._temp_row_id
        self._temp_row_id -= 1
        self._new_rows.add(row_id)

        columns = ["id", "product_name", "article", "variant_name"]
        values = {
            "id": str(row_id),
            "product_name": selected_product_name,
            "article": "",
            "variant_name": "",
        }

        self._pending_changes[row_id] = {
            "Product name": selected_product_name,
            "Article": "",
            "Product name (variant)": "",
        }

        for j, col in enumerate(columns):
            item = QTableWidgetItem(values[col])

            if col == "id":
                item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                item.setTextAlignment(Qt.AlignCenter)
            else:
                item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable)
                item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)

            self.table.setItem(0, j, item)

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
            short_text = "Произошла ошибка. Подробности ниже (используйте кнопку 'Show Details')"
            msg.setText(short_text)
            msg.setDetailedText(text)
        else:
            msg.setText(text)

        copy_button = msg.addButton("Copy", QMessageBox.ActionRole)
        msg.addButton(QMessageBox.Ok)

        def copy_text():
            QApplication.clipboard().setText(text)

        copy_button.clicked.connect(copy_text)
        msg.exec_()