from pathlib import Path

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
    QComboBox,
    QFileDialog,
    QLineEdit,
)
from PySide6.QtCore import Qt, QFile, QTimer
from PySide6.QtGui import QDesktopServices
from PySide6.QtUiTools import QUiLoader

from app.db.models import Product, ProductArticle
from app.db.db import SessionLocal
from app.ui.table_style import *
from app.imports.product_article_importer import ProductArticleImporter
from app.exports.product_article_exporter import ProductArticleExporter


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
        setup_data_table(self.table, sorting=True)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_context_menu)
        self.table.cellDoubleClicked.connect(self.on_cell_double_clicked)

    def setup_connections(self):
        self.table.itemChanged.connect(self.on_item_changed)

        self.ui.line_Brand.currentTextChanged.connect(self.fill_in_prod_fam_list)
        self.ui.line_Prod_Fam.currentTextChanged.connect(self.fill_in_prod_name_list)

        self.ui.btn_Search.clicked.connect(self.find_product_articles)
        self.ui.btn_AddLine.clicked.connect(self.add_line)
        self.ui.btn_Save.clicked.connect(self.apply_pending_changes)
        if hasattr(self.ui, "btn_SaveExcel"):
            self.ui.btn_SaveExcel.clicked.connect(self.save_to_excel)

        if hasattr(self.ui, "btn_DownFile"):
            self.ui.btn_DownFile.clicked.connect(self.download_template)
        if hasattr(self.ui, "btn_Import"):
            self.ui.btn_Import.clicked.connect(self.import_articles)
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

    def on_cell_double_clicked(self, row, column):
        if self._updating_table:
            return

        if column < 0:
            return

        header_item = self.table.horizontalHeaderItem(column)
        if not header_item:
            return

        header = header_item.text()
        if header == "Product name":
            self.start_product_edit(row)

    def start_product_edit(self, row: int):
        id_item = self.table.item(row, 0)
        if not id_item:
            return

        row_id_text = id_item.text().strip()
        if not row_id_text:
            return

        row_id = int(row_id_text)
        product_col = 1

        current_item = self.table.item(row, product_col)
        current_value = current_item.text().strip() if current_item else ""

        combo = QComboBox()
        combo.setEditable(False)
        combo.setInsertPolicy(QComboBox.NoInsert)
        combo.setFrame(False)

        product_names = self._get_product_name_values()

        combo.addItem("")
        combo.addItems(product_names)

        if current_value and combo.findText(current_value) < 0:
            combo.addItem(current_value)

        combo.setCurrentText(current_value)

        combo.setProperty("edit_row", row)
        combo.setProperty("edit_row_id", row_id)

        combo.activated.connect(lambda *_, c=combo: self.finish_product_edit_from_combo(c))

        self._updating_table = True
        self.table.setCellWidget(row, product_col, combo)
        self._updating_table = False

        combo.setFocus()
        QTimer.singleShot(0, combo.showPopup)

    def finish_product_edit_from_combo(self, combo: QComboBox):
        row = combo.property("edit_row")
        row_id = combo.property("edit_row_id")

        if row is None or row_id is None:
            return

        product_col = 1
        text = self.clean_multi_spaces(combo.currentText())

        if row_id not in self._pending_changes:
            self._pending_changes[row_id] = {}

        self._pending_changes[row_id]["Product name"] = text

        self._updating_table = True
        self.table.removeCellWidget(row, product_col)

        item = QTableWidgetItem(text)
        item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable)
        item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.table.setItem(row, product_col, item)

        self._updating_table = False
        self.table.resizeColumnsToContents()

    def _get_product_name_values(self):
        with self.get_session() as session:
            products = session.query(Product).filter(Product.name.isnot(None), Product.name != "").all()

        products.sort(
            key=lambda x: (
                (x.family or "").lower(),
                -(float(x.pack) if str(x.pack).replace(".", "", 1).isdigit() else -999999),
                (x.name or "").lower(),
            )
        )

        return [row.name for row in products if row.name]

    def clean_multi_spaces(self, text: str) -> str:
        return " ".join((text or "").split())

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

    def _normalize_text(self, value) -> str:
        return self.clean_multi_spaces(value)

    def apply_pending_changes(self):
        if not self._pending_changes and not self._pending_deletes:
            self.show_message("Нет изменений для применения")
            return

        try:
            with self.get_session() as session:
                skipped_duplicates = []

                if self._pending_deletes:
                    session.query(ProductArticle).filter(
                        ProductArticle.id.in_(self._pending_deletes)
                    ).delete(synchronize_session=False)

                for row_id, changes in self._pending_changes.items():
                    product_article = None
                    current_product_id = None
                    current_article_value = None
                    current_variant_value = None

                    if row_id not in self._new_rows:
                        product_article = (
                            session.query(ProductArticle)
                            .filter(ProductArticle.id == row_id)
                            .first()
                        )
                        if not product_article:
                            raise Exception(f"Не найден ProductArticle id={row_id}")

                        current_product_id = product_article.product_id
                        current_article_value = product_article.article
                        current_variant_value = product_article.name

                    raw_product_name = changes.get("Product name")
                    raw_article = changes.get("Article")
                    raw_variant_name = changes.get("Product name (variant)")

                    product_name = self._normalize_text(
                        raw_product_name if raw_product_name is not None else ""
                    )
                    article = self._normalize_text(
                        raw_article if raw_article is not None else (
                            current_article_value if current_article_value is not None else ""
                        )
                    )
                    variant_name = self._normalize_text(
                        raw_variant_name if raw_variant_name is not None else (
                            current_variant_value if current_variant_value is not None else ""
                        )
                    )

                    article_value = article if article else None
                    variant_value = variant_name if variant_name else None

                    if row_id in self._new_rows:
                        if not product_name:
                            raise Exception("Для новой строки поле Product name обязательно")

                        product = session.query(Product).filter(Product.name == product_name).first()
                        if not product:
                            raise Exception(f"Продукт с name '{product_name}' не найден")

                        existing_duplicate = (
                            session.query(ProductArticle)
                            .filter(
                                ProductArticle.product_id == product.id,
                                ProductArticle.article == article_value,
                                ProductArticle.name == variant_value,
                            )
                            .first()
                        )
                        if existing_duplicate:
                            skipped_duplicates.append(
                                f"{product_name} | {article_value or ''} | {variant_value or ''}"
                            )
                            continue

                        product_article = ProductArticle(
                            product_id=product.id,
                            article=article_value,
                            name=variant_value,
                        )
                        session.add(product_article)

                    else:
                        final_product_id = current_product_id
                        if "Product name" in changes:
                            if not product_name:
                                raise Exception("Поле Product name не может быть пустым")

                            product = session.query(Product).filter(Product.name == product_name).first()
                            if not product:
                                raise Exception(f"Продукт с name '{product_name}' не найден")

                            final_product_id = product.id
                        else:
                            product = session.query(Product).filter(Product.id == current_product_id).first()
                            product_name = product.name if product else ""

                        existing_duplicate = (
                            session.query(ProductArticle)
                            .filter(
                                ProductArticle.product_id == final_product_id,
                                ProductArticle.article == article_value,
                                ProductArticle.name == variant_value,
                                ProductArticle.id != row_id,
                            )
                            .first()
                        )
                        if existing_duplicate:
                            skipped_duplicates.append(
                                f"{product_name} | {article_value or ''} | {variant_value or ''}"
                            )
                            continue

                        if "Product name" in changes:
                            product_article.product_id = final_product_id

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

            if skipped_duplicates:
                self.show_message(
                    f"Данные сохранены. Пропущено дублей: {len(skipped_duplicates)}"
                )
            else:
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

    def _standalone_search_line_edits(self) -> list[QLineEdit]:
        line_edits = []
        for widget in self.ui.findChildren(QLineEdit):
            if isinstance(widget.parent(), QComboBox):
                continue
            line_edits.append(widget)

        line_edits.sort(
            key=lambda w: (
                w.mapTo(self.ui, w.rect().topLeft()).y(),
                w.mapTo(self.ui, w.rect().topLeft()).x(),
                w.objectName(),
            )
        )
        return line_edits

    def _get_search_line_edit(self, role: str):
        if role == "name":
            possible_names = [
                "line_FindProduct",
                "line_FindName",
                "line_SearchName",
                "line_SearchProduct",
                "line_SearchProductName",
                "line_ProductNameSearch",
                "line_ProdNameSearch",
                "line_FindProductName",
            ]
            keywords = ("name", "product", "prod")
        else:
            possible_names = [
                "line_FindArticle",
                "line_SearchArticle",
                "line_ArticleSearch",
                "line_FindArt",
                "line_SearchArt",
            ]
            keywords = ("article", "art")

        for name in possible_names:
            widget = getattr(self.ui, name, None)
            if isinstance(widget, QLineEdit):
                return widget

        for widget in self._standalone_search_line_edits():
            object_name = (widget.objectName() or "").lower()
            if any(keyword in object_name for keyword in keywords):
                return widget

        line_edits = self._standalone_search_line_edits()
        if role == "name" and len(line_edits) >= 1:
            return line_edits[0]
        if role == "article" and len(line_edits) >= 2:
            return line_edits[1]
        return None

    def _get_name_search_text(self) -> str:
        widget = self._get_search_line_edit("name")
        return self.clean_multi_spaces(widget.text()) if widget is not None else ""

    def _get_article_search_text(self) -> str:
        widget = self._get_search_line_edit("article")
        return self.clean_multi_spaces(widget.text()) if widget is not None else ""

    def _clear_search_fields(self):
        for widget in (self._get_search_line_edit("name"), self._get_search_line_edit("article")):
            if widget is not None:
                widget.clear()

    def _apply_current_filters(self, article_data):
        brand = self.ui.line_Brand.currentText()
        family = self.ui.line_Prod_Fam.currentText()
        product_name = self.ui.line_Prod_name.currentText()
        name_search = self._get_name_search_text().lower()
        article_search = self._get_article_search_text().lower()

        if brand != "-":
            article_data = [row for row in article_data if (row["brand"] or "") == brand]
        if family != "-":
            article_data = [row for row in article_data if (row["family"] or "") == family]
        if product_name != "-":
            article_data = [row for row in article_data if (row["product_name"] or "") == product_name]
        if name_search:
            article_data = [
                row for row in article_data
                if name_search in (row["product_name"] or "").lower()
                or name_search in (row["variant_name"] or "").lower()
            ]
        if article_search:
            article_data = [
                row for row in article_data
                if article_search in str(row["article"] or "").lower()
            ]

        return article_data

    def find_product_articles(self):
        self.table.clearContents()
        self.table.setRowCount(0)

        article_data = self.get_product_articles_from_db()

        if article_data:
            article_data = self._apply_current_filters(article_data)
            self._display_data(article_data)

            if not article_data:
                self.show_message("Нет данных по заданным фильтрам")
        else:
            self.show_message("Нет данных для отображения")

    def get_filtered_product_articles(self):
        article_data = self.get_product_articles_from_db()

        if not article_data:
            return []

        return self._apply_current_filters(article_data)

    def save_to_excel(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Сохранить Excel",
            "product_articles.xlsx",
            "Excel Files (*.xlsx)"
        )

        if not file_path:
            return

        if not file_path.lower().endswith(".xlsx"):
            file_path += ".xlsx"

        rows = self.get_filtered_product_articles()
        if not rows:
            self.show_message("Нет данных для выгрузки")
            return

        try:
            self._export_product_articles_to_excel(file_path, rows)
            self.show_message("Данные сохранены в Excel")
        except PermissionError:
            self.show_error_message(
                "Не удалось сохранить файл Excel. Возможно, файл уже открыт."
            )
        except Exception as e:
            self.show_error_message(f"Ошибка при сохранении Excel: {str(e)}")

    def _export_product_articles_to_excel(self, file_path: str, rows: list[dict]):
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
                "Article",
                "Product name (variant)",
            ]

            for col_index, header in enumerate(headers, start=1):
                ws.Cells(1, col_index).Value = header

            ws.Cells.Font.Name = "Aptos Narrow"
            ws.Cells.Font.Size = 11

            header_range = ws.Range("A1:D1")
            header_range.Font.Name = "Aptos Narrow"
            header_range.Font.Size = 11
            header_range.Font.Bold = True
            header_range.Interior.Color = 0xCDCDCD
            header_range.WrapText = True
            header_range.HorizontalAlignment = -4108
            header_range.VerticalAlignment = -4160

            ws.Rows(1).EntireRow.AutoFit()

            try:
                ws.Range("A1:D1").AutoFilter(1)
            except Exception:
                pass

            ws.Columns("A:A").ColumnWidth = 10
            ws.Columns("B:B").ColumnWidth = 34
            ws.Columns("C:C").ColumnWidth = 22
            ws.Columns("D:D").ColumnWidth = 34

            ws.Columns("C:C").NumberFormat = "@"

            row_num = 2
            for row in rows:
                ws.Cells(row_num, 1).Value = row.get("id")
                ws.Cells(row_num, 2).Value = row.get("product_name", "") or ""

                article_value = row.get("article", "")
                if article_value is None:
                    article_value = ""
                ws.Cells(row_num, 3).Value = str(article_value)

                ws.Cells(row_num, 4).Value = row.get("variant_name", "") or ""
                row_num += 1

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

    def download_template(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Сохранить шаблон",
            str(Path.home() / "ProductArticleTemplate.xlsx"),
            "Excel files (*.xlsx)",
        )
        if not file_path:
            return

        if not file_path.lower().endswith(".xlsx"):
            file_path += ".xlsx"

        try:
            exporter = ProductArticleExporter()
            exporter.export_template(file_path)
            QDesktopServices.openUrl(Path(file_path).as_uri())
            self.show_message("Шаблон сохранен")
        except PermissionError:
            self.show_error_message(
                "Не удалось сохранить файл Excel. Возможно, файл уже открыт."
            )
        except Exception as e:
            self.show_error_message(f"Ошибка при создании шаблона: {str(e)}")

    def import_articles(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите файл",
            "",
            "Excel files (*.xlsx *.xls)",
        )
        if not file_path:
            return

        try:
            importer = ProductArticleImporter()
            rows = importer.read_excel(file_path)
            imported_count, skipped_names = self._load_imported_articles(rows)

            if skipped_names:
                names_text = "\n".join(skipped_names)
                self.show_error_message(
                    "В БД не найдены продукты. Эти строки пропущены:\n\n" + names_text
                )

            self.show_message(f"Импортировано строк: {imported_count}")
        except Exception as e:
            self.show_error_message(str(e))

    def _load_imported_articles(self, rows):
        self._updating_table = True
        self.table.setSortingEnabled(False)

        self._pending_changes.clear()
        self._pending_deletes.clear()
        self._new_rows.clear()
        self._original_values.clear()
        self.table.clear()

        columns = ["id", "product_name", "article", "variant_name"]
        headers = ["id", "Product name", "Article", "Product name (variant)"]

        valid_rows = []
        skipped_names = []

        with self.get_session() as session:
            existing_products = {
                (product.name or "").strip().upper(): product
                for product in session.query(Product).filter(Product.name.isnot(None), Product.name != "").all()
            }

            seen_skipped = set()
            for imported in rows:
                product_name = self.clean_multi_spaces(imported.get("product_name", "")).upper()
                if not product_name:
                    continue

                if product_name not in existing_products:
                    if product_name not in seen_skipped:
                        skipped_names.append(product_name)
                        seen_skipped.add(product_name)
                    continue

                valid_rows.append(
                    {
                        "product_name": product_name,
                        "article": self.normalize_article_value(imported.get("article", "")),
                        "variant_name": self.clean_multi_spaces(imported.get("variant_name", "")).upper(),
                    }
                )

        unique_valid_rows = []
        seen_keys = set()

        for row in valid_rows:
            key = (
                row["product_name"],
                row["article"] or "",
                row["variant_name"] or "",
            )
            if key in seen_keys:
                continue
            seen_keys.add(key)
            unique_valid_rows.append(row)

        valid_rows = unique_valid_rows

        self.table.setColumnCount(len(headers))
        self.table.setRowCount(len(valid_rows))
        self.table.setHorizontalHeaderLabels(headers)

        for i, row in enumerate(valid_rows):
            row_id = self._temp_row_id
            self._temp_row_id -= 1
            self._new_rows.add(row_id)

            self._pending_changes[row_id] = {
                "Product name": row["product_name"],
                "Article": row["article"],
                "Product name (variant)": row["variant_name"],
            }

            values = {
                "id": str(row_id),
                "product_name": row["product_name"],
                "article": row["article"],
                "variant_name": row["variant_name"],
            }
            self._original_values[row_id] = values.copy()

            for j, col in enumerate(columns):
                item = QTableWidgetItem(values[col])
                if col == "id":
                    item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                    item.setTextAlignment(Qt.AlignCenter)
                else:
                    item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable)
                    item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                self.table.setItem(i, j, item)

        self.table.resizeColumnsToContents()
        for i in range(self.table.columnCount()):
            if self.table.columnWidth(i) < 120:
                self.table.setColumnWidth(i, 120)

        self._updating_table = False
        return len(valid_rows), skipped_names

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
        self.clear_message()
        self._updating_table = True

        self.table.setSortingEnabled(False)

        columns = ["id", "product_name", "article", "variant_name"]
        headers = ["id", "Product name", "Article", "Product name (variant)"]

        if self.table.columnCount() == 0:
            self.table.setColumnCount(len(headers))
            self.table.setHorizontalHeaderLabels(headers)

        self.table.insertRow(0)

        selected_product_name = self.ui.line_Prod_name.currentText()
        if selected_product_name == "-":
            selected_product_name = ""

        row_id = self._temp_row_id
        self._temp_row_id -= 1
        self._new_rows.add(row_id)

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

        for i in range(self.table.columnCount()):
            if self.table.columnWidth(i) < 120:
                self.table.setColumnWidth(i, 120)

        self._updating_table = False
        self.table.setCurrentCell(0, 1)
        self.show_message("Добавлена новая строка")

    def has_active_filters(self):
        return (
            self.ui.line_Brand.currentText() != "-"
            or self.ui.line_Prod_Fam.currentText() != "-"
            or self.ui.line_Prod_name.currentText() != "-"
            or bool(self._get_name_search_text())
            or bool(self._get_article_search_text())
        )

    def normalize_article_value(self, value) -> str:
        if value is None:
            return ""

        text = str(value).strip()
        if not text or text.lower() == "nan":
            return ""

        if text.endswith(".0"):
            left = text[:-2]
            if left.isdigit():
                return left

        return self.clean_multi_spaces(text)

    def reset_form(self):
        try:
            self._pending_changes.clear()
            self._pending_deletes.clear()
            self._new_rows.clear()
            self._original_values.clear()
            self._temp_row_id = -1

            self._clear_search_fields()
            self.refresh_all_comboboxes()
            self.table.clearContents()
            self.table.setRowCount(0)
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