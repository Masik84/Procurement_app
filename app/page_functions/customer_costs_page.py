from __future__ import annotations

from datetime import datetime
from pathlib import Path

from PySide6.QtCore import QFile, Qt, QUrl, QTimer
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QHBoxLayout,
    QMenu,
    QMessageBox,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)
from PySide6.QtUiTools import QUiLoader

from app.db.db import SessionLocal
from app.db.models import Product, TempCustomerCostImport, TempCustomerCostOption
from app.exports.customer_cost_export import CustomerCostExport
from app.services.customer_cost_import import CustomerCostImportService
from app.utils.batch import get_current_username
from app.imports.customer_cost_importer import CustomerCostImporter
from app.utils.parsers import parse_loose_number
from app.utils.text import clean_multi_spaces
from app.ui.table_style import *


BASE_DIR = Path(__file__).resolve().parents[2]
UI_PATH = BASE_DIR / "app" / "ui" / "windows" / "customer_costs.ui"
TEMPLATE_PATH = BASE_DIR / "Price request_template.xlsx"


COL_SUPPLIER_OPTION = 0
COL_PRODUCT = 1
COL_BRAND = 12


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


class CustomerCostsPage(QWidget):
    def __init__(self):
        super().__init__()
        self.ui = load_ui(UI_PATH)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.ui)

        self._updating_table = False
        self._batch_id = ""
        self._imported_by = get_current_username()
        self._current_file_path = ""
        self._table_row_ids: list[int] = []

        self.columns = [
            "selected_option_id",
            "selected_product_id",
            "manager_name",
            "customer_name",
            "supplier_article",
            "product_name",
            "pack",
            "qty_pcs",
            "volume_l",
            "purchase_type",
            "payment_terms",
            "new_product_name",
            "new_brand",
            "new_pack",
            "new_is_excise",
        ]
        self.headers = [
            "final Supplier",
            "Product name",
            "Manager name",
            "Customer name",
            "Article",
            "Product (request)",
            "Pack (req)",
            "Qty, pcs",
            "Volume, Lt",
            "Purchase Type (request)",
            "Payment Terms (request)",
            "Product name (for new)",
            "Brand (for new)",
            "Pack (for new)",
            "Excise duty (for new)",
        ]
        self.numeric_headers = {"Pack (req)", "Qty, pcs", "Volume, Lt", "Pack (for new)"}
        self.editable_field_map = {
            "Manager name": "manager_name",
            "Customer name": "customer_name",
            "Article": "supplier_article",
            "Product (request)": "product_name",
            "Pack (req)": "pack",
            "Qty, pcs": "qty_pcs",
            "Volume, Lt": "volume_l",
            "Purchase Type (request)": "purchase_type",
            "Payment Terms (request)": "payment_terms",
            "Product name (for new)": "new_product_name",
            "Pack (for new)": "new_pack",
        }

        self.setup_ui()
        self.setup_connections()
        self.start_new_batch()
        self.refresh_filters()
        self.clear_message()

    def setup_ui(self):
        self.table = self.ui.table
        setup_data_table(self.table, sorting=False)
        self.table.horizontalHeader().setSectionsMovable(False)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)

        self.ui.label_msg.setText("")
        self.ui.line_FindProduct.setToolTip("Часть названия продукта")
        self.ui.line_FindProduct.textChanged.connect(self.refresh_product_combos)
        self.ui.cbo_FindBrand.currentTextChanged.connect(self.refresh_product_combos)

    def setup_connections(self):
        self.table.cellDoubleClicked.connect(self.start_cell_edit)
        self.table.itemChanged.connect(self.on_item_changed)
        self.table.customContextMenuRequested.connect(self.show_context_menu)

        self.ui.btn_DownFile.clicked.connect(self.download_template)
        self.ui.btn_Import.clicked.connect(self.import_file)
        self.ui.btn_AddLine.clicked.connect(self.add_line)
        self.ui.btn_CalcCost.clicked.connect(self.calculate_costs)
        self.ui.btn_Save.clicked.connect(self.save_all)
        self.ui.btn_Reset.clicked.connect(self.reset_all)

    def get_session(self):
        return SessionLocal()

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
        self.ui.label_msg.setVisible(True)

    def show_error_message(self, text: str):
        self.clear_message()

        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Warning)
        msg.setWindowTitle("Ошибка")
        msg.setText(text)

        copy_btn = msg.addButton("Copy", QMessageBox.ActionRole)
        msg.addButton(QMessageBox.Ok)

        msg.exec()

        if msg.clickedButton() == copy_btn:
            QApplication.clipboard().setText(text)

    def start_new_batch(self):
        self._batch_id = datetime.now().strftime("CC_%Y%m%d_%H%M%S_%f")
        self._imported_by = get_current_username()
        self.table.clearContents()
        self.table.setRowCount(0)
        self._table_row_ids = []
    
    def refresh_filters(self):
        with self.get_session() as session:
            brands = (
                session.query(Product.brand)
                .filter(Product.brand.isnot(None), Product.brand != "")
                .distinct()
                .order_by(Product.brand.asc())
                .all()
            )

        self.ui.cbo_FindBrand.blockSignals(True)
        self.ui.cbo_FindBrand.clear()
        self.ui.cbo_FindBrand.addItem("-")
        for row in brands:
            self.ui.cbo_FindBrand.addItem(row[0])
        self.ui.cbo_FindBrand.blockSignals(False)

    def _get_filtered_products(self) -> list[Product]:
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

    def _get_brand_values(self) -> list[str]:
        with self.get_session() as session:
            rows = (
                session.query(Product.brand)
                .filter(Product.brand.isnot(None), Product.brand != "")
                .distinct()
                .order_by(Product.brand.asc())
                .all()
            )
        return [r[0] for r in rows if r[0]]

    def _get_product_name_by_id(self, product_id: int | None) -> str:
        if not product_id:
            return ""
        with self.get_session() as session:
            product = session.query(Product).filter(Product.id == product_id).first()
            return product.name if product else ""

    def _get_supplier_option_name(self, row_id: int, option_id: int | None) -> str:
        if not option_id:
            return ""
        with self.get_session() as session:
            option = (
                session.query(TempCustomerCostOption)
                .filter(
                    TempCustomerCostOption.id == option_id,
                    TempCustomerCostOption.temp_import_id == row_id,
                )
                .first()
            )
            return option.supplier_name if option else ""

    def _get_row_selected_product_id(self, row_id: int) -> int | None:
        with self.get_session() as session:
            row = session.query(TempCustomerCostImport).filter(TempCustomerCostImport.id == row_id).first()
            return row.selected_product_id if row else None

    def _get_row_selected_option_id(self, row_id: int) -> int | None:
        with self.get_session() as session:
            row = session.query(TempCustomerCostImport).filter(TempCustomerCostImport.id == row_id).first()
            return row.selected_option_id if row else None

    def _get_row_brand(self, row_id: int) -> str:
        with self.get_session() as session:
            row = session.query(TempCustomerCostImport).filter(TempCustomerCostImport.id == row_id).first()
            return row.new_brand or "" if row else ""

    def import_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Выберите файл", "", "Excel files (*.xls *.xlsx)")
        if not file_path:
            return

        try:
            importer = CustomerCostImporter()
            rows = importer.read_excel(file_path)
            self._current_file_path = file_path

            with self.get_session() as session:
                service = CustomerCostImportService(session)
                self.start_new_batch()
                service.import_rows(rows=rows, batch_id=self._batch_id, imported_by=self._imported_by)
                service.automatch_temp_rows(batch_id=self._batch_id, imported_by=self._imported_by)
                session.commit()

            self.load_table()
            self.show_message("Данные импортированы")
        except Exception as e:
            self.show_error_message(str(e))

    def load_table(self):
        with self.get_session() as session:
            rows = (
                session.query(TempCustomerCostImport)
                .filter(
                    TempCustomerCostImport.batch_id == self._batch_id,
                    TempCustomerCostImport.imported_by == self._imported_by,
                )
                .order_by(TempCustomerCostImport.import_row_no.asc(), TempCustomerCostImport.id.asc())
                .all()
            )

        self.display_rows(rows)

    def display_rows(self, rows: list[TempCustomerCostImport]):
        self._updating_table = True
        try:
            self._table_row_ids = [row.id for row in rows]

            self.table.clear()
            self.table.setColumnCount(len(self.headers))
            self.table.setRowCount(len(rows))
            self.table.setHorizontalHeaderLabels(self.headers)

            for row_index, row in enumerate(rows):
                row_id = row.id

                supplier_option_name = self._get_supplier_option_name(row_id, row.selected_option_id)
                if not supplier_option_name:
                    supplier_option_name = "-"

                self.table.setItem(
                    row_index,
                    COL_SUPPLIER_OPTION,
                    self.build_display_item(
                        row_id,
                        "selected_option_id",
                        supplier_option_name,
                    ),
                )

                self.table.setItem(
                    row_index,
                    COL_PRODUCT,
                    self.build_display_item(
                        row_id,
                        "selected_product_id",
                        self._get_product_name_by_id(row.selected_product_id),
                    ),
                )

                self.table.setItem(row_index, 2, self.build_table_item(row_id, "manager_name", self._clean_table_text(row.manager_name), align_left=True))
                self.table.setItem(row_index, 3, self.build_table_item(row_id, "customer_name", self._clean_table_text(row.customer_name), align_left=True))
                self.table.setItem(row_index, 4, self.build_table_item(row_id, "supplier_article", self._format_article_text(row.supplier_article), align_left=True))
                self.table.setItem(row_index, 5, self.build_table_item(row_id, "product_name", self._clean_table_text(row.product_name), align_left=True))
                self.table.setItem(row_index, 6, self.build_table_item(row_id, "pack", self._format_number_text(row.pack), align_left=False))
                self.table.setItem(row_index, 7, self.build_table_item(row_id, "qty_pcs", self._format_number_text(row.qty_pcs), align_left=False))
                self.table.setItem(row_index, 8, self.build_table_item(row_id, "volume_l", self._format_volume_text(row.volume_l), align_left=False))
                self.table.setItem(row_index, 9, self.build_table_item(row_id, "purchase_type", self._clean_table_text(row.purchase_type), align_left=True))
                self.table.setItem(row_index, 10, self.build_table_item(row_id, "payment_terms", self._clean_table_text(row.payment_terms), align_left=True))
                self.table.setItem(row_index, 11, self.build_table_item(row_id, "new_product_name", self._clean_table_text(row.new_product_name), align_left=True))

                self.table.setItem(
                    row_index,
                    COL_BRAND,
                    self.build_display_item(row_id, "new_brand", self._clean_table_text(row.new_brand)),
                )

                self.table.setItem(row_index, 13, self.build_table_item(row_id, "new_pack", self._format_number_text(row.new_pack), align_left=False))
                self.table.setCellWidget(
                    row_index,
                    14,
                    self.build_checkbox_widget(row_id, bool(row.new_is_excise) if row.new_is_excise is not None else False),
                )

            self.table.resizeColumnsToContents()
            self.table.resizeRowsToContents()
        finally:
            self._updating_table = False

    def _clean_table_text(self, value) -> str:
        if value is None:
            return ""
        text = str(value).strip()
        if text.lower() == "nan":
            return ""
        return text

    def _format_article_text(self, value) -> str:
        text = self._clean_table_text(value)
        if not text:
            return ""
        if text.endswith(".0"):
            try:
                return str(int(float(text)))
            except Exception:
                return text
        return text

    def _format_number_text(self, value) -> str:
        if value is None:
            return ""
        text = self._clean_table_text(value)
        if not text:
            return ""
        parsed = parse_loose_number(value)
        if parsed is None:
            return text
        return str(parsed).replace(".", ",")

    def _format_volume_text(self, value) -> str:
        if value is None:
            return ""
        parsed = parse_loose_number(value)
        if parsed is None:
            return self._clean_table_text(value)
        try:
            return f"{float(parsed):,.1f}".replace(",", " ").replace(".", ",")
        except Exception:
            return str(parsed).replace(".", ",")

    def build_table_item(self, row_id: int, column_name: str, value: str, align_left: bool = False) -> QTableWidgetItem:
        item = QTableWidgetItem(value)
        item.setData(Qt.UserRole, column_name)
        item.setData(Qt.UserRole + 1, row_id)
        item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable)
        item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter if align_left else Qt.AlignCenter)
        return item

    def build_display_item(self, row_id: int, column_name: str, value: str) -> QTableWidgetItem:
        item = QTableWidgetItem(value)
        item.setData(Qt.UserRole, column_name)
        item.setData(Qt.UserRole + 1, row_id)
        item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
        item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        return item

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

    def start_cell_edit(self, row: int, column: int):
        if self._updating_table:
            return

        if column not in (COL_SUPPLIER_OPTION, COL_PRODUCT, COL_BRAND):
            return

        if row < 0 or row >= len(self._table_row_ids):
            return

        row_id = self._table_row_ids[row]

        if column == COL_SUPPLIER_OPTION:
            current_option_id = self._get_row_selected_option_id(row_id)
            combo = self._build_supplier_option_combo(row_id, current_option_id)
            combo.activated.connect(
                lambda _, r=row, rid=row_id, c=combo: self.finish_supplier_option_edit(r, rid, c)
            )
            self.table.setCellWidget(row, column, combo)
            combo.setFocus()
            QTimer.singleShot(0, combo.showPopup)

        elif column == COL_PRODUCT:
            current_product_id = self._get_row_selected_product_id(row_id)
            combo = self._build_product_combo(row_id, current_product_id)
            combo.activated.connect(
                lambda _, r=row, rid=row_id, c=combo: self.finish_product_edit(r, rid, c)
            )
            self.table.setCellWidget(row, column, combo)
            combo.setFocus()
            QTimer.singleShot(0, combo.showPopup)

        elif column == COL_BRAND:
            current_brand = self._get_row_brand(row_id)
            combo = self.build_brand_combo(row_id, current_brand)
            combo.activated.connect(
                lambda _, r=row, rid=row_id, c=combo: self.finish_brand_edit(r, rid, c)
            )
            if combo.lineEdit() is not None:
                combo.lineEdit().returnPressed.connect(
                    lambda r=row, rid=row_id, c=combo: self.finish_brand_edit(r, rid, c)
                )
            self.table.setCellWidget(row, column, combo)
            combo.setFocus()
            QTimer.singleShot(0, combo.showPopup)

    def _build_product_combo(self, row_id: int, selected_product_id: int | None) -> QComboBox:
        combo = QComboBox()
        combo.setProperty("row_id", row_id)
        combo.setProperty("combo_role", "product_combo")
        combo.setToolTip("Выберите продукт из базы")
        self.populate_product_combo(
            combo,
            keep_current=False,
            selected_product_id=selected_product_id,
        )
        return combo

    def populate_product_combo(
        self,
        combo: QComboBox,
        keep_current: bool,
        selected_product_id: int | None = None,
    ):
        current_id = combo.currentData() if keep_current else selected_product_id

        if current_id is not None:
            try:
                current_id = int(current_id)
            except (TypeError, ValueError):
                current_id = None

        products = self._get_filtered_products()

        combo.blockSignals(True)
        combo.clear()
        combo.addItem("", None)

        for product in products:
            combo.addItem(product.name, int(product.id))

        if current_id is not None and combo.findData(current_id) < 0:
            with self.get_session() as session:
                product = session.query(Product).filter(Product.id == current_id).first()
                if product:
                    combo.addItem(product.name, int(product.id))

        idx = combo.findData(current_id)
        if idx >= 0:
            combo.setCurrentIndex(idx)
        else:
            combo.setCurrentIndex(0)

        combo.blockSignals(False)

    def _build_supplier_option_combo(self, row_id: int, selected_option_id: int | None) -> QComboBox:
        combo = QComboBox()
        combo.setProperty("row_id", row_id)
        combo.setProperty("combo_role", "supplier_option_combo")
        self.populate_supplier_option_combo(combo, row_id=row_id, keep_current=False, selected_option_id=selected_option_id)
        return combo

    def populate_supplier_option_combo(
        self,
        combo: QComboBox,
        row_id: int,
        keep_current: bool,
        selected_option_id: int | None = None,
    ):
        current_id = combo.currentData() if keep_current else selected_option_id

        if current_id is not None:
            try:
                current_id = int(current_id)
            except (TypeError, ValueError):
                current_id = None

        with self.get_session() as session:
            options = (
                session.query(TempCustomerCostOption)
                .filter(
                    TempCustomerCostOption.batch_id == self._batch_id,
                    TempCustomerCostOption.imported_by == self._imported_by,
                    TempCustomerCostOption.temp_import_id == row_id,
                )
                .order_by(
                    TempCustomerCostOption.opt_rank.asc(),
                    TempCustomerCostOption.full_cost_msk.asc(),
                    TempCustomerCostOption.id.asc(),
                )
                .all()
            )

        combo.blockSignals(True)
        combo.clear()
        combo.addItem("", None)

        for option in options:
            combo.addItem(option.supplier_name or "", int(option.id))

        idx = combo.findData(current_id)
        if idx >= 0:
            combo.setCurrentIndex(idx)
        else:
            combo.setCurrentIndex(0)

        combo.blockSignals(False)

    def build_brand_combo(self, row_id: int, brand_name: str) -> QComboBox:
        combo = QComboBox()
        combo.setEditable(True)
        combo.setInsertPolicy(QComboBox.NoInsert)
        combo.setProperty("row_id", row_id)
        combo.setProperty("combo_role", "brand_combo")
        self.populate_brand_combo(combo, keep_current=False, current_text=brand_name)
        return combo

    def populate_brand_combo(self, combo: QComboBox, keep_current: bool, current_text: str = ""):
        brand_value = combo.currentText().strip() if keep_current else current_text
        brands = self._get_brand_values()

        combo.blockSignals(True)
        combo.clear()
        combo.addItem("")
        if brands:
            combo.addItems(brands)
        if brand_value and combo.findText(brand_value) < 0:
            combo.addItem(brand_value)
        combo.setCurrentText(brand_value)
        combo.blockSignals(False)

    def finish_product_edit(self, row: int, row_id: int, combo: QComboBox):
        product_id = combo.currentData()

        try:
            product_id = int(product_id)
        except (TypeError, ValueError):
            product_id = None

        self.update_temp_field(row_id, "selected_product_id", product_id)
        product_name = self._get_product_name_by_id(product_id)

        self._updating_table = True
        self.table.removeCellWidget(row, COL_PRODUCT)
        self.table.setItem(
            row,
            COL_PRODUCT,
            self.build_display_item(row_id, "selected_product_id", product_name),
        )
        self._updating_table = False
        self.table.resizeColumnsToContents()

    def finish_supplier_option_edit(self, row: int, row_id: int, combo: QComboBox):
        option_id = combo.currentData()

        try:
            option_id = int(option_id)
        except (TypeError, ValueError):
            option_id = None

        self.update_temp_field(row_id, "selected_option_id", option_id)
        option_name = self._get_supplier_option_name(row_id, option_id) or "-"
        
        self._updating_table = True
        self.table.removeCellWidget(row, COL_SUPPLIER_OPTION)
        self.table.setItem(
            row,
            COL_SUPPLIER_OPTION,
            self.build_display_item(row_id, "selected_option_id", option_name),
        )
        self._updating_table = False
        self.table.resizeColumnsToContents()

    def finish_brand_edit(self, row: int, row_id: int, combo: QComboBox):
        brand_text = clean_multi_spaces(combo.currentText()) or None
        self.update_temp_field(row_id, "new_brand", brand_text)

        self._updating_table = True
        self.table.removeCellWidget(row, COL_BRAND)
        self.table.setItem(
            row,
            COL_BRAND,
            self.build_display_item(row_id, "new_brand", brand_text or ""),
        )
        self._updating_table = False
        self.table.resizeColumnsToContents()

    def on_checkbox_changed(self, row_id: int, checked: bool):
        if self._updating_table:
            return
        self.update_temp_field(row_id, "new_is_excise", bool(checked))

    def on_item_changed(self, item: QTableWidgetItem):
        if self._updating_table:
            return

        row_id = item.data(Qt.UserRole + 1)
        if row_id is None:
            return

        column_name = item.data(Qt.UserRole)
        if not column_name:
            return

        if column_name in {"selected_option_id", "selected_product_id", "new_brand"}:
            return

        value = clean_multi_spaces(item.text())
        header = self.headers[item.column()]

        if header in self.numeric_headers:
            value = parse_loose_number(value)
        elif value == "":
            value = None

        self.update_temp_field(row_id, column_name, value)

    def update_temp_field(self, row_id: int, field_name: str, value):
        if self._updating_table:
            return

        try:
            with self.get_session() as session:
                row = session.query(TempCustomerCostImport).filter(TempCustomerCostImport.id == row_id).first()
                if row is None:
                    return
                setattr(row, field_name, value)
                session.commit()

            if field_name == "selected_product_id":
                self.show_message("Продукт выбран")
            elif field_name == "selected_option_id":
                self.show_message("Поставщик выбран")
        except Exception as e:
            self.show_error_message(str(e))

    def refresh_product_combos(self):
        row = self.table.currentRow()
        if row < 0:
            return

        combo = self.table.cellWidget(row, COL_PRODUCT)
        if isinstance(combo, QComboBox):
            self.populate_product_combo(combo, keep_current=True)

    def add_line(self):
        try:
            with self.get_session() as session:
                service = CustomerCostImportService(session)
                next_row_no = 1
                last_row = (
                    session.query(TempCustomerCostImport)
                    .filter(
                        TempCustomerCostImport.batch_id == self._batch_id,
                        TempCustomerCostImport.imported_by == self._imported_by,
                    )
                    .order_by(TempCustomerCostImport.import_row_no.desc(), TempCustomerCostImport.id.desc())
                    .first()
                )
                if last_row and last_row.import_row_no:
                    next_row_no = int(last_row.import_row_no) + 1

                service.import_rows(
                    [{
                        "import_row_no": next_row_no,
                        "RequestDate": None,
                        "ManagerName": None,
                        "CustomerName": None,
                        "SupplierArticle": None,
                        "ProductName": None,
                        "Pack": None,
                        "QtyPcs": None,
                        "VolumeL": None,
                        "PurchaseType": None,
                        "PaymentTerms": None,
                        "Comments": None,
                    }],
                    batch_id=self._batch_id,
                    imported_by=self._imported_by,
                    replace_existing=False,
                )
                session.commit()

            self.load_table()
            if self.table.rowCount() > 0:
                self.table.setCurrentCell(0, 0)
            self.show_message("Добавлена строка")
        except Exception as e:
            self.show_error_message(str(e))

    def calculate_costs(self):
        try:
            default_name = f"CustCostCalc_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
            save_path, _ = QFileDialog.getSaveFileName(
                self,
                "Сохранить расчет",
                default_name,
                "Excel files (*.xlsx)"
            )
            if not save_path:
                return

            with self.get_session() as session:
                service = CustomerCostImportService(session)
                exporter = CustomerCostExport(session)
                service.run_calculation(self._batch_id, self._imported_by)
                exporter.export_calculated(self._batch_id, self._imported_by, save_path)
                session.commit()

            self.load_table()
            self.show_message("Расчет выполнен")
        except Exception as e:
            self.show_error_message(str(e))

    def save_all(self):
        try:
            folder = QFileDialog.getExistingDirectory(self, "Папка для файлов менеджеров")
            if not folder:
                return

            with self.get_session() as session:
                service = CustomerCostImportService(session)
                exporter = CustomerCostExport(session)
                service.save_calculations(self._batch_id, self._imported_by)
                exporter.export_kam_files(self._batch_id, self._imported_by, folder)
                session.commit()

            self.show_message("Данные сохранены")
        except Exception as e:
            self.show_error_message(str(e))

    def reset_all(self):
        try:
            with self.get_session() as session:
                service = CustomerCostImportService(session)
                service.delete_temp_rows(self._batch_id, self._imported_by)
                session.commit()

            self.start_new_batch()
            self._current_file_path = ""
            self.show_message("Форма очищена")
        except Exception as e:
            self.show_error_message(str(e))

    def download_template(self):
        if not TEMPLATE_PATH.exists():
            self.show_error_message("Шаблон не найден")
            return

        save_path, _ = QFileDialog.getSaveFileName(self, "Сохранить шаблон", TEMPLATE_PATH.name, "Excel files (*.xlsx)")
        if not save_path:
            return

        target = Path(save_path)
        if target.suffix.lower() != ".xlsx":
            target = target.with_suffix(".xlsx")

        target.write_bytes(TEMPLATE_PATH.read_bytes())
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(target)))
        self.show_message("Шаблон сохранен")

    def show_context_menu(self, position):
        menu = QMenu()
        copy_action = menu.addAction("Копировать")
        action = menu.exec_(self.table.viewport().mapToGlobal(position))
        if action == copy_action:
            self.copy_cell_content()

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
            text = "\n".join("\t".join(cols[c] for c in sorted(cols)) for _, cols in sorted(rows.items()))
            clipboard.setText(text)

        self.show_message("Скопировано")