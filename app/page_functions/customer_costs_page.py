from __future__ import annotations

from datetime import datetime
from pathlib import Path

from PySide6.QtCore import QFile, Qt, QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QHBoxLayout,
    QLineEdit,
    QMenu,
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


class HintLineEdit(QLineEdit):
    def __init__(self, hint_text: str, parent=None):
        super().__init__(parent)
        self.setToolTip(hint_text)

    def focusInEvent(self, event):
        super().focusInEvent(event)
        self.setToolTip(self.toolTip())


class CustomerCostsPage(QWidget):
    def __init__(self):
        super().__init__()
        self.ui = load_ui(UI_PATH)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.ui)
        self.setStyleSheet(load_stylesheet("app/ui/styles/app_styles.qss"))

        self._updating_table = False
        self._batch_id = ""
        self._imported_by = get_current_username()
        self._current_file_path = ""
        self._table_row_ids: list[int] = []

        self.columns = [
            "selected_option_id", "selected_product_id", "manager_name", "customer_name",
            "supplier_article", "product_name", "pack", "qty_pcs", "volume_l",
            "purchase_type", "payment_terms", "new_product_name", "new_brand",
            "new_pack", "new_is_excise",
        ]
        self.headers = [
            "final Supplier", "Product name", "Manager name", "Customer name",
            "Article", "Product (request)", "Pack (req)", "Qty, pcs", "Volume, Lt",
            "Purchase Type (request)", "Payment Terms (request)",
            "Product name (for new)", "Brand (for new)", "Pack (for new)", "Excise duty (for new)",
        ]
        self.numeric_headers = {"Pack (req)", "Qty, pcs", "Volume, Lt", "Pack (for new)"}
        self.header_to_column = dict(zip(self.headers, self.columns))

        self.setup_ui()
        self.setup_connections()
        self.start_new_batch()
        self.refresh_filters()
        self.show_message("")

    def setup_ui(self):
        self.table = self.ui.table
        setup_data_table(self.table, sorting=False)
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

    def show_message(self, text: str):
        self.ui.label_msg.setText(text)

    def show_error_message(self, text: str):
        self.ui.label_msg.setText(text)

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
                .order_by(Product.brand)
                .all()
            )
        self.ui.cbo_FindBrand.blockSignals(True)
        self.ui.cbo_FindBrand.clear()
        self.ui.cbo_FindBrand.addItem("-")
        for row in brands:
            self.ui.cbo_FindBrand.addItem(row[0])
        self.ui.cbo_FindBrand.blockSignals(False)

    def _get_filtered_products(self):
        with self.get_session() as session:
            query = session.query(Product).filter(Product.name.isnot(None), Product.name != "")
            brand = self.ui.cbo_FindBrand.currentText().strip()
            find_text = self.ui.line_FindProduct.text().strip()
            if brand and brand != "-":
                query = query.filter(Product.brand == brand)
            products = query.order_by(Product.name.asc()).all()
            if find_text:
                products = [p for p in products if find_text.lower() in (p.name or "").lower()]
            return products

    def _get_brand_values(self):
        with self.get_session() as session:
            rows = (
                session.query(Product.brand)
                .filter(Product.brand.isnot(None), Product.brand != "")
                .distinct()
                .order_by(Product.brand)
                .all()
            )
        return [r[0] for r in rows]

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
            option = session.query(TempCustomerCostOption).filter(
                TempCustomerCostOption.id == option_id,
                TempCustomerCostOption.temp_import_id == row_id,
            ).first()
            return option.supplier_name if option else ""

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
            data = []
            for row in rows:
                data.append({
                    "id": row.id,
                    "selected_option_id": row.selected_option_id,
                    "selected_option_name": self._get_supplier_option_name(row.id, row.selected_option_id),
                    "selected_product_id": row.selected_product_id,
                    "selected_product_name": self._get_product_name_by_id(row.selected_product_id),
                    "manager_name": row.manager_name,
                    "customer_name": row.customer_name,
                    "supplier_article": row.supplier_article,
                    "product_name": row.product_name,
                    "pack": row.pack,
                    "qty_pcs": row.qty_pcs,
                    "volume_l": row.volume_l,
                    "purchase_type": row.purchase_type,
                    "payment_terms": row.payment_terms,
                    "new_product_name": row.new_product_name,
                    "new_brand": row.new_brand,
                    "new_pack": row.new_pack,
                    "new_is_excise": bool(row.new_is_excise) if row.new_is_excise is not None else False,
                })
        self.display_table(data)

    def display_table(self, data):
        self._updating_table = True
        try:
            self._table_row_ids = [row_data["id"] for row_data in data]
            self.table.clear()
            self.table.setColumnCount(len(self.headers))
            self.table.setHorizontalHeaderLabels(self.headers)
            self.table.setRowCount(len(data))

            for row_index, row_data in enumerate(data):
                row_id = row_data["id"]
                for col_index, col_name in enumerate(self.columns):
                    if col_name == "new_is_excise":
                        self.table.setCellWidget(row_index, col_index, self._build_checkbox(row_id, bool(row_data[col_name])))
                        continue

                    if col_name == "selected_option_id":
                        item = self._build_item(row_data.get("selected_option_name") or "", row_id, align_left=True)
                        self.table.setItem(row_index, col_index, item)
                        continue

                    if col_name == "selected_product_id":
                        item = self._build_item(row_data.get("selected_product_name") or "", row_id, align_left=True)
                        self.table.setItem(row_index, col_index, item)
                        continue

                    if col_name == "new_brand":
                        item = self._build_item(row_data.get(col_name) or "", row_id, align_left=True)
                        self.table.setItem(row_index, col_index, item)
                        continue

                    value = row_data[col_name]
                    item = self._build_item(format_table_value(value), row_id, align_left=(col_name not in {"pack", "qty_pcs", "volume_l", "new_pack"}))
                    self.table.setItem(row_index, col_index, item)

            self.table.resizeColumnsToContents()
        finally:
            self._updating_table = False

    def _build_item(self, value, row_id: int, align_left: bool = False):
        text = "" if value is None else str(value)
        item = QTableWidgetItem(text)
        item.setData(Qt.UserRole, row_id)
        item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable)
        item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter if align_left else Qt.AlignCenter)
        return item

    def start_cell_edit(self, row: int, column: int):
        if self._updating_table:
            return

        if row < 0 or row >= len(self._table_row_ids):
            return

        row_id = self._table_row_ids[row]

        if column == COL_SUPPLIER_OPTION:
            current_option_id = self._get_row_selected_option_id(row_id)
            combo = self._build_supplier_option_combo(row_id, current_option_id)
            combo.activated.connect(lambda _=None, r=row, rid=row_id, cb=combo: self.finish_supplier_option_edit(r, rid, cb))
            self.table.setCellWidget(row, column, combo)
            combo.setFocus()
            combo.showPopup()

        elif column == COL_PRODUCT:
            current_name = self._get_cell_text(row, column)
            current_product_id = self._get_row_selected_product_id(row_id)
            combo = self._build_product_combo(row_id, current_product_id, current_name)
            combo.activated.connect(lambda _=None, r=row, rid=row_id, cb=combo: self.finish_product_edit(r, rid, cb))
            self.table.setCellWidget(row, column, combo)
            combo.setFocus()
            combo.showPopup()

        elif column == COL_BRAND:
            current_text = self._get_cell_text(row, column)
            combo = self._build_brand_combo(row_id, current_text)
            combo.activated.connect(lambda _=None, r=row, rid=row_id, cb=combo: self.finish_brand_edit(r, rid, cb))
            if combo.lineEdit() is not None:
                combo.lineEdit().editingFinished.connect(lambda r=row, rid=row_id, cb=combo: self.finish_brand_edit(r, rid, cb))
            self.table.setCellWidget(row, column, combo)
            combo.setFocus()
            combo.showPopup()

    def _get_cell_text(self, row: int, column: int) -> str:
        item = self.table.item(row, column)
        return item.text().strip() if item else ""

    def _get_row_selected_product_id(self, row_id: int) -> int | None:
        with self.get_session() as session:
            row = session.query(TempCustomerCostImport).filter(TempCustomerCostImport.id == row_id).first()
            return row.selected_product_id if row else None

    def _get_row_selected_option_id(self, row_id: int) -> int | None:
        with self.get_session() as session:
            row = session.query(TempCustomerCostImport).filter(TempCustomerCostImport.id == row_id).first()
            return row.selected_option_id if row else None

    def _build_brand_combo(self, row_id: int, value: str | None):
        combo = QComboBox()
        combo.setEditable(True)
        combo.setInsertPolicy(QComboBox.NoInsert)
        combo.addItem("")
        brands = self._get_brand_values()
        combo.addItems(brands)
        if value and combo.findText(value) < 0:
            combo.addItem(value)
        combo.setCurrentText(value or "")
        return combo

    def _build_product_combo(self, row_id: int, selected_product_id: int | None, selected_name: str = ""):
        combo = QComboBox()
        combo.addItem("", None)
        selected_present = False
        for product in self._get_filtered_products():
            combo.addItem(product.name, product.id)
            if selected_product_id == product.id:
                selected_present = True
        if selected_product_id and not selected_present:
            with self.get_session() as session:
                product = session.query(Product).filter(Product.id == selected_product_id).first()
                if product:
                    combo.addItem(product.name, product.id)
        idx = combo.findData(selected_product_id)
        if idx >= 0:
            combo.setCurrentIndex(idx)
        elif selected_name and combo.findText(selected_name) >= 0:
            combo.setCurrentText(selected_name)
        return combo

    def _build_supplier_option_combo(self, row_id: int, selected_option_id: int | None):
        combo = QComboBox()
        combo.addItem("", None)
        with self.get_session() as session:
            options = (
                session.query(TempCustomerCostOption)
                .filter(
                    TempCustomerCostOption.batch_id == self._batch_id,
                    TempCustomerCostOption.imported_by == self._imported_by,
                    TempCustomerCostOption.temp_import_id == row_id,
                )
                .order_by(TempCustomerCostOption.opt_rank.asc(), TempCustomerCostOption.full_cost_msk.asc())
                .all()
            )
            for option in options:
                combo.addItem(option.supplier_name, option.id)
        idx = combo.findData(selected_option_id)
        if idx >= 0:
            combo.setCurrentIndex(idx)
        return combo

    def finish_product_edit(self, row: int, row_id: int, combo: QComboBox):
        product_id = combo.currentData()
        product_name = combo.currentText().strip()
        self.update_temp_field(row_id, "selected_product_id", product_id)
        self.table.removeCellWidget(row, COL_PRODUCT)
        self.table.setItem(row, COL_PRODUCT, self._build_item(product_name, row_id, align_left=True))
        self.table.resizeColumnsToContents()

    def finish_supplier_option_edit(self, row: int, row_id: int, combo: QComboBox):
        option_id = combo.currentData()
        option_name = combo.currentText().strip()
        self.update_temp_field(row_id, "selected_option_id", option_id)
        self.table.removeCellWidget(row, COL_SUPPLIER_OPTION)
        self.table.setItem(row, COL_SUPPLIER_OPTION, self._build_item(option_name, row_id, align_left=True))
        self.table.resizeColumnsToContents()

    def finish_brand_edit(self, row: int, row_id: int, combo: QComboBox):
        brand_text = clean_multi_spaces(combo.currentText()) or None
        self.update_temp_field(row_id, "new_brand", brand_text)
        self.table.removeCellWidget(row, COL_BRAND)
        self.table.setItem(row, COL_BRAND, self._build_item(brand_text or "", row_id, align_left=True))
        self.table.resizeColumnsToContents()

    def on_item_changed(self, item):
        if self._updating_table:
            return
        row_id = item.data(Qt.UserRole)
        if row_id is None:
            return
        header = self.headers[item.column()]
        field_map = {
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
        field_name = field_map.get(header)
        if field_name:
            value = item.text().strip()
            if header in self.numeric_headers:
                value = parse_loose_number(value)
            elif value == "":
                value = None
            self.update_temp_field(row_id, field_name, value)

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
        self.load_table()

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
                service.import_rows([{
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
                }], batch_id=self._batch_id, imported_by=self._imported_by, replace_existing=False)
                session.commit()
            self.load_table()
            self.show_message("Добавлена строка")
        except Exception as e:
            self.show_error_message(str(e))

    def calculate_costs(self):
        try:
            save_path, _ = QFileDialog.getSaveFileName(self, "Сохранить расчет", "Price request.xlsx", "Excel files (*.xlsx)")
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
