from __future__ import annotations

from datetime import datetime
from pathlib import Path

from PySide6.QtCore import QFile, Qt, QDate, QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QHBoxLayout,
    QHeaderView,
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
from app.ui.table_style import setup_data_table


BASE_DIR = Path(__file__).resolve().parents[2]
UI_PATH = BASE_DIR / "app" / "ui" / "windows" / "customer_costs.ui"
TEMPLATE_PATH = BASE_DIR / "Price request_template.xlsx"


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
        self._current_product_combo_row: int | None = None

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

    def refresh_filters(self):
        with self.get_session() as session:
            brands = session.query(Product.brand).filter(Product.brand.isnot(None), Product.brand != "").distinct().order_by(Product.brand).all()
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
            rows = session.query(Product.brand).filter(Product.brand.isnot(None), Product.brand != "").distinct().order_by(Product.brand).all()
        return [r[0] for r in rows]

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
            rows = session.query(TempCustomerCostImport).filter(
                TempCustomerCostImport.batch_id == self._batch_id,
                TempCustomerCostImport.imported_by == self._imported_by,
            ).order_by(TempCustomerCostImport.import_row_no.asc(), TempCustomerCostImport.id.asc()).all()
            data = []
            for row in rows:
                data.append({
                    "id": row.id,
                    "selected_option_id": row.selected_option_id,
                    "selected_product_id": row.selected_product_id,
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
            self.table.clear()
            self.table.setColumnCount(len(self.headers))
            self.table.setHorizontalHeaderLabels(self.headers)
            self.table.setRowCount(len(data))
            brand_values = self._get_brand_values()

            for row_index, row_data in enumerate(data):
                row_id = row_data["id"]
                for col_index, col_name in enumerate(self.columns):
                    if col_name == "selected_option_id":
                        combo = self._build_supplier_option_combo(row_id, row_data[col_name])
                        self.table.setCellWidget(row_index, col_index, combo)
                        continue
                    if col_name == "selected_product_id":
                        combo = self._build_product_combo(row_id, row_data[col_name])
                        self.table.setCellWidget(row_index, col_index, combo)
                        continue
                    if col_name == "new_brand":
                        combo = self._build_brand_combo(row_id, row_data[col_name], brand_values)
                        self.table.setCellWidget(row_index, col_index, combo)
                        continue
                    if col_name == "new_is_excise":
                        self.table.setCellWidget(row_index, col_index, self._build_checkbox(row_id, bool(row_data[col_name])))
                        continue

                    value = row_data[col_name]
                    text = "" if value is None else str(value)
                    item = QTableWidgetItem(text)
                    item.setData(Qt.UserRole, row_id)
                    self.table.setItem(row_index, col_index, item)
        finally:
            self._updating_table = False

    def _build_checkbox(self, row_id: int, checked: bool):
        checkbox = QCheckBox()
        checkbox.setChecked(checked)
        checkbox.toggled.connect(lambda state, rid=row_id: self.update_temp_field(rid, "new_is_excise", bool(state)))
        container = QWidget()
        layout = QHBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(checkbox)
        layout.setAlignment(Qt.AlignCenter)
        return container

    def _build_brand_combo(self, row_id: int, value: str | None, brands: list[str]):
        combo = QComboBox()
        combo.setEditable(True)
        combo.addItem("")
        combo.addItems(brands)
        combo.setCurrentText(value or "")
        combo.currentTextChanged.connect(lambda text, rid=row_id: self.update_temp_field(rid, "new_brand", clean_multi_spaces(text) or None))
        return combo

    def _build_product_combo(self, row_id: int, selected_product_id: int | None):
        combo = QComboBox()
        combo.setEditable(False)
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
        combo.currentIndexChanged.connect(lambda _, rid=row_id, c=combo: self.update_temp_field(rid, "selected_product_id", c.currentData()))
        return combo

    def _build_supplier_option_combo(self, row_id: int, selected_option_id: int | None):
        combo = QComboBox()
        combo.addItem("", None)
        with self.get_session() as session:
            options = session.query(TempCustomerCostOption).filter(
                TempCustomerCostOption.batch_id == self._batch_id,
                TempCustomerCostOption.imported_by == self._imported_by,
                TempCustomerCostOption.temp_import_id == row_id,
            ).order_by(TempCustomerCostOption.opt_rank.asc(), TempCustomerCostOption.full_cost_msk.asc()).all()
            for option in options:
                combo.addItem(option.supplier_name, option.id)
        idx = combo.findData(selected_option_id)
        if idx >= 0:
            combo.setCurrentIndex(idx)
        combo.currentIndexChanged.connect(lambda _, rid=row_id, c=combo: self.update_temp_field(rid, "selected_option_id", c.currentData()))
        return combo

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
                last_row = session.query(TempCustomerCostImport).filter(
                    TempCustomerCostImport.batch_id == self._batch_id,
                    TempCustomerCostImport.imported_by == self._imported_by,
                ).order_by(TempCustomerCostImport.import_row_no.desc(), TempCustomerCostImport.id.desc()).first()
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
