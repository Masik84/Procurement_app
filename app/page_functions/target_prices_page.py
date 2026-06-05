from __future__ import annotations

from datetime import datetime
from decimal import Decimal
from pathlib import Path

from PySide6.QtCore import QFile, Qt, QEvent, QPoint, QTimer, QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QComboBox, QFileDialog, QHBoxLayout, QLineEdit,
    QMenu, QMessageBox, QTableWidgetItem, QToolTip, QVBoxLayout, QWidget,
)
from PySide6.QtUiTools import QUiLoader

from app.db.db import SessionLocal
from app.db.models import ExchangeRate, Product, Supplier, TempTargetPriceImport, TempTargetPriceOption
from app.exports.target_price_exporter import TargetPriceExporter
from app.imports.target_price_importer import TargetPriceImporter
from app.services.supplier_service import SupplierService, SupplierUpsertData
from app.services.target_price_service import TargetPriceService
from app.utils.batch import get_current_username
from app.utils.parsers import parse_loose_number
from app.utils.text import clean_multi_spaces
from app.ui.table_style import *
from app.workers.excel_export_worker import start_excel_export


BASE_DIR = Path(__file__).resolve().parents[2]
UI_PATH = BASE_DIR / "app" / "ui" / "windows" / "target_prices.ui"

COL_SUPPLIER_OPTION = 0
COL_PRODUCT = 1
COL_BRAND = 5


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


class TargetPricesPage(QWidget):
    def __init__(self):
        super().__init__()
        self.ui = load_ui(UI_PATH)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.ui)

        self.imported_by = get_current_username()
        self.batch_id = ""
        self._updating_table = False
        self._table_row_ids: list[int] = []
        self._showing_options = False
        self._excel_export_thread = None
        self._excel_export_worker = None
        self._manual_full_cost_visible = False
        self._manual_full_costs: dict[int, Decimal] = {}

        self.import_headers = [
            "Our Product Name",
            "Supplier Article",
            "Supplier Product Name",
            "Product name (for new)",
            "Brand (for new)",
            "Pack (for new)",
            "Excise duty (for new)",
        ]
        self.calc_headers = [
            "final Supplier",
            "Our Product Name",
            "Supplier Article",
            "Supplier Product Name",
        ]
        self.numeric_headers = {"Pack (for new)"}

        self.setup_ui()
        self.setup_connections()
        self.load_initial_state()

    def setup_ui(self):
        self.table = self.ui.table
        setup_data_table(self.table, sorting=False)
        self.table.horizontalHeader().setSectionsMovable(False)
        from app.utils.gui_table_actions import install_standard_table_context_menu
        install_standard_table_context_menu(self, self.table)
        self.ui.label_msg.setText("Сообщений нет")
        if hasattr(self.ui, "spb_SuppPriceAge"):
            self.ui.spb_SuppPriceAge.setMinimum(0)
            self.ui.spb_SuppPriceAge.setMaximum(120)
            self.ui.spb_SuppPriceAge.setValue(6)
        self.ui.line_NewSupplier.setEnabled(False)
        self.ui.line_NewSupplier.setStyleSheet("background-color: #f2f2f2;")
        for widget, tip in [
            (self.ui.line_ExchangeRate, "Формат: 82,0000"),
            (self.ui.line_Transport, "Формат: 1,2500"),
            (self.ui.line_AgentFee, "Формат: 0,2500"),
            (self.ui.line_Reexport, "Формат: 3,5% / 0,24%"),
            (self.ui.line_FXMarkup, "Формат: 3,5% / 0,24%"),
        ]:
            widget.setToolTip(tip)
            widget.installEventFilter(self)

    def setup_connections(self):
        self.table.cellDoubleClicked.connect(self.start_cell_edit)
        self.table.itemChanged.connect(self.on_item_changed)
        self.ui.cbo_SupplName.currentIndexChanged.connect(self.on_supplier_changed)
        self.ui.cbx_NewSupplier.toggled.connect(self.on_new_supplier_toggled)
        self.ui.cbo_Currency.currentIndexChanged.connect(self.on_currency_changed)
        self.ui.cbo_FindBrand.currentTextChanged.connect(self.refresh_current_product_combo)
        self.ui.line_FindProduct.textChanged.connect(self.refresh_current_product_combo)
        self.ui.line_ExchangeRate.editingFinished.connect(self.normalize_exchange_rate)
        self.ui.line_Transport.editingFinished.connect(self.normalize_transport)
        self.ui.line_AgentFee.editingFinished.connect(self.normalize_agent_fee)
        self.ui.line_Reexport.editingFinished.connect(self.normalize_reexport)
        self.ui.line_FXMarkup.editingFinished.connect(self.normalize_fx_markup)
        self.ui.btn_DownFile.clicked.connect(self.download_template)
        self.ui.btn_Import.clicked.connect(self.import_file)
        self.ui.btn_AddLine.clicked.connect(self.add_line)
        self.ui.btn_CalcCost.clicked.connect(self.calculate_costs)
        if hasattr(self.ui, "btn_ManualPrice"):
            self.ui.btn_ManualPrice.clicked.connect(self.add_manual_price_column)
        self.ui.btn_Save.clicked.connect(self.save_all)
        self.ui.btn_Reset.clicked.connect(self.reset_all)

    def get_session(self):
        return SessionLocal()

    def get_supplier_price_age_months(self) -> int:
        widget = getattr(self.ui, "spb_SuppPriceAge", None)
        if widget is None:
            return 6
        return int(widget.value())

    def load_initial_state(self):
        self.start_new_batch()
        self.load_suppliers()
        self.load_currencies()
        self.load_find_brands()
        self.apply_default_values()
        with self.get_session() as session:
            TargetPriceService(session).cleanup_old_temp_rows(self.imported_by)
            session.commit()

    def eventFilter(self, watched, event):
        if watched in {self.ui.line_ExchangeRate, self.ui.line_Transport, self.ui.line_AgentFee, self.ui.line_Reexport, self.ui.line_FXMarkup}:
            if event.type() in {QEvent.Enter, QEvent.FocusIn, QEvent.MouseButtonPress}:
                QToolTip.showText(watched.mapToGlobal(QPoint(0, watched.height())), watched.toolTip(), watched)
        return super().eventFilter(watched, event)

    def show_message(self, text):
        self.ui.label_msg.setText(text)
        self.ui.label_msg.setProperty("active", True)
        self.ui.label_msg.style().unpolish(self.ui.label_msg)
        self.ui.label_msg.style().polish(self.ui.label_msg)
        self.ui.label_msg.setVisible(True)

    def show_error_message(self, text: str):
        text = str(text or "Неизвестная ошибка").strip() or "Неизвестная ошибка"
        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Warning)
        msg.setWindowTitle("Ошибка")
        msg.setTextFormat(Qt.PlainText)
        msg.setText(text)
        msg.setMinimumWidth(520)
        msg.setStyleSheet("QMessageBox { background-color: #fffaf4; } QMessageBox QLabel { color: #262626; } QPushButton { color: #262626; }")
        copy_btn = msg.addButton("Copy", QMessageBox.ActionRole)
        msg.addButton(QMessageBox.Ok)
        msg.exec()
        if msg.clickedButton() == copy_btn:
            QApplication.clipboard().setText(text)

    def set_combo_text(self, combo: QComboBox, value: str):
        idx = combo.findText(value)
        if idx >= 0:
            combo.setCurrentIndex(idx)
        elif combo.isEditable():
            combo.setCurrentText(value)

    def start_new_batch(self):
        with self.get_session() as session:
            self.batch_id = TargetPriceService(session).start_batch()
        self._showing_options = False
        self.table.clearContents()
        self.table.setRowCount(0)
        self._table_row_ids = []

    def load_suppliers(self):
        current_text = self.ui.cbo_SupplName.currentText().strip()
        with self.get_session() as session:
            suppliers = SupplierService(session).get_all_suppliers()
        self.ui.cbo_SupplName.blockSignals(True)
        self.ui.cbo_SupplName.clear()
        self.ui.cbo_SupplName.addItem("-", None)
        for supplier in suppliers:
            if (supplier.name or "").strip().lower() == "manual":
                continue
            self.ui.cbo_SupplName.addItem(supplier.name, supplier.id)
        self.ui.cbo_SupplName.blockSignals(False)
        if current_text:
            self.set_combo_text(self.ui.cbo_SupplName, current_text)

    def load_currencies(self):
        current_text = self.ui.cbo_Currency.currentText().strip()
        with self.get_session() as session:
            rows = session.query(ExchangeRate.currency_code).order_by(ExchangeRate.currency_code.asc()).all()
        self.ui.cbo_Currency.blockSignals(True)
        self.ui.cbo_Currency.clear()
        self.ui.cbo_Currency.addItem("-")
        for row in rows:
            if row[0]:
                self.ui.cbo_Currency.addItem(row[0])
        self.ui.cbo_Currency.blockSignals(False)
        if current_text:
            self.set_combo_text(self.ui.cbo_Currency, current_text)

    def load_find_brands(self):
        current_text = self.ui.cbo_FindBrand.currentText().strip()
        brands = self.get_brand_names()
        self.ui.cbo_FindBrand.blockSignals(True)
        self.ui.cbo_FindBrand.clear()
        self.ui.cbo_FindBrand.addItem("-")
        self.ui.cbo_FindBrand.addItems(brands)
        self.ui.cbo_FindBrand.blockSignals(False)
        if current_text:
            self.set_combo_text(self.ui.cbo_FindBrand, current_text)

    def get_brand_names(self) -> list[str]:
        with self.get_session() as session:
            rows = session.query(Product.brand).filter(Product.brand.isnot(None), Product.brand != "").distinct().order_by(Product.brand.asc()).all()
        return [r[0] for r in rows if r[0]]

    def apply_default_values(self):
        self.set_combo_text(self.ui.cbo_SupplName, "-")
        self.ui.cbx_NewSupplier.setChecked(False)
        self.ui.line_NewSupplier.clear()
        self.set_combo_text(self.ui.cbo_SupplierRF, "нет")
        self.set_combo_text(self.ui.cbo_Currency, "-")
        self.ui.line_ExchangeRate.clear()
        self.ui.line_Transport.clear()
        self.ui.line_AgentFee.clear()
        self.set_combo_text(self.ui.cbo_viaNovo, "через Ново")
        self.ui.line_Reexport.setText("0,0%")
        self.ui.line_FXMarkup.setText("0,0%")
        self.set_combo_text(self.ui.cbo_Customs, "да")
        self.set_combo_text(self.ui.cbo_Marking, "Феникс")
        self.set_combo_text(self.ui.cbo_History, "да")
        self.ui.line_FindProduct.clear()
        self.set_combo_text(self.ui.cbo_FindBrand, "-")
        self.toggle_new_supplier_field(False)

    def on_new_supplier_toggled(self, checked: bool):
        if checked:
            self.set_combo_text(self.ui.cbo_SupplName, "-")
            self.ui.line_NewSupplier.clear()
            self.ui.line_ExchangeRate.clear()
            self.ui.line_Transport.clear()
            self.ui.line_AgentFee.clear()
            self.ui.line_Reexport.setText("0,0%")
            self.ui.line_FXMarkup.setText("0,0%")
            self.set_combo_text(self.ui.cbo_Currency, "-")
            self.set_combo_text(self.ui.cbo_viaNovo, "через Ново")
            self.set_combo_text(self.ui.cbo_Customs, "да")
            self.set_combo_text(self.ui.cbo_Marking, "Феникс")
            self.set_combo_text(self.ui.cbo_SupplierRF, "нет")
        self.toggle_new_supplier_field(checked)

    def toggle_new_supplier_field(self, enabled: bool):
        self.ui.line_NewSupplier.setEnabled(enabled)
        self.ui.line_NewSupplier.setStyleSheet("" if enabled else "background-color: #f2f2f2;")
        if enabled:
            self.ui.line_NewSupplier.setFocus()

    def on_supplier_changed(self):
        supplier_id = self.ui.cbo_SupplName.currentData()
        if supplier_id is None:
            return
        with self.get_session() as session:
            service = SupplierService(session)
            data = service.load_supplier_snapshot(int(supplier_id))
            rate = service.get_rate_to_rub(data.base_currency)
        self.ui.cbx_NewSupplier.blockSignals(True)
        self.ui.cbx_NewSupplier.setChecked(False)
        self.ui.cbx_NewSupplier.blockSignals(False)
        self.toggle_new_supplier_field(False)
        self.ui.line_NewSupplier.setText(data.name)
        self.set_combo_text(self.ui.cbo_SupplierRF, "да" if data.is_rf else "нет")
        self.set_combo_text(self.ui.cbo_Currency, data.base_currency or "-")
        self.ui.line_ExchangeRate.setText(self.format_number(rate, 4) if rate is not None else "")
        self.ui.line_Transport.setText(self.format_number(data.transport_cost_per_l, 4))
        self.ui.line_AgentFee.setText(self.format_number(data.agent_fee, 4))
        self.set_combo_text(self.ui.cbo_viaNovo, "через Ново" if data.is_via_novo else "в Мск")
        self.ui.line_Reexport.setText(self.format_percent(data.reexport_percent))
        self.ui.line_FXMarkup.setText(self.format_percent(data.fx_rate_markup))
        self.set_combo_text(self.ui.cbo_Customs, "да" if data.has_import_duty else "нет")
        self.set_combo_text(self.ui.cbo_Marking, "Поставщик" if data.marks_for_us else "Феникс")

    def on_currency_changed(self):
        currency = clean_multi_spaces(self.ui.cbo_Currency.currentText()).upper()
        if not currency or currency == "-":
            return
        with self.get_session() as session:
            rate = SupplierService(session).get_rate_to_rub(currency)
        if rate is not None:
            self.ui.line_ExchangeRate.setText(self.format_number(rate, 4))

    def get_supplier_form_data(self) -> SupplierUpsertData:
        supplier_name = clean_multi_spaces(self.ui.line_NewSupplier.text())
        if not supplier_name and not self.ui.cbx_NewSupplier.isChecked():
            supplier_name = clean_multi_spaces(self.ui.cbo_SupplName.currentText())
        currency = clean_multi_spaces(self.ui.cbo_Currency.currentText()).upper()
        if currency == "-":
            currency = ""
        return SupplierUpsertData(
            name=supplier_name,
            base_currency=currency,
            transport_cost_per_l=self.parse_decimal_field(self.ui.line_Transport, "Транспорт"),
            agent_fee=self.parse_decimal_field(self.ui.line_AgentFee, "Agent fee"),
            reexport_percent=self.parse_percent_field(self.ui.line_Reexport, "Реэкспорт"),
            fx_rate_markup=self.parse_percent_field(self.ui.line_FXMarkup, "FX markup"),
            is_via_novo=self.ui.cbo_viaNovo.currentText() == "через Ново",
            has_import_duty=self.ui.cbo_Customs.currentText() == "да",
            rating_calc=True,
            marks_for_us=self.ui.cbo_Marking.currentText() == "Поставщик",
            is_rf=self.ui.cbo_SupplierRF.currentText() == "да",
        )

    def ensure_supplier(self) -> int:
        supplier_id = self.ui.cbo_SupplName.currentData()
        if self.ui.cbx_NewSupplier.isChecked():
            supplier_id = None
            if not clean_multi_spaces(self.ui.line_NewSupplier.text()):
                raise ValueError("Введите название нового поставщика.")
        elif supplier_id is None:
            raise ValueError("Выберите поставщика.")
        currency = clean_multi_spaces(self.ui.cbo_Currency.currentText()).upper()
        if not currency or currency == "-":
            raise ValueError("Выберите валюту поставщика.")
        with self.get_session() as session:
            service = SupplierService(session)
            supplier = service.ensure_supplier(
                supplier_id=int(supplier_id) if supplier_id is not None else None,
                data=self.get_supplier_form_data(),
            )
            fx_rate = self.parse_decimal_field(self.ui.line_ExchangeRate, "Курс")
            if fx_rate is not None and supplier.base_currency:
                service.save_exchange_rate(supplier.base_currency, float(fx_rate))
            session.commit()
            supplier_id = supplier.id
        self.load_suppliers()
        idx = self.ui.cbo_SupplName.findData(supplier_id)
        if idx >= 0:
            self.ui.cbo_SupplName.setCurrentIndex(idx)
        self.ui.cbx_NewSupplier.setChecked(False)
        return int(supplier_id)

    def parse_decimal_field(self, widget: QLineEdit, field_name: str) -> Decimal:
        text = clean_multi_spaces(widget.text())
        if not text:
            return Decimal("0")
        value = parse_loose_number(text)
        if value is None:
            raise ValueError(f"Некорректное поле: {field_name}")
        return Decimal(str(value))

    def parse_percent_field(self, widget: QLineEdit, field_name: str) -> Decimal:
        text = clean_multi_spaces(widget.text())
        if not text:
            return Decimal("0")
        has_percent = "%" in text
        value = parse_loose_number(text.replace("%", ""))
        if value is None:
            raise ValueError(f"Некорректное поле: {field_name}")
        d = Decimal(str(value))
        if has_percent or abs(d) >= Decimal("1"):
            d = d / Decimal("100")
        return d

    def format_number(self, value: object, digits: int = 4) -> str:
        number = parse_loose_number(value)
        if number is None:
            return ""
        return f"{float(number):,.{digits}f}".replace(",", "_").replace(".", ",").replace("_", "")

    def format_percent(self, value: object) -> str:
        number = parse_loose_number(value)
        if number is None:
            return "0,0%"
        return f"{float(number) * 100:.1f}".replace(".", ",") + "%"

    def normalize_exchange_rate(self): self._normalize_number_widget(self.ui.line_ExchangeRate, 4)
    def normalize_transport(self): self._normalize_number_widget(self.ui.line_Transport, 4)
    def normalize_agent_fee(self): self._normalize_number_widget(self.ui.line_AgentFee, 4)
    def normalize_reexport(self): self._normalize_percent_widget(self.ui.line_Reexport)
    def normalize_fx_markup(self): self._normalize_percent_widget(self.ui.line_FXMarkup)

    def _normalize_number_widget(self, widget: QLineEdit, digits: int):
        text = clean_multi_spaces(widget.text())
        if not text:
            return
        value = parse_loose_number(text)
        if value is None:
            self.show_error_message("Проверь число")
            return
        widget.setText(self.format_number(value, digits))

    def _normalize_percent_widget(self, widget: QLineEdit):
        text = clean_multi_spaces(widget.text())
        if not text:
            widget.setText("0,0%")
            return
        has_percent = "%" in text
        value = parse_loose_number(text.replace("%", ""))
        if value is None:
            self.show_error_message("Проверь процент")
            return
        d = Decimal(str(value))
        if has_percent or abs(d) >= Decimal("1"):
            d = d / Decimal("100")
        widget.setText(self.format_percent(d))

    def import_file(self):
        try:
            supplier_id = self.ensure_supplier()
        except Exception as e:
            self.show_error_message(str(e))
            return
        file_path, _ = QFileDialog.getOpenFileName(self, "Выберите файл", str(BASE_DIR), "Excel files (*.xls *.xlsx)")
        if not file_path:
            return
        try:
            rows = TargetPriceImporter().read_excel(file_path)
            with self.get_session() as session:
                service = TargetPriceService(session)
                self.start_new_batch()
                service.import_rows(rows=rows, batch_id=self.batch_id, imported_by=self.imported_by, supplier_id=supplier_id)
                service.automatch_temp_rows(self.batch_id, self.imported_by)
                session.commit()
            self._showing_options = False
            self.load_table()
            self.show_message("Данные импортированы")
        except Exception as e:
            self.show_error_message(str(e))

    def load_table(self):
        with self.get_session() as session:
            rows = session.query(TempTargetPriceImport).filter(
                TempTargetPriceImport.batch_id == self.batch_id,
                TempTargetPriceImport.imported_by == self.imported_by,
            ).order_by(TempTargetPriceImport.import_row_no.asc(), TempTargetPriceImport.id.asc()).all()
        self.display_rows(rows)

    def display_rows(self, rows: list[TempTargetPriceImport]):
        self._updating_table = True
        try:
            self._table_row_ids = [r.id for r in rows]
            headers = list(self.calc_headers if self._showing_options else self.import_headers)
            if self._showing_options and self._manual_full_cost_visible and "Manual Full Cost Msk" not in headers:
                headers.append("Manual Full Cost Msk")
            self.table.clear()
            self.table.setColumnCount(len(headers))
            self.table.setRowCount(len(rows))
            self.table.setHorizontalHeaderLabels(headers)
            for row_index, row in enumerate(rows):
                row_id = row.id
                if self._showing_options:
                    supplier_option_name = self._get_supplier_option_name(row_id, row.selected_option_id) or "-"
                    self.table.setItem(row_index, 0, self.build_display_item(row_id, "selected_option_id", supplier_option_name))
                    self.table.setItem(row_index, 1, self.build_display_item(row_id, "selected_product_id", self._get_product_name_by_id(row.selected_product_id)))
                    self.table.setItem(row_index, 2, self.build_table_item(row_id, "supplier_article", self._format_article_text(row.supplier_article), True))
                    self.table.setItem(row_index, 3, self.build_table_item(row_id, "product_name", self._clean_table_text(row.product_name), True))
                    if self._manual_full_cost_visible:
                        manual_value = self._manual_full_costs.get(row_id)
                        if manual_value is None:
                            manual_value = self._get_manual_full_cost(row_id)
                        self.table.setItem(row_index, 4, self.build_table_item(row_id, "manual_full_cost_msk", self._format_number_text(manual_value), False))
                else:
                    self.table.setItem(row_index, 0, self.build_display_item(row_id, "selected_product_id", self._get_product_name_by_id(row.selected_product_id)))
                    self.table.setItem(row_index, 1, self.build_table_item(row_id, "supplier_article", self._format_article_text(row.supplier_article), True))
                    self.table.setItem(row_index, 2, self.build_table_item(row_id, "product_name", self._clean_table_text(row.product_name), True))
                    self.table.setItem(row_index, 3, self.build_table_item(row_id, "new_product_name", self._clean_table_text(row.new_product_name), True))
                    self.table.setItem(row_index, 4, self.build_display_item(row_id, "new_brand", self._clean_table_text(row.new_brand)))
                    self.table.setItem(row_index, 5, self.build_table_item(row_id, "new_pack", self._format_number_text(row.new_pack), False))
                    self.table.setCellWidget(row_index, 6, self.build_checkbox_widget(row_id, bool(row.new_is_excise) if row.new_is_excise is not None else False))
            self.table.resizeColumnsToContents()
            self.table.resizeRowsToContents()
        finally:
            self._updating_table = False

    def _clean_table_text(self, value) -> str:
        if value is None:
            return ""
        text = str(value).strip()
        return "" if text.lower() == "nan" else text

    def _format_article_text(self, value) -> str:
        text = self._clean_table_text(value)
        if text.endswith(".0"):
            try: return str(int(float(text)))
            except Exception: return text
        return text

    def _format_number_text(self, value) -> str:
        parsed = parse_loose_number(value)
        return self._clean_table_text(value) if parsed is None else str(parsed).replace(".", ",")

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
        checkbox.setStyleSheet("QCheckBox {background: transparent;} QCheckBox::indicator {width: 14px; height: 14px;}")
        checkbox.toggled.connect(lambda state, rid=row_id: self.on_checkbox_changed(rid, state))
        container = QWidget()
        layout = QHBoxLayout(container)
        layout.addWidget(checkbox)
        layout.setAlignment(Qt.AlignCenter)
        layout.setContentsMargins(0, 0, 0, 0)
        return container

    def _get_filtered_products(self) -> list[Product]:
        brand_filter = clean_multi_spaces(self.ui.cbo_FindBrand.currentText())
        text_filter = clean_multi_spaces(self.ui.line_FindProduct.text())
        with self.get_session() as session:
            q = session.query(Product).filter(Product.name.isnot(None), Product.name != "")
            if brand_filter and brand_filter != "-":
                q = q.filter(Product.brand == brand_filter)
            if text_filter:
                q = q.filter(Product.name.ilike(f"%{text_filter}%"))
            return q.order_by(Product.name.asc()).all()

    def _get_product_name_by_id(self, product_id: int | None) -> str:
        if not product_id:
            return ""
        with self.get_session() as session:
            p = session.query(Product).filter(Product.id == product_id).first()
            return p.name if p else ""

    def _get_supplier_option_name(self, row_id: int, option_id: int | None) -> str:
        if not option_id:
            return ""
        with self.get_session() as session:
            opt = session.query(TempTargetPriceOption).filter(TempTargetPriceOption.id == option_id, TempTargetPriceOption.temp_import_id == row_id).first()
            return opt.supplier_name if opt else ""

    def _get_row_selected_product_id(self, row_id: int) -> int | None:
        with self.get_session() as session:
            row = session.query(TempTargetPriceImport).filter(TempTargetPriceImport.id == row_id).first()
            return row.selected_product_id if row else None

    def _get_row_selected_option_id(self, row_id: int) -> int | None:
        with self.get_session() as session:
            row = session.query(TempTargetPriceImport).filter(TempTargetPriceImport.id == row_id).first()
            return row.selected_option_id if row else None

    def _get_row_brand(self, row_id: int) -> str:
        with self.get_session() as session:
            row = session.query(TempTargetPriceImport).filter(TempTargetPriceImport.id == row_id).first()
            return row.new_brand or "" if row else ""

    def start_cell_edit(self, row: int, column: int):
        if self._updating_table or row < 0 or row >= len(self._table_row_ids):
            return
        row_id = self._table_row_ids[row]
        if self._showing_options and column == COL_SUPPLIER_OPTION:
            combo = self._build_supplier_option_combo(row_id, self._get_row_selected_option_id(row_id))
            combo.activated.connect(lambda _, r=row, rid=row_id, c=combo: self.finish_supplier_option_edit(r, rid, c))
            self.table.setCellWidget(row, column, combo)
            combo.setFocus(); QTimer.singleShot(0, combo.showPopup)
        elif (not self._showing_options) and column == 0:
            combo = self._build_product_combo(row_id, self._get_row_selected_product_id(row_id))
            combo.activated.connect(lambda _, r=row, rid=row_id, c=combo: self.finish_product_edit(r, rid, c))
            self.table.setCellWidget(row, column, combo)
            combo.setFocus(); QTimer.singleShot(0, combo.showPopup)
        elif (not self._showing_options) and column == 4:
            combo = self.build_brand_combo(row_id, self._get_row_brand(row_id))
            combo.activated.connect(lambda _, r=row, rid=row_id, c=combo: self.finish_brand_edit(r, rid, c))
            if combo.lineEdit() is not None:
                combo.lineEdit().returnPressed.connect(lambda r=row, rid=row_id, c=combo: self.finish_brand_edit(r, rid, c))
            self.table.setCellWidget(row, column, combo)
            combo.setFocus(); QTimer.singleShot(0, combo.showPopup)

    def _build_product_combo(self, row_id: int, selected_product_id: int | None) -> QComboBox:
        combo = QComboBox(); combo.setProperty("row_id", row_id); combo.setProperty("combo_role", "product_combo")
        self.populate_product_combo(combo, False, selected_product_id)
        return combo

    def populate_product_combo(self, combo: QComboBox, keep_current: bool, selected_product_id: int | None = None):
        current_id = combo.currentData() if keep_current else selected_product_id
        try: current_id = int(current_id) if current_id is not None else None
        except Exception: current_id = None
        products = self._get_filtered_products()
        combo.blockSignals(True); combo.clear(); combo.addItem("", None)
        for product in products:
            combo.addItem(product.name, int(product.id))
        if current_id is not None and combo.findData(current_id) < 0:
            with self.get_session() as session:
                p = session.query(Product).filter(Product.id == current_id).first()
                if p: combo.addItem(p.name, int(p.id))
        idx = combo.findData(current_id)
        combo.setCurrentIndex(idx if idx >= 0 else 0)
        combo.blockSignals(False)

    def _build_supplier_option_combo(self, row_id: int, selected_option_id: int | None) -> QComboBox:
        combo = QComboBox(); combo.setProperty("row_id", row_id); combo.setProperty("combo_role", "supplier_option_combo")
        with self.get_session() as session:
            options = session.query(TempTargetPriceOption).filter(
                TempTargetPriceOption.batch_id == self.batch_id,
                TempTargetPriceOption.imported_by == self.imported_by,
                TempTargetPriceOption.temp_import_id == row_id,
            ).order_by(TempTargetPriceOption.opt_rank.asc(), TempTargetPriceOption.cost_novo_wvat.asc(), TempTargetPriceOption.id.asc()).all()
        combo.addItem("", None)
        for opt in options:
            combo.addItem(opt.supplier_name or "", int(opt.id))
        idx = combo.findData(selected_option_id)
        combo.setCurrentIndex(idx if idx >= 0 else 0)
        return combo

    def build_brand_combo(self, row_id: int, brand_name: str) -> QComboBox:
        combo = QComboBox(); combo.setEditable(True); combo.setInsertPolicy(QComboBox.NoInsert)
        brands = self.get_brand_names(); combo.addItem(""); combo.addItems(brands)
        if brand_name and combo.findText(brand_name) < 0: combo.addItem(brand_name)
        combo.setCurrentText(brand_name); return combo

    def finish_product_edit(self, row: int, row_id: int, combo: QComboBox):
        product_id = combo.currentData()
        try: product_id = int(product_id)
        except Exception: product_id = None
        self.update_temp_field(row_id, "selected_product_id", product_id)
        self._updating_table = True
        self.table.removeCellWidget(row, 0)
        self.table.setItem(row, 0, self.build_display_item(row_id, "selected_product_id", self._get_product_name_by_id(product_id)))
        self._updating_table = False
        self.table.resizeColumnsToContents()

    def finish_supplier_option_edit(self, row: int, row_id: int, combo: QComboBox):
        option_id = combo.currentData()
        try: option_id = int(option_id)
        except Exception: option_id = None
        self.update_temp_field(row_id, "selected_option_id", option_id)
        self._updating_table = True
        self.table.removeCellWidget(row, COL_SUPPLIER_OPTION)
        self.table.setItem(row, COL_SUPPLIER_OPTION, self.build_display_item(row_id, "selected_option_id", self._get_supplier_option_name(row_id, option_id) or "-"))
        self._updating_table = False
        self.table.resizeColumnsToContents()

    def finish_brand_edit(self, row: int, row_id: int, combo: QComboBox):
        brand = clean_multi_spaces(combo.currentText()) or None
        self.update_temp_field(row_id, "new_brand", brand)
        self._updating_table = True
        self.table.removeCellWidget(row, 4)
        self.table.setItem(row, 4, self.build_display_item(row_id, "new_brand", brand or ""))
        self._updating_table = False
        self.table.resizeColumnsToContents()

    def on_checkbox_changed(self, row_id: int, checked: bool):
        if not self._updating_table:
            self.update_temp_field(row_id, "new_is_excise", bool(checked))

    def on_item_changed(self, item: QTableWidgetItem):
        if self._updating_table:
            return
        row_id = item.data(Qt.UserRole + 1); column_name = item.data(Qt.UserRole)
        if row_id is None or not column_name or column_name in {"selected_option_id", "selected_product_id", "new_brand"}:
            return
        value = clean_multi_spaces(item.text())
        if column_name == "manual_full_cost_msk":
            parsed = parse_loose_number(value) if value else None
            if parsed is None or Decimal(str(parsed)) == 0:
                self._manual_full_costs.pop(int(row_id), None)
            else:
                manual_value = Decimal(str(parsed))
                self._manual_full_costs[int(row_id)] = manual_value
                self._sync_manual_full_cost_to_temp(int(row_id), manual_value)
            return
        header = self.table.horizontalHeaderItem(item.column()).text()
        if header in self.numeric_headers:
            value = parse_loose_number(value)
        elif value == "":
            value = None
        self.update_temp_field(row_id, column_name, value)

    def _sync_manual_full_cost_to_temp(self, row_id: int, value: Decimal):
        """Persist Manual Full Cost immediately so final Supplier is really Manual in temp DB."""
        try:
            with self.get_session() as session:
                service = TargetPriceService(session)
                service.apply_manual_full_costs(self.batch_id, self.imported_by, {int(row_id): value})
                session.commit()
            if self._showing_options and row_id in self._table_row_ids:
                table_row = self._table_row_ids.index(row_id)
                self._updating_table = True
                self.table.setItem(table_row, COL_SUPPLIER_OPTION, self.build_display_item(row_id, "selected_option_id", "Manual"))
                self._updating_table = False
        except Exception as e:
            self.show_error_message(str(e))

    def update_temp_field(self, row_id: int, field_name: str, value):
        if self._updating_table:
            return
        try:
            with self.get_session() as session:
                row = session.query(TempTargetPriceImport).filter(TempTargetPriceImport.id == row_id).first()
                if row is None: return
                setattr(row, field_name, value)
                if field_name == "selected_product_id":
                    if value is not None:
                        row.new_product_name = row.new_brand = row.new_pack = row.new_is_excise = None
                    row.selected_option_id = None
                elif field_name in {"new_product_name", "new_brand", "new_pack", "new_is_excise"}:
                    if value not in (None, ""):
                        row.selected_product_id = None
                        if row.new_is_excise is None: row.new_is_excise = False
                    row.selected_option_id = None
                elif field_name in {"supplier_article", "product_name"}:
                    row.selected_option_id = None
                session.commit()
        except Exception as e:
            self.show_error_message(str(e))

    def refresh_current_product_combo(self):
        row = self.table.currentRow()
        if row < 0: return
        col = COL_PRODUCT if self._showing_options else 0
        combo = self.table.cellWidget(row, col)
        if isinstance(combo, QComboBox) and combo.property("combo_role") == "product_combo":
            self.populate_product_combo(combo, True)

    def _commit_open_editors(self):
        for row in range(self.table.rowCount()):
            for column in range(self.table.columnCount()):
                widget = self.table.cellWidget(row, column)
                if not isinstance(widget, QComboBox) or row >= len(self._table_row_ids):
                    continue
                row_id = self._table_row_ids[row]
                role = widget.property("combo_role")
                if role == "supplier_option_combo": self.finish_supplier_option_edit(row, row_id, widget)
                elif role == "product_combo": self.finish_product_edit(row, row_id, widget)
                elif role == "brand_combo": self.finish_brand_edit(row, row_id, widget)

    def add_manual_price_column(self):
        self._commit_open_editors()
        self._manual_full_cost_visible = True
        self._showing_options = True
        self.load_table()
        self.show_message("Добавлена колонка Manual Full Cost Msk")

    def _get_manual_full_cost(self, row_id: int):
        with self.get_session() as session:
            opt = session.query(TempTargetPriceOption).filter(
                TempTargetPriceOption.batch_id == self.batch_id,
                TempTargetPriceOption.imported_by == self.imported_by,
                TempTargetPriceOption.temp_import_id == row_id,
                TempTargetPriceOption.supplier_name == "Manual",
            ).first()
            return opt.full_cost_msk if opt else None

    def collect_manual_full_costs_from_table(self) -> dict[int, Decimal]:
        result = dict(self._manual_full_costs)
        for row in range(self.table.rowCount()):
            if row >= len(self._table_row_ids):
                continue
            row_id = int(self._table_row_ids[row])
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                if item is None or item.data(Qt.UserRole) != "manual_full_cost_msk":
                    continue
                text = clean_multi_spaces(item.text())
                parsed = parse_loose_number(text) if text else None
                if parsed is None or Decimal(str(parsed)) == 0:
                    result.pop(row_id, None)
                else:
                    result[row_id] = Decimal(str(parsed))
        self._manual_full_costs = result
        return result

    def add_line(self):
        supplier_id = self.ensure_supplier()
        try:
            with self.get_session() as session:
                service = TargetPriceService(session)
                last = session.query(TempTargetPriceImport).filter(TempTargetPriceImport.batch_id == self.batch_id, TempTargetPriceImport.imported_by == self.imported_by).order_by(TempTargetPriceImport.import_row_no.desc(), TempTargetPriceImport.id.desc()).first()
                next_row = int(last.import_row_no or 0) + 1 if last else 1
                service.import_rows(rows=[{"import_row_no": next_row}], batch_id=self.batch_id, imported_by=self.imported_by, supplier_id=supplier_id, replace_existing=False)
                session.commit()
            self._showing_options = False
            self.load_table(); self.show_message("Добавлена строка")
        except Exception as e:
            self.show_error_message(str(e))

    def calculate_costs(self):
        self._commit_open_editors()
        try:
            supplier_id = self.ensure_supplier()
            supplier_name = clean_multi_spaces(self.ui.cbo_SupplName.currentText()) or clean_multi_spaces(self.ui.line_NewSupplier.text()) or "NoName"
            safe_supplier_name = "".join(ch if ch not in r'<>:"/\|?*' else "_" for ch in supplier_name)
            default = f"TargetPriceCalc_{safe_supplier_name}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
            manual_full_costs = self.collect_manual_full_costs_from_table()
            supplier_price_age_months = self.get_supplier_price_age_months()
            save_path, _ = QFileDialog.getSaveFileName(self, "Сохранить расчет", str(BASE_DIR / default), "Excel files (*.xlsx)")
            if not save_path: return
            batch_id = self.batch_id
            imported_by = self.imported_by

            def do_export():
                with self.get_session() as session:
                    from app.utils.gui_table_actions import apply_pending_table_deletes_to_db
                    apply_pending_table_deletes_to_db(session, self)
                    service = TargetPriceService(session)
                    exporter = TargetPriceExporter(session)
                    service.run_calculation(
                        batch_id,
                        imported_by,
                        manual_full_costs=manual_full_costs,
                        supplier_price_age_months=supplier_price_age_months,
                    )
                    output = exporter.export_calculated(batch_id, imported_by, save_path)
                    session.commit()
                    return output

            def done(output_path):
                self._showing_options = True
                self.load_table()
                QDesktopServices.openUrl(QUrl.fromLocalFile(str(output_path)))
                self.show_message("Расчет выполнен")

            if not start_excel_export(self, do_export, on_finished=done, on_error=lambda text: self.show_error_message(str(text))):
                self.show_message("Excel файл уже формируется. Можно продолжать работать в программе.")
            else:
                self.show_message("Расчет и Excel файл формируются в фоновом режиме. Можно продолжать работать в программе.")
        except Exception as e:
            self.show_error_message(str(e))

    def save_all(self):
        self._commit_open_editors()
        try:
            if not self._showing_options:
                self.show_error_message("Сначала запустите расчет")
                return
            supplier_id = self.ensure_supplier()
            supplier_name = clean_multi_spaces(self.ui.cbo_SupplName.currentText()) or clean_multi_spaces(self.ui.line_NewSupplier.text()) or "NoName"
            safe_supplier_name = "".join(ch if ch not in r'<>:"/\|?*' else "_" for ch in supplier_name)
            default = f"TargetPrice_{safe_supplier_name}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
            manual_full_costs = self.collect_manual_full_costs_from_table()
            save_path, _ = QFileDialog.getSaveFileName(self, "Сохранить target price", str(BASE_DIR / default), "Excel files (*.xlsx)")
            if not save_path: return
            currency = clean_multi_spaces(self.ui.cbo_Currency.currentText()).upper()
            if currency == "-": currency = ""
            batch_id = self.batch_id
            imported_by = self.imported_by
            fx_rate = self.parse_decimal_field(self.ui.line_ExchangeRate, "Курс")
            transport = self.parse_decimal_field(self.ui.line_Transport, "Транспорт")
            reexport = self.parse_percent_field(self.ui.line_Reexport, "Реэкспорт")
            fx_markup = self.parse_percent_field(self.ui.line_FXMarkup, "FX markup")
            has_customs = self.ui.cbo_Customs.currentText() == "да"
            via_novo = self.ui.cbo_viaNovo.currentText() == "через Ново"

            def do_export():
                with self.get_session() as session:
                    from app.utils.gui_table_actions import apply_pending_table_deletes_to_db
                    apply_pending_table_deletes_to_db(session, self)
                    service = TargetPriceService(session)
                    exporter = TargetPriceExporter(session)
                    service.save_target_calculations(
                        batch_id=batch_id,
                        imported_by=imported_by,
                        target_supplier_id=supplier_id,
                        currency_code=currency,
                        fx_rate=fx_rate,
                        transport=transport,
                        reexport=reexport,
                        fx_markup=fx_markup,
                        has_customs=has_customs,
                        via_novo=via_novo,
                        manual_full_costs=manual_full_costs,
                    )
                    output = exporter.export_final(batch_id, imported_by, save_path)
                    session.commit()
                    return output

            def done(output_path):
                QDesktopServices.openUrl(QUrl.fromLocalFile(str(output_path)))
                self.show_message("Target price сохранен")

            if not start_excel_export(self, do_export, on_finished=done, on_error=lambda text: self.show_error_message(str(text))):
                self.show_message("Excel файл уже формируется. Можно продолжать работать в программе.")
            else:
                self.show_message("Target price сохраняется в фоновом режиме. Можно продолжать работать в программе.")
        except Exception as e:
            self.show_error_message(str(e))

    def reset_all(self):
        try:
            with self.get_session() as session:
                TargetPriceService(session).delete_temp_rows(self.batch_id, self.imported_by)
                session.commit()
            self.start_new_batch(); self.apply_default_values(); self.show_message("Форма очищена")
        except Exception as e:
            self.show_error_message(str(e))

    def download_template(self):
        save_path, _ = QFileDialog.getSaveFileName(self, "Сохранить шаблон", str(BASE_DIR / "Target price_template.xlsx"), "Excel files (*.xlsx)")
        if not save_path: return
        try:
            with self.get_session() as session:
                output = TargetPriceExporter(session).export_template(save_path)
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(output)))
            self.show_message("Шаблон сформирован")
        except Exception as e:
            self.show_error_message(str(e))

