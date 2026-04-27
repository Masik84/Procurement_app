from __future__ import annotations

import os
from datetime import datetime, date
from decimal import Decimal
from pathlib import Path

from PySide6.QtCore import Qt, QFile, QEvent, QPoint, QDate, QTimer
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
    QMessageBox,
    QTableWidget,
    QTableWidgetItem,
    QToolTip,
    QVBoxLayout,
    QWidget,
    QPushButton,
)
from PySide6.QtUiTools import QUiLoader
from sqlalchemy.orm import joinedload

from app.db.db import SessionLocal
from app.db.models import ExchangeRate, Product, Supplier, TempPriceImport
from app.imports.supplier_price_importer import SupplierPriceImporter
from app.exports.supplier_price_exporter import SupplierPriceExporter
from app.services.supplier_service import SupplierService, SupplierUpsertData
from app.services.supplier_price_service import SupplierPriceService
from app.utils.batch import get_current_username
from app.utils.parsers import parse_flexible_date, parse_loose_number
from app.utils.text import clean_multi_spaces
from app.ui.table_style import *


BASE_DIR = Path(__file__).resolve().parents[2]
SUPPLIER_PRICES_UI = BASE_DIR / "app" / "ui" / "windows" / "supplier_prices.ui"


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


class SupplierPricesPage(QWidget):
    def __init__(self):
        super().__init__()

        self.ui = load_ui(SUPPLIER_PRICES_UI)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.ui)

        self.imported_by = get_current_username()
        self.batch_id = ""
        self.selected_file_path = ""
        self.rf_prices_include_vat = False

        self._updating_table = False
        self._pending_changes: dict[int, dict] = {}
        self._pending_deletes: set[int] = set()
        self._new_rows: set[int] = set()
        self._table_row_ids: list[int] = []

        self.columns = [
            "selected_product_id",
            "supplier_article",
            "product_name",
            "price",
            "price_pack",
            "qty_pcs",
            "volume_l",
            "new_product_name",
            "new_brand",
            "new_pack",
            "new_is_excise",
        ]
        self.headers = [
            "Product name",
            "Article",
            "Supplier Product name",
            "Price, Lt",
            "Price, pack",
            "Qty, pcs",
            "Volume, L",
            "Product name (for new)",
            "Brand (for new)",
            "Pack (for new)",
            "Excise duty (for new)",
        ]
        self.numeric_columns = {"price", "price_pack", "qty_pcs", "volume_l", "new_pack"}

        self.setup_ui()
        self.setup_connections()
        self.load_initial_state()

    def setup_ui(self):
        self.table = self.ui.table
        setup_data_table(self.table, sorting=False)
        self.table.horizontalHeader().setSectionsMovable(False)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)

        self.ui.label_msg.setText("Сообщений нет")
        self.ui.line_NewSupplier.setEnabled(False)
        self.ui.line_NewSupplier.setStyleSheet("background-color: #f2f2f2;")

        self.ui.date_Price.setCalendarPopup(True)
        self.ui.date_Price.setDisplayFormat("dd.MM.yyyy")
        self.ui.date_Price.setSpecialValueText("")

        self._setup_number_field(self.ui.line_ExchangeRate, "Формат: 82,0000")
        self._setup_number_field(self.ui.line_Transport, "Формат: 1,2500")
        if hasattr(self.ui, "line_AgentFee"):
            self._setup_number_field(self.ui.line_AgentFee, "Формат: 0,2500")
        self._setup_number_field(self.ui.line_Reexport, "Формат: 3,5% / 0,24%")
        self._setup_number_field(self.ui.line_FXMarkup, "Формат: 3,5% / 0,24%")

    def setup_connections(self):
        self.table.cellDoubleClicked.connect(self.start_cell_edit)
        self.table.itemChanged.connect(self.on_item_changed)
        self.table.customContextMenuRequested.connect(self.show_context_menu)

        self.ui.cbo_SupplName.currentIndexChanged.connect(self.on_supplier_changed)
        self.ui.cbx_NewSupplier.toggled.connect(self.on_new_supplier_toggled)
        self.ui.cbo_Currency.currentIndexChanged.connect(self.on_currency_changed)

        self.ui.btn_DownFile.clicked.connect(self.download_template)
        self.ui.btn_Import.clicked.connect(self.import_file)
        self.ui.btn_AddLine.clicked.connect(self.add_line)
        self.ui.btn_Save.clicked.connect(self.apply_pending_changes)
        self.ui.btn_Reset.clicked.connect(self.reset_form)

        self.ui.cbo_FindBrand.currentTextChanged.connect(self.refresh_current_product_combo)
        self.ui.line_FindProduct.textChanged.connect(self.refresh_current_product_combo)

        self.ui.line_ExchangeRate.editingFinished.connect(self.normalize_exchange_rate)
        self.ui.line_Transport.editingFinished.connect(self.normalize_transport)
        if hasattr(self.ui, "line_AgentFee"):
            self.ui.line_AgentFee.editingFinished.connect(self.normalize_agent_fee)
        self.ui.line_Reexport.editingFinished.connect(self.normalize_reexport)
        self.ui.line_FXMarkup.editingFinished.connect(self.normalize_fx_markup)

    def get_session(self):
        return SessionLocal()

    def load_initial_state(self):
        self.start_new_batch()
        self.load_suppliers()
        self.load_currencies()
        self.load_find_brands()
        self.apply_default_values()
        self.cleanup_old_temp_rows()
        self.load_table_rows()

    def _setup_number_field(self, widget: QLineEdit, tooltip_text: str):
        widget.setToolTip(tooltip_text)
        widget.installEventFilter(self)

    def eventFilter(self, watched, event):
        if watched in {
            self.ui.line_ExchangeRate,
            self.ui.line_Transport,
            *([self.ui.line_AgentFee] if hasattr(self.ui, "line_AgentFee") else []),
            self.ui.line_Reexport,
            self.ui.line_FXMarkup,
        }:
            if event.type() in {QEvent.Enter, QEvent.FocusIn, QEvent.MouseButtonPress}:
                QToolTip.showText(
                    watched.mapToGlobal(QPoint(0, watched.height())),
                    watched.toolTip(),
                    watched,
                )

        if isinstance(watched, QComboBox):
            row_id = watched.property("row_id")
            role = watched.property("combo_role")
            if row_id and role and event.type() in {QEvent.FocusIn, QEvent.MouseButtonPress}:
                if role == "brand_combo":
                    self.populate_brand_combo(watched, keep_current=True)

        return super().eventFilter(watched, event)

    def start_new_batch(self):
        with self.get_session() as session:
            service = SupplierPriceService(session)
            self.batch_id = service.start_batch()

    def cleanup_old_temp_rows(self):
        with self.get_session() as session:
            service = SupplierPriceService(session)
            service.cleanup_old_temp_rows(imported_by=self.imported_by)
            session.commit()

    def apply_default_values(self):
        self.set_combo_text(self.ui.cbo_SupplName, "-")
        self.ui.cbx_NewSupplier.setChecked(False)
        self.ui.line_NewSupplier.clear()
        self.set_combo_text(self.ui.cbo_SupplierRF, "нет")
        self.set_combo_text(self.ui.cbo_Currency, "-")
        self.ui.line_ExchangeRate.clear()
        self.ui.line_Transport.clear()
        self.set_combo_text(self.ui.cbo_viaNovo, "через Ново")
        self.ui.line_Reexport.setText("0,0%")
        self.ui.line_FXMarkup.setText("0,0%")
        self.set_combo_text(self.ui.cbo_Customs, "да")
        self.set_combo_text(self.ui.cbo_Marking, "Феникс")
        self.set_combo_text(self.ui.cbo_History, "да")
        self.ui.date_Price.setDate(QDate.currentDate())
        self.set_combo_text(self.ui.cbo_Rating, "да")
        self.ui.line_FindProduct.clear()
        self.set_combo_text(self.ui.cbo_FindBrand, "-")
        self.selected_file_path = ""
        self.toggle_new_supplier_field(False)

    def set_combo_text(self, combo: QComboBox, value: str):
        index = combo.findText(value)
        if index >= 0:
            combo.setCurrentIndex(index)
        elif combo.isEditable():
            combo.setCurrentText(value)

    def ask_rf_prices_include_vat(self) -> bool:
        return (
            QMessageBox.question(
                self,
                "Поставщик РФ",
                "Цены поставщика с НДС?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes,
            )
            == QMessageBox.Yes
        )

    def ask_export_calculated_excel(self) -> bool:
        return (
            QMessageBox.question(
                self,
                "Excel",
                "Сохранить рассчитанные данные в Excel?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes,
            )
            == QMessageBox.Yes
        )

    def export_calculated_excel(self, supplier_id: int):
        supplier_name = (
            clean_multi_spaces(self.ui.line_NewSupplier.text())
            or clean_multi_spaces(self.ui.cbo_SupplName.currentText())
            or "Supplier"
        )
        safe_supplier_name = supplier_name
        for ch in ['\\', '/', ':', '*', '?', '"', '<', '>', '|']:
            safe_supplier_name = safe_supplier_name.replace(ch, '_')

        default_name = f"CostCalc_{safe_supplier_name}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Сохранить Excel файл",
            str(Path.home() / default_name),
            "Excel files (*.xlsx)",
        )
        if not file_path:
            return
        if not file_path.lower().endswith('.xlsx'):
            file_path += '.xlsx'

        with self.get_session() as session:
            exporter = SupplierPriceExporter(session)
            output_path = exporter.export_calculated(
                batch_id=self.batch_id,
                imported_by=self.imported_by,
                supplier_id=supplier_id,
                output_path=file_path,
                source_file_path=self.selected_file_path or None,
            )

        QDesktopServices.openUrl(Path(output_path).as_uri())
        self.show_message("Excel файл сохранен")

    def load_suppliers(self):
        current_text = self.ui.cbo_SupplName.currentText().strip()

        with self.get_session() as session:
            supplier_service = SupplierService(session)
            suppliers = supplier_service.get_all_suppliers()

        self.ui.cbo_SupplName.blockSignals(True)
        self.ui.cbo_SupplName.clear()
        self.ui.cbo_SupplName.addItem("-", None)
        for supplier in suppliers:
            self.ui.cbo_SupplName.addItem(supplier.name, supplier.id)
        self.ui.cbo_SupplName.blockSignals(False)

        if current_text:
            self.set_combo_text(self.ui.cbo_SupplName, current_text if current_text != "" else "-")

    def load_currencies(self):
        current_text = self.ui.cbo_Currency.currentText().strip()

        with self.get_session() as session:
            rows = (
                session.query(ExchangeRate.currency_code)
                .order_by(ExchangeRate.currency_code.asc())
                .all()
            )

        self.ui.cbo_Currency.blockSignals(True)
        self.ui.cbo_Currency.clear()
        self.ui.cbo_Currency.addItem("-")
        for row in rows:
            if row[0]:
                self.ui.cbo_Currency.addItem(row[0])
        self.ui.cbo_Currency.blockSignals(False)

        if current_text:
            self.set_combo_text(self.ui.cbo_Currency, current_text if current_text != "" else "-")

    def load_find_brands(self):
        current_text = self.ui.cbo_FindBrand.currentText().strip()
        brands = self.get_brand_names()

        self.ui.cbo_FindBrand.blockSignals(True)
        self.ui.cbo_FindBrand.clear()
        self.ui.cbo_FindBrand.addItem("-")
        if brands:
            self.ui.cbo_FindBrand.addItems(brands)
        self.ui.cbo_FindBrand.blockSignals(False)

        if current_text:
            self.set_combo_text(self.ui.cbo_FindBrand, current_text if current_text != "" else "-")

    def get_brand_names(self) -> list[str]:
        with self.get_session() as session:
            rows = (
                session.query(Product.brand)
                .filter(Product.brand.isnot(None), Product.brand != "")
                .distinct()
                .order_by(Product.brand.asc())
                .all()
            )
        return [row[0] for row in rows if row[0]]

    def on_new_supplier_toggled(self, checked: bool):
        if checked:
            self.set_combo_text(self.ui.cbo_SupplName, "-")
            self.ui.line_NewSupplier.clear()
            self.ui.line_ExchangeRate.clear()
            self.ui.line_Transport.clear()
            if hasattr(self.ui, "line_AgentFee"):
                self.ui.line_AgentFee.clear()
            self.ui.line_Reexport.setText("0,0%")
            self.ui.line_FXMarkup.setText("0,0%")
            self.set_combo_text(self.ui.cbo_Currency, "-")
            self.set_combo_text(self.ui.cbo_viaNovo, "через Ново")
            self.set_combo_text(self.ui.cbo_Customs, "да")
            self.set_combo_text(self.ui.cbo_Marking, "Феникс")
            self.set_combo_text(self.ui.cbo_SupplierRF, "нет")
            self.set_combo_text(self.ui.cbo_Rating, "да")

        self.toggle_new_supplier_field(checked)

    def toggle_new_supplier_field(self, enabled: bool):
        self.ui.line_NewSupplier.setEnabled(enabled)
        if enabled:
            self.ui.line_NewSupplier.setStyleSheet("")
            self.ui.line_NewSupplier.setFocus()
        else:
            self.ui.line_NewSupplier.setStyleSheet("background-color: #f2f2f2;")

    def on_supplier_changed(self):
        supplier_id = self.ui.cbo_SupplName.currentData()
        if supplier_id is None:
            return

        with self.get_session() as session:
            supplier_service = SupplierService(session)
            supplier_data = supplier_service.load_supplier_snapshot(int(supplier_id))
            rate = supplier_service.get_rate_to_rub(supplier_data.base_currency)

        self.ui.cbx_NewSupplier.blockSignals(True)
        self.ui.cbx_NewSupplier.setChecked(False)
        self.ui.cbx_NewSupplier.blockSignals(False)
        self.toggle_new_supplier_field(False)

        self.ui.line_NewSupplier.setText(supplier_data.name)
        self.set_combo_text(self.ui.cbo_SupplierRF, "да" if supplier_data.is_rf else "нет")
        self.set_combo_text(self.ui.cbo_Currency, supplier_data.base_currency or "-")
        self.ui.line_ExchangeRate.setText(self.format_number(rate, 4) if rate is not None else "")
        self.ui.line_Transport.setText(self.format_number(supplier_data.transport_cost_per_l, 4))
        if hasattr(self.ui, "line_AgentFee"):
            self.ui.line_AgentFee.setText(self.format_number(supplier_data.agent_fee, 4))
        self.set_combo_text(self.ui.cbo_viaNovo, "через Ново" if supplier_data.is_via_novo else "в Мск")
        self.ui.line_Reexport.setText(self.format_percent(supplier_data.reexport_percent))
        self.ui.line_FXMarkup.setText(self.format_percent(supplier_data.fx_rate_markup))
        self.set_combo_text(self.ui.cbo_Customs, "да" if supplier_data.has_import_duty else "нет")
        self.set_combo_text(self.ui.cbo_Marking, "Поставщик" if supplier_data.marks_for_us else "Феникс")
        self.set_combo_text(self.ui.cbo_Rating, "да" if supplier_data.rating_calc else "нет")

    def on_currency_changed(self):
        currency_code = clean_multi_spaces(self.ui.cbo_Currency.currentText()).upper()
        if not currency_code or currency_code == "-":
            return

        with self.get_session() as session:
            supplier_service = SupplierService(session)
            rate = supplier_service.get_rate_to_rub(currency_code)

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
            agent_fee=self.parse_decimal_field(self.ui.line_AgentFee, "Agent fee") if hasattr(self.ui, "line_AgentFee") else Decimal("0"),
            reexport_percent=self.parse_percent_field(self.ui.line_Reexport, "Реэкспорт"),
            fx_rate_markup=self.parse_percent_field(self.ui.line_FXMarkup, "FX markup"),
            is_via_novo=self.ui.cbo_viaNovo.currentText() == "через Ново",
            has_import_duty=self.ui.cbo_Customs.currentText() == "да",
            rating_calc=self.ui.cbo_Rating.currentText() == "да",
            marks_for_us=self.ui.cbo_Marking.currentText() == "Поставщик",
            is_rf=self.ui.cbo_SupplierRF.currentText() == "да",
        )

    def ensure_supplier(self) -> int:
        supplier_id = self.ui.cbo_SupplName.currentData()
        if self.ui.cbx_NewSupplier.isChecked():
            supplier_id = None

        with self.get_session() as session:
            supplier_service = SupplierService(session)
            supplier = supplier_service.ensure_supplier(
                supplier_id=int(supplier_id) if supplier_id is not None else None,
                data=self.get_supplier_form_data(),
            )
            fx_rate = self.parse_decimal_field(self.ui.line_ExchangeRate, "Курс")
            if fx_rate is not None and supplier.base_currency:
                supplier_service.save_exchange_rate(supplier.base_currency, float(fx_rate))
            session.commit()
            supplier_id = supplier.id

        self.load_suppliers()
        index = self.ui.cbo_SupplName.findData(supplier_id)
        if index >= 0:
            self.ui.cbo_SupplName.setCurrentIndex(index)
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

        cleaned = text.replace("%", "")
        value = parse_loose_number(cleaned)
        if value is None:
            raise ValueError(f"Некорректное поле: {field_name}")

        decimal_value = Decimal(str(value))
        if abs(decimal_value) > Decimal("1"):
            decimal_value = decimal_value / Decimal("100")
        return decimal_value

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

    def normalize_exchange_rate(self):
        self._normalize_number_widget(self.ui.line_ExchangeRate, digits=4)

    def normalize_transport(self):
        self._normalize_number_widget(self.ui.line_Transport, digits=4)

    def normalize_agent_fee(self):
        if hasattr(self.ui, "line_AgentFee"):
            self._normalize_number_widget(self.ui.line_AgentFee, digits=4)

    def normalize_reexport(self):
        self._normalize_percent_widget(self.ui.line_Reexport)

    def normalize_fx_markup(self):
        self._normalize_percent_widget(self.ui.line_FXMarkup)

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
        cleaned = text.replace("%", "")
        value = parse_loose_number(cleaned)
        if value is None:
            self.show_error_message("Проверь процент")
            return
        decimal_value = Decimal(str(value))
        if abs(decimal_value) > Decimal("1"):
            decimal_value = decimal_value / Decimal("100")
        widget.setText(self.format_percent(decimal_value))

    def get_price_date(self) -> datetime:
        qdate = self.ui.date_Price.date()
        if qdate.isValid():
            return datetime(qdate.year(), qdate.month(), qdate.day())

        parsed = parse_flexible_date(self.ui.date_Price.text())
        if parsed is None:
            parsed = date.today()
        return datetime(parsed.year, parsed.month, parsed.day)

    def download_template(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Сохранить шаблон",
            str(Path.home() / "ImportTemplate.xlsx"),
            "Excel files (*.xlsx)",
        )
        if not file_path:
            return

        if not file_path.lower().endswith(".xlsx"):
            file_path += ".xlsx"

        try:
            with self.get_session() as session:
                exporter = SupplierPriceExporter(session)
                output_path = exporter.export_template(file_path)
            QDesktopServices.openUrl(Path(output_path).as_uri())
            self.show_message("Шаблон сформирован")
        except Exception as e:
            self.show_error_message(str(e))

    def import_file(self):
        try:
            supplier_id = self.ensure_supplier()
        except Exception as e:
            self.show_error_message(str(e))
            return

        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите файл прайс-листа",
            "",
            "Excel files (*.xls *.xlsx)",
        )
        if not file_path:
            return

        try:
            importer = SupplierPriceImporter()
            rows = importer.read_excel(file_path)

            with self.get_session() as session:
                service = SupplierPriceService(session)
                service.import_rows_to_temp(
                    supplier_id=supplier_id,
                    batch_id=self.batch_id,
                    imported_by=self.imported_by,
                    rows=rows,
                    import_date=self.get_price_date(),
                    replace_existing_batch_rows=True,
                )
                service.automatch_temp_rows(self.batch_id, self.imported_by)
                session.commit()

            self.selected_file_path = file_path
            self.load_table_rows()
            self.show_message("Данные импортированы")
        except Exception as e:
            self.show_error_message(str(e))

    def load_table_rows(self):
        with self.get_session() as session:
            rows = (
                session.query(TempPriceImport)
                .options(joinedload(TempPriceImport.selected_product))
                .filter(
                    TempPriceImport.batch_id == self.batch_id,
                    TempPriceImport.imported_by == self.imported_by,
                )
                .order_by(TempPriceImport.import_row_no.asc(), TempPriceImport.id.asc())
                .all()
            )

        self.display_rows(rows)

    def display_rows(self, rows: list[TempPriceImport]):
        self._updating_table = True
        self._table_row_ids = [row.id for row in rows]
        self._pending_deletes.clear()
        self._new_rows.clear()

        self.table.clear()
        self.table.setColumnCount(len(self.headers))
        self.table.setRowCount(len(rows))
        self.table.setHorizontalHeaderLabels(self.headers)

        for row_index, row in enumerate(rows):
            self._pending_changes.setdefault(row.id, {})

            product_text = row.selected_product.name if row.selected_product else ""
            product_item = self.build_display_item(row.id, "selected_product_id", product_text)
            self.table.setItem(row_index, 0, product_item)

            article_text = self._clean_table_text(row.supplier_article)
            supplier_product_name_text = self._clean_table_text(row.product_name)
            new_product_name_text = self._clean_table_text(row.new_product_name)
            brand_text = self._clean_table_text(row.new_brand)

            self.table.setItem(row_index, 1, self.build_table_item("supplier_article", article_text))
            self.table.setItem(row_index, 2, self.build_table_item("product_name", supplier_product_name_text))
            self.table.setItem(row_index, 3, self.build_table_item("price", self.value_to_text(row.price)))
            self.table.setItem(row_index, 4, self.build_table_item("price_pack", self.value_to_text(row.price_pack)))
            self.table.setItem(row_index, 5, self.build_table_item("qty_pcs", self.value_to_text(row.qty_pcs)))
            self.table.setItem(row_index, 6, self.build_table_item("volume_l", self.value_to_text(row.volume_l)))
            self.table.setItem(row_index, 7, self.build_table_item("new_product_name", new_product_name_text))

            brand_item = self.build_display_item(row.id, "new_brand", brand_text)
            self.table.setItem(row_index, 8, brand_item)

            self.table.setItem(row_index, 9, self.build_table_item("new_pack", self.value_to_text(row.new_pack)))
            self.table.setCellWidget(
                row_index,
                10,
                self.build_checkbox_widget(row.id, bool(row.new_is_excise)),
            )

        self.table.resizeColumnsToContents()
        self._updating_table = False

    def start_cell_edit(self, row: int, column: int):
        if self._updating_table:
            return

        if column not in (0, 8):
            return

        if row < 0 or row >= len(self._table_row_ids):
            return

        row_id = self._table_row_ids[row]

        if column == 0:
            current_product_id = self._get_row_selected_product_id(row_id)
            combo = self._build_product_combo(row_id, current_product_id)
            combo.activated.connect(
                lambda _, r=row, rid=row_id, c=combo: self.finish_product_edit(r, rid, c)
            )
            self.table.setCellWidget(row, column, combo)
            combo.setFocus()
            QTimer.singleShot(0, combo.showPopup)

        elif column == 8:
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
            combo.lineEdit().selectAll()

    def _get_row_selected_product_id(self, row_id: int):
        with self.get_session() as session:
            row = (
                session.query(TempPriceImport)
                .options(joinedload(TempPriceImport.selected_product))
                .filter(TempPriceImport.id == row_id)
                .first()
            )
            return row.selected_product_id if row else None

    def _get_row_selected_product_name(self, row_id: int) -> str:
        with self.get_session() as session:
            row = (
                session.query(TempPriceImport)
                .options(joinedload(TempPriceImport.selected_product))
                .filter(TempPriceImport.id == row_id)
                .first()
            )
            if row and row.selected_product:
                return row.selected_product.name or ""
            return ""

    def _get_row_brand(self, row_id: int) -> str:
        with self.get_session() as session:
            row = session.query(TempPriceImport).filter(TempPriceImport.id == row_id).first()
            return row.new_brand or "" if row else ""

    def finish_product_edit(self, row: int, row_id: int, combo: QComboBox):
        product_id = combo.currentData()

        # Пустой пункт в комбобоксе = снять привязку продукта
        if product_id in (None, "", 0):
            self._pending_changes.setdefault(row_id, {})
            self._pending_changes[row_id]["selected_product_id"] = None

            self._updating_table = True
            self.table.removeCellWidget(row, 0)
            self.table.setItem(
                row,
                0,
                self.build_display_item(row_id, "selected_product_id", ""),
            )
            self._updating_table = False
            self.table.resizeColumnsToContents()
            return

        try:
            product_id = int(product_id)
        except (TypeError, ValueError):
            self._updating_table = True
            self.table.removeCellWidget(row, 0)
            self._updating_table = False
            return

        self._pending_changes.setdefault(row_id, {})
        self._pending_changes[row_id]["selected_product_id"] = product_id
        self._pending_changes[row_id]["new_product_name"] = None
        self._pending_changes[row_id]["new_brand"] = None
        self._pending_changes[row_id]["new_pack"] = None
        self._pending_changes[row_id]["new_is_excise"] = False

        with self.get_session() as session:
            product = session.query(Product).filter(Product.id == product_id).first()
            product_name = product.name if product else ""

        self._updating_table = True
        self.table.removeCellWidget(row, 0)
        self.table.setItem(
            row,
            0,
            self.build_display_item(row_id, "selected_product_id", product_name),
        )
        self._updating_table = False
        self.table.resizeColumnsToContents()

    def finish_brand_edit(self, row: int, row_id: int, combo: QComboBox):
        text = clean_multi_spaces(combo.currentText()).upper() or None

        self._pending_changes.setdefault(row_id, {})
        self._pending_changes[row_id]["new_brand"] = text

        self._updating_table = True
        self.table.removeCellWidget(row, 8)
        self.table.setItem(
            row,
            8,
            self.build_display_item(row_id, "new_brand", text or ""),
        )
        self._updating_table = False

        self.table.resizeColumnsToContents()

    def value_to_text(self, value: object) -> str:
        if value is None:
            return ""

        text = str(value).strip()
        if text.lower() == "nan":
            return ""

        number = parse_loose_number(value)
        if number is None:
            return text

        formatted = f"{float(number):.4f}".replace(".", ",")
        formatted = formatted.rstrip("0").rstrip(",")

        return formatted

    def build_table_item(self, column_name: str, value: str) -> QTableWidgetItem:
        item = QTableWidgetItem(value)
        item.setData(Qt.UserRole, column_name)
        item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable)
        if column_name in self.numeric_columns:
            item.setTextAlignment(Qt.AlignCenter)
        else:
            item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        return item

    def build_display_item(self, row_id: int, column_name: str, value: str) -> QTableWidgetItem:
        item = QTableWidgetItem(value)
        item.setData(Qt.UserRole, column_name)
        item.setData(Qt.UserRole + 1, row_id)
        item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
        item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        return item

    def _build_product_combo(self, row_id: int, selected_product_id: int | None) -> QComboBox:
        combo = QComboBox()
        combo.setProperty("row_id", row_id)
        combo.setProperty("combo_role", "product_combo")
        combo.setToolTip("Выберите продукт из базы")
        self.populate_product_combo(
            combo,
            row_id=row_id,
            keep_current=False,
            selected_product_id=selected_product_id,
        )
        return combo

    def populate_product_combo(
        self,
        combo: QComboBox,
        row_id: int,
        keep_current: bool,
        selected_product_id: int | None = None,
    ):
        current_id = combo.currentData() if keep_current else selected_product_id

        if current_id is not None:
            try:
                current_id = int(current_id)
            except (TypeError, ValueError):
                current_id = None

        products = self.get_filtered_products()

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
        brands = self.get_brand_names()

        combo.blockSignals(True)
        combo.clear()
        combo.addItem("")
        if brands:
            combo.addItems(brands)
        if brand_value and combo.findText(brand_value) < 0:
            combo.addItem(brand_value)
        combo.setCurrentText(brand_value)
        combo.blockSignals(False)

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

    def on_product_combo_changed(self, row_id: int, combo: QComboBox):
        if self._updating_table:
            return

        value = combo.currentData()
        try:
            value = int(value)
        except (TypeError, ValueError):
            return

        self._pending_changes.setdefault(row_id, {})
        self._pending_changes[row_id]["selected_product_id"] = value

    def on_brand_combo_changed(self, row_id: int, combo: QComboBox):
        if self._updating_table:
            return
        self._pending_changes.setdefault(row_id, {})
        self._pending_changes[row_id]["new_brand"] = clean_multi_spaces(combo.currentText()) or None

    def on_checkbox_changed(self, row_id: int, checked: bool):
        if self._updating_table:
            return
        self._pending_changes.setdefault(row_id, {})
        self._pending_changes[row_id]["new_is_excise"] = bool(checked)

    def on_item_changed(self, item: QTableWidgetItem):
        if self._updating_table:
            return

        row = item.row()
        if row < 0 or row >= len(self._table_row_ids):
            return

        row_id = self._table_row_ids[row]
        column_name = item.data(Qt.UserRole)
        if not column_name:
            return

        if column_name == "selected_product_id":
            return

        value = clean_multi_spaces(item.text()).upper()
        self._pending_changes.setdefault(row_id, {})

        if column_name in self.numeric_columns:
            self._pending_changes[row_id][column_name] = value or None
        else:
            self._pending_changes[row_id][column_name] = value or None

        if column_name in {"supplier_article", "product_name"}:
            self._pending_changes[row_id]["selected_product_id"] = None

    def show_context_menu(self, position):
        item = self.table.itemAt(position)
        if item is not None:
            self.table.setCurrentCell(item.row(), item.column())

        menu = QMenu()
        copy_action = menu.addAction("Копировать")
        delete_action = menu.addAction("Удалить строку")
        apply_action = menu.addAction("Применить изменения")
        revert_action = menu.addAction("Обновить")

        copy_action.triggered.connect(self.copy_cell_content)
        delete_action.triggered.connect(self.delete_selected_row)
        apply_action.triggered.connect(self.apply_pending_changes)
        revert_action.triggered.connect(self.load_table_rows)

        menu.exec_(self.table.viewport().mapToGlobal(position))

    def copy_cell_content(self):
        selected_items = self.table.selectedItems()
        if not selected_items:
            return

        clipboard = QApplication.clipboard()
        if len(selected_items) == 1:
            item = selected_items[0]
            clipboard.setText(item.text())
        else:
            rows = {}
            for item in selected_items:
                rows.setdefault(item.row(), {})[item.column()] = item.text()
            text_rows = []
            for _, cols in sorted(rows.items()):
                text_rows.append("\t".join(value for _, value in sorted(cols.items())))
            clipboard.setText("\n".join(text_rows).strip())

        self.show_message("Скопировано")

    def delete_selected_row(self):
        row = self.table.currentRow()
        if row < 0 or row >= len(self._table_row_ids):
            self.show_error_message("Не выбрана строка")
            return

        row_id = self._table_row_ids[row]

        self._pending_deletes.add(row_id)
        self._pending_changes.pop(row_id, None)

        self._updating_table = True
        try:
            self.table.removeRow(row)
            del self._table_row_ids[row]
        finally:
            self._updating_table = False

        self.show_message("Строка помечена на удаление. Нажми 'Сохранить' для применения.")

    def refresh_current_product_combo(self):
        row = self.table.currentRow()
        if row < 0:
            return

        combo = self.table.cellWidget(row, 0)
        if isinstance(combo, QComboBox):
            row_id = combo.property("row_id")
            if row_id is not None:
                self.populate_product_combo(
                    combo,
                    row_id=int(row_id),
                    keep_current=True,
                )

    def get_filtered_products(self) -> list[Product]:
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

    def add_line(self):
        try:
            self.save_pending_changes_to_temp()

            supplier_id = self.ensure_supplier()
            import_date = self.get_price_date()

            with self.get_session() as session:
                service = SupplierPriceService(session)
                service.create_empty_temp_row(
                    supplier_id=supplier_id,
                    batch_id=self.batch_id,
                    imported_by=self.imported_by,
                    import_date=import_date,
                )
                session.commit()

            self.load_table_rows()
            if self.table.rowCount() > 0:
                self.table.setCurrentCell(0, 0)
            self.show_message("Добавлена строка")
        except Exception as e:
            self.show_error_message(str(e))


    def _commit_open_editors(self):
        for row in range(self.table.rowCount()):
            for column in (0, 8):
                widget = self.table.cellWidget(row, column)
                if not isinstance(widget, QComboBox):
                    continue
                if row < 0 or row >= len(self._table_row_ids):
                    continue
                row_id = self._table_row_ids[row]
                if column == 0:
                    self.finish_product_edit(row, row_id, widget)
                elif column == 8:
                    self.finish_brand_edit(row, row_id, widget)

    def save_pending_changes_to_temp(self):
        self._commit_open_editors()

        if not self._pending_changes and not self._pending_deletes:
            return

        supplier_id = self.ensure_supplier()
        import_date = self.get_price_date()

        with self.get_session() as session:
            for row_id, changes in self._pending_changes.items():
                row = session.query(TempPriceImport).filter(TempPriceImport.id == row_id).first()
                if row is None:
                    continue

                row.supplier_id = supplier_id
                row.import_date = import_date

                for key, value in changes.items():
                    if key == "selected_product_id":
                        if value in (None, "", 0):
                            row.selected_product_id = None
                        else:
                            try:
                                row.selected_product_id = int(value)
                            except (TypeError, ValueError):
                                continue
                    elif key in {"price", "price_pack", "qty_pcs", "volume_l", "new_pack"}:
                        parsed = parse_loose_number(value)
                        setattr(row, key, parsed if parsed is not None else None)
                    else:
                        setattr(row, key, value)

                if row.selected_product_id is not None:
                    row.new_product_name = None
                    row.new_brand = None
                    row.new_pack = None
                    row.new_is_excise = False
                else:
                    has_new_product_data = any([
                        bool(clean_multi_spaces(row.new_product_name)),
                        bool(clean_multi_spaces(row.new_brand)),
                        row.new_pack is not None,
                    ])
                    if has_new_product_data and row.new_is_excise is None:
                        row.new_is_excise = False

            if self._pending_deletes:
                session.query(TempPriceImport).filter(
                    TempPriceImport.id.in_(self._pending_deletes)
                ).delete(synchronize_session=False)

            session.commit()

        self._pending_changes.clear()
        self._pending_deletes.clear()

    def _has_rows_in_current_batch(self) -> bool:
        with self.get_session() as session:
            exists_row = (
                session.query(TempPriceImport.id)
                .filter(
                    TempPriceImport.batch_id == self.batch_id,
                    TempPriceImport.imported_by == self.imported_by,
                )
                .first()
            )
        return exists_row is not None

    def apply_pending_changes(self):
        self._commit_open_editors()

        has_batch_rows = self._has_rows_in_current_batch()
        if not self._pending_changes and not self._pending_deletes and not has_batch_rows:
            self.show_message("Нет изменений")
            return

        try:
            supplier_id = self.ensure_supplier()
            currency_code = clean_multi_spaces(self.ui.cbo_Currency.currentText()).upper()
            if not currency_code or currency_code == "-":
                raise ValueError("Выбери валюту")

            with self.get_session() as session:
                service = SupplierPriceService(session)

                for row_id, changes in self._pending_changes.items():
                    row = session.query(TempPriceImport).filter(TempPriceImport.id == row_id).first()
                    if row is None:
                        continue

                    row.supplier_id = supplier_id
                    row.import_date = self.get_price_date()

                    for key, value in changes.items():
                        if key == "selected_product_id":
                            if value is None:
                                setattr(row, key, None)
                                continue
                            try:
                                setattr(row, key, int(value))
                            except (TypeError, ValueError):
                                continue
                        elif key in {"price", "price_pack", "qty_pcs", "volume_l", "new_pack"}:
                            parsed = parse_loose_number(value)
                            setattr(row, key, parsed if parsed is not None else None)
                        else:
                            setattr(row, key, value)

                    if row.selected_product_id is not None:
                        row.new_product_name = None
                        row.new_brand = None
                        row.new_pack = None
                        row.new_is_excise = False
                    else:
                        has_new_product_data = any([
                            bool(clean_multi_spaces(row.new_product_name)),
                            bool(clean_multi_spaces(row.new_brand)),
                            row.new_pack is not None,
                        ])
                        if has_new_product_data and row.new_is_excise is None:
                            row.new_is_excise = False

                if self._pending_deletes:
                    session.query(TempPriceImport).filter(TempPriceImport.id.in_(self._pending_deletes)).delete(
                        synchronize_session=False
                    )

                service.validate_new_products_before_save(self.batch_id, self.imported_by)
                service.create_products_from_temp(self.batch_id, self.imported_by)
                service.create_or_update_product_articles(self.batch_id, self.imported_by)
                service.fill_price_from_price_pack(self.batch_id, self.imported_by)

                rf_prices_include_vat = False
                if self.ui.cbo_SupplierRF.currentText() == "да":
                    rf_prices_include_vat = self.ask_rf_prices_include_vat()
                self.rf_prices_include_vat = rf_prices_include_vat

                if self.ui.cbo_History.currentText() == "да":
                    service.save_prices_to_history_and_current(
                        batch_id=self.batch_id,
                        imported_by=self.imported_by,
                        currency_code=currency_code,
                        rf_prices_include_vat=rf_prices_include_vat,
                    )

                fx_rate = self.parse_decimal_field(self.ui.line_ExchangeRate, "Курс")
                service.save_supplier_price_calculations(
                    batch_id=self.batch_id,
                    imported_by=self.imported_by,
                    fx_rate=fx_rate,
                    currency_code=currency_code,
                    rf_prices_include_vat=rf_prices_include_vat,
                )
                session.commit()

            export_error_text = None
            if self.ask_export_calculated_excel():
                try:
                    self.export_calculated_excel(supplier_id)
                except Exception as export_error:
                    export_error_text = str(export_error)

            self.cleanup_current_batch(start_new_batch_after=True)
            self.load_find_brands()
            self.show_message("Данные сохранены")

            if export_error_text:
                self.show_error_message(f"Данные сохранены, но Excel не удалось выгрузить:\n{export_error_text}")
        except Exception as e:
            self.show_error_message(str(e))


    def cleanup_current_batch(self, start_new_batch_after: bool = False):
        current_batch_id = self.batch_id
        try:
            if current_batch_id:
                with self.get_session() as session:
                    service = SupplierPriceService(session)
                    service.reset_batch(current_batch_id, self.imported_by)
                    session.commit()
        except Exception:
            pass

        self._pending_changes.clear()
        self._pending_deletes.clear()
        self._new_rows.clear()
        self._table_row_ids.clear()
        self.selected_file_path = ""

        if start_new_batch_after:
            self.start_new_batch()
            self.load_table_rows()
        else:
            self.batch_id = ""

    def hideEvent(self, event):
        # При простом переходе на другое окно страницу не сбрасываем.
        # Иначе пользователь теряет данные, хотя окно фактически не закрывал.
        super().hideEvent(event)

    def closeEvent(self, event):
        self.cleanup_current_batch(start_new_batch_after=False)
        super().closeEvent(event)

    def reset_form(self):
        try:
            with self.get_session() as session:
                service = SupplierPriceService(session)
                service.reset_batch(self.batch_id, self.imported_by)
                session.commit()

            self._pending_changes.clear()
            self._pending_deletes.clear()
            self.start_new_batch()
            self.apply_default_values()
            self.load_table_rows()
            self.show_message("Форма очищена")
        except Exception as e:
            self.show_error_message(str(e))

    def _clean_table_text(self, value) -> str:
        if value is None:
            return ""

        text = str(value).strip()
        if text.lower() == "nan":
            return ""

        return text

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

    def show_error_message(self, text: str):
        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Warning)
        msg.setWindowTitle("Ошибка")
        msg.setText(text)

        copy_btn = msg.addButton("Copy", QMessageBox.ActionRole)
        msg.addButton(QMessageBox.Ok)

        msg.exec()

        if msg.clickedButton() == copy_btn:
            QApplication.clipboard().setText(text)



