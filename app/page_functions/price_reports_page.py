from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Dict, List, Optional, Sequence

from PySide6.QtCore import QFile, Qt
from PySide6.QtGui import QColor, QDesktopServices
from PySide6.QtUiTools import QUiLoader
from PySide6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QFileDialog,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from app.db.db import SessionLocal
from app.db.models import (
    CurrentSupplierPrice,
    ExchangeRate,
    FixedCosts,
    MarkingRate,
    PackType,
    PriceHistory,
    Product,
    ProductStock,
    Supplier,
)
from app.ui.table_style import *
from app.exports.price_report_exporter import PriceReportExporter


BASE_DIR = Path(__file__).resolve().parents[2]
PRICE_REPORTS_UI = BASE_DIR / "app" / "ui" / "windows" / "price_reports.ui"



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


@dataclass
class SupplierOption:
    supplier_id: int
    supplier_name: str
    supplier_price: Optional[Decimal]
    price_date: Optional[datetime]
    currency: str
    cost_novo: Optional[Decimal]
    full_cost: Optional[Decimal]


class PriceReportsPage(QWidget):
    def __init__(self):
        super().__init__()

        self.ui = load_ui(PRICE_REPORTS_UI)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.ui)

        self._product_name_combo = self.ui.cbo_ProductName
        self._preview_headers: List[str] = []
        self._preview_rows: List[List[object]] = []
        self._export_headers: List[str] = []
        self._export_rows: List[List[object]] = []
        self._updating_fx_table = False

        self.setup_ui()
        self.setup_connections()
        self.load_initial_data()

    def get_session(self):
        return SessionLocal()

    def setup_ui(self):
        self.preview_table = self.ui.table_ReportPreview
        self.fx_table = self.ui.table_FXRates

        setup_data_table(self.preview_table, sorting=True)
        setup_data_table(self.fx_table, sorting=False)

        self.preview_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.preview_table.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.preview_table.setSelectionMode(QAbstractItemView.ExtendedSelection)

        self.fx_table.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.fx_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.fx_table.setColumnCount(2)
        self.fx_table.setHorizontalHeaderLabels(["Currency", "Rate to RUB"])

        self.ui.radio_ByProduct.setChecked(True)
        self.ui.cbx_ShowPrevPrice.setChecked(False)
        self.ui.cbx_ShowPrevPrice.setEnabled(False)

        self.ui.lst_Brand.setSelectionMode(QAbstractItemView.MultiSelection)
        self.ui.lst_ProductFamily.setSelectionMode(QAbstractItemView.MultiSelection)

    def setup_connections(self):
        self.ui.radio_ByProduct.toggled.connect(self.on_mode_changed)
        self.ui.radio_BySupplier.toggled.connect(self.on_mode_changed)
        self.ui.cbo_Supplier.currentIndexChanged.connect(self.on_supplier_changed)
        self.ui.btn_BuildReport.clicked.connect(self.build_report)
        self.ui.btn_Reset.clicked.connect(self.reset_filters)
        self.ui.btn_ExportExcel.clicked.connect(self.export_excel)

        self.ui.lst_Brand.itemSelectionChanged.connect(self.on_brand_selection_changed)
        self.ui.lst_ProductFamily.itemSelectionChanged.connect(self.on_family_selection_changed)

    def load_initial_data(self):
        self.fill_suppliers()
        self.fill_brand_list()
        self.fill_product_family_list()
        self.fill_product_name_list()
        self.load_fx_rates_table()

    def on_mode_changed(self):
        by_supplier = self.ui.radio_BySupplier.isChecked()
        self.ui.cbo_Supplier.setEnabled(by_supplier)
        self.ui.cbx_ShowPrevPrice.setEnabled(by_supplier)
        if not by_supplier:
            self.ui.cbx_ShowPrevPrice.setChecked(False)

        self.fill_brand_list()
        self.fill_product_family_list()
        self.fill_product_name_list()
        self.clear_preview_table()
        self.clear_message()

    def on_supplier_changed(self):
        if self.ui.radio_BySupplier.isChecked():
            self.fill_brand_list()
            self.fill_product_family_list()
            self.fill_product_name_list()
            self.clear_preview_table()
            self.clear_message()

    def on_brand_selection_changed(self):
        self._normalize_multiselect(self.ui.lst_Brand)
        self.fill_product_family_list()
        self.fill_product_name_list()

    def on_family_selection_changed(self):
        self._normalize_multiselect(self.ui.lst_ProductFamily)
        self.fill_product_name_list()

    def _normalize_multiselect(self, list_widget: QListWidget):
        selected_texts = self._get_selected_list_values(list_widget)
        if "-" in selected_texts and len(selected_texts) > 1:
            for i in range(list_widget.count()):
                item = list_widget.item(i)
                if item and item.text() == "-":
                    item.setSelected(False)
                    break

    def fill_suppliers(self):
        try:
            with self.get_session() as session:
                suppliers = session.query(Supplier).order_by(Supplier.name).all()

            self.ui.cbo_Supplier.blockSignals(True)
            self.ui.cbo_Supplier.clear()
            self.ui.cbo_Supplier.addItem("-", None)
            for supplier in suppliers:
                self.ui.cbo_Supplier.addItem(supplier.name, supplier.id)
            self.ui.cbo_Supplier.blockSignals(False)
        except Exception as e:
            self.show_error_message(f"Ошибка при получении поставщиков: {str(e)}")

    def fill_brand_list(self):
        try:
            with self.get_session() as session:
                products = self._get_available_products(session)
                brands = sorted({(p.brand or "").strip() for p in products if (p.brand or "").strip()})
            self._fill_list_widget(self.ui.lst_Brand, brands)
        except Exception as e:
            self.show_error_message(f"Ошибка при получении брендов: {str(e)}")

    def fill_product_family_list(self):
        selected_brands = self._get_selected_list_values(self.ui.lst_Brand)
        try:
            with self.get_session() as session:
                products = self._get_available_products(session)
                if selected_brands:
                    products = [p for p in products if (p.brand or "") in selected_brands]
                families = sorted({(p.family or "").strip() for p in products if (p.family or "").strip()})
            self._fill_list_widget(self.ui.lst_ProductFamily, families)
        except Exception as e:
            self.show_error_message(f"Ошибка при получении Product Family: {str(e)}")

    def fill_product_name_list(self):
        selected_brands = self._get_selected_list_values(self.ui.lst_Brand)
        selected_families = self._get_selected_list_values(self.ui.lst_ProductFamily)
        current_name = self._product_name_combo.currentText().strip()

        try:
            with self.get_session() as session:
                products = self._get_available_products(session)
                if selected_brands:
                    products = [p for p in products if (p.brand or "") in selected_brands]
                if selected_families:
                    products = [p for p in products if (p.family or "") in selected_families]
                names = sorted({(p.name or "").strip() for p in products if (p.name or "").strip()})

            self._product_name_combo.blockSignals(True)
            self._product_name_combo.clear()
            self._product_name_combo.addItem("-")
            self._product_name_combo.addItems(names)
            if current_name and current_name in names:
                self._product_name_combo.setCurrentText(current_name)
            self._product_name_combo.blockSignals(False)
        except Exception as e:
            self.show_error_message(f"Ошибка при получении названий продуктов: {str(e)}")

    def _fill_list_widget(self, widget: QListWidget, values: Sequence[str]):
        selected_before = set(self._get_selected_list_values(widget, include_dash=True))
        widget.blockSignals(True)
        widget.clear()
        dash_item = QListWidgetItem("-")
        widget.addItem(dash_item)
        if not selected_before or selected_before == {"-"}:
            dash_item.setSelected(True)

        for value in values:
            item = QListWidgetItem(value)
            widget.addItem(item)
            if value in selected_before and value != "-":
                item.setSelected(True)

        widget.blockSignals(False)

    def _get_selected_list_values(self, widget: QListWidget, include_dash: bool = False) -> List[str]:
        values = [item.text().strip() for item in widget.selectedItems() if item and item.text().strip()]
        if include_dash:
            return values
        return [value for value in values if value != "-"]

    def _get_available_products(self, session) -> List[Product]:
        by_supplier = self.ui.radio_BySupplier.isChecked()
        supplier_id = self.ui.cbo_Supplier.currentData()

        query = session.query(Product).order_by(Product.brand, Product.family, Product.name)
        products = query.all()

        if not by_supplier or not supplier_id:
            return products

        valid_product_ids = set()
        current_ids = (
            session.query(CurrentSupplierPrice.product_id)
            .filter(CurrentSupplierPrice.supplier_id == supplier_id)
            .all()
        )
        history_ids = (
            session.query(PriceHistory.product_id)
            .filter(PriceHistory.supplier_id == supplier_id)
            .all()
        )
        valid_product_ids.update(row[0] for row in current_ids)
        valid_product_ids.update(row[0] for row in history_ids)

        if not valid_product_ids:
            return []

        return [product for product in products if product.id in valid_product_ids]

    def load_fx_rates_table(self):
        try:
            with self.get_session() as session:
                rates = session.query(ExchangeRate).order_by(ExchangeRate.currency_code).all()

            self._updating_fx_table = True
            self.fx_table.clearContents()
            self.fx_table.setRowCount(len(rates))

            for row_index, rate in enumerate(rates):
                cur_item = QTableWidgetItem(rate.currency_code or "")
                cur_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                cur_item.setTextAlignment(Qt.AlignCenter)
                self.fx_table.setItem(row_index, 0, cur_item)

                rate_item = QTableWidgetItem(self._format_decimal(rate.rate_to_rub))
                rate_item.setTextAlignment(Qt.AlignCenter)
                self.fx_table.setItem(row_index, 1, rate_item)

            self.fx_table.resizeColumnsToContents()
            self._updating_fx_table = False
        except Exception as e:
            self._updating_fx_table = False
            self.show_error_message(f"Ошибка загрузки курсов валют: {str(e)}")

    def _get_fx_rate_map(self) -> Dict[str, Decimal]:
        result: Dict[str, Decimal] = {}
        for row in range(self.fx_table.rowCount()):
            cur_item = self.fx_table.item(row, 0)
            rate_item = self.fx_table.item(row, 1)
            if not cur_item:
                continue
            currency = (cur_item.text() or "").strip()
            if not currency:
                continue
            rate = self._to_decimal(rate_item.text() if rate_item else None)
            if rate is not None:
                result[currency] = rate
        return result

    def build_report(self):
        try:
            self.clear_message()
            fx_rates = self._get_fx_rate_map()

            if self.ui.radio_BySupplier.isChecked() and not self.ui.cbo_Supplier.currentData():
                self.show_error_message("Выбери поставщика")
                return

            if self.ui.radio_ByProduct.isChecked():
                preview_headers, preview_rows, export_headers, export_rows = self._build_product_report(fx_rates)
            else:
                preview_headers, preview_rows, export_headers, export_rows = self._build_supplier_report(fx_rates)

            self._preview_headers = preview_headers
            self._preview_rows = preview_rows
            self._export_headers = export_headers
            self._export_rows = export_rows

            self._display_preview(preview_headers, preview_rows)

            if preview_rows:
                self.show_message("Отчет сформирован")
            else:
                self.show_message("Нет данных по заданным фильтрам")
        except Exception as e:
            self.show_error_message(f"Ошибка формирования отчета: {str(e)}")

    def _get_filtered_products(self, session) -> List[Product]:
        products = self._get_available_products(session)
        selected_brands = self._get_selected_list_values(self.ui.lst_Brand)
        selected_families = self._get_selected_list_values(self.ui.lst_ProductFamily)
        selected_product_name = self._product_name_combo.currentText().strip()

        if selected_brands:
            products = [p for p in products if (p.brand or "") in selected_brands]
        if selected_families:
            products = [p for p in products if (p.family or "") in selected_families]
        if selected_product_name and selected_product_name != "-":
            products = [p for p in products if (p.name or "") == selected_product_name]

        return sorted(products, key=lambda p: ((p.brand or ""), (p.family or ""), (p.name or ""), self._pack_sort_key(p.pack)))

    def _build_product_report(self, fx_rates: Dict[str, Decimal]):
        with self.get_session() as session:
            products = self._get_filtered_products(session)
            fixed_costs = session.query(FixedCosts).first()
            preview_rows: List[List[object]] = []
            export_rows: List[List[object]] = []
            max_export_suppliers = 0
            product_export_data = []

            for product in products:
                stock = session.query(ProductStock).filter(ProductStock.product_id == product.id).first()
                options = self._get_all_supplier_options_for_product(
                    session=session,
                    product=product,
                    fx_rates=fx_rates,
                    fixed_costs=fixed_costs,
                    include_supplier_without_rating=False,
                )
                product_export_data.append((product, stock, options))
                if len(options) > max_export_suppliers:
                    max_export_suppliers = len(options)

            preview_headers = self._build_product_headers(supplier_count=4)
            export_headers = self._build_product_headers(supplier_count=max_export_suppliers)

            for product, stock, options in product_export_data:
                preview_rows.append(self._build_product_row(product, stock, options[:4], 4))
                export_rows.append(self._build_product_row(product, stock, options, max_export_suppliers))

            return preview_headers, preview_rows, export_headers, export_rows

    def _build_supplier_report(self, fx_rates: Dict[str, Decimal]):
        with self.get_session() as session:
            supplier_id = self.ui.cbo_Supplier.currentData()
            supplier = session.query(Supplier).filter(Supplier.id == supplier_id).first()
            if not supplier:
                raise Exception("Не найден выбранный поставщик")

            products = self._get_filtered_products(session)
            fixed_costs = session.query(FixedCosts).first()
            show_prev = self.ui.cbx_ShowPrevPrice.isChecked()
            preview_rows: List[List[object]] = []
            export_rows: List[List[object]] = []
            max_other_suppliers = 0
            report_data = []

            for product in products:
                stock = session.query(ProductStock).filter(ProductStock.product_id == product.id).first()
                chosen = self._build_supplier_option_for_specific_supplier(
                    session=session,
                    supplier=supplier,
                    product=product,
                    fx_rates=fx_rates,
                    fixed_costs=fixed_costs,
                )
                if chosen is None or chosen.supplier_price is None:
                    continue

                alternatives = self._get_all_supplier_options_for_product(
                    session=session,
                    product=product,
                    fx_rates=fx_rates,
                    fixed_costs=fixed_costs,
                    exclude_supplier_id=supplier.id,
                    include_supplier_without_rating=False,
                )
                prev = self._get_previous_supplier_option(
                    session=session,
                    supplier=supplier,
                    product=product,
                    current_price_date=chosen.price_date,
                    fx_rates=fx_rates,
                    fixed_costs=fixed_costs,
                ) if show_prev else None

                report_data.append((product, stock, chosen, prev, alternatives))
                if len(alternatives) > max_other_suppliers:
                    max_other_suppliers = len(alternatives)

            preview_headers = self._build_supplier_headers(show_prev=show_prev, other_count=4)
            export_headers = self._build_supplier_headers(show_prev=show_prev, other_count=max_other_suppliers)

            for product, stock, chosen, prev, alternatives in report_data:
                preview_rows.append(self._build_supplier_row(product, stock, chosen, prev, alternatives[:4], show_prev, 4))
                export_rows.append(self._build_supplier_row(product, stock, chosen, prev, alternatives, show_prev, max_other_suppliers))

            return preview_headers, preview_rows, export_headers, export_rows

    def _build_product_headers(self, supplier_count: int) -> List[str]:
        headers = [
            "Brand",
            "Product Name",
            "Pack",
            "Дистр цена",
            "Промо цена",
            "curr LPC",
            "curr Landed cost",
            "Stock",
            "Transit",
            "Purchase Order",
            "Order IS",
            "Stock IS",
            "Reserve cust",
            "Damaged",
        ]
        for idx in range(1, supplier_count + 1):
            headers.extend([
                f"Cost Novo with VAT_{idx}",
                f"Full Cost Msk_{idx}",
                f"Supplier_{idx}",
                f"last update_{idx}",
                f"Currency_{idx}",
            ])
        return headers

    def _build_supplier_headers(self, show_prev: bool, other_count: int) -> List[str]:
        headers = [
            "Our Product Name",
            "Pack",
            "last update",
            "Price, L",
            "Price (Pack)",
            "Currency",
            "Cost Novo with VAT",
            "Full Cost Msk",
        ]
        if show_prev:
            headers.extend([
                "Price, L (prev)",
                "Cost Novo with VAT (prev)",
                "Full Cost Msk (prev)",
            ])
        headers.extend([
            "Дистр цена",
            "Промо цена",
            "curr LPC",
            "curr Landed cost",
            "Best Suppl",
            "Best full Price, L",
            "last update Best1",
            "Best Suppl 2",
            "Best full Price, L 2",
            "last update Best2",
            "Stock",
            "Transit",
            "Purchase Order",
            "Order IS",
            "Stock IS",
            "Reserve cust",
            "Damaged",
        ])
        for idx in range(1, other_count + 1):
            headers.extend([
                f"Cost Novo with VAT_{idx + 2}",
                f"Full Cost Msk_{idx + 2}",
                f"Supplier_{idx + 2}",
                f"last update_{idx + 2}",
                f"Currency_{idx + 2}",
            ])
        return headers

    def _build_product_row(self, product: Product, stock: Optional[ProductStock], options: Sequence[SupplierOption], supplier_count: int) -> List[object]:
        row: List[object] = [
            product.brand or "",
            product.name or "",
            self._display_pack(product.pack),
            self._decimal_or_empty(getattr(stock, "distr_price", None)),
            self._decimal_or_empty(getattr(stock, "promo_price", None)),
            self._decimal_or_empty(getattr(stock, "lpc", None)),
            self._decimal_or_empty(getattr(stock, "landed_cost", None)),
            self._decimal_or_empty(getattr(stock, "stock_qty", None)),
            self._decimal_or_empty(getattr(stock, "transit_qty", None)),
            self._decimal_or_empty(getattr(stock, "order_qty", None)),
            self._decimal_or_empty(getattr(stock, "is_confirmed_order_qty", None)),
            self._decimal_or_empty(getattr(stock, "is_stock_qty", None)),
            self._decimal_or_empty(getattr(stock, "reserve_qty", None)),
            self._decimal_or_empty(getattr(stock, "markdown_qty", None)),
        ]

        normalized = list(options[:supplier_count])
        while len(normalized) < supplier_count:
            normalized.append(None)

        for option in normalized:
            if option is None:
                row.extend(["", "", "", "", ""])
            else:
                row.extend([
                    self._decimal_or_empty(option.cost_novo),
                    self._decimal_or_empty(option.full_cost),
                    option.supplier_name,
                    self._date_or_empty(option.price_date),
                    option.currency,
                ])
        return row

    def _build_supplier_row(
        self,
        product: Product,
        stock: Optional[ProductStock],
        chosen: SupplierOption,
        prev: Optional[SupplierOption],
        alternatives: Sequence[SupplierOption],
        show_prev: bool,
        other_count: int,
    ) -> List[object]:
        row: List[object] = [
            product.name or "",
            self._display_pack(product.pack),
            self._date_or_empty(chosen.price_date),
            self._decimal_or_empty(chosen.supplier_price),
            self._decimal_or_empty(self._pack_price(chosen.supplier_price, product.pack)),
            chosen.currency,
            self._decimal_or_empty(chosen.cost_novo),
            self._decimal_or_empty(chosen.full_cost),
        ]
        if show_prev:
            row.extend([
                self._decimal_or_empty(prev.supplier_price if prev else None),
                self._decimal_or_empty(prev.cost_novo if prev else None),
                self._decimal_or_empty(prev.full_cost if prev else None),
            ])

        best1 = alternatives[0] if len(alternatives) >= 1 else None
        best2 = alternatives[1] if len(alternatives) >= 2 else None

        row.extend([
            self._decimal_or_empty(getattr(stock, "distr_price", None)),
            self._decimal_or_empty(getattr(stock, "promo_price", None)),
            self._decimal_or_empty(getattr(stock, "lpc", None)),
            self._decimal_or_empty(getattr(stock, "landed_cost", None)),
            best1.supplier_name if best1 else "",
            self._decimal_or_empty(best1.full_cost if best1 else None),
            self._date_or_empty(best1.price_date if best1 else None),
            best2.supplier_name if best2 else "",
            self._decimal_or_empty(best2.full_cost if best2 else None),
            self._date_or_empty(best2.price_date if best2 else None),
            self._decimal_or_empty(getattr(stock, "stock_qty", None)),
            self._decimal_or_empty(getattr(stock, "transit_qty", None)),
            self._decimal_or_empty(getattr(stock, "order_qty", None)),
            self._decimal_or_empty(getattr(stock, "is_confirmed_order_qty", None)),
            self._decimal_or_empty(getattr(stock, "is_stock_qty", None)),
            self._decimal_or_empty(getattr(stock, "reserve_qty", None)),
            self._decimal_or_empty(getattr(stock, "markdown_qty", None)),
        ])

        normalized = list(alternatives[:other_count])
        while len(normalized) < other_count:
            normalized.append(None)

        for option in normalized:
            if option is None:
                row.extend(["", "", "", "", ""])
            else:
                row.extend([
                    self._decimal_or_empty(option.cost_novo),
                    self._decimal_or_empty(option.full_cost),
                    option.supplier_name,
                    self._date_or_empty(option.price_date),
                    option.currency,
                ])

        return row

    def _display_preview(self, headers: Sequence[str], rows: Sequence[Sequence[object]]):
        self.preview_table.clear()
        self.preview_table.setColumnCount(len(headers))
        self.preview_table.setHorizontalHeaderLabels(list(headers))
        self.preview_table.setRowCount(len(rows))

        for row_index, row in enumerate(rows):
            for col_index, value in enumerate(row):
                header = headers[col_index]
                item = QTableWidgetItem(self._to_display_text(value, header))
                if self._is_numeric_header(header):
                    item.setTextAlignment(Qt.AlignCenter)
                else:
                    item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                self.preview_table.setItem(row_index, col_index, item)

        self.preview_table.resizeColumnsToContents()
        for i in range(self.preview_table.columnCount()):
            if self.preview_table.columnWidth(i) < 110:
                self.preview_table.setColumnWidth(i, 110)

    def clear_preview_table(self):
        self._preview_headers = []
        self._preview_rows = []
        self._export_headers = []
        self._export_rows = []
        self.preview_table.clear()
        self.preview_table.setRowCount(0)
        self.preview_table.setColumnCount(0)

    def export_excel(self):
        if not self._export_headers or not self._export_rows:
            self.show_error_message("Сначала сформируй отчет")
            return

        try:
            default_name = self._build_export_file_name()
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "Сохранить Excel файл",
                str(BASE_DIR / default_name),
                "Excel Files (*.xlsx)",
            )
            if not file_path:
                return

            report_mode = "supplier" if self.ui.radio_BySupplier.isChecked() else "product"
            exporter = PriceReportExporter()
            output_path = exporter.export_report(
                headers=self._export_headers,
                rows=self._export_rows,
                output_path=file_path,
                report_mode=report_mode,
            )

            QDesktopServices.openUrl(Path(output_path).as_uri())
            self.show_message("Excel файл сохранен")
        except Exception as e:
            self.show_error_message(f"Ошибка экспорта в Excel: {str(e)}")

    def _build_export_file_name(self) -> str:
        now_text = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        if self.ui.radio_BySupplier.isChecked() and self.ui.cbo_Supplier.currentData():
            supplier_name = self.ui.cbo_Supplier.currentText().strip().replace("/", "_")
            return f"SupplierPrices_{supplier_name}_{now_text}.xlsx"
        return f"ProductPrices_{now_text}.xlsx"

    def reset_filters(self):
        self.ui.radio_ByProduct.setChecked(True)
        self.ui.cbo_Supplier.setCurrentIndex(0)
        self.ui.cbx_ShowPrevPrice.setChecked(False)
        self.fill_brand_list()
        self.fill_product_family_list()
        self.fill_product_name_list()
        self.load_fx_rates_table()
        self.clear_preview_table()
        self.show_message("Форма очищена")

    def _get_latest_price_record(self, session, supplier_id: int, product_id: int):
        current = (
            session.query(CurrentSupplierPrice)
            .filter(
                CurrentSupplierPrice.supplier_id == supplier_id,
                CurrentSupplierPrice.product_id == product_id,
            )
            .first()
        )
        if current:
            return {
                "price": self._to_decimal(current.price),
                "price_date": current.last_update,
                "currency": current.currency or "",
            }

        history = (
            session.query(PriceHistory)
            .filter(
                PriceHistory.supplier_id == supplier_id,
                PriceHistory.product_id == product_id,
            )
            .order_by(PriceHistory.price_date.desc())
            .first()
        )
        if history:
            return {
                "price": self._to_decimal(history.price),
                "price_date": history.price_date,
                "currency": history.currency or "",
            }
        return None

    def _get_previous_price_record(self, session, supplier_id: int, product_id: int, current_price_date: Optional[datetime]):
        if current_price_date is None:
            return None

        history = (
            session.query(PriceHistory)
            .filter(
                PriceHistory.supplier_id == supplier_id,
                PriceHistory.product_id == product_id,
                PriceHistory.price_date < current_price_date,
            )
            .order_by(PriceHistory.price_date.desc())
            .first()
        )
        if history:
            return {
                "price": self._to_decimal(history.price),
                "price_date": history.price_date,
                "currency": history.currency or "",
            }
        return None

    def _build_supplier_option_for_specific_supplier(
        self,
        session,
        supplier: Supplier,
        product: Product,
        fx_rates: Dict[str, Decimal],
        fixed_costs: Optional[FixedCosts],
    ) -> Optional[SupplierOption]:
        latest = self._get_latest_price_record(session, supplier.id, product.id)
        if not latest:
            return None
        cost_novo = self._calc_cost_novo(
            product=product,
            supplier=supplier,
            supplier_price=latest["price"],
            fx_rates=fx_rates,
            fixed_costs=fixed_costs,
            session=session,
        )
        full_cost = self._calc_full_cost(
            product=product,
            supplier=supplier,
            supplier_price=latest["price"],
            fx_rates=fx_rates,
            fixed_costs=fixed_costs,
            session=session,
            cost_novo=cost_novo,
        )
        return SupplierOption(
            supplier_id=supplier.id,
            supplier_name=supplier.name or "",
            supplier_price=latest["price"],
            price_date=latest["price_date"],
            currency=latest["currency"],
            cost_novo=cost_novo,
            full_cost=full_cost,
        )

    def _get_previous_supplier_option(
        self,
        session,
        supplier: Supplier,
        product: Product,
        current_price_date: Optional[datetime],
        fx_rates: Dict[str, Decimal],
        fixed_costs: Optional[FixedCosts],
    ) -> Optional[SupplierOption]:
        previous = self._get_previous_price_record(session, supplier.id, product.id, current_price_date)
        if not previous:
            return None
        cost_novo = self._calc_cost_novo(
            product=product,
            supplier=supplier,
            supplier_price=previous["price"],
            fx_rates=fx_rates,
            fixed_costs=fixed_costs,
            session=session,
        )
        full_cost = self._calc_full_cost(
            product=product,
            supplier=supplier,
            supplier_price=previous["price"],
            fx_rates=fx_rates,
            fixed_costs=fixed_costs,
            session=session,
            cost_novo=cost_novo,
        )
        return SupplierOption(
            supplier_id=supplier.id,
            supplier_name=supplier.name or "",
            supplier_price=previous["price"],
            price_date=previous["price_date"],
            currency=previous["currency"],
            cost_novo=cost_novo,
            full_cost=full_cost,
        )

    def _get_all_supplier_options_for_product(
        self,
        session,
        product: Product,
        fx_rates: Dict[str, Decimal],
        fixed_costs: Optional[FixedCosts],
        exclude_supplier_id: Optional[int] = None,
        include_supplier_without_rating: bool = False,
    ) -> List[SupplierOption]:
        supplier_ids = set(
            row[0] for row in session.query(CurrentSupplierPrice.supplier_id).filter(CurrentSupplierPrice.product_id == product.id).all()
        )
        supplier_ids.update(
            row[0] for row in session.query(PriceHistory.supplier_id).filter(PriceHistory.product_id == product.id).all()
        )
        if exclude_supplier_id:
            supplier_ids.discard(exclude_supplier_id)

        options: List[SupplierOption] = []
        for supplier_id in supplier_ids:
            supplier = session.query(Supplier).filter(Supplier.id == supplier_id).first()
            if not supplier:
                continue
            if not include_supplier_without_rating and not bool(getattr(supplier, "rating_calc", True)):
                continue

            option = self._build_supplier_option_for_specific_supplier(
                session=session,
                supplier=supplier,
                product=product,
                fx_rates=fx_rates,
                fixed_costs=fixed_costs,
            )
            if option and option.supplier_price is not None:
                options.append(option)

        options.sort(key=lambda opt: (self._sort_cost_key(opt.full_cost), opt.supplier_name.lower()))
        return options

    def _calc_cost_novo(
        self,
        product: Product,
        supplier: Supplier,
        supplier_price: Optional[Decimal],
        fx_rates: Dict[str, Decimal],
        fixed_costs: Optional[FixedCosts],
        session,
    ) -> Optional[Decimal]:
        if supplier_price is None or supplier_price == 0:
            return None

        fx_rate = fx_rates.get((supplier.base_currency or "").strip())
        if fx_rate is None or fx_rate == 0:
            return None

        transport = self._to_decimal(getattr(supplier, "transport_cost_per_l", None), Decimal("0"))
        reexport = self._to_decimal(getattr(supplier, "reexport_percent", None), Decimal("0"))
        fx_markup = self._to_decimal(getattr(supplier, "fx_rate_markup", None), Decimal("0"))
        agent_fee = self._to_decimal(getattr(supplier, "agent_fee", None), Decimal("0"))

        customs_clearance = self._fixed_cost(fixed_costs, "customs_clearance")
        additional_customs = self._fixed_cost(fixed_costs, "additional_customs")
        excise = self._fixed_cost(fixed_costs, "excise")
        eco_fee = self._fixed_cost(fixed_costs, "eco_fee")
        vat = self._fixed_cost(fixed_costs, "vat")
        customs_fee = self._fixed_cost(fixed_costs, "customs_fee")
        bank_fee = self._fixed_cost(fixed_costs, "bank_fee")

        if bool(getattr(supplier, "marks_for_us", False)):
            marking = Decimal("0")
        else:
            marking = self._get_marking_cost(session, product)

        customs_multiplier = Decimal("1") + customs_clearance if bool(getattr(supplier, "has_import_duty", False)) else Decimal("1")

        if bool(getattr(supplier, "is_rf", False)):
            base_before_add = (
                (supplier_price + transport)
                * (Decimal("1") + reexport)
                * customs_multiplier
                * fx_rate
                * (Decimal("1") + fx_markup)
            )
            base = base_before_add + marking + (agent_fee * fx_rate)
        else:
            base_before_add = (
                (supplier_price + transport)
                * (Decimal("1") + reexport)
                * customs_multiplier
                * (Decimal("1") + bank_fee)
                * fx_rate
                * (Decimal("1") + fx_markup)
            )
            base = base_before_add + additional_customs + marking + (agent_fee * fx_rate)
            base = base + customs_fee + (excise if bool(getattr(product, "is_excise", False)) else Decimal("0")) + eco_fee

        return self._round4(base * (Decimal("1") + vat))

    def _calc_full_cost(
        self,
        product: Product,
        supplier: Supplier,
        supplier_price: Optional[Decimal],
        fx_rates: Dict[str, Decimal],
        fixed_costs: Optional[FixedCosts],
        session,
        cost_novo: Optional[Decimal] = None,
    ) -> Optional[Decimal]:
        if cost_novo is None:
            cost_novo = self._calc_cost_novo(product, supplier, supplier_price, fx_rates, fixed_costs, session)
        if cost_novo is None:
            return None

        money = self._fixed_cost(fixed_costs, "money")
        storage = self._fixed_cost(fixed_costs, "storage")
        move_novo = self._fixed_cost(fixed_costs, "move_novo_tamozh")
        move_msk = self._fixed_cost(fixed_costs, "move_tamozh_chekhov")
        vat = self._fixed_cost(fixed_costs, "vat")

        logistics = storage
        if not bool(getattr(supplier, "is_rf", False)):
            logistics += move_msk
            if bool(getattr(supplier, "is_via_novo", False)):
                logistics += move_novo

        result = cost_novo * (Decimal("1") + money) + logistics * (Decimal("1") + vat)
        return self._round4(result)

    def _get_marking_cost(self, session, product: Product) -> Decimal:
        if product.pack is None:
            return Decimal("0")
        pack_type = session.query(PackType).filter(PackType.volume == product.pack).first()
        if not pack_type:
            return Decimal("0")
        rate = session.query(MarkingRate).filter(MarkingRate.pack_type == pack_type.name).first()
        if not rate:
            return Decimal("0")
        return self._to_decimal(rate.cost_per_l, Decimal("0")) or Decimal("0")

    def _fixed_cost(self, fixed_costs: Optional[FixedCosts], field_name: str) -> Decimal:
        if fixed_costs is None:
            return Decimal("0")
        return self._to_decimal(getattr(fixed_costs, field_name, None), Decimal("0")) or Decimal("0")

    def _pack_price(self, price_per_l: Optional[Decimal], pack: Optional[Decimal]) -> Optional[Decimal]:
        if price_per_l is None or pack is None:
            return None
        return self._round4(price_per_l * self._to_decimal(pack, Decimal("0")))

    def _sort_cost_key(self, value: Optional[Decimal]):
        if value is None:
            return (1, Decimal("0"))
        return (0, value)

    def _pack_sort_key(self, value):
        decimal_value = self._to_decimal(value)
        return decimal_value if decimal_value is not None else Decimal("0")

    def _display_pack(self, value) -> object:
        decimal_value = self._to_decimal(value)
        return decimal_value if decimal_value is not None else ""

    def _to_decimal(self, value, default: Optional[Decimal] = None) -> Optional[Decimal]:
        if value is None or value == "":
            return default
        if isinstance(value, Decimal):
            return value
        try:
            return Decimal(str(value).replace(",", "."))
        except (InvalidOperation, ValueError, TypeError):
            return default

    def _round4(self, value: Decimal) -> Decimal:
        return value.quantize(Decimal("0.0001"))

    def _decimal_or_empty(self, value):
        decimal_value = self._to_decimal(value)
        return "" if decimal_value is None else decimal_value

    def _date_or_empty(self, value: Optional[datetime]):
        return value if value else ""

    def _format_decimal(self, value) -> str:
        decimal_value = self._to_decimal(value)
        if decimal_value is None:
            return ""
        text = format(decimal_value.normalize(), "f") if decimal_value != decimal_value.to_integral() else format(decimal_value, "f")
        if "." in text:
            text = text.rstrip("0").rstrip(".")
        return text

    def _is_date_header(self, header: str) -> bool:
        h = (header or "").strip().lower()
        return "last update" in h

    def _is_money_header(self, header: str) -> bool:
        h = (header or "").strip().lower()
        return (
            h.startswith("cost novo with vat_")
            or h.startswith("full cost msk_")
            or h in {
                "curr lpc",
                "curr landed cost",
                "cost novo with vat",
                "full cost msk",
                "cost novo with vat (prev)",
                "full cost msk (prev)",
                "best full price, l",
                "best full price, l 2",
            }
        )

    def _is_integer_header(self, header: str) -> bool:
        h = (header or "").strip().lower()
        return h in {"stock", "transit", "purchase order", "order is", "stock is", "reserve cust", "damaged"}

    def _is_decimal1_header(self, header: str) -> bool:
        h = (header or "").strip().lower()
        return h in {"дистр цена", "промо цена", "price, l", "price (pack)", "price, l (prev)"}

    def _is_numeric_header(self, header: str) -> bool:
        return self._is_money_header(header) or self._is_integer_header(header) or self._is_decimal1_header(header) or self._is_date_header(header)

    def _format_int_like_text(self, value, blank_zero: bool = False) -> str:
        decimal_value = self._to_decimal(value)
        if decimal_value is None:
            return ""
        rounded = int(decimal_value.quantize(Decimal("1")))
        if blank_zero and rounded == 0:
            return ""
        return f"{rounded:,}".replace(",", " ")

    def _format_decimal1_text(self, value, blank_zero: bool = False) -> str:
        decimal_value = self._to_decimal(value)
        if decimal_value is None:
            return ""
        if blank_zero and decimal_value == 0:
            return ""
        text = f"{float(decimal_value):,.1f}"
        return text.replace(",", "_").replace(".", ",").replace("_", " ")

    def _to_display_text(self, value, header: str = "") -> str:
        if value is None or value == "":
            return ""
        if self._is_date_header(header):
            if isinstance(value, datetime):
                return value.strftime("%d.%m.%y")
            return str(value)
        if self._is_decimal1_header(header):
            return self._format_decimal1_text(value, blank_zero=True)
        if self._is_integer_header(header) or self._is_money_header(header):
            return self._format_int_like_text(value)
        if isinstance(value, datetime):
            return value.strftime("%d.%m.%y")
        if isinstance(value, Decimal):
            text = format(value, "f")
            if "." in text:
                text = text.rstrip("0").rstrip(".")
            return text.replace(".", ",")
        return str(value)

    def _write_excel_value(self, cell, header: str, value):
        if value is None or value == "":
            cell.value = ""
            return

        if self._is_date_header(header):
            if isinstance(value, datetime):
                cell.value = value
                cell.number_format = DATE_FORMAT
            else:
                cell.value = value
            return

        if self._is_decimal1_header(header):
            decimal_value = self._to_decimal(value)
            if decimal_value is None or decimal_value == 0:
                cell.value = ""
            else:
                cell.value = float(decimal_value)
                cell.number_format = DECIMAL1_FORMAT
            return

        if self._is_money_header(header):
            decimal_value = self._to_decimal(value)
            if decimal_value is None:
                cell.value = ""
            else:
                cell.value = float(decimal_value)
                cell.number_format = MONEY_FORMAT
            return

        if self._is_integer_header(header):
            decimal_value = self._to_decimal(value)
            if decimal_value is None:
                cell.value = ""
            else:
                cell.value = int(decimal_value.quantize(Decimal("1")))
                cell.number_format = INTEGER_FORMAT
            return

        if isinstance(value, Decimal):
            cell.value = float(value)
        else:
            cell.value = value

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
