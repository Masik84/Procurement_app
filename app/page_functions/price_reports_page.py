from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal, InvalidOperation, ROUND_CEILING, ROUND_HALF_UP
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence

from PySide6.QtCore import QFile, Qt, QThread
from PySide6.QtGui import QDesktopServices
from PySide6.QtUiTools import QUiLoader
from PySide6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QMessageBox,
    QInputDialog,
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
    OrderPlanningCalculation,
)
from app.ui.table_style import *
from app.utils.checked_filter_dialog import CheckedFilterDialog, FilterOption
from app.exports.price_report_exporter import PriceReportExporter
from app.workers.excel_export_worker import ExcelExportWorker


BASE_DIR = Path(__file__).resolve().parents[2]
PRICE_REPORTS_UI = BASE_DIR / "app" / "ui" / "windows" / "price_reports.ui"
CHECKED_FILTER_DIALOG_UI = BASE_DIR / "app" / "ui" / "windows" / "checked_filter_dialog.ui"



def load_ui(ui_path: Path, parent=None):
    loader = QUiLoader()
    ui_file = QFile(str(ui_path))
    if not ui_file.open(QFile.ReadOnly):
        raise RuntimeError(f"Не удалось открыть UI: {ui_path}")
    try:
        widget = loader.load(ui_file, parent)
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
    fx_rate: Optional[Decimal]
    cost_novo: Optional[Decimal]
    full_cost: Optional[Decimal]


def _export_price_report_file(
    *,
    headers: Sequence[str],
    rows: Sequence[Sequence[object]],
    output_path: str,
    report_mode: str,
    quick_order_months: int | None,
    safe_stock_months: int | None,
) -> Path:
    return PriceReportExporter().export_report(
        headers=headers,
        rows=rows,
        output_path=output_path,
        report_mode=report_mode,
        quick_order_months=quick_order_months,
        safe_stock_months=safe_stock_months,
    )


class PriceReportsPage(QWidget):
    def __init__(self):
        super().__init__()

        self.ui = load_ui(PRICE_REPORTS_UI)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.ui)

        self._product_name_combo = getattr(self.ui, "cbo_ProductName", None)
        self._name_search_widget = getattr(self.ui, "line_NameSearch", None) or getattr(self.ui, "lineEdit", None)
        self._selected_brand_values: Optional[set[str]] = None
        self._selected_family_values: Optional[set[str]] = None
        self._selected_product_ids: Optional[set[int]] = None
        self._preview_headers: List[str] = []
        self._preview_rows: List[List[object]] = []
        self._export_headers: List[str] = []
        self._export_rows: List[List[object]] = []
        self._updating_fx_table = False
        self._export_quick_order_months = None
        self._export_safe_stock_months = None
        self._excel_export_thread: QThread | None = None
        self._excel_export_worker: ExcelExportWorker | None = None
        self._export_button_text = ""

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
        self._export_button_text = self.ui.btn_ExportExcel.text()

        self._refresh_filter_buttons()

    def setup_connections(self):
        self.ui.radio_ByProduct.toggled.connect(self.on_mode_changed)
        self.ui.radio_BySupplier.toggled.connect(self.on_mode_changed)
        self.ui.cbo_Supplier.currentIndexChanged.connect(self.on_supplier_changed)
        self.ui.btn_BuildReport.clicked.connect(self.build_report)
        self.ui.btn_Reset.clicked.connect(self.reset_filters)
        self.ui.btn_ExportExcel.clicked.connect(self.export_excel)

        self.ui.btn_FilterBrand.clicked.connect(self.open_brand_filter)
        self.ui.btn_FilterProductFamily.clicked.connect(self.open_family_filter)
        self.ui.btn_FilterProduct.clicked.connect(self.open_product_filter)

        if self._product_name_combo is not None and hasattr(self._product_name_combo, "currentIndexChanged"):
            self._product_name_combo.currentIndexChanged.connect(self.clear_preview_table)
        if self._name_search_widget is not None and hasattr(self._name_search_widget, "textChanged"):
            self._name_search_widget.textChanged.connect(self.on_name_search_changed)

    def load_initial_data(self):
        self.fill_suppliers()
        self._refresh_filter_buttons(prune=True)
        self.load_fx_rates_table()

    def on_mode_changed(self):
        by_supplier = self.ui.radio_BySupplier.isChecked()
        self.ui.cbo_Supplier.setEnabled(by_supplier)
        self.ui.cbx_ShowPrevPrice.setEnabled(by_supplier)
        if not by_supplier:
            self.ui.cbx_ShowPrevPrice.setChecked(False)

        self._refresh_filter_buttons(prune=True)
        self.clear_preview_table()
        self.clear_message()

    def on_supplier_changed(self):
        if self.ui.radio_BySupplier.isChecked():
            self._refresh_filter_buttons(prune=True)
            self.clear_preview_table()
            self.clear_message()

    def on_name_search_changed(self):
        self._refresh_filter_buttons()
        self.clear_preview_table()

    def fill_suppliers(self):
        try:
            with self.get_session() as session:
                suppliers = session.query(Supplier).filter(Supplier.name != "Manual").order_by(Supplier.name).all()

            self.ui.cbo_Supplier.blockSignals(True)
            self.ui.cbo_Supplier.clear()
            self.ui.cbo_Supplier.addItem("-", None)
            for supplier in suppliers:
                self.ui.cbo_Supplier.addItem(supplier.name, supplier.id)
            self.ui.cbo_Supplier.blockSignals(False)
        except Exception as e:
            self.show_error_message(f"Ошибка при получении поставщиков: {str(e)}")

    def open_brand_filter(self):
        options = self._get_brand_filter_options()
        accepted, selected = self._open_checked_filter_dialog(
            title="Фильтр по брендам",
            options=options,
            selected_keys=self._selected_brand_values,
        )
        if not accepted:
            return

        self._selected_brand_values = None if selected is None else {str(value) for value in selected}
        self._refresh_filter_buttons(prune=True)
        self.clear_preview_table()

    def open_family_filter(self):
        options = self._get_family_filter_options()
        accepted, selected = self._open_checked_filter_dialog(
            title="Фильтр по Product Family",
            options=options,
            selected_keys=self._selected_family_values,
        )
        if not accepted:
            return

        self._selected_family_values = None if selected is None else {str(value) for value in selected}
        self._refresh_filter_buttons(prune=True)
        self.clear_preview_table()

    def open_product_filter(self):
        options = self._get_product_filter_options()
        accepted, selected = self._open_checked_filter_dialog(
            title="Фильтр по продуктам",
            options=options,
            selected_keys=self._selected_product_ids,
        )
        if not accepted:
            return

        self._selected_product_ids = None if selected is None else {int(value) for value in selected}
        self._refresh_filter_buttons(prune=True)
        self.clear_preview_table()

    def _open_checked_filter_dialog(
        self,
        *,
        title: str,
        options: Sequence[FilterOption],
        selected_keys: Optional[set[Any]],
    ) -> tuple[bool, Optional[set[Any]]]:
        dialog = CheckedFilterDialog(
            self,
            title=title,
            options=options,
            selected_keys=selected_keys,
        )
        return dialog.exec_and_get_selection()

    def _get_brand_filter_options(self) -> List[FilterOption]:
        try:
            with self.get_session() as session:
                products = self._get_available_products(session)
                products = self._apply_selected_filters_to_products(
                    products,
                    use_brand=False,
                    use_family=True,
                    use_product=True,
                    use_name=True,
                )
                brands = sorted({self._clean_text(product.brand) for product in products if self._clean_text(product.brand)})
            return [FilterOption(key=brand, label=brand, search_text=brand) for brand in brands]
        except Exception as e:
            self.show_error_message(f"Ошибка при получении брендов: {str(e)}")
            return []

    def _get_family_filter_options(self) -> List[FilterOption]:
        try:
            with self.get_session() as session:
                products = self._get_available_products(session)
                products = self._apply_selected_filters_to_products(
                    products,
                    use_brand=True,
                    use_family=False,
                    use_product=True,
                    use_name=True,
                )
                families = sorted({self._clean_text(product.family) for product in products if self._clean_text(product.family)})
            return [FilterOption(key=family, label=family, search_text=family) for family in families]
        except Exception as e:
            self.show_error_message(f"Ошибка при получении Product Family: {str(e)}")
            return []

    def _get_product_filter_options(self) -> List[FilterOption]:
        try:
            with self.get_session() as session:
                products = self._get_available_products(session)
                products = self._apply_selected_filters_to_products(
                    products,
                    use_brand=True,
                    use_family=True,
                    use_product=False,
                    use_name=True,
                )
            products = self._sort_products(products)
            return [
                FilterOption(
                    key=int(product.id),
                    label=self._product_filter_label(product),
                    search_text=self._product_filter_search_text(product),
                )
                for product in products
                if product.id is not None and self._clean_text(product.name)
            ]
        except Exception as e:
            self.show_error_message(f"Ошибка при получении продуктов: {str(e)}")
            return []

    def _refresh_filter_buttons(self, prune: bool = False) -> None:
        if prune:
            self._prune_filter_selections()

        self._set_filter_button_text(
            self.ui.btn_FilterBrand,
            all_text="все Бренды",
            selected=self._selected_brand_values,
        )
        self._set_filter_button_text(
            self.ui.btn_FilterProductFamily,
            all_text="все Product Family",
            selected=self._selected_family_values,
        )
        self._set_filter_button_text(
            self.ui.btn_FilterProduct,
            all_text="все Продукты",
            selected=self._selected_product_ids,
        )

    def _set_filter_button_text(self, button, *, all_text: str, selected: set[Any] | None) -> None:
        if selected is None:
            button.setText(all_text)
            return

        button.setText(f"{all_text} ({len(selected)})")

    def _prune_filter_selections(self) -> None:
        try:
            with self.get_session() as session:
                products = self._get_available_products(session)

            available_brands = {self._clean_text(product.brand) for product in products if self._clean_text(product.brand)}
            if self._selected_brand_values is not None:
                self._selected_brand_values = {value for value in self._selected_brand_values if value in available_brands}

            products_for_family = self._apply_selected_filters_to_products(
                products,
                use_brand=True,
                use_family=False,
                use_product=False,
                use_name=True,
            )
            available_families = {self._clean_text(product.family) for product in products_for_family if self._clean_text(product.family)}
            if self._selected_family_values is not None:
                self._selected_family_values = {value for value in self._selected_family_values if value in available_families}

            products_for_product = self._apply_selected_filters_to_products(
                products,
                use_brand=True,
                use_family=True,
                use_product=False,
                use_name=True,
            )
            available_product_ids = {int(product.id) for product in products_for_product if product.id is not None}
            if self._selected_product_ids is not None:
                self._selected_product_ids = {int(value) for value in self._selected_product_ids if int(value) in available_product_ids}
        except Exception as e:
            self.show_error_message(f"Ошибка обновления фильтров: {str(e)}")

    def _apply_selected_filters_to_products(
        self,
        products: Sequence[Product],
        *,
        use_brand: bool = True,
        use_family: bool = True,
        use_product: bool = True,
        use_name: bool = True,
    ) -> List[Product]:
        filtered = list(products)

        if use_brand and self._selected_brand_values is not None:
            filtered = [product for product in filtered if self._clean_text(product.brand) in self._selected_brand_values]

        if use_family and self._selected_family_values is not None:
            filtered = [product for product in filtered if self._clean_text(product.family) in self._selected_family_values]

        if use_product and self._selected_product_ids is not None:
            selected_ids = {int(value) for value in self._selected_product_ids}
            filtered = [product for product in filtered if product.id is not None and int(product.id) in selected_ids]

        if use_name:
            name_search = self._get_name_search_text()
            if name_search:
                filtered = [product for product in filtered if self._matches_product_name_search(product.name or "", name_search)]

        return filtered

    def _clean_text(self, value: object) -> str:
        return " ".join(str(value or "").split())

    def _sort_products(self, products: Sequence[Product]) -> List[Product]:
        return sorted(products, key=lambda p: ((p.brand or ""), (p.family or ""), (p.name or ""), self._pack_sort_key(p.pack)))

    def _product_filter_label(self, product: Product) -> str:
        return self._clean_text(product.name)

    def _product_filter_search_text(self, product: Product) -> str:
        return " ".join(
            part
            for part in [
                str(product.id or ""),
                self._clean_text(product.name),
                self._clean_text(product.brand),
                self._clean_text(product.family),
                self._format_decimal(product.pack),
            ]
            if part
        )

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
            currency = (cur_item.text() or "").strip().upper()
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

    def _get_name_search_text(self) -> str:
        widget = self._name_search_widget
        if widget is None:
            return ""

        if hasattr(widget, "text"):
            value = widget.text()
        elif hasattr(widget, "toPlainText"):
            value = widget.toPlainText()
        elif hasattr(widget, "currentText"):
            value = widget.currentText()
        else:
            value = ""

        return " ".join(str(value or "").split())

    def _matches_product_name_search(self, product_name: str, search_text: str) -> bool:
        normalized_name = " ".join(str(product_name or "").split()).casefold()
        normalized_search = " ".join(str(search_text or "").split()).casefold()
        return bool(normalized_search) and normalized_search in normalized_name

    def _get_filtered_products(self, session) -> List[Product]:
        products = self._get_available_products(session)
        products = self._apply_selected_filters_to_products(products)
        return self._sort_products(products)

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
            "Reserve E-Comm",
            "Damaged",
        ]
        for idx in range(1, supplier_count + 1):
            headers.extend([
                f"Cost Novo with VAT_{idx}",
                f"Full Cost Msk_{idx}",
                f"Supplier_{idx}",
                f"last update_{idx}",
                f"FX rate_{idx}",
                f"Currency_{idx}",
            ])
        return headers

    def _build_supplier_headers(self, show_prev: bool, other_count: int) -> List[str]:
        headers = [
            "Our Product Name",
            "Pack",
            "last update",
            "Price, L",
            "Price, pack",
            "Currency",
            "FX rate",
            "Cost Novo with VAT",
            "Full Cost Msk",
        ]
        if show_prev:
            headers.extend([
                "last update (prev)",
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
            "FX rate Best1",
            "Currency Best1",
            "Best Suppl 2",
            "Best full Price, L 2",
            "last update Best2",
            "FX rate Best2",
            "Currency Best2",
            "Stock",
            "Transit",
            "Purchase Order",
            "Order IS",
            "Stock IS",
            "Reserve cust",
            "Reserve E-Comm",
            "Damaged",
        ])
        for idx in range(1, other_count + 1):
            headers.extend([
                f"Cost Novo with VAT_{idx + 2}",
                f"Full Cost Msk_{idx + 2}",
                f"Supplier_{idx + 2}",
                f"last update_{idx + 2}",
                f"FX rate_{idx + 2}",
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
            self._decimal_or_empty(getattr(stock, "is_order_qty", None)),
            self._decimal_or_empty(getattr(stock, "is_stock_qty", None)),
            self._decimal_or_empty(getattr(stock, "reserve_qty", None)),
            self._decimal_or_empty(getattr(stock, "reserve_ecomm_qty", None)),
            self._decimal_or_empty(getattr(stock, "markdown_qty", None)),
        ]

        normalized = list(options[:supplier_count])
        while len(normalized) < supplier_count:
            normalized.append(None)

        for option in normalized:
            if option is None:
                row.extend(["", "", "", "", "", ""])
            else:
                row.extend([
                    self._decimal_or_empty(option.cost_novo),
                    self._decimal_or_empty(option.full_cost),
                    option.supplier_name,
                    self._date_or_empty(option.price_date),
                    self._round_fx_rate(option.fx_rate),
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
            self._round_fx_rate(chosen.fx_rate),
            self._decimal_or_empty(chosen.cost_novo),
            self._decimal_or_empty(chosen.full_cost),
        ]
        if show_prev:
            row.extend([
                self._date_or_empty(prev.price_date if prev else None),
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
            self._round_fx_rate(best1.fx_rate if best1 else None),
            best1.currency if best1 else "",
            best2.supplier_name if best2 else "",
            self._decimal_or_empty(best2.full_cost if best2 else None),
            self._date_or_empty(best2.price_date if best2 else None),
            self._round_fx_rate(best2.fx_rate if best2 else None),
            best2.currency if best2 else "",
            self._decimal_or_empty(getattr(stock, "stock_qty", None)),
            self._decimal_or_empty(getattr(stock, "transit_qty", None)),
            self._decimal_or_empty(getattr(stock, "order_qty", None)),
            self._decimal_or_empty(getattr(stock, "is_order_qty", None)),
            self._decimal_or_empty(getattr(stock, "is_stock_qty", None)),
            self._decimal_or_empty(getattr(stock, "reserve_qty", None)),
            self._decimal_or_empty(getattr(stock, "reserve_ecomm_qty", None)),
            self._decimal_or_empty(getattr(stock, "markdown_qty", None)),
        ])

        normalized = list(alternatives[:other_count])
        while len(normalized) < other_count:
            normalized.append(None)

        for option in normalized:
            if option is None:
                row.extend(["", "", "", "", "", ""])
            else:
                row.extend([
                    self._decimal_or_empty(option.cost_novo),
                    self._decimal_or_empty(option.full_cost),
                    option.supplier_name,
                    self._date_or_empty(option.price_date),
                    self._round_fx_rate(option.fx_rate),
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
        if self._excel_export_thread is not None:
            self.show_message("Excel файл уже формируется. Дождись окончания экспорта.")
            return

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

            quick_months, order_months = self._ask_order_plan_months()
            export_headers, export_rows = self._add_order_plan_columns_for_export(
                headers=self._export_headers,
                rows=self._export_rows,
                quick_months=quick_months,
                order_months=order_months,
            )

            self._export_quick_order_months = quick_months
            self._export_safe_stock_months = order_months

            report_mode = "supplier" if self.ui.radio_BySupplier.isChecked() else "product"
            self._start_excel_export(
                headers=list(export_headers),
                rows=[list(row) for row in export_rows],
                output_path=file_path,
                report_mode=report_mode,
                quick_order_months=self._export_quick_order_months,
                safe_stock_months=self._export_safe_stock_months,
            )
        except Exception as e:
            self.show_error_message(f"Ошибка экспорта в Excel: {str(e)}")

    def _start_excel_export(
        self,
        *,
        headers: list[str],
        rows: list[list[object]],
        output_path: str,
        report_mode: str,
        quick_order_months: int | None,
        safe_stock_months: int | None,
    ) -> None:
        self.ui.btn_ExportExcel.setEnabled(False)
        self.ui.btn_ExportExcel.setText("Формируется...")
        self.show_message("Excel файл формируется в фоновом режиме. Можно продолжать работать в программе.")

        self._excel_export_thread = QThread(self)
        self._excel_export_worker = ExcelExportWorker(
            _export_price_report_file,
            headers=headers,
            rows=rows,
            output_path=output_path,
            report_mode=report_mode,
            quick_order_months=quick_order_months,
            safe_stock_months=safe_stock_months,
        )
        self._excel_export_worker.moveToThread(self._excel_export_thread)

        self._excel_export_thread.started.connect(self._excel_export_worker.run)
        self._excel_export_worker.finished.connect(self._on_excel_export_finished)
        self._excel_export_worker.error.connect(self._on_excel_export_error)
        self._excel_export_worker.finished.connect(self._excel_export_thread.quit)
        self._excel_export_worker.error.connect(self._excel_export_thread.quit)
        self._excel_export_worker.finished.connect(self._excel_export_worker.deleteLater)
        self._excel_export_worker.error.connect(self._excel_export_worker.deleteLater)
        self._excel_export_thread.finished.connect(self._excel_export_thread.deleteLater)
        self._excel_export_thread.finished.connect(self._clear_excel_export_refs)

        self._excel_export_thread.start()

    def _finish_excel_export_ui(self) -> None:
        self.ui.btn_ExportExcel.setEnabled(True)
        self.ui.btn_ExportExcel.setText(self._export_button_text or "Export Excel")

    def _on_excel_export_finished(self, output_path: object) -> None:
        self._finish_excel_export_ui()
        path = Path(output_path)
        self.show_message(f"Excel файл сохранен: {path}")
        QDesktopServices.openUrl(path.as_uri())

    def _on_excel_export_error(self, error_text: str) -> None:
        self._finish_excel_export_ui()
        self.show_error_message(f"Ошибка экспорта в Excel: {error_text}")

    def _clear_excel_export_refs(self) -> None:
        self._excel_export_thread = None
        self._excel_export_worker = None

    def _ask_order_plan_months(self) -> tuple[int | None, int | None]:
        quick_months, ok = QInputDialog.getInt(
            self,
            "Быстрый заказ",
            "Кол-во месяцев к Быстрому заказу:",
            3,
            0,
            120,
            1,
        )
        if not ok:
            return None, None

        order_months, ok = QInputDialog.getInt(
            self,
            "Заказ",
            "Кол-во месяцев к Стандартному заказу:",
            5,
            0,
            120,
            1,
        )
        if not ok:
            return None, None

        return int(quick_months), int(order_months)

    def _add_order_plan_columns_for_export(
        self,
        *,
        headers: Sequence[str],
        rows: Sequence[Sequence[object]],
        quick_months: int | None,
        order_months: int | None,
    ) -> tuple[List[str], List[List[object]]]:
        if "Damaged" not in headers:
            return list(headers), [list(row) for row in rows]

        insert_at = list(headers).index("Damaged") + 1
        new_headers = list(headers)
        extra_headers = ["Ср.Продажи мес", "к Быстрому заказу, л", "к Заказу, л"]

        # Не дублируем колонки, если отчет уже был подготовлен с ними.
        if not all(header in new_headers for header in extra_headers):
            for offset, header in enumerate(extra_headers):
                if header not in new_headers:
                    new_headers.insert(insert_at + offset, header)

        product_name_header = "Our Product Name" if self.ui.radio_BySupplier.isChecked() else "Product Name"
        if product_name_header not in headers:
            return new_headers, [list(row) for row in rows]

        product_name_idx = list(headers).index(product_name_header)
        order_plan_by_name = self._load_order_plan_export_values_by_product_name(
            quick_months=quick_months,
            order_months=order_months,
        )

        out_rows: List[List[object]] = []
        for src_row in rows:
            row = list(src_row)
            product_name = str(row[product_name_idx] or "").strip() if product_name_idx < len(row) else ""
            values = order_plan_by_name.get(product_name, {
                "Ср.Продажи мес": "",
                "к Быстрому заказу, л": "",
                "к Заказу, л": "",
            })

            # Если колонки уже есть, обновляем значения; если нет — вставляем после Damaged.
            if all(header in headers for header in extra_headers):
                for header in extra_headers:
                    idx = list(headers).index(header)
                    if idx < len(row):
                        row[idx] = values.get(header, "")
                out_rows.append(row)
            else:
                row[insert_at:insert_at] = [
                    values.get("Ср.Продажи мес", ""),
                    values.get("к Быстрому заказу, л", ""),
                    values.get("к Заказу, л", ""),
                ]
                out_rows.append(row)

        return new_headers, out_rows

    def _load_order_plan_export_values_by_product_name(
        self,
        *,
        quick_months: int | None,
        order_months: int | None,
    ) -> Dict[str, dict]:
        with self.get_session() as session:
            products = session.query(Product).all()
            product_by_id = {int(product.id): product for product in products}
            stock_by_product_id = {
                int(stock.product_id): stock
                for stock in session.query(ProductStock).all()
                if stock.product_id is not None
            }

            # Берем последний сохраненный расчет Ср.Продажи мес по каждому продукту.
            calc_rows = (
                session.query(OrderPlanningCalculation)
                .order_by(
                    OrderPlanningCalculation.product_id.asc(),
                    OrderPlanningCalculation.period_to.desc(),
                    OrderPlanningCalculation.period_from.desc(),
                    OrderPlanningCalculation.id.desc(),
                )
                .all()
            )

            latest_calc_by_product_id: Dict[int, OrderPlanningCalculation] = {}
            for calc in calc_rows:
                product_id = int(calc.product_id)
                if product_id not in latest_calc_by_product_id:
                    latest_calc_by_product_id[product_id] = calc

            result: Dict[str, dict] = {}
            for product_id, product in product_by_id.items():
                product_name = (product.name or "").strip()
                if not product_name:
                    continue

                calc = latest_calc_by_product_id.get(product_id)
                avg_sales_month = self._to_decimal(getattr(calc, "avg_sales_month", None)) if calc else None
                stock = stock_by_product_id.get(product_id)

                if avg_sales_month is None:
                    result[product_name] = {
                        "Ср.Продажи мес": "",
                        "к Быстрому заказу, л": "",
                        "к Заказу, л": "",
                    }
                    continue

                stock_qty = self._to_decimal(getattr(stock, "stock_qty", None), Decimal("0")) or Decimal("0")
                transit_qty = self._to_decimal(getattr(stock, "transit_qty", None), Decimal("0")) or Decimal("0")
                order_qty = self._to_decimal(getattr(stock, "order_qty", None), Decimal("0")) or Decimal("0")
                is_order_qty = self._to_decimal(getattr(stock, "is_order_qty", None), Decimal("0")) or Decimal("0")
                free_base = stock_qty

                # Safe Stock (st+tr) = Stock + Transit
                # Safe Stock (+ord) = Stock + Transit + Purchase Order + Order IS
                free_st_tr = free_base + transit_qty
                free_plus_ord = free_base + transit_qty + order_qty + is_order_qty

                result[product_name] = {
                    "Ср.Продажи мес": avg_sales_month,
                    "к Быстрому заказу, л": self._calc_order_liters_for_export(
                        months=quick_months,
                        avg_sales_month=avg_sales_month,
                        free_qty=free_st_tr,
                        pack=getattr(product, "pack", None),
                    ),
                    "к Заказу, л": self._calc_order_liters_for_export(
                        months=order_months,
                        avg_sales_month=avg_sales_month,
                        free_qty=free_plus_ord,
                        pack=getattr(product, "pack", None),
                    ),
                }

            return result

    def _calc_order_liters_for_export(
        self,
        *,
        months: int | None,
        avg_sales_month: Decimal,
        free_qty: Decimal,
        pack: object,
    ) -> Decimal:
        if months is None:
            return Decimal("0")

        avg_sales_month = self._to_decimal(avg_sales_month, Decimal("0")) or Decimal("0")
        if avg_sales_month <= 0:
            return Decimal("0")

        pack_value = self._to_decimal(pack, Decimal("0")) or Decimal("0")
        if pack_value <= 0:
            return Decimal("0")

        target_liters = (Decimal(str(months)) * avg_sales_month) - free_qty
        if target_liters <= 0:
            return Decimal("0")

        pieces = (target_liters / pack_value).to_integral_value(rounding=ROUND_CEILING)
        return pieces * pack_value

    def _build_export_file_name(self) -> str:
        now_text = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        if self.ui.radio_BySupplier.isChecked() and self.ui.cbo_Supplier.currentData():
            supplier_name = self.ui.cbo_Supplier.currentText().strip().replace("/", "_")
            return f"SupplierPrices_{supplier_name}_{now_text}.xlsx"
        return f"ProductPrices_{now_text}.xlsx"

    def reset_filters(self):
        self._selected_brand_values = None
        self._selected_family_values = None
        self._selected_product_ids = None
        self.ui.radio_ByProduct.setChecked(True)
        self.ui.cbo_Supplier.setCurrentIndex(0)
        self.ui.cbx_ShowPrevPrice.setChecked(False)
        if self._name_search_widget is not None:
            if hasattr(self._name_search_widget, "clear"):
                self._name_search_widget.clear()
            elif hasattr(self._name_search_widget, "setText"):
                self._name_search_widget.setText("")
        self._refresh_filter_buttons(prune=True)
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
            price_currency=latest["currency"],
        )
        full_cost = self._calc_full_cost(
            product=product,
            supplier=supplier,
            supplier_price=latest["price"],
            fx_rates=fx_rates,
            fixed_costs=fixed_costs,
            session=session,
            cost_novo=cost_novo,
            price_currency=latest["currency"],
        )
        return SupplierOption(
            supplier_id=supplier.id,
            supplier_name=supplier.name or "",
            supplier_price=latest["price"],
            price_date=latest["price_date"],
            currency=latest["currency"],
            fx_rate=fx_rates.get((latest["currency"] or "").strip().upper()),
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
            price_currency=previous["currency"],
        )
        full_cost = self._calc_full_cost(
            product=product,
            supplier=supplier,
            supplier_price=previous["price"],
            fx_rates=fx_rates,
            fixed_costs=fixed_costs,
            session=session,
            cost_novo=cost_novo,
            price_currency=previous["currency"],
        )
        return SupplierOption(
            supplier_id=supplier.id,
            supplier_name=supplier.name or "",
            supplier_price=previous["price"],
            price_date=previous["price_date"],
            currency=previous["currency"],
            fx_rate=fx_rates.get((previous["currency"] or "").strip().upper()),
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

    @staticmethod
    def _currency(value: object) -> str:
        return str(value or "").strip().upper()

    def _convert_amount_by_rates(
        self,
        amount: object,
        from_currency: object,
        to_currency: object,
        fx_rates: Dict[str, Decimal],
    ) -> Optional[Decimal]:
        value = self._to_decimal(amount, Decimal("0")) or Decimal("0")
        source = self._currency(from_currency)
        target = self._currency(to_currency)
        if value == 0 or not source or not target or source == target:
            return value
        source_rate = fx_rates.get(source)
        target_rate = fx_rates.get(target)
        if source_rate is None or target_rate is None or source_rate == 0 or target_rate == 0:
            return None
        return self._round4(value * source_rate / target_rate)

    def _calc_cost_novo(
        self,
        product: Product,
        supplier: Supplier,
        supplier_price: Optional[Decimal],
        fx_rates: Dict[str, Decimal],
        fixed_costs: Optional[FixedCosts],
        session,
        price_currency: Optional[str] = None,
    ) -> Optional[Decimal]:
        if supplier_price is None or supplier_price == 0:
            return None

        supplier_currency = self._currency(getattr(supplier, "base_currency", None))
        calc_currency = self._currency(price_currency) or supplier_currency
        fx_rate = fx_rates.get(calc_currency)
        if fx_rate is None or fx_rate == 0:
            return None

        transport = self._to_decimal(getattr(supplier, "transport_cost_per_l", None), Decimal("0")) or Decimal("0")
        agent_fee = self._to_decimal(getattr(supplier, "agent_fee", None), Decimal("0")) or Decimal("0")
        if supplier_currency and calc_currency and supplier_currency != calc_currency:
            transport = self._convert_amount_by_rates(transport, supplier_currency, calc_currency, fx_rates)
            agent_fee = self._convert_amount_by_rates(agent_fee, supplier_currency, calc_currency, fx_rates)
            if transport is None or agent_fee is None:
                return None

        reexport = self._to_decimal(getattr(supplier, "reexport_percent", None), Decimal("0"))
        fx_markup = self._to_decimal(getattr(supplier, "fx_rate_markup", None), Decimal("0"))

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
        price_currency: Optional[str] = None,
    ) -> Optional[Decimal]:
        if cost_novo is None:
            cost_novo = self._calc_cost_novo(
                product,
                supplier,
                supplier_price,
                fx_rates,
                fixed_costs,
                session,
                price_currency=price_currency,
            )
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

    def _round_fx_rate(self, value: object):
        if value is None or value == "":
            return ""
        try:
            return int(self._to_decimal(value, Decimal("0")).to_integral_value(rounding=ROUND_HALF_UP))
        except Exception:
            return ""

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
