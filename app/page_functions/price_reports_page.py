from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal, InvalidOperation, ROUND_CEILING, ROUND_HALF_UP
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence

from PySide6.QtCore import QFile, Qt, QThread
from PySide6.QtGui import QColor, QDesktopServices
from PySide6.QtUiTools import QUiLoader
from PySide6.QtWidgets import (
    QAbstractItemView,
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
from app.services.price_repository import PriceRepository
from app.services.cost_calculation_service import CostCalculationService
from app.utils.excel_format_rules import FORMATS
from app.workers.excel_export_worker import ExcelExportWorker
from app.utils.money import to_decimal as _shared_to_decimal, round4 as money_round4
from app.utils.message_dialogs import show_error


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
    gray_row_indexes: Sequence[int] | None = None,
) -> Path:
    return PriceReportExporter().export_report(
        headers=headers,
        rows=rows,
        output_path=output_path,
        report_mode=report_mode,
        quick_order_months=quick_order_months,
        safe_stock_months=safe_stock_months,
        gray_row_indexes=gray_row_indexes,
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
        self._selected_supplier_ids: Optional[set[int]] = None
        self._selected_country_values: Optional[set[str]] = None
        self._selected_brand_values: Optional[set[str]] = None
        self._selected_family_values: Optional[set[str]] = None
        self._selected_product_ids: Optional[set[int]] = None
        self._preview_headers: List[str] = []
        self._preview_rows: List[List[object]] = []
        self._export_headers: List[str] = []
        self._export_rows: List[List[object]] = []
        self._preview_gray_rows: set[int] = set()
        self._export_gray_rows: set[int] = set()
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
        # Блок курсов оставляем в .ui и всю его обработку сохраняем,
        # но в текущем интерфейсе он не показывается.
        if hasattr(self.ui, "frame_FXRates"):
            self.ui.frame_FXRates.setVisible(False)

        self.ui.radio_ByProduct.setChecked(True)
        self.ui.btn_FilterSupplier.setEnabled(False)
        self.ui.btn_FilterCountry.setEnabled(False)
        self.ui.cbx_ShowPrevPrice.setChecked(False)
        self.ui.cbx_ShowPrevPrice.setEnabled(False)
        if hasattr(self.ui, "spb_SuppPriceAge"):
            self.ui.spb_SuppPriceAge.setMinimum(0)
            self.ui.spb_SuppPriceAge.setMaximum(120)
            self.ui.spb_SuppPriceAge.setValue(3)
        self._export_button_text = self.ui.btn_ExportExcel.text()

        self._refresh_filter_buttons()

    def setup_connections(self):
        self.ui.radio_ByProduct.toggled.connect(self.on_mode_changed)
        self.ui.radio_BySupplier.toggled.connect(self.on_mode_changed)
        self.ui.btn_FilterSupplier.clicked.connect(self.open_supplier_filter)
        self.ui.btn_FilterCountry.clicked.connect(self.open_country_filter)
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
        self._refresh_filter_buttons(prune=True)
        self.load_fx_rates_table()

    def on_mode_changed(self):
        by_supplier = self.ui.radio_BySupplier.isChecked()
        self.ui.btn_FilterSupplier.setEnabled(by_supplier)
        self.ui.btn_FilterCountry.setEnabled(by_supplier)
        self.ui.cbx_ShowPrevPrice.setEnabled(by_supplier)
        if not by_supplier:
            self.ui.cbx_ShowPrevPrice.setChecked(False)

        self._refresh_filter_buttons(prune=True)
        self.clear_preview_table()
        self.clear_message()

    def on_name_search_changed(self):
        self._refresh_filter_buttons()
        self.clear_preview_table()

    def get_supplier_price_age_months(self) -> int:
        widget = getattr(self.ui, "spb_SuppPriceAge", None)
        if widget is None:
            return 3
        return int(widget.value())

    def _get_all_report_suppliers(self, session) -> List[Supplier]:
        return (
            session.query(Supplier)
            .filter(Supplier.name != "Manual")
            .order_by(Supplier.name.asc())
            .all()
        )

    def _get_selected_suppliers(self, session) -> List[Supplier]:
        suppliers = self._get_all_report_suppliers(session)
        if self._selected_country_values is not None:
            suppliers = [
                supplier for supplier in suppliers
                if self._clean_text(supplier.country) in self._selected_country_values
            ]
        if self._selected_supplier_ids is not None:
            selected_ids = {int(value) for value in self._selected_supplier_ids}
            suppliers = [
                supplier for supplier in suppliers
                if supplier.id is not None and int(supplier.id) in selected_ids
            ]
        return suppliers

    def open_supplier_filter(self):
        options = self._get_supplier_filter_options()
        accepted, selected = self._open_checked_filter_dialog(
            title="Фильтр по поставщикам",
            options=options,
            selected_keys=self._selected_supplier_ids,
        )
        if not accepted:
            return

        self._selected_supplier_ids = None if selected is None else {int(value) for value in selected}
        self._refresh_filter_buttons(prune=True)
        self.clear_preview_table()
        self.clear_message()

    def open_country_filter(self):
        options = self._get_country_filter_options()
        accepted, selected = self._open_checked_filter_dialog(
            title="Фильтр по странам",
            options=options,
            selected_keys=self._selected_country_values,
        )
        if not accepted:
            return

        self._selected_country_values = None if selected is None else {str(value) for value in selected}
        self._refresh_filter_buttons(prune=True)
        self.clear_preview_table()
        self.clear_message()

    def _get_supplier_filter_options(self) -> List[FilterOption]:
        try:
            with self.get_session() as session:
                suppliers = self._get_all_report_suppliers(session)
            if self._selected_country_values is not None:
                suppliers = [
                    supplier for supplier in suppliers
                    if self._clean_text(supplier.country) in self._selected_country_values
                ]
            return [
                FilterOption(
                    key=int(supplier.id),
                    label=self._clean_text(supplier.name),
                    search_text=" ".join(
                        part for part in [
                            self._clean_text(supplier.name),
                            self._clean_text(supplier.country),
                        ]
                        if part
                    ),
                )
                for supplier in suppliers
                if supplier.id is not None and self._clean_text(supplier.name)
            ]
        except Exception as e:
            self.show_error_message(f"Ошибка при получении поставщиков: {str(e)}")
            return []

    def _get_country_filter_options(self) -> List[FilterOption]:
        try:
            with self.get_session() as session:
                suppliers = self._get_all_report_suppliers(session)
            countries = sorted({
                self._clean_text(supplier.country)
                for supplier in suppliers
                if self._clean_text(supplier.country)
            })
            return [FilterOption(key=country, label=country, search_text=country) for country in countries]
        except Exception as e:
            self.show_error_message(f"Ошибка при получении стран: {str(e)}")
            return []

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
            self.ui.btn_FilterSupplier,
            all_text="все Поставщики",
            selected=self._selected_supplier_ids,
        )
        self._set_filter_button_text(
            self.ui.btn_FilterCountry,
            all_text="все Страны",
            selected=self._selected_country_values,
        )
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
                suppliers = self._get_all_report_suppliers(session)

                available_countries = {
                    self._clean_text(supplier.country)
                    for supplier in suppliers
                    if self._clean_text(supplier.country)
                }
                if self._selected_country_values is not None:
                    self._selected_country_values = {
                        value for value in self._selected_country_values
                        if value in available_countries
                    }

                suppliers_for_filter = suppliers
                if self._selected_country_values is not None:
                    suppliers_for_filter = [
                        supplier for supplier in suppliers
                        if self._clean_text(supplier.country) in self._selected_country_values
                    ]
                available_supplier_ids = {
                    int(supplier.id)
                    for supplier in suppliers_for_filter
                    if supplier.id is not None
                }
                if self._selected_supplier_ids is not None:
                    self._selected_supplier_ids = {
                        int(value) for value in self._selected_supplier_ids
                        if int(value) in available_supplier_ids
                    }

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
        def sort_key(product: Product):
            prod_group = self._clean_text(getattr(product, "prod_group", None) or product.family).casefold()
            family = self._clean_text(product.family).casefold()
            pack = self._pack_sort_key(product.pack)
            return (prod_group, family, -pack, self._clean_text(product.name).casefold())

        return sorted(products, key=sort_key)

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
        products = session.query(Product).order_by(Product.brand, Product.family, Product.name).all()

        if not by_supplier:
            return products

        suppliers = self._get_selected_suppliers(session)
        supplier_ids = [int(supplier.id) for supplier in suppliers if supplier.id is not None]
        if not supplier_ids:
            return []

        min_price_date = PriceRepository.supplier_price_cutoff_from_months(self.get_supplier_price_age_months())
        valid_product_ids = set()
        current_query = session.query(CurrentSupplierPrice.product_id).filter(
            CurrentSupplierPrice.supplier_id.in_(supplier_ids),
            CurrentSupplierPrice.price.isnot(None),
        )
        history_query = session.query(PriceHistory.product_id).filter(
            PriceHistory.supplier_id.in_(supplier_ids),
            PriceHistory.price.isnot(None),
        )
        if min_price_date is not None:
            current_query = current_query.filter(CurrentSupplierPrice.last_update >= min_price_date)
            history_query = history_query.filter(PriceHistory.price_date >= min_price_date)
        valid_product_ids.update(row[0] for row in current_query.all())
        valid_product_ids.update(row[0] for row in history_query.all())

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

            resize_columns_for_multiline_headers(self.fx_table)
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
            min_price_date = PriceRepository.supplier_price_cutoff_from_months(self.get_supplier_price_age_months())

            self._preview_gray_rows.clear()
            self._export_gray_rows.clear()

            if self.ui.radio_ByProduct.isChecked():
                preview_headers, preview_rows, export_headers, export_rows = self._build_product_report(fx_rates, min_price_date)
            else:
                preview_headers, preview_rows, export_headers, export_rows = self._build_supplier_report(fx_rates, min_price_date)

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

    def _build_group_candidates_by_source(
        self,
        session,
        source_products: Sequence[Product],
    ) -> Dict[int, List[Product]]:
        source_product_ids = {
            int(product.id)
            for product in source_products
            if product.id is not None
        }
        if not source_product_ids:
            return {}

        cost_calculation = CostCalculationService(session)
        result: Dict[int, List[Product]] = {}

        for source_product in source_products:
            if source_product.id is None:
                continue
            prod_group = self._clean_text(
                getattr(source_product, "prod_group", None) or source_product.family
            ).upper()
            if not prod_group or source_product.pack is None:
                continue

            source_pack_type = cost_calculation.get_pack_type_by_volume(source_product.pack)
            if source_pack_type is None:
                continue

            source_pack = self._to_decimal(source_product.pack)
            source_pack_type_name = self._clean_text(source_pack_type.name).casefold()
            candidates = (
                session.query(Product)
                .filter(
                    Product.prod_group == prod_group,
                    Product.id != int(source_product.id),
                )
                .order_by(Product.name.asc(), Product.id.asc())
                .all()
            )

            group_products: List[Product] = []
            for candidate in candidates:
                if candidate.id is None or int(candidate.id) in source_product_ids:
                    continue
                try:
                    if self._to_decimal(candidate.pack) != source_pack:
                        continue
                except Exception:
                    continue
                candidate_pack_type = cost_calculation.get_pack_type_by_volume(candidate.pack)
                if candidate_pack_type is None:
                    continue
                if self._clean_text(candidate_pack_type.name).casefold() != source_pack_type_name:
                    continue
                group_products.append(candidate)

            if group_products:
                group_products = sorted(
                    group_products,
                    key=lambda product: (
                        self._clean_text(product.family).casefold(),
                        -self._pack_sort_key(product.pack),
                        self._clean_text(product.name).casefold(),
                    ),
                )
                result[int(source_product.id)] = group_products

        return result

    def _build_product_report(self, fx_rates: Dict[str, Decimal], min_price_date: Optional[datetime] = None):
        with self.get_session() as session:
            products = self._get_filtered_products(session)
            group_candidates_by_source = self._build_group_candidates_by_source(session, products)

            expanded_products: List[tuple[Product, bool]] = []
            emitted_group_product_ids: set[int] = set()
            for product in products:
                expanded_products.append((product, False))
                for group_product in group_candidates_by_source.get(int(product.id), []):
                    group_product_id = int(group_product.id)
                    if group_product_id in emitted_group_product_ids:
                        continue
                    expanded_products.append((group_product, True))
                    emitted_group_product_ids.add(group_product_id)

            all_products = [product for product, _is_group in expanded_products]
            fixed_costs = session.query(FixedCosts).first()
            stock_by_product = self._prepare_report_caches(session, all_products, min_price_date)
            preview_rows: List[List[object]] = []
            export_rows: List[List[object]] = []
            preview_gray_rows: set[int] = set()
            export_gray_rows: set[int] = set()
            max_export_suppliers = 0
            product_export_data = []

            for product, is_group_generated in expanded_products:
                stock = stock_by_product.get(int(product.id))
                options = self._get_all_supplier_options_for_product(
                    session=session,
                    product=product,
                    fx_rates=fx_rates,
                    fixed_costs=fixed_costs,
                    include_supplier_without_rating=False,
                    min_price_date=min_price_date,
                )
                product_export_data.append((product, stock, options, is_group_generated))
                if len(options) > max_export_suppliers:
                    max_export_suppliers = len(options)

            preview_headers = self._build_product_headers(supplier_count=4)
            export_headers = self._build_product_headers(supplier_count=max_export_suppliers)

            for product, stock, options, is_group_generated in product_export_data:
                preview_index = len(preview_rows)
                export_index = len(export_rows)
                preview_rows.append(self._build_product_row(product, stock, options[:4], 4))
                export_rows.append(self._build_product_row(product, stock, options, max_export_suppliers))
                if is_group_generated:
                    preview_gray_rows.add(preview_index)
                    export_gray_rows.add(export_index)

            self._preview_gray_rows = preview_gray_rows
            self._export_gray_rows = export_gray_rows
            return preview_headers, preview_rows, export_headers, export_rows

    def _build_supplier_report(self, fx_rates: Dict[str, Decimal], min_price_date: Optional[datetime] = None):
        with self.get_session() as session:
            suppliers = self._get_selected_suppliers(session)
            if not suppliers:
                raise Exception("Нет поставщиков по заданному фильтру")

            products = self._get_filtered_products(session)
            fixed_costs = session.query(FixedCosts).first()

            # Сначала кешируем только базовые строки, чтобы определить, какие
            # продукты реально имеют цену у каждого выбранного поставщика.
            self._prepare_report_caches(session, products, min_price_date)
            latest_base_records = getattr(self, "_report_latest_price_records", {})

            supplier_sources: Dict[int, List[Product]] = {}
            supplier_groups: Dict[int, Dict[int, List[Product]]] = {}
            all_products_by_id: Dict[int, Product] = {
                int(product.id): product
                for product in products
                if product.id is not None
            }

            for supplier in suppliers:
                source_products = [
                    product for product in products
                    if product.id is not None
                    and (int(supplier.id), int(product.id)) in latest_base_records
                ]
                supplier_sources[int(supplier.id)] = source_products
                group_map = self._build_group_candidates_by_source(session, source_products)
                supplier_groups[int(supplier.id)] = group_map
                for candidates in group_map.values():
                    for candidate in candidates:
                        if candidate.id is not None:
                            all_products_by_id[int(candidate.id)] = candidate

            # Второй пакетный кеш включает и автоматически добавленные аналоги.
            all_products = list(all_products_by_id.values())
            stock_by_product = self._prepare_report_caches(session, all_products, min_price_date)
            show_prev = self.ui.cbx_ShowPrevPrice.isChecked()
            preview_rows: List[List[object]] = []
            export_rows: List[List[object]] = []
            preview_gray_rows: set[int] = set()
            export_gray_rows: set[int] = set()
            max_other_suppliers = 0
            report_data = []

            for supplier in suppliers:
                source_products = supplier_sources.get(int(supplier.id), [])
                group_map = supplier_groups.get(int(supplier.id), {})
                if show_prev:
                    self._prepare_previous_price_cache(
                        session,
                        int(supplier.id),
                        all_products_by_id.keys(),
                    )

                emitted_group_product_ids: set[int] = set()
                for product in source_products:
                    stock = stock_by_product.get(int(product.id))
                    chosen = self._build_supplier_option_for_specific_supplier(
                        session=session,
                        supplier=supplier,
                        product=product,
                        fx_rates=fx_rates,
                        fixed_costs=fixed_costs,
                        min_price_date=min_price_date,
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
                        min_price_date=min_price_date,
                    )
                    prev = self._get_previous_supplier_option(
                        session=session,
                        supplier=supplier,
                        product=product,
                        current_price_date=chosen.price_date,
                        fx_rates=fx_rates,
                        fixed_costs=fixed_costs,
                    ) if show_prev else None

                    report_data.append((supplier, product, stock, chosen, prev, alternatives, False))
                    if len(alternatives) > max_other_suppliers:
                        max_other_suppliers = len(alternatives)

                    for group_product in group_map.get(int(product.id), []):
                        group_product_id = int(group_product.id)
                        if group_product_id in emitted_group_product_ids:
                            continue

                        group_stock = stock_by_product.get(group_product_id)
                        group_chosen = self._build_supplier_option_for_specific_supplier(
                            session=session,
                            supplier=supplier,
                            product=group_product,
                            fx_rates=fx_rates,
                            fixed_costs=fixed_costs,
                            min_price_date=min_price_date,
                        )
                        group_alternatives = self._get_all_supplier_options_for_product(
                            session=session,
                            product=group_product,
                            fx_rates=fx_rates,
                            fixed_costs=fixed_costs,
                            exclude_supplier_id=supplier.id,
                            include_supplier_without_rating=False,
                            min_price_date=min_price_date,
                        )
                        group_prev = self._get_previous_supplier_option(
                            session=session,
                            supplier=supplier,
                            product=group_product,
                            current_price_date=group_chosen.price_date if group_chosen else None,
                            fx_rates=fx_rates,
                            fixed_costs=fixed_costs,
                        ) if show_prev and group_chosen is not None else None

                        report_data.append((
                            supplier,
                            group_product,
                            group_stock,
                            group_chosen,
                            group_prev,
                            group_alternatives,
                            True,
                        ))
                        if len(group_alternatives) > max_other_suppliers:
                            max_other_suppliers = len(group_alternatives)
                        emitted_group_product_ids.add(group_product_id)

            preview_headers = self._build_supplier_headers(show_prev=show_prev, other_count=4)
            export_headers = self._build_supplier_headers(show_prev=show_prev, other_count=max_other_suppliers)

            for supplier, product, stock, chosen, prev, alternatives, is_group_generated in report_data:
                preview_index = len(preview_rows)
                export_index = len(export_rows)
                preview_rows.append(
                    self._build_supplier_row(
                        supplier, product, stock, chosen, prev, alternatives[:4], show_prev, 4
                    )
                )
                export_rows.append(
                    self._build_supplier_row(
                        supplier, product, stock, chosen, prev, alternatives, show_prev, max_other_suppliers
                    )
                )
                if is_group_generated:
                    preview_gray_rows.add(preview_index)
                    export_gray_rows.add(export_index)

            self._preview_gray_rows = preview_gray_rows
            self._export_gray_rows = export_gray_rows
            return preview_headers, preview_rows, export_headers, export_rows

    def _build_product_headers(self, supplier_count: int) -> List[str]:
        headers = [
            "Prod Group",
            "Brand",
            "Product Name",
            "Pack",
            "Категория ABC",
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
            "Supplier",
            "Prod Group",
            "Our Product Name",
            "Pack",
            "Категория ABC",
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
            getattr(product, "prod_group", None) or product.family or "",
            product.brand or "",
            product.name or "",
            self._display_pack(product.pack),
            product.abc_category or "-",
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
        report_supplier: Supplier,
        product: Product,
        stock: Optional[ProductStock],
        chosen: Optional[SupplierOption],
        prev: Optional[SupplierOption],
        alternatives: Sequence[SupplierOption],
        show_prev: bool,
        other_count: int,
    ) -> List[object]:
        row: List[object] = [
            report_supplier.name or "",
            getattr(product, "prod_group", None) or product.family or "",
            product.name or "",
            self._display_pack(product.pack),
            product.abc_category or "-",
            self._date_or_empty(chosen.price_date if chosen else None),
            self._decimal_or_empty(chosen.supplier_price if chosen else None),
            self._decimal_or_empty(
                self._pack_price(chosen.supplier_price, product.pack) if chosen else None
            ),
            chosen.currency if chosen else "",
            self._round_fx_rate(chosen.fx_rate if chosen else None),
            self._decimal_or_empty(chosen.cost_novo if chosen else None),
            self._decimal_or_empty(chosen.full_cost if chosen else None),
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
        # При построении отчета сохраняем рассчитанный порядок строк: основная
        # позиция -> ее групповые аналоги. Сортировку пользователь сможет
        # включить кликом по заголовку уже после заполнения таблицы.
        self.preview_table.setSortingEnabled(False)
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
                if row_index in self._preview_gray_rows:
                    item.setBackground(QColor("#E8E8E8"))
                self.preview_table.setItem(row_index, col_index, item)

        resize_columns_for_multiline_headers(self.preview_table)
        self.preview_table.horizontalHeader().setSortIndicator(-1, Qt.SortOrder.AscendingOrder)
        self.preview_table.setSortingEnabled(True)

    def clear_preview_table(self):
        self._preview_headers = []
        self._preview_rows = []
        self._export_headers = []
        self._export_rows = []
        self._preview_gray_rows.clear()
        self._export_gray_rows.clear()
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
                gray_row_indexes=sorted(self._export_gray_rows),
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
        gray_row_indexes: list[int],
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
            gray_row_indexes=gray_row_indexes,
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
        if self.ui.radio_BySupplier.isChecked():
            with self.get_session() as session:
                suppliers = self._get_selected_suppliers(session)
            if len(suppliers) == 1:
                supplier_name = self._clean_text(suppliers[0].name).replace("/", "_")
                return f"SupplierPrices_{supplier_name}_{now_text}.xlsx"
            if suppliers:
                return f"SupplierPrices_{len(suppliers)}_suppliers_{now_text}.xlsx"
            return f"SupplierPrices_{now_text}.xlsx"
        return f"ProductPrices_{now_text}.xlsx"

    def reset_filters(self):
        self._selected_supplier_ids = None
        self._selected_country_values = None
        self._selected_brand_values = None
        self._selected_family_values = None
        self._selected_product_ids = None
        self.ui.radio_ByProduct.setChecked(True)
        self.ui.btn_FilterSupplier.setEnabled(False)
        self.ui.btn_FilterCountry.setEnabled(False)
        self.ui.cbx_ShowPrevPrice.setChecked(False)
        if hasattr(self.ui, "spb_SuppPriceAge"):
            self.ui.spb_SuppPriceAge.setValue(3)
        if self._name_search_widget is not None:
            if hasattr(self._name_search_widget, "clear"):
                self._name_search_widget.clear()
            elif hasattr(self._name_search_widget, "setText"):
                self._name_search_widget.setText("")
        self._refresh_filter_buttons(prune=True)
        self.load_fx_rates_table()
        self.clear_preview_table()
        self.show_message("Форма очищена")

    def _prepare_report_caches(
        self,
        session,
        products: Sequence[Product],
        min_price_date: Optional[datetime],
    ) -> dict[int, ProductStock]:
        """Load report inputs once instead of querying inside every product loop."""
        product_ids = [int(product.id) for product in products]
        stocks = (
            session.query(ProductStock).filter(ProductStock.product_id.in_(product_ids)).all()
            if product_ids else []
        )

        prices_by_product = PriceRepository(session).get_supplier_prices_for_products(
            product_ids,
            only_rating_calc=False,
            min_price_date=min_price_date,
            exclude_manual=False,
        )
        latest_records: dict[tuple[int, int], dict[str, object]] = {}
        supplier_ids_by_product: dict[int, set[int]] = {}
        supplier_ids: set[int] = set()
        for product_id, prices in prices_by_product.items():
            for price in prices:
                supplier_id = int(price.supplier_id)
                supplier_ids.add(supplier_id)
                supplier_ids_by_product.setdefault(int(product_id), set()).add(supplier_id)
                latest_records[(supplier_id, int(product_id))] = {
                    "price": self._to_decimal(price.price),
                    "price_date": price.price_date,
                    "currency": price.currency_code or "",
                }

        suppliers = (
            session.query(Supplier).filter(Supplier.id.in_(supplier_ids)).all()
            if supplier_ids else []
        )
        pack_types = session.query(PackType).all()
        marking_rates = session.query(MarkingRate).all()
        rate_by_pack_name = {
            str(rate.pack_type): self._to_decimal(rate.cost_per_l, Decimal("0")) or Decimal("0")
            for rate in marking_rates
        }
        marking_by_volume = {
            pack_type.volume: rate_by_pack_name.get(str(pack_type.name), Decimal("0"))
            for pack_type in pack_types
            if pack_type.volume is not None
        }

        self._report_latest_price_records = latest_records
        self._report_supplier_ids_by_product = supplier_ids_by_product
        self._report_suppliers_by_id = {int(supplier.id): supplier for supplier in suppliers}
        self._report_marking_by_volume = marking_by_volume
        return {int(stock.product_id): stock for stock in stocks}

    def _prepare_previous_price_cache(
        self,
        session,
        supplier_id: int,
        product_ids,
    ) -> None:
        ids = list(product_ids)
        rows = (
            session.query(PriceHistory)
            .filter(
                PriceHistory.supplier_id == supplier_id,
                PriceHistory.product_id.in_(ids),
                PriceHistory.price.isnot(None),
            )
            .order_by(
                PriceHistory.product_id.asc(),
                PriceHistory.price_date.desc(),
                PriceHistory.id.desc(),
            )
            .all()
            if ids else []
        )
        histories_by_product: dict[int, list[PriceHistory]] = {}
        for row in rows:
            histories_by_product.setdefault(int(row.product_id), []).append(row)

        previous_records: dict[tuple[int, int], dict[str, object] | None] = {}
        latest_records = getattr(self, "_report_latest_price_records", {})
        for product_id in ids:
            current = latest_records.get((supplier_id, int(product_id)))
            current_date = current.get("price_date") if current else None
            previous = None
            if current_date is not None:
                previous = next(
                    (
                        row for row in histories_by_product.get(int(product_id), [])
                        if row.price_date < current_date
                    ),
                    None,
                )
            previous_records[(supplier_id, int(product_id))] = (
                {
                    "price": self._to_decimal(previous.price),
                    "price_date": previous.price_date,
                    "currency": previous.currency or "",
                }
                if previous is not None else None
            )
        self._report_previous_price_records = previous_records

    def _get_latest_price_record(
        self,
        session,
        supplier_id: int,
        product_id: int,
        min_price_date: Optional[datetime] = None,
    ):
        cached_records = getattr(self, "_report_latest_price_records", None)
        if cached_records is not None:
            return cached_records.get((int(supplier_id), int(product_id)))

        current_query = session.query(CurrentSupplierPrice).filter(
            CurrentSupplierPrice.supplier_id == supplier_id,
            CurrentSupplierPrice.product_id == product_id,
            CurrentSupplierPrice.price.isnot(None),
        )
        if min_price_date is not None:
            current_query = current_query.filter(CurrentSupplierPrice.last_update >= min_price_date)
        current = current_query.first()
        if current:
            return {
                "price": self._to_decimal(current.price),
                "price_date": current.last_update,
                "currency": current.currency or "",
            }

        history_query = session.query(PriceHistory).filter(
            PriceHistory.supplier_id == supplier_id,
            PriceHistory.product_id == product_id,
            PriceHistory.price.isnot(None),
        )
        if min_price_date is not None:
            history_query = history_query.filter(PriceHistory.price_date >= min_price_date)
        history = history_query.order_by(PriceHistory.price_date.desc(), PriceHistory.id.desc()).first()
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

        cached_records = getattr(self, "_report_previous_price_records", None)
        if cached_records is not None:
            return cached_records.get((int(supplier_id), int(product_id)))

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
        min_price_date: Optional[datetime] = None,
    ) -> Optional[SupplierOption]:
        latest = self._get_latest_price_record(session, supplier.id, product.id, min_price_date)
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
        min_price_date: Optional[datetime] = None,
    ) -> List[SupplierOption]:
        cached_supplier_ids = getattr(self, "_report_supplier_ids_by_product", None)
        if cached_supplier_ids is not None:
            supplier_ids = set(cached_supplier_ids.get(int(product.id), set()))
        else:
            current_query = session.query(CurrentSupplierPrice.supplier_id).filter(
                CurrentSupplierPrice.product_id == product.id,
                CurrentSupplierPrice.price.isnot(None),
            )
            history_query = session.query(PriceHistory.supplier_id).filter(
                PriceHistory.product_id == product.id,
                PriceHistory.price.isnot(None),
            )
            if min_price_date is not None:
                current_query = current_query.filter(CurrentSupplierPrice.last_update >= min_price_date)
                history_query = history_query.filter(PriceHistory.price_date >= min_price_date)
            supplier_ids = {row[0] for row in current_query.all()}
            supplier_ids.update(row[0] for row in history_query.all())
        if exclude_supplier_id:
            supplier_ids.discard(exclude_supplier_id)

        options: List[SupplierOption] = []
        cached_suppliers = getattr(self, "_report_suppliers_by_id", None)
        for supplier_id in supplier_ids:
            supplier = (
                cached_suppliers.get(int(supplier_id))
                if cached_suppliers is not None
                else session.query(Supplier).filter(Supplier.id == supplier_id).first()
            )
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
                min_price_date=min_price_date,
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
        insurance = self._to_decimal(getattr(supplier, "insurance_percent", None), Decimal("0"))
        fx_markup = self._to_decimal(getattr(supplier, "fx_rate_markup", None), Decimal("0"))
        fx_markup_abs = self._to_decimal(getattr(supplier, "fx_rate_markup_abs", None), Decimal("0"))
        effective_fx_rate = fx_rate * (Decimal("1") + fx_markup) + fx_markup_abs

        customs_clearance = self._fixed_cost(fixed_costs, "customs_clearance")
        additional_customs = self._fixed_cost(fixed_costs, "additional_customs")
        excise = self._fixed_cost(fixed_costs, "excise")
        eco_fee = self._fixed_cost(fixed_costs, "eco_fee")
        vat = self._fixed_cost(fixed_costs, "vat")
        customs_fee = self._fixed_cost(fixed_costs, "customs_fee")
        bank_fee = self._fixed_cost(fixed_costs, "bank_fee")
        move = self._fixed_cost(fixed_costs, "move")

        if bool(getattr(supplier, "marks_for_us", False)):
            marking = Decimal("0")
        else:
            marking = self._get_marking_cost(session, product)

        customs_multiplier = Decimal("1") + customs_clearance if bool(getattr(supplier, "has_import_duty", False)) else Decimal("1")
        customs_and_insurance_multiplier = customs_multiplier + insurance

        if bool(getattr(supplier, "is_rf", False)):
            base_before_add = (
                (supplier_price + transport)
                * (Decimal("1") + reexport)
                * customs_and_insurance_multiplier
                * effective_fx_rate
            )
            base = base_before_add + marking + (agent_fee * fx_rate)
        else:
            base_before_add = (
                (supplier_price + transport)
                * (Decimal("1") + reexport)
                * customs_and_insurance_multiplier
                * (Decimal("1") + bank_fee)
                * effective_fx_rate
            )
            base = base_before_add + additional_customs + marking + (agent_fee * fx_rate)
            base = base + customs_fee + (excise if bool(getattr(product, "is_excise", False)) else Decimal("0")) + eco_fee

        if bool(getattr(supplier, "is_via_novo", False)):
            base += move

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
        vat = self._fixed_cost(fixed_costs, "vat")

        return CostCalculationService.calc_full_cost_from_cost_novo(
            cost_novo=cost_novo,
            money=money,
            storage=storage,
            vat=vat,
        )

    def _get_marking_cost(self, session, product: Product) -> Decimal:
        if product.pack is None:
            return Decimal("0")
        cached_marking = getattr(self, "_report_marking_by_volume", None)
        if cached_marking is not None:
            return cached_marking.get(product.pack, Decimal("0"))
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
        # Delegates to the shared app.utils.money.to_decimal implementation.
        # Kept as a thin wrapper (instead of `_to_decimal = staticmethod(...)`)
        # because this file's default is `None` (used to mean "blank cell"),
        # while the shared helper's own default is Decimal("0").
        return _shared_to_decimal(value, default)

    def _round4(self, value: Decimal) -> Decimal:
        # Delegates to the single project-wide rounding rule (ROUND_HALF_UP).
        # Previously this quantized without a rounding mode, which silently
        # used ROUND_HALF_EVEN and could round the same amount differently
        # than every other module in the app.
        return money_round4(value)

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
        rounded = int(decimal_value.quantize(Decimal("1"), rounding=ROUND_HALF_UP))
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
                cell.number_format = FORMATS.DATE
            else:
                cell.value = value
            return

        if self._is_decimal1_header(header):
            decimal_value = self._to_decimal(value)
            if decimal_value is None or decimal_value == 0:
                cell.value = ""
            else:
                cell.value = float(decimal_value)
                cell.number_format = FORMATS.DECIMAL_1
            return

        if self._is_money_header(header):
            decimal_value = self._to_decimal(value)
            if decimal_value is None:
                cell.value = ""
            else:
                cell.value = float(decimal_value)
                cell.number_format = FORMATS.MONEY_RUB
            return

        if self._is_integer_header(header):
            decimal_value = self._to_decimal(value)
            if decimal_value is None:
                cell.value = ""
            else:
                cell.value = int(decimal_value.quantize(Decimal("1"), rounding=ROUND_HALF_UP))
                cell.number_format = FORMATS.INTEGER
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
        show_error(self, text)
