from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Sequence

from PySide6.QtCore import QFile, Qt, QDate, QTimer
from PySide6.QtGui import QColor, QDesktopServices
from PySide6.QtUiTools import QUiLoader
from PySide6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QHBoxLayout,
    QMessageBox,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from app.db.db import SessionLocal
from app.db.models import Product
from app.exports.order_planning_exporter import OrderPlanningExporter
from app.services.order_planning_service import OrderPlanningService
from app.ui.table_style import *
from app.utils.checked_filter_dialog import CheckedFilterDialog, FilterOption
from app.utils.text import clean_multi_spaces
from app.workers.excel_export_worker import start_excel_export
from app.utils.output_headers import display_headers


BASE_DIR = Path(__file__).resolve().parents[2]
ORDER_PLANNING_UI = BASE_DIR / "app" / "ui" / "windows" / "order_planning.ui"


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


class OrderPlanningPage(QWidget):
    COL_PRODUCT = 0
    TABLE_DELETE_MESSAGE = "Строки удалены из текущей таблицы. Для фиксации результата нажми «Сохранить»."

    CALC_COLUMNS = [
        "product_name",
        "sales_product_name",
        "brand",
        "pack",
        "avg_sales_month",
        "safe_stock_st_month",
        "safe_stock_st_tr_month",
        "safe_stock_ord_month",
        "quick_order_pcs",
        "quick_order_l",
        "std_order_pcs",
        "std_order_l",
        "distr_price",
        "promo_price",
        "free_stock_st",
        "free_stock_st_tr",
        "free_stock_ord",
        "stock",
        "transit",
        "purchase_order",
        "order_is",
        "stock_is",
        "reserve",
        "reserve_ecomm",
        "markdown",
    ]
    CALC_HEADERS = [
        "Product Name",
        "Продукт_упаковка",
        "Brand",
        "Pack",
        "Ср.Продажи мес",
        "Safe Stock (st), mnth",
        "Safe Stock (st+tr), mnth",
        "Safe Stock (+ord), mnth",
        "к Быстрому Заказу, шт",
        "к Быстрому Заказу, л",
        "к Заказу, шт",
        "к Заказу, л",
        "Дистр цена",
        "Промо цена",
        "Free Stock (st)",
        "Free Stock (st+tr)",
        "Free Stock (+ord)",
        "Stock",
        "Transit",
        "Purchase Order",
        "Order IS",
        "Stock IS",
        "Резервы",
        "Reserve E-Comm",
        "УГ",
    ]

    CHECK_COLUMNS = [
        "sales_product_name",
        "product_name",
        "sales_pack",
        "sales_brand",
        "sales_is_excise",
    ]
    CHECK_HEADERS = [
        "Продукт_упаковка",
        "Product Name",
        "Упаковка",
        "Бренд",
        "Акциз",
    ]

    NUMERIC_COLUMNS = {
        "pack", "avg_sales_month", "safe_stock_st_month", "safe_stock_st_tr_month", "safe_stock_ord_month",
        "quick_order_pcs", "quick_order_l", "std_order_pcs", "std_order_l", "distr_price", "promo_price",
        "free_stock_st", "free_stock_st_tr", "free_stock_ord", "stock", "transit",
        "purchase_order", "order_is", "stock_is", "reserve", "reserve_ecomm", "markdown", "sales_pack",
    }

    def __init__(self):
        super().__init__()
        self.ui = load_ui(ORDER_PLANNING_UI)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.ui)

        self.table = self.ui.table
        self._mode = "calc"
        self._rows: list[dict] = []
        self._base_rows: list[dict] = []
        self._period_from: date | None = None
        self._period_to: date | None = None
        self._updating_table = False
        self._excel_export_thread = None
        self._excel_export_worker = None
        self._selected_brand_values: set[str] | None = None
        self._selected_family_values: set[str] | None = None
        self._selected_product_ids: set[int] | None = None

        self.setup_ui()
        self.setup_connections()
        self.load_initial_data()

    def get_session(self):
        return SessionLocal()

    def service(self, session) -> OrderPlanningService:
        return OrderPlanningService(session)

    def setup_ui(self):
        setup_data_table(self.table, sorting=True)
        self.table.horizontalHeader().setSectionsMovable(False)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)

        from app.utils.gui_table_actions import install_standard_table_context_menu
        install_standard_table_context_menu(self, self.table)

        self.ui.date_SalesFrom.setCalendarPopup(True)
        self.ui.date_SalesTo.setCalendarPopup(True)
        self.ui.date_SalesFrom.setDisplayFormat("dd.MM.yyyy")
        self.ui.date_SalesTo.setDisplayFormat("dd.MM.yyyy")
        today = QDate.currentDate()
        self.ui.date_SalesFrom.setDate(today)
        self.ui.date_SalesTo.setDate(today)

        self.ui.spin_QuickOrd.setValue(3)
        self.ui.spin_SafeStock.setValue(5)
        if hasattr(self.ui, "spb_SuppPriceAge"):
            self.ui.spb_SuppPriceAge.setMinimum(0)
            self.ui.spb_SuppPriceAge.setMaximum(120)
            self.ui.spb_SuppPriceAge.setValue(3)
        self.ui.lbl_CalcPeriod.setText("Текущий период расчета: -")
        self.ui.lbl_QuickVolResult.setText("0")
        self.ui.lbl_StandVolResult.setText("0")
        self._refresh_filter_buttons()

    def setup_connections(self):
        self.ui.btn_FilterBrand.clicked.connect(self.open_brand_filter)
        self.ui.btn_FilterProductFamily.clicked.connect(self.open_family_filter)
        self.ui.btn_FilterProduct.clicked.connect(self.open_product_filter)
        self.ui.radio_VolNotNull.toggled.connect(self.recalculate_current_rows)
        self.ui.spin_QuickOrd.valueChanged.connect(self.recalculate_current_rows)
        self.ui.spin_SafeStock.valueChanged.connect(self.recalculate_current_rows)

        self.ui.btn_Search.clicked.connect(self.search_saved)
        self.ui.btn_Calculate.clicked.connect(self.calculate)
        self.ui.btn_CheckProducts.clicked.connect(self.check_products)
        self.ui.btn_Reset.clicked.connect(self.reset_form)
        self.ui.btn_SaveExcel.clicked.connect(self.save_excel)
        self.ui.btn_Save.clicked.connect(self.save)
        self.table.cellDoubleClicked.connect(self.start_product_edit)
        self.table.itemChanged.connect(self.on_item_changed)

    def load_initial_data(self):
        self._refresh_filter_buttons(prune=True)

    def get_supplier_price_age_months(self) -> int:
        widget = getattr(self.ui, "spb_SuppPriceAge", None)
        if widget is None:
            return 3
        return int(widget.value())

    # ------------------------------------------------------------------
    # Filters
    # ------------------------------------------------------------------
    def open_brand_filter(self):
        accepted, selected = self._open_checked_filter_dialog(
            title="Фильтр по брендам",
            options=self._get_brand_filter_options(),
            selected_keys=self._selected_brand_values,
        )
        if not accepted:
            return
        self._selected_brand_values = None if selected is None else {str(value) for value in selected}
        self._refresh_filter_buttons(prune=True)
        self.recalculate_current_rows()

    def open_family_filter(self):
        accepted, selected = self._open_checked_filter_dialog(
            title="Фильтр по Product Family",
            options=self._get_family_filter_options(),
            selected_keys=self._selected_family_values,
        )
        if not accepted:
            return
        self._selected_family_values = None if selected is None else {str(value) for value in selected}
        self._refresh_filter_buttons(prune=True)
        self.recalculate_current_rows()

    def open_product_filter(self):
        accepted, selected = self._open_checked_filter_dialog(
            title="Фильтр по продуктам",
            options=self._get_product_filter_options(),
            selected_keys=self._selected_product_ids,
        )
        if not accepted:
            return
        self._selected_product_ids = None if selected is None else {int(value) for value in selected}
        self._refresh_filter_buttons(prune=True)
        self.recalculate_current_rows()

    def _open_checked_filter_dialog(
        self,
        *,
        title: str,
        options: Sequence[FilterOption],
        selected_keys: set[Any] | None,
    ) -> tuple[bool, set[Any] | None]:
        dialog = CheckedFilterDialog(self, title=title, options=options, selected_keys=selected_keys)
        return dialog.exec_and_get_selection()

    def _get_brand_filter_options(self) -> list[FilterOption]:
        try:
            with self.get_session() as session:
                products = self._get_all_products(session)
            products = self._apply_selected_filters_to_products(products, use_brand=False, use_family=True, use_product=True)
            brands = sorted({self._clean_text(product.brand) for product in products if self._clean_text(product.brand)})
            return [FilterOption(key=brand, label=brand, search_text=brand) for brand in brands]
        except Exception as e:
            self.show_error_message(f"Ошибка загрузки брендов: {e}")
            return []

    def _get_family_filter_options(self) -> list[FilterOption]:
        try:
            with self.get_session() as session:
                products = self._get_all_products(session)
            products = self._apply_selected_filters_to_products(products, use_brand=True, use_family=False, use_product=True)
            families = sorted({self._clean_text(product.family) for product in products if self._clean_text(product.family)})
            return [FilterOption(key=family, label=family, search_text=family) for family in families]
        except Exception as e:
            self.show_error_message(f"Ошибка загрузки Product Family: {e}")
            return []

    def _get_product_filter_options(self) -> list[FilterOption]:
        try:
            with self.get_session() as session:
                products = self._get_all_products(session)
            products = self._apply_selected_filters_to_products(products, use_brand=True, use_family=True, use_product=False)
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
            self.show_error_message(f"Ошибка загрузки продуктов: {e}")
            return []

    def _get_all_products(self, session) -> list[Product]:
        return session.query(Product).order_by(Product.brand.asc(), Product.family.asc(), Product.name.asc()).all()

    def _apply_selected_filters_to_products(
        self,
        products: Sequence[Product],
        *,
        use_brand: bool = True,
        use_family: bool = True,
        use_product: bool = True,
    ) -> list[Product]:
        filtered = list(products)
        if use_brand and self._selected_brand_values is not None:
            filtered = [product for product in filtered if self._clean_text(product.brand) in self._selected_brand_values]
        if use_family and self._selected_family_values is not None:
            filtered = [product for product in filtered if self._clean_text(product.family) in self._selected_family_values]
        if use_product and self._selected_product_ids is not None:
            selected_ids = {int(value) for value in self._selected_product_ids}
            filtered = [product for product in filtered if product.id is not None and int(product.id) in selected_ids]
        return filtered

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
                products = self._get_all_products(session)

            available_brands = {self._clean_text(product.brand) for product in products if self._clean_text(product.brand)}
            if self._selected_brand_values is not None:
                self._selected_brand_values = {value for value in self._selected_brand_values if value in available_brands}

            products_for_family = self._apply_selected_filters_to_products(products, use_brand=True, use_family=False, use_product=False)
            available_families = {self._clean_text(product.family) for product in products_for_family if self._clean_text(product.family)}
            if self._selected_family_values is not None:
                self._selected_family_values = {value for value in self._selected_family_values if value in available_families}

            products_for_product = self._apply_selected_filters_to_products(products, use_brand=True, use_family=True, use_product=False)
            available_product_ids = {int(product.id) for product in products_for_product if product.id is not None}
            if self._selected_product_ids is not None:
                self._selected_product_ids = {int(value) for value in self._selected_product_ids if int(value) in available_product_ids}
        except Exception as e:
            self.show_error_message(f"Ошибка обновления фильтров: {e}")

    def _clean_text(self, value: object) -> str:
        return " ".join(str(value or "").split())

    def _sort_products(self, products: Sequence[Product]) -> list[Product]:
        return sorted(products, key=lambda p: ((p.brand or ""), (p.family or ""), (p.name or ""), str(p.pack or "")))

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
                str(product.pack or ""),
            ]
            if part
        )

    # ------------------------------------------------------------------
    # Date / number helpers
    # ------------------------------------------------------------------
    def _qdate_to_date(self, qdate: QDate) -> date:
        return date(qdate.year(), qdate.month(), qdate.day())

    def _today_dates_selected(self) -> bool:
        today = QDate.currentDate()
        return self.ui.date_SalesFrom.date() == today and self.ui.date_SalesTo.date() == today

    def _selected_period(self) -> tuple[date, date]:
        return self._qdate_to_date(self.ui.date_SalesFrom.date()), self._qdate_to_date(self.ui.date_SalesTo.date())

    @staticmethod
    def _to_decimal(value: object) -> Decimal:
        if value is None or value == "":
            return Decimal("0")
        if isinstance(value, Decimal):
            return value
        try:
            return Decimal(str(value))
        except (InvalidOperation, ValueError, TypeError):
            return Decimal("0")

    def _format_decimal(self, value: object, digits: int = 0, blank_zero: bool = False) -> str:
        d = self._to_decimal(value)
        if blank_zero and d == 0:
            return ""
        if digits == 0:
            return f"{int(d.quantize(Decimal('1'))):,}".replace(",", " ")
        text = f"{float(d):,.{digits}f}"
        return text.replace(",", "_").replace(".", ",").replace("_", " ")

    def _format_period(self, period_from: date | None, period_to: date | None) -> str:
        if not period_from or not period_to:
            return "-"
        return f"{period_from.strftime('%d.%m.%y')}-{period_to.strftime('%d.%m.%y')}"

    def _set_period_label(self):
        self.ui.lbl_CalcPeriod.setText(f"Текущий период расчета: {self._format_period(self._period_from, self._period_to)}")

    # ------------------------------------------------------------------
    # Display
    # ------------------------------------------------------------------
    def display_rows(self, rows: list[dict], mode: str):
        self._mode = mode
        self._rows = rows
        columns = self.CHECK_COLUMNS if mode == "check" else self.CALC_COLUMNS
        headers = self.CHECK_HEADERS if mode == "check" else self.CALC_HEADERS

        self._updating_table = True
        self.table.setSortingEnabled(False)
        self.table.clear()
        self.table.setColumnCount(len(headers))
        self.table.setRowCount(len(rows))
        self.table.setHorizontalHeaderLabels(display_headers(headers))

        for row_index, row in enumerate(rows):
            for col_index, key in enumerate(columns):
                value = row.get(key, "")

                if key == "sales_product_name" and mode != "check" and row.get("product_id"):
                    value = ""

                if key == "sales_is_excise":
                    placeholder = QTableWidgetItem("")
                    placeholder.setData(Qt.UserRole, key)
                    placeholder.setData(Qt.UserRole + 1, row_index)
                    placeholder.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                    placeholder.setTextAlignment(Qt.AlignCenter)
                    self.table.setItem(row_index, col_index, placeholder)
                    self.table.setCellWidget(row_index, col_index, self.build_checkbox_widget(row_index, bool(value)))
                    continue

                if key in self.NUMERIC_COLUMNS:
                    show_zero_columns = {
                        "quick_order_pcs",
                        "quick_order_l",
                        "std_order_pcs",
                        "std_order_l",
                    }
                    text = self._format_decimal(
                        value,
                        digits=2 if key.startswith("safe_stock") or key == "avg_sales_month" else 0,
                        blank_zero=key not in show_zero_columns,
                    )
                else:
                    text = str(value or "")

                item = QTableWidgetItem(text)
                item.setData(Qt.UserRole, key)
                item.setData(Qt.UserRole + 1, row_index)
                item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                item.setTextAlignment(Qt.AlignCenter if key in self.NUMERIC_COLUMNS else Qt.AlignLeft | Qt.AlignVCenter)

                if row.get("is_auto_matched") and key == "product_name":
                    item.setBackground(QColor(255, 242, 204))
                if not row.get("product_id") and key == "product_name":
                    item.setBackground(QColor(255, 199, 206))

                self.table.setItem(row_index, col_index, item)

        self.table.resizeColumnsToContents()
        for col in range(self.table.columnCount()):
            if self.table.columnWidth(col) < 90:
                self.table.setColumnWidth(col, 90)

        if mode != "check" and "sales_product_name" in columns:
            sales_name_col = columns.index("sales_product_name")
            show_sales_name_col = any(
                (not row.get("product_id"))
                or bool(row.get("is_auto_matched"))
                or bool(row.get("is_new"))
                for row in rows
            )
            self.table.setColumnHidden(sales_name_col, not show_sales_name_col)

        self.table.setSortingEnabled(True)
        self._updating_table = False
        self.update_volume_labels()

    def _row_index_from_table_row(self, table_row: int) -> int | None:
        for col in range(self.table.columnCount()):
            item = self.table.item(table_row, col)
            if item is not None:
                idx = item.data(Qt.UserRole + 1)
                if idx is not None:
                    try:
                        return int(idx)
                    except (TypeError, ValueError):
                        return None
        return None

    def build_checkbox_widget(self, row_index: int, checked: bool) -> QWidget:
        checkbox = QCheckBox()
        checkbox.setChecked(checked)
        checkbox.setStyleSheet(
            """
            QCheckBox { background: transparent; }
            QCheckBox::indicator { width: 14px; height: 14px; }
            """
        )
        checkbox.toggled.connect(lambda state, idx=row_index: self.on_excise_checkbox_changed(idx, state))

        container = QWidget()
        layout = QHBoxLayout(container)
        layout.addWidget(checkbox)
        layout.setAlignment(Qt.AlignCenter)
        layout.setContentsMargins(0, 0, 0, 0)
        return container

    def on_excise_checkbox_changed(self, row_index: int, checked: bool):
        if self._updating_table or row_index < 0 or row_index >= len(self._rows):
            return
        self._rows[row_index]["sales_is_excise"] = bool(checked)

    def on_item_changed(self, item: QTableWidgetItem):
        if self._updating_table:
            return
        key = item.data(Qt.UserRole)
        if key != "sales_is_excise":
            return
        source_row = item.data(Qt.UserRole + 1)
        try:
            source_row = int(source_row)
        except (TypeError, ValueError):
            return
        if 0 <= source_row < len(self._rows):
            self._rows[source_row]["sales_is_excise"] = item.checkState() == Qt.Checked

    def update_volume_labels(self):
        quick_total = sum(self._to_decimal(row.get("quick_order_l")) for row in self._rows)
        std_total = sum(self._to_decimal(row.get("std_order_l")) for row in self._rows)
        self.ui.lbl_QuickVolResult.setText(self._format_decimal(quick_total, 0))
        self.ui.lbl_StandVolResult.setText(self._format_decimal(std_total, 0))

    # ------------------------------------------------------------------
    # Standard table context menu integration
    # ------------------------------------------------------------------
    def before_standard_table_rows_deleted(self, table, table_rows: list[int], snapshots: list[dict]):
        """Keep Order Planning source rows in sync with the generic visual delete.

        The shared context-menu helper removes rows from QTableWidget only. This
        page saves/exports data from self._rows, so deleted rows must be removed
        from the backing list too. The original row dict is stored in the helper
        snapshot to make Undo restore the data, not just the visual cells.
        """
        if table is not self.table:
            return

        source_rows: list[int] = []
        for table_row, snapshot in zip(table_rows, snapshots):
            source_row = self._row_index_from_table_row(table_row)
            if source_row is None or source_row < 0 or source_row >= len(self._rows):
                continue
            snapshot["_order_source_row"] = source_row
            snapshot["_order_row_data"] = dict(self._rows[source_row])
            source_rows.append(source_row)

        for source_row in sorted(set(source_rows), reverse=True):
            if 0 <= source_row < len(self._rows):
                self._rows.pop(source_row)

    def after_standard_table_rows_deleted(self, table, table_rows: list[int], snapshots: list[dict]):
        if table is not self.table:
            return
        self._refresh_after_context_rows_changed()

    def after_standard_table_rows_restored(self, table, snapshots: list[dict]):
        if table is not self.table:
            return

        rows_to_restore = []
        for snapshot in snapshots:
            row_data = snapshot.get("_order_row_data")
            source_row = snapshot.get("_order_source_row")
            if row_data is None or source_row is None:
                continue
            try:
                source_row = int(source_row)
            except (TypeError, ValueError):
                source_row = len(self._rows)
            rows_to_restore.append((source_row, dict(row_data)))

        for source_row, row_data in sorted(rows_to_restore, key=lambda item: item[0]):
            self._rows.insert(min(max(source_row, 0), len(self._rows)), row_data)

        self._refresh_after_context_rows_changed()

    def _refresh_after_context_rows_changed(self):
        if self._mode in {"calc", "search"}:
            self._base_rows = [dict(row) for row in self._rows]
        self.display_rows(self._rows, self._mode)

    def recalculate_current_rows(self):
        if self._mode not in {"calc", "search"} or not self._base_rows:
            return
        try:
            with self.get_session() as session:
                rows = self.service(session).build_display_rows(
                    self._base_rows,
                    quick_months=self.ui.spin_QuickOrd.value(),
                    safe_months=self.ui.spin_SafeStock.value(),
                    brand_filter=None if self._selected_brand_values is None else list(self._selected_brand_values),
                    family_filter=None if self._selected_family_values is None else list(self._selected_family_values),
                    product_filter=None if self._selected_product_ids is None else list(self._selected_product_ids),
                    vol_not_null=self.ui.radio_VolNotNull.isChecked(),
                )
            self.display_rows(rows, self._mode)
        except Exception as e:
            self.show_error_message(str(e))

    # ------------------------------------------------------------------
    # Actions
    # ------------------------------------------------------------------
    def check_products(self):
        try:
            with self.get_session() as session:
                result = self.service(session).check_products()
            if not result.rows:
                self._rows = []
                self.display_rows([], "check")
                self.show_message("Изменений не найдено")
                return
            self.display_rows(result.rows, "check")
            self.show_message(f"Найдено изменений: {len(result.rows)}. Новые Код: {result.new_count}. Автоподбор: {result.auto_matched_count}.")
        except Exception as e:
            self.show_error_message(str(e))

    def calculate(self):
        if self._today_dates_selected():
            self.show_error_message("Выбери период расчета. Он должен быть не менее месяца")
            return
        period_from, period_to = self._selected_period()
        if period_to < period_from:
            self.show_error_message("Дата 'Продажи по' не может быть меньше даты 'Продажи с'")
            return
        if (period_to - period_from).days + 1 < 30:
            self.show_error_message("Выбери период расчета. Он должен быть не менее месяца")
            return
        try:
            with self.get_session() as session:
                result = self.service(session).calculate(period_from, period_to)
                self._base_rows = result.rows
                self._period_from = result.period_from
                self._period_to = result.period_to
            self._set_period_label()
            self.recalculate_current_rows()
            msg = f"Расчет выполнен. Строк: {len(self._rows)}."
            if result.auto_matched_count:
                msg += f" Автоматически подобрано продуктов: {result.auto_matched_count}, проверь."
            if result.unmatched_count:
                msg += f" Не найдено продуктов: {result.unmatched_count}."
            self.show_message(msg)
        except Exception as e:
            self.show_error_message(str(e))

    def search_saved(self):
        period_from, period_to = self._selected_period()
        try:
            with self.get_session() as session:
                base_rows, saved_from, saved_to = self.service(session).load_saved_base_rows()
            if not base_rows:
                self.show_message("Сохраненный расчет не найден")
                self.display_rows([], "search")
                return

            if not self._today_dates_selected() and saved_from and saved_to and (period_from != saved_from or period_to != saved_to):
                answer = QMessageBox.question(
                    self,
                    "Период не совпадает",
                    f"Период Ср.Продаж и Период фильтра не совпадает. В БД рассчитанный период {self._format_period(saved_from, saved_to)}. Продолжить?",
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.No,
                )
                if answer != QMessageBox.Yes:
                    return

            self._base_rows = base_rows
            self._period_from = saved_from
            self._period_to = saved_to
            self._set_period_label()
            self._mode = "search"
            self.recalculate_current_rows()
            self.show_message("Сохраненный расчет загружен")
        except Exception as e:
            self.show_error_message(str(e))

    def reset_form(self):
        today = QDate.currentDate()
        self.ui.date_SalesFrom.setDate(today)
        self.ui.date_SalesTo.setDate(today)
        self.ui.spin_QuickOrd.setValue(3)
        self.ui.spin_SafeStock.setValue(5)
        if hasattr(self.ui, "spb_SuppPriceAge"):
            self.ui.spb_SuppPriceAge.setValue(3)
        self.ui.radio_VolNotNull.setChecked(False)
        self._base_rows = []
        self._rows = []
        self._period_from = None
        self._period_to = None
        self._selected_brand_values = None
        self._selected_family_values = None
        self._selected_product_ids = None
        self._refresh_filter_buttons(prune=True)
        self.display_rows([], "calc")
        self._set_period_label()
        self.show_message("Форма очищена")

    def save(self):
        try:
            self._commit_open_product_editors()
            if self._mode == "check":
                with self.get_session() as session:
                    service = self.service(session)
                    count = service.save_product_links(self._rows)
                    session.commit()

                # После сохранения сразу перечитываем проверку: сохраненные/совпавшие строки должны исчезнуть.
                with self.get_session() as session:
                    result = self.service(session).check_products()

                self.display_rows(result.rows, "check")
                if result.rows:
                    self.show_message(f"Сопоставления сохранены: {count}. Осталось изменений: {len(result.rows)}")
                else:
                    self.show_message(f"Сопоставления сохранены: {count}. Изменений не найдено")
                return

            if not self._period_from or not self._period_to:
                self.show_error_message("Сначала сделай расчет")
                return
            missing = [row for row in self._rows if not row.get("product_id") and not clean_multi_spaces(row.get("product_name"))]
            if missing:
                self.show_error_message("Есть строки без Product Name. Выбери продукт или впиши новый Product Name перед сохранением")
                return
            with self.get_session() as session:
                count = self.service(session).save_calculation(self._rows, self._period_from, self._period_to)
                session.commit()
            self.show_message(f"Расчет сохранен. Продуктов: {count}")
        except Exception as e:
            self.show_error_message(str(e))

    def save_excel(self):
        if not self._rows or self._mode == "check":
            self.show_error_message("Сначала сформируй расчет")
            return
        try:
            default_name = f"OrderPlanning_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
            file_path, _ = QFileDialog.getSaveFileName(self, "Сохранить Excel файл", str(BASE_DIR / default_name), "Excel Files (*.xlsx)")
            if not file_path:
                return
            rows = [dict(row) for row in self._rows]
            supplier_price_age_months = self.get_supplier_price_age_months()

            def do_export():
                with self.get_session() as session:
                    exporter = OrderPlanningExporter(session)
                    return exporter.export_report(
                        display_rows=rows,
                        output_path=file_path,
                        supplier_price_age_months=supplier_price_age_months,
                    )

            def done(output_path):
                QDesktopServices.openUrl(Path(output_path).as_uri())
                self.show_message("Excel файл сохранен")

            if not start_excel_export(self, do_export, on_finished=done, on_error=lambda text: self.show_error_message(str(text))):
                self.show_message("Excel файл уже формируется. Можно продолжать работать в программе.")
            else:
                self.show_message("Excel файл формируется в фоновом режиме. Можно продолжать работать в программе.")
        except Exception as e:
            self.show_error_message(str(e))

    # ------------------------------------------------------------------
    # Product combo editing
    # ------------------------------------------------------------------
    def start_product_edit(self, row: int, column: int):
        if self._updating_table:
            return
        product_col = self.CHECK_COLUMNS.index("product_name") if self._mode == "check" else self.CALC_COLUMNS.index("product_name")
        if column != product_col:
            return
        source_row = self._row_index_from_table_row(row)
        if source_row is None or source_row < 0 or source_row >= len(self._rows):
            return
        current_id = self._rows[source_row].get("product_id")
        combo = self.build_product_combo(current_id, self._rows[source_row])
        combo.activated.connect(lambda _=None, tr=row, sr=source_row, c=combo: self.finish_product_edit(tr, sr, c))
        if combo.lineEdit() is not None:
            combo.lineEdit().returnPressed.connect(lambda tr=row, sr=source_row, c=combo: self.finish_product_edit(tr, sr, c))
        self.table.setCellWidget(row, column, combo)
        combo.setFocus()
        QTimer.singleShot(0, combo.showPopup)

    def build_product_combo(self, selected_product_id: int | None, row_data: dict | None = None) -> QComboBox:
        combo = QComboBox()
        combo.setEditable(True)
        combo.setInsertPolicy(QComboBox.NoInsert)
        combo.setToolTip("Выберите продукт из базы или впишите новый Product Name вручную")
        combo.addItem("", None)

        row_data = row_data or {}
        brand_filter = clean_multi_spaces(row_data.get("brand") or row_data.get("sales_brand") or "")

        with self.get_session() as session:
            query = session.query(Product).filter(Product.name.isnot(None), Product.name != "")
            if brand_filter:
                query = query.filter(Product.brand == brand_filter)
            products = query.order_by(Product.name.asc()).all()

            # Если текущий выбранный продукт другого бренда, обязательно добавляем его в список,
            # чтобы при повторном открытии редактора значение не потерялось.
            selected_product = None
            if selected_product_id:
                selected_product = session.query(Product).filter(Product.id == selected_product_id).first()

            added_ids: set[int] = set()
            for product in products:
                combo.addItem(product.name, product.id)
                added_ids.add(int(product.id))

            if selected_product and int(selected_product.id) not in added_ids:
                combo.addItem(selected_product.name, selected_product.id)

        index = combo.findData(selected_product_id)
        if index >= 0:
            combo.setCurrentIndex(index)
        else:
            combo.setCurrentText(clean_multi_spaces(row_data.get("product_name") or ""))

        return combo

    def finish_product_edit(self, table_row: int, source_row: int, combo: QComboBox):
        product_id = combo.currentData()
        product_name = clean_multi_spaces(combo.currentText()).upper()
        if source_row < 0 or source_row >= len(self._rows):
            return

        # Editable combo: if user typed a name that is not an existing item,
        # keep it as a new product name. It will be created on Save.
        matched_existing_id = None
        if product_name:
            for index in range(combo.count()):
                item_text = clean_multi_spaces(combo.itemText(index)).upper()
                if item_text == product_name:
                    matched_existing_id = combo.itemData(index)
                    product_name = combo.itemText(index)
                    break

        if matched_existing_id not in (None, "", 0):
            product_id = int(matched_existing_id)
            with self.get_session() as session:
                product = session.query(Product).filter(Product.id == product_id).first()
                if product:
                    product_name = product.name
                    self._rows[source_row]["brand"] = product.brand
                    self._rows[source_row]["family"] = product.family
                    self._rows[source_row]["pack"] = product.pack
        elif not product_name:
            product_id = None
            product_name = ""
        else:
            product_id = None
            # For a manually typed new product keep source brand/pack/excise from sales DB.
            self._rows[source_row]["brand"] = self._rows[source_row].get("sales_brand") or self._rows[source_row].get("brand") or ""
            self._rows[source_row]["pack"] = self._rows[source_row].get("sales_pack") or self._rows[source_row].get("pack")

        self._rows[source_row]["product_id"] = product_id
        self._rows[source_row]["product_name"] = product_name
        self._rows[source_row]["is_auto_matched"] = False

        product_col = self.CHECK_COLUMNS.index("product_name") if self._mode == "check" else self.CALC_COLUMNS.index("product_name")
        self.table.removeCellWidget(table_row, product_col)
        self.display_rows(self._rows, self._mode)
        if self._mode in {"calc", "search"}:
            self._base_rows = [dict(r) for r in self._rows]
            self.recalculate_current_rows()

    def _commit_open_product_editors(self):
        product_col = self.CHECK_COLUMNS.index("product_name") if self._mode == "check" else self.CALC_COLUMNS.index("product_name")
        for table_row in range(self.table.rowCount()):
            widget = self.table.cellWidget(table_row, product_col)
            if isinstance(widget, QComboBox):
                source_row = self._row_index_from_table_row(table_row)
                if source_row is not None:
                    self.finish_product_edit(table_row, source_row, widget)

    # ------------------------------------------------------------------
    # Messages
    # ------------------------------------------------------------------
    def show_message(self, text: str):
        if hasattr(self.ui, "label_msg"):
            self.ui.label_msg.setText(text)
            self.ui.label_msg.setProperty("active", True)
            self.ui.label_msg.style().unpolish(self.ui.label_msg)
            self.ui.label_msg.style().polish(self.ui.label_msg)
            self.ui.label_msg.setVisible(True)
        else:
            QMessageBox.information(self, "Сообщение", text)

    def show_error_message(self, text: str):
        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Warning)
        msg.setWindowTitle("Ошибка")
        msg.setText(text if len(text) <= 500 else "Произошла ошибка. Подробности ниже.")
        if len(text) > 500:
            msg.setDetailedText(text)
        copy_btn = msg.addButton("Copy", QMessageBox.ActionRole)
        msg.addButton(QMessageBox.Ok)
        msg.exec()
        if msg.clickedButton() == copy_btn:
            QApplication.clipboard().setText(text)
