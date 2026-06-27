from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, time
from decimal import Decimal
from pathlib import Path
from typing import Any, Optional, Sequence

from sqlalchemy import and_, func
from sqlalchemy.orm import joinedload
from PySide6.QtCore import QFile, Qt, QDate, QThread, QTimer
from PySide6.QtGui import QColor, QDesktopServices
from PySide6.QtUiTools import QUiLoader
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QComboBox,
    QFileDialog,
    QMenu,
    QMessageBox,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from app.db.db import SessionLocal
from app.db.models import CustomerPriceCalculation, Product, Supplier
from app.exports.customer_cost_report_exporter import CustomerCostReportExporter
from app.workers.excel_export_worker import ExcelExportWorker
from app.ui.table_style import build_table_item, setup_data_table
from app.utils.checked_filter_dialog import CheckedFilterDialog, FilterOption


BASE_DIR = Path(__file__).resolve().parents[2]
CUSTOMER_COSTS_REPORTS_UI = BASE_DIR / "app" / "ui" / "windows" / "customer_costs_reports.ui"


@dataclass(frozen=True)
class ProductOption:
    id: int
    name: str
    brand: str
    family: str
    pack: Decimal | None


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


def _export_customer_cost_report_file(
    *,
    headers: list[str],
    rows: list[list[Any]],
    output_path: str,
) -> Path:
    return CustomerCostReportExporter().export_report(headers=headers, rows=rows, output_path=output_path)


class CustomerCostsReportsPage(QWidget):
    """Report page for saved customer cost calculations."""

    HEADERS = [
        "id",
        "Дата",
        "Менеджер",
        "Клиент",
        "Customer Product Name",
        "Our Product Name",
        "Pack",
        "Qty, pcs",
        "Volume, L",
        "Supplier",
        "Supplier Article",
        "Supplier Price, L",
        "Currency",
        "FX rate",
        "Cost Novo with VAT",
        "Full Cost Msk",
        "FX markup",
        "Transport",
        "Re-export",
        "Agent fee",
        "Has customs",
        "Via Novo",
        "Bank fee",
        "Customs fee",
        "Additional customs",
        "Storage",
        "Move Novo",
        "Move Msk",
        "Marking",
        "Excise duty",
        "Price date",
        "Comments",
    ]

    def __init__(self):
        super().__init__()
        self.ui = load_ui(CUSTOMER_COSTS_REPORTS_UI)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.ui)

        self.table = self.ui.table_ReportPreview
        self._updating_table = False
        self._updating_filters = False
        self._pending_changes: dict[int, dict[str, Any]] = {}
        self._pending_deletes: set[int] = set()
        self._row_ids: list[int] = []
        self._product_options: list[ProductOption] = []
        self._product_by_id: dict[int, ProductOption] = {}
        self._selected_brand_values: set[str] | None = None
        self._selected_family_values: set[str] | None = None
        self._selected_product_ids: set[int] | None = None
        self._excel_export_thread: QThread | None = None
        self._excel_export_worker: ExcelExportWorker | None = None
        self._export_button_text = ""

        self.setup_ui()
        self.setup_connections()
        self.reset_filters(initial=True)

    def get_session(self):
        return SessionLocal()

    def setup_ui(self):
        setup_data_table(self.table, sorting=True)
        self._setup_date_edits()
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        from app.utils.gui_table_actions import install_standard_table_context_menu
        install_standard_table_context_menu(self, self.table)
        self._export_button_text = self.ui.btn_ExportExcel.text()
        self._hide_report_header()

    def _setup_date_edits(self):
        for date_edit in (self.ui.line_Start_date, self.ui.line_End_date):
            date_edit.setCalendarPopup(True)
            date_edit.setDisplayFormat("dd.MM.yyyy")
            date_edit.setSpecialValueText("")

    def setup_connections(self):
        self.ui.cbo_Manager.currentIndexChanged.connect(self.on_manager_changed)
        self.ui.cbo_Customer.currentIndexChanged.connect(self.on_customer_changed)
        self.ui.btn_FilterBrand.clicked.connect(self.open_brand_filter)
        self.ui.btn_FilterProductFamily.clicked.connect(self.open_family_filter)
        self.ui.btn_FilterProduct.clicked.connect(self.open_product_filter)
        self.ui.btn_BuildReport.clicked.connect(self.build_report)
        self.ui.btn_Reset.clicked.connect(self.reset_filters)
        self.ui.btn_ExportExcel.clicked.connect(self.export_excel)
        self.ui.btn_Save.clicked.connect(self.save_changes)
        self.table.itemChanged.connect(self.on_item_changed)
        self.table.cellDoubleClicked.connect(self.start_product_cell_edit)

    def reset_filters(self, initial: bool = False):
        today = QDate.currentDate()
        self._updating_filters = True
        self.ui.line_Start_date.setDate(today)
        self.ui.line_End_date.setDate(today)
        self.ui.chb_CalcToday.setChecked(False)
        self._updating_filters = False

        self._pending_changes.clear()
        self._pending_deletes.clear()
        self._selected_brand_values = None
        self._selected_family_values = None
        self._selected_product_ids = None
        self.load_filter_values()
        self.clear_report_table()
        self.show_message("" if initial else "Фильтры сброшены")

    def load_filter_values(self):
        self.fill_managers()
        self.fill_customers()
        self._refresh_filter_buttons(prune=True)

    def on_manager_changed(self):
        if self._updating_filters:
            return
        self.fill_customers()
        self._refresh_filter_buttons(prune=True)
        self.clear_report_table()

    def on_customer_changed(self):
        if self._updating_filters:
            return
        self._refresh_filter_buttons(prune=True)
        self.clear_report_table()

    def _combo_value(self, combo: QComboBox) -> str | None:
        text = combo.currentText().strip()
        return None if not text or text == "-" else text

    def _fill_combo(self, combo: QComboBox, values: list[str], keep_current: bool = True):
        current = combo.currentText().strip() if keep_current else "-"
        combo.blockSignals(True)
        combo.clear()
        combo.addItem("-")
        combo.addItems(values)
        if current and current in values:
            combo.setCurrentText(current)
        else:
            combo.setCurrentIndex(0)
        combo.blockSignals(False)

    def _base_query(self, session, *, include_customer: bool = True, include_brand: bool = True, include_family: bool = True, include_product: bool = True):
        query = session.query(CustomerPriceCalculation).outerjoin(Product, CustomerPriceCalculation.product_id == Product.id)
        manager = self._combo_value(self.ui.cbo_Manager)
        customer = self._combo_value(self.ui.cbo_Customer) if include_customer else None
        brand_values = self._selected_brand_values if include_brand else None
        family_values = self._selected_family_values if include_family else None
        product_ids = self._selected_product_ids if include_product else None

        if manager:
            query = query.filter(CustomerPriceCalculation.manager_name == manager)
        if customer:
            query = query.filter(CustomerPriceCalculation.customer_name == customer)
        if brand_values is not None:
            query = query.filter(Product.brand.in_(list(brand_values)))
        if family_values is not None:
            query = query.filter(Product.family.in_(list(family_values)))
        if product_ids is not None:
            query = query.filter(Product.id.in_([int(value) for value in product_ids]))
        return query

    def fill_managers(self):
        try:
            with self.get_session() as session:
                rows = (
                    session.query(CustomerPriceCalculation.manager_name)
                    .filter(CustomerPriceCalculation.manager_name.isnot(None), CustomerPriceCalculation.manager_name != "")
                    .distinct()
                    .order_by(CustomerPriceCalculation.manager_name.asc())
                    .all()
                )
            self._fill_combo(self.ui.cbo_Manager, [r[0] for r in rows if r[0]], keep_current=True)
        except Exception as e:
            self.show_error_message(f"Ошибка при получении менеджеров: {e}")

    def fill_customers(self):
        try:
            with self.get_session() as session:
                query = session.query(CustomerPriceCalculation.customer_name)
                manager = self._combo_value(self.ui.cbo_Manager)
                if manager:
                    query = query.filter(CustomerPriceCalculation.manager_name == manager)
                rows = (
                    query.filter(CustomerPriceCalculation.customer_name.isnot(None), CustomerPriceCalculation.customer_name != "")
                    .distinct()
                    .order_by(CustomerPriceCalculation.customer_name.asc())
                    .all()
                )
            self._fill_combo(self.ui.cbo_Customer, [r[0] for r in rows if r[0]], keep_current=True)
        except Exception as e:
            self.show_error_message(f"Ошибка при получении клиентов: {e}")

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
        self.clear_report_table()

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
        self.clear_report_table()

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
        self.clear_report_table()

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
                query = self._base_query(session, include_brand=False, include_family=True, include_product=True)
                rows = (
                    query.with_entities(Product.brand)
                    .filter(Product.brand.isnot(None), Product.brand != "")
                    .distinct()
                    .order_by(Product.brand.asc())
                    .all()
                )
            brands = [r[0] for r in rows if r[0]]
            return [FilterOption(key=brand, label=brand, search_text=brand) for brand in brands]
        except Exception as e:
            self.show_error_message(f"Ошибка при получении брендов: {e}")
            return []

    def _get_family_filter_options(self) -> list[FilterOption]:
        try:
            with self.get_session() as session:
                query = self._base_query(session, include_brand=True, include_family=False, include_product=True)
                rows = (
                    query.with_entities(Product.family)
                    .filter(Product.family.isnot(None), Product.family != "")
                    .distinct()
                    .order_by(Product.family.asc())
                    .all()
                )
            families = [r[0] for r in rows if r[0]]
            return [FilterOption(key=family, label=family, search_text=family) for family in families]
        except Exception as e:
            self.show_error_message(f"Ошибка при получении Product Family: {e}")
            return []

    def _get_product_filter_options(self) -> list[FilterOption]:
        try:
            with self.get_session() as session:
                query = self._base_query(session, include_brand=True, include_family=True, include_product=False)
                rows = (
                    query.with_entities(Product.id, Product.name, Product.brand, Product.family, Product.pack)
                    .filter(Product.id.isnot(None), Product.name.isnot(None), Product.name != "")
                    .distinct()
                    .order_by(Product.name.asc())
                    .all()
                )
            options: list[FilterOption] = []
            for product_id, name, brand, family, pack in rows:
                if product_id is None or not self._clean_text(name):
                    continue
                options.append(
                    FilterOption(
                        key=int(product_id),
                        label=self._clean_text(name),
                        search_text=self._product_filter_search_text(product_id, name, brand, family, pack),
                    )
                )
            return options
        except Exception as e:
            self.show_error_message(f"Ошибка при получении продуктов: {e}")
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
            available_brands = {option.key for option in self._get_brand_filter_options()}
            if self._selected_brand_values is not None:
                self._selected_brand_values = {value for value in self._selected_brand_values if value in available_brands}

            available_families = {option.key for option in self._get_family_filter_options()}
            if self._selected_family_values is not None:
                self._selected_family_values = {value for value in self._selected_family_values if value in available_families}

            available_product_ids = {int(option.key) for option in self._get_product_filter_options()}
            if self._selected_product_ids is not None:
                self._selected_product_ids = {int(value) for value in self._selected_product_ids if int(value) in available_product_ids}
        except Exception as e:
            self.show_error_message(f"Ошибка обновления фильтров: {e}")

    def _clean_text(self, value: object) -> str:
        return " ".join(str(value or "").split())

    def _product_filter_search_text(self, product_id: object, name: object, brand: object, family: object, pack: object) -> str:
        return " ".join(
            part
            for part in [
                str(product_id or ""),
                self._clean_text(name),
                self._clean_text(brand),
                self._clean_text(family),
                str(pack or ""),
            ]
            if part
        )

    def _qdate_to_date(self, qdate: QDate, field_name: str) -> date:
        if not qdate.isValid():
            raise ValueError(f"Некорректная дата в поле {field_name}. Нужно ДД.ММ.ГГГГ")
        return date(qdate.year(), qdate.month(), qdate.day())

    def _date_period(self) -> tuple[datetime, datetime]:
        if self.ui.chb_CalcToday.isChecked():
            start = end = date.today()
        else:
            start = self._qdate_to_date(self.ui.line_Start_date.date(), "Дата с")
            end = self._qdate_to_date(self.ui.line_End_date.date(), "Дата по")
            if start > end:
                raise ValueError("Дата начала периода не может быть больше даты окончания периода")
        return datetime.combine(start, time.min), datetime.combine(end, time.max)

    def _get_customer_product_name(self, calc: CustomerPriceCalculation) -> str:
        return getattr(calc, "customer_product_name", "") or ""

    def build_report(self):
        try:
            start_dt, end_dt = self._date_period()
            self._load_all_products()
            with self.get_session() as session:
                query = self._base_query(session)
                query = query.options(
                    joinedload(CustomerPriceCalculation.product),
                    joinedload(CustomerPriceCalculation.supplier),
                )
                query = query.filter(CustomerPriceCalculation.calc_date >= start_dt, CustomerPriceCalculation.calc_date <= end_dt)
                rows = query.order_by(CustomerPriceCalculation.calc_date.desc(), CustomerPriceCalculation.id.asc()).all()
            self._populate_table(rows)
            self.show_message(f"Сформировано строк: {len(rows)}")
        except Exception as e:
            self.show_error_message(str(e))

    def _load_all_products(self):
        with self.get_session() as session:
            products = session.query(Product).order_by(Product.name.asc()).all()
        self._product_options = [
            ProductOption(p.id, p.name or "", p.brand or "", p.family or "", p.pack)
            for p in products
        ]
        self._product_by_id = {p.id: p for p in self._product_options}

    def _row_values(self, calc: CustomerPriceCalculation) -> list[Any]:
        product = calc.product
        supplier = calc.supplier
        return [
            calc.id,
            calc.calc_date,
            calc.manager_name,
            calc.customer_name,
            self._get_customer_product_name(calc),
            product.name if product else "",
            calc.pack,
            calc.qty_pcs,
            calc.volume_l,
            supplier.name if supplier else "",
            calc.supplier_article,
            calc.supplier_price,
            calc.currency_code,
            calc.fx_rate_used,
            calc.cost_novo_wvat,
            calc.full_cost_msk,
            calc.fx_markup_used,
            calc.transport_used,
            calc.reexport_used,
            calc.agent_fee_used,
            calc.has_customs_used,
            calc.via_novo_used,
            calc.bank_fee_used,
            calc.customs_fee_used,
            calc.additional_customs_used,
            calc.storage_used,
            calc.move_novo_used,
            calc.move_msk_used,
            calc.marking_used,
            calc.is_excise_used,
            calc.price_date_used,
            calc.comments,
        ]

    def _show_report_header(self):
        self.table.setColumnCount(len(self.HEADERS))
        self.table.setHorizontalHeaderLabels(self.HEADERS)
        self.table.setColumnHidden(0, True)
        self.table.horizontalHeader().setVisible(True)

    def _hide_report_header(self):
        self.table.setSortingEnabled(False)
        self.table.clearContents()
        self.table.setRowCount(0)
        self.table.setColumnCount(0)
        self.table.horizontalHeader().setVisible(False)
        self.table.setSortingEnabled(True)

    def _populate_table(self, rows: list[CustomerPriceCalculation]):
        self._updating_table = True
        self.table.setSortingEnabled(False)
        self._show_report_header()
        self.table.clearContents()
        self.table.setRowCount(len(rows))
        self._row_ids = []
        self._pending_changes.clear()
        self._pending_deletes.clear()

        editable_headers = {"Менеджер", "Клиент"}
        for row_idx, calc in enumerate(rows):
            values = self._row_values(calc)
            row_id = int(calc.id)
            self._row_ids.append(row_id)
            for col_idx, value in enumerate(values):
                header = self.HEADERS[col_idx]
                if header == "Our Product Name":
                    item = self._build_product_display_item(value, calc.product_id)
                else:
                    editable = header in editable_headers
                    item = build_table_item(value, editable=editable, align_left=header in {"Менеджер", "Клиент", "Customer Product Name", "Supplier", "Comments"})
                if header in {"Дата", "Price date"} and isinstance(value, datetime):
                    item.setText(value.strftime("%d.%m.%Y"))
                if isinstance(value, bool):
                    item.setText("Да" if value else "Нет")
                self.table.setItem(row_idx, col_idx, item)
        self.table.setColumnHidden(0, True)
        self.table.resizeColumnsToContents()
        self.table.setSortingEnabled(True)
        self._updating_table = False

    def _build_product_display_item(self, product_name: object, product_id: int | None) -> QTableWidgetItem:
        item = build_table_item(product_name, editable=False, align_left=True)
        item.setData(Qt.UserRole, product_id)
        item.setToolTip("Дважды кликните, чтобы выбрать другой продукт")
        return item

    def start_product_cell_edit(self, row: int, column: int):
        if self._updating_table or column != self.HEADERS.index("Our Product Name"):
            return

        row_id = self._id_for_visual_row(row)
        if row_id is None:
            return

        item = self.table.item(row, column)
        current_product_id = item.data(Qt.UserRole) if item else None
        try:
            current_product_id = int(current_product_id)
        except (TypeError, ValueError):
            current_product_id = None

        combo = self._build_product_combo(current_product_id)
        combo.activated.connect(lambda _=None, rid=row_id, cb=combo: self.finish_product_edit(rid, cb))
        self.table.setCellWidget(row, column, combo)
        combo.setFocus()
        QTimer.singleShot(0, combo.showPopup)

    def _build_product_combo(self, selected_product_id: int | None) -> QComboBox:
        combo = QComboBox(self.table)
        combo.setEditable(False)
        for option in self._product_options:
            combo.addItem(option.name, option.id)

        if selected_product_id is not None:
            idx = combo.findData(selected_product_id)
            if idx >= 0:
                combo.setCurrentIndex(idx)

        return combo

    def finish_product_edit(self, row_id: int, combo: QComboBox):
        if self._updating_table:
            return

        product_id = combo.currentData()
        try:
            product_id = int(product_id)
        except (TypeError, ValueError):
            product_id = None

        row = self._find_row_by_id(row_id)
        product_col = self.HEADERS.index("Our Product Name")
        if row is None or product_id is None:
            return

        option = self._product_by_id.get(product_id)
        product_name = option.name if option else combo.currentText()

        self._updating_table = True
        sorting_enabled = self.table.isSortingEnabled()
        self.table.setSortingEnabled(False)
        self.table.removeCellWidget(row, product_col)
        self.table.setItem(row, product_col, self._build_product_display_item(product_name, product_id))

        if option:
            pack_item = self.table.item(row, self.HEADERS.index("Pack"))
            if pack_item:
                pack_item.setText(str(option.pack or "").replace(".", ","))

        self.table.resizeColumnsToContents()
        self.table.setSortingEnabled(sorting_enabled)
        self._updating_table = False

        self._pending_changes.setdefault(row_id, {})["product_id"] = product_id

    def on_item_changed(self, item: QTableWidgetItem):
        if self._updating_table or item is None:
            return
        row_id = self._id_for_visual_row(item.row())
        if row_id is None:
            return
        header = self.HEADERS[item.column()]
        if header == "Менеджер":
            self._pending_changes.setdefault(row_id, {})["manager_name"] = item.text().strip() or None
        elif header == "Клиент":
            self._pending_changes.setdefault(row_id, {})["customer_name"] = item.text().strip() or None

    def _id_for_visual_row(self, row: int) -> int | None:
        item = self.table.item(row, 0)
        if not item or not item.text().strip():
            return None
        try:
            return int(item.text())
        except ValueError:
            return None

    def _find_row_by_id(self, row_id: int) -> int | None:
        for row in range(self.table.rowCount()):
            if self._id_for_visual_row(row) == row_id:
                return row
        return None

    def _selected_row_ids(self) -> set[int]:
        ids: set[int] = set()
        for index in self.table.selectionModel().selectedRows():
            row_id = self._id_for_visual_row(index.row())
            if row_id is not None:
                ids.add(row_id)
        if not ids:
            for item in self.table.selectedItems():
                row_id = self._id_for_visual_row(item.row())
                if row_id is not None:
                    ids.add(row_id)
        return ids

    def _cell_text(self, row: int, col: int) -> str:
        if self.HEADERS[col] == "Our Product Name":
            combo = self.table.cellWidget(row, col)
            if combo:
                return combo.currentText()
        item = self.table.item(row, col)
        return item.text() if item else ""

    def save_changes(self):
        if not self._pending_changes and not self._pending_deletes:
            self.show_message("Нет изменений для сохранения")
            return
        try:
            with self.get_session() as session:
                if self._pending_deletes:
                    session.query(CustomerPriceCalculation).filter(CustomerPriceCalculation.id.in_(self._pending_deletes)).delete(synchronize_session=False)
                for row_id, changes in self._pending_changes.items():
                    if row_id in self._pending_deletes:
                        continue
                    row = session.query(CustomerPriceCalculation).filter(CustomerPriceCalculation.id == row_id).first()
                    if not row:
                        continue
                    for field, value in changes.items():
                        if field == "product_id":
                            row.product_id = int(value)
                            product = session.query(Product).filter(Product.id == int(value)).first()
                            if product:
                                row.pack = product.pack
                        elif field in {"manager_name", "customer_name"}:
                            setattr(row, field, value)
                session.commit()
            deleted = len(self._pending_deletes)
            changed = len(self._pending_changes)
            self._pending_changes.clear()
            self._pending_deletes.clear()
            self.rebuild_current_report()
            self.load_filter_values()
            self.show_message(f"Сохранено. Изменено строк: {changed}. Удалено строк: {deleted}.")
        except Exception as e:
            self.show_error_message(f"Ошибка сохранения: {e}")

    def rebuild_current_report(self):
        self.build_report()

    def clear_report_table(self):
        self._updating_table = True
        self._hide_report_header()
        self._row_ids = []
        self._pending_changes.clear()
        self._pending_deletes.clear()
        self._updating_table = False

    def export_excel(self):
        if self._excel_export_thread is not None:
            self.show_message("Excel файл уже формируется. Дождись окончания экспорта.")
            return

        rows = self._collect_export_rows()
        if not rows:
            self.show_error_message("Сначала сформируй отчет")
            return
        default_name = f"CustCostReport_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
        path, _ = QFileDialog.getSaveFileName(self, "Сохранить отчет", str(Path.home() / default_name), "Excel files (*.xlsx)")
        if not path:
            return
        try:
            headers = [h for h in self.HEADERS if h != "id"]
            self._start_excel_export(headers=headers, rows=[list(row) for row in rows], output_path=path)
        except Exception as e:
            self.show_error_message(f"Ошибка экспорта в Excel: {e}")

    def _start_excel_export(self, *, headers: list[str], rows: list[list[Any]], output_path: str) -> None:
        self.ui.btn_ExportExcel.setEnabled(False)
        self.ui.btn_ExportExcel.setText("Формируется...")
        self.show_message("Excel файл формируется в фоновом режиме. Можно продолжать работать в программе.")

        self._excel_export_thread = QThread(self)
        self._excel_export_worker = ExcelExportWorker(
            _export_customer_cost_report_file,
            headers=headers,
            rows=rows,
            output_path=output_path,
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
        self.show_message(f"Файл сохранен: {path}")
        QDesktopServices.openUrl(path.as_uri())

    def _on_excel_export_error(self, error_text: str) -> None:
        self._finish_excel_export_ui()
        self.show_error_message(f"Ошибка экспорта в Excel: {error_text}")

    def _clear_excel_export_refs(self) -> None:
        self._excel_export_thread = None
        self._excel_export_worker = None

    def _collect_export_rows(self) -> list[list[Any]]:
        rows: list[list[Any]] = []
        for row in range(self.table.rowCount()):
            row_id = self._id_for_visual_row(row)
            if row_id in self._pending_deletes:
                continue
            values = []
            for col, header in enumerate(self.HEADERS):
                if header == "id":
                    continue
                values.append(self._cell_text(row, col))
            rows.append(values)
        return rows

    def show_message(self, text: str):
        if hasattr(self.ui, "label_msg"):
            self.ui.label_msg.setText(text)

    def show_error_message(self, text: str):
        self.show_message(text)
        QMessageBox.critical(self, "Ошибка", text)
