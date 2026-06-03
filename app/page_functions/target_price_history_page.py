from __future__ import annotations

from datetime import date, datetime, time
from decimal import Decimal
from pathlib import Path
from typing import Any

from sqlalchemy.orm import aliased, joinedload
from PySide6.QtCore import QFile, QDate, Qt
from PySide6.QtGui import QDesktopServices
from PySide6.QtUiTools import QUiLoader
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QComboBox,
    QFileDialog,
    QMessageBox,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from app.db.db import SessionLocal
from app.db.models import Product, Supplier, TargetPriceCalculation
from app.exports.excel_column_format import (
    BOOL_HEADERS,
    DATE_HEADERS,
    INTEGER_NUMERIC_HEADERS,
    NUMERIC_HEADERS,
    excel_value_by_header,
    normalize_header,
)
from app.exports.target_price_history_exporter import TargetPriceHistoryExporter
from app.ui.table_style import setup_data_table
from app.utils.text import clean_multi_spaces
from app.workers.excel_export_worker import start_excel_export


BASE_DIR = Path(__file__).resolve().parents[2]
TARGET_PRICE_HISTORY_UI = BASE_DIR / "app" / "ui" / "windows" / "target_price_history.ui"
PRICE_HISTORY_UI = BASE_DIR / "app" / "ui" / "windows" / "price_history.ui"


class SortableTableWidgetItem(QTableWidgetItem):
    def __lt__(self, other):
        if isinstance(other, QTableWidgetItem):
            left = self.data(Qt.UserRole + 1)
            right = other.data(Qt.UserRole + 1)
            if left is not None and right is not None:
                try:
                    return left < right
                except Exception:
                    pass
        return super().__lt__(other)


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


def _export_target_price_history_file(*, headers: list[str], rows: list[list[Any]], output_path: str) -> Path:
    return TargetPriceHistoryExporter().export_report(headers=headers, rows=rows, output_path=output_path)


class TargetPriceHistoryPage(QWidget):
    """Report page for saved target price calculations."""

    HEADERS = [
        "id",
        "Calc date",
        "Supplier",
        "Brand",
        "Our Product Name",
        "Pack",
        "Target Price, L",
        "Target Price, pack",
        "Currency",
        "FX rate",
        "Supplier (donor)",
        "Full Cost Msk",
        "Cost Novo with VAT",
        "Currency (donor)",
        "FX rate (donor)",
        "Price date",
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
        "VAT",
        "Money",
    ]

    def __init__(self):
        super().__init__()
        ui_path = TARGET_PRICE_HISTORY_UI if TARGET_PRICE_HISTORY_UI.exists() else PRICE_HISTORY_UI
        self.ui = load_ui(ui_path)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.ui)

        self.table = self.ui.table
        self._updating_filters = False
        self._updating_table = False
        self._date_filter_changed = False
        self._pending_changes: dict[int, dict[str, Any]] = {}
        self._pending_deletes: set[int] = set()
        self._deleted_row_snapshots: list[dict[str, Any]] = []
        self._table_row_ids: list[int] = []
        self._excel_export_thread = None
        self._excel_export_worker = None

        self.setup_ui()
        self.setup_connections()
        self.reset_filters(initial=True)

    def get_session(self):
        return SessionLocal()

    def setup_ui(self):
        setup_data_table(self.table, sorting=True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self._setup_date_edits()

        # These controls belong to the source price_history.ui, but are not used here.
        for name in ("line_TableName", "btn_AddLine", "btn_DownFile", "btn_Import"):
            widget = getattr(self.ui, name, None)
            if widget is not None:
                widget.setVisible(False)

        from app.utils.gui_table_actions import install_standard_table_context_menu

        install_standard_table_context_menu(self, self.table)
        self.clear_message()

    def _setup_date_edits(self):
        for date_edit in (self.ui.line_Start_date, self.ui.line_End_date):
            date_edit.setCalendarPopup(True)
            date_edit.setDisplayFormat("dd.MM.yyyy")
            date_edit.setSpecialValueText("")

    def setup_connections(self):
        self.ui.btn_Search.clicked.connect(self.build_report)
        self.ui.btn_SaveExcel.clicked.connect(self.save_excel)
        self.ui.btn_Save.clicked.connect(self.apply_pending_changes)

        self.ui.line_SupplName.currentTextChanged.connect(self.on_supplier_changed)
        self.ui.line_Brand.currentTextChanged.connect(self.on_product_filter_changed)
        self.ui.line_Prod_Fam.currentTextChanged.connect(self.on_family_changed)
        self.ui.line_Prod_name.currentTextChanged.connect(self.clear_report_table)
        if hasattr(self.ui, "line_NameSearch"):
            self.ui.line_NameSearch.textChanged.connect(self.clear_report_table)

        self.ui.line_Start_date.dateChanged.connect(self.on_date_filter_changed)
        self.ui.line_End_date.dateChanged.connect(self.on_date_filter_changed)
        if hasattr(self.ui, "chb_CalcToday"):
            self.ui.chb_CalcToday.stateChanged.connect(self.on_calc_today_changed)

    def on_date_filter_changed(self):
        if self._updating_filters:
            return
        self._date_filter_changed = True
        self.clear_report_table()

    def on_calc_today_changed(self):
        if self._updating_filters:
            return
        self.clear_report_table()

    def reset_filters(self, initial: bool = False):
        today = QDate.currentDate()
        self._updating_filters = True
        self.ui.line_Start_date.setDate(today)
        self.ui.line_End_date.setDate(today)
        self._date_filter_changed = False
        if hasattr(self.ui, "chb_CalcToday"):
            self.ui.chb_CalcToday.setChecked(False)
        self._pending_changes.clear()
        self._pending_deletes.clear()
        self._deleted_row_snapshots.clear()
        self.load_filter_values()
        self._updating_filters = False
        self.clear_report_table()
        self.show_message("" if initial else "Фильтры сброшены")

    def load_filter_values(self):
        self.fill_suppliers()
        self.fill_brands()
        self.fill_families()
        self.fill_products()

    def on_supplier_changed(self):
        if self._updating_filters:
            return
        self.fill_brands()
        self.fill_families()
        self.fill_products()
        self.clear_report_table()

    def on_product_filter_changed(self):
        if self._updating_filters:
            return
        self.fill_families()
        self.fill_products()
        self.clear_report_table()

    def on_family_changed(self):
        if self._updating_filters:
            return
        self.fill_products()
        self.clear_report_table()

    def _combo_value(self, combo: QComboBox) -> str | None:
        text = combo.currentText().strip()
        return None if not text or text == "-" else text

    def _fill_combo(self, combo: QComboBox, values: list[str], keep_current: bool = True):
        current = combo.currentText().strip() if keep_current else "-"
        combo.blockSignals(True)
        combo.clear()
        combo.addItem("-")
        combo.addItems(sorted(values))
        if current and current in values:
            combo.setCurrentText(current)
        else:
            combo.setCurrentIndex(0)
        combo.blockSignals(False)

    def _base_query(self, session, *, include_supplier: bool = True, include_brand: bool = True, include_family: bool = True, include_product: bool = True):
        TargetSupplier = aliased(Supplier)
        query = (
            session.query(TargetPriceCalculation)
            .join(Product, TargetPriceCalculation.product_id == Product.id)
            .join(TargetSupplier, TargetPriceCalculation.target_supplier_id == TargetSupplier.id)
        )

        supplier_name = self._combo_value(self.ui.line_SupplName) if include_supplier else None
        brand = self._combo_value(self.ui.line_Brand) if include_brand else None
        family = self._combo_value(self.ui.line_Prod_Fam) if include_family else None
        product_name = self._combo_value(self.ui.line_Prod_name) if include_product else None
        name_search = clean_multi_spaces(self.ui.line_NameSearch.text()) if hasattr(self.ui, "line_NameSearch") else ""

        if supplier_name:
            query = query.filter(TargetSupplier.name == supplier_name)
        if brand:
            query = query.filter(Product.brand == brand)
        if family:
            query = query.filter(Product.family == family)
        if product_name:
            query = query.filter(Product.name == product_name)
        if name_search:
            query = query.filter(Product.name.ilike(f"%{name_search}%"))
        return query

    def fill_suppliers(self):
        try:
            with self.get_session() as session:
                rows = (
                    session.query(Supplier.name)
                    .join(TargetPriceCalculation, TargetPriceCalculation.target_supplier_id == Supplier.id)
                    .filter(Supplier.name.isnot(None), Supplier.name != "")
                    .distinct()
                    .order_by(Supplier.name.asc())
                    .all()
                )
            self._fill_combo(self.ui.line_SupplName, [r[0] for r in rows if r[0]], keep_current=True)
        except Exception as e:
            self.show_error_message(f"Ошибка при получении поставщиков: {e}")

    def fill_brands(self):
        try:
            with self.get_session() as session:
                query = self._base_query(session, include_brand=False)
                rows = (
                    query.with_entities(Product.brand)
                    .filter(Product.brand.isnot(None), Product.brand != "")
                    .distinct()
                    .order_by(Product.brand.asc())
                    .all()
                )
            self._fill_combo(self.ui.line_Brand, [r[0] for r in rows if r[0]], keep_current=True)
        except Exception as e:
            self.show_error_message(f"Ошибка при получении брендов: {e}")

    def fill_families(self):
        try:
            with self.get_session() as session:
                query = self._base_query(session, include_family=False)
                rows = (
                    query.with_entities(Product.family)
                    .filter(Product.family.isnot(None), Product.family != "")
                    .distinct()
                    .order_by(Product.family.asc())
                    .all()
                )
            self._fill_combo(self.ui.line_Prod_Fam, [r[0] for r in rows if r[0]], keep_current=True)
        except Exception as e:
            self.show_error_message(f"Ошибка при получении Product Family: {e}")

    def fill_products(self):
        try:
            with self.get_session() as session:
                query = self._base_query(session, include_product=False)
                rows = (
                    query.with_entities(Product.name)
                    .filter(Product.name.isnot(None), Product.name != "")
                    .distinct()
                    .order_by(Product.name.asc())
                    .all()
                )
            self._fill_combo(self.ui.line_Prod_name, [r[0] for r in rows if r[0]], keep_current=True)
        except Exception as e:
            self.show_error_message(f"Ошибка при получении продуктов: {e}")

    def _qdate_to_date(self, qdate: QDate, field_name: str) -> date:
        if not qdate.isValid():
            raise ValueError(f"Некорректная дата в поле {field_name}. Нужно ДД.ММ.ГГГГ")
        return date(qdate.year(), qdate.month(), qdate.day())

    def _date_period(self) -> tuple[datetime, datetime] | None:
        # If chb_CalcToday is checked, the report is forced to today regardless
        # of selected dates. If it is unchecked and the user did not touch dates,
        # dates are not used as a filter.
        if hasattr(self.ui, "chb_CalcToday") and self.ui.chb_CalcToday.isChecked():
            start = end = date.today()
        elif self._date_filter_changed:
            start = self._qdate_to_date(self.ui.line_Start_date.date(), "Дата с")
            end = self._qdate_to_date(self.ui.line_End_date.date(), "Дата по")
            if start > end:
                raise ValueError("Дата начала периода не может быть больше даты окончания периода")
        else:
            return None
        return datetime.combine(start, time.min), datetime.combine(end, time.max)

    def get_rows_from_db(self) -> list[TargetPriceCalculation]:
        period = self._date_period()
        with self.get_session() as session:
            query = self._base_query(session)
            query = query.options(
                joinedload(TargetPriceCalculation.product),
                joinedload(TargetPriceCalculation.target_supplier),
                joinedload(TargetPriceCalculation.donor_supplier),
            )
            if period is not None:
                start_dt, end_dt = period
                query = query.filter(TargetPriceCalculation.calc_date >= start_dt, TargetPriceCalculation.calc_date <= end_dt)
            return query.order_by(TargetPriceCalculation.calc_date.desc(), TargetPriceCalculation.id.asc()).all()

    def build_report(self):
        try:
            rows = self.get_rows_from_db()
            self._populate_table(rows)
            self.show_message(f"Сформировано строк: {len(rows)}")
        except Exception as e:
            self.show_error_message(str(e))

    @staticmethod
    def _calc_attr(calc: TargetPriceCalculation, *names: str, default: Any = "") -> Any:
        for name in names:
            if hasattr(calc, name):
                value = getattr(calc, name)
                if value is not None:
                    return value
        return default

    def _row_values(self, calc: TargetPriceCalculation) -> list[Any]:
        product = calc.product
        target_supplier = calc.target_supplier
        donor_supplier = calc.donor_supplier
        return [
            calc.id,
            calc.calc_date,
            target_supplier.name if target_supplier else "",
            product.brand if product else "",
            product.name if product else "",
            product.pack if product else None,
            calc.target_price_l,
            calc.target_price_pack,
            calc.currency_code,
            calc.fx_rate_used,
            donor_supplier.name if donor_supplier else "",
            calc.full_cost_msk_source,
            self._calc_attr(calc, "cost_novo_wvat", "cost_novo_wvat_recalculated"),
            self._calc_attr(calc, "donor_currency_code"),
            self._calc_attr(calc, "donor_fx_rate_used"),
            calc.price_date_used,
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
            calc.vat_used,
            calc.money_used,
        ]

    def _show_report_header(self):
        self.table.setColumnCount(len(self.HEADERS))
        self.table.setHorizontalHeaderLabels(self.HEADERS)
        self.table.setColumnHidden(0, True)
        self.table.horizontalHeader().setVisible(True)

    def clear_report_table(self):
        self.table.setSortingEnabled(False)
        self.table.clearContents()
        self.table.setRowCount(0)
        self.table.setColumnCount(0)
        self.table.horizontalHeader().setVisible(False)
        self.table.setSortingEnabled(True)
        self._table_row_ids = []
        self._pending_changes.clear()
        self._pending_deletes.clear()
        self._deleted_row_snapshots.clear()
        self.clear_message()

    def _sort_value(self, value: Any) -> Any:
        if isinstance(value, Decimal):
            return float(value)
        if isinstance(value, datetime):
            return value
        if isinstance(value, date):
            return datetime(value.year, value.month, value.day)
        if isinstance(value, bool):
            return 1 if value else 0
        if value is None:
            return ""
        text = str(value)
        return text.casefold()

    @staticmethod
    def _format_display_date(value: Any) -> str:
        parsed = excel_value_by_header("Calc date", value)
        if isinstance(parsed, datetime):
            return parsed.strftime("%d.%m.%Y")
        if isinstance(parsed, date):
            return parsed.strftime("%d.%m.%Y")
        return "" if parsed is None else str(parsed)

    def _display_text(self, header: str, value: Any) -> str:
        base = normalize_header(header)
        if value is None:
            return ""
        if base in DATE_HEADERS:
            return self._format_display_date(value)
        prepared = excel_value_by_header(header, value)
        if prepared is None:
            return ""
        if base in BOOL_HEADERS:
            return str(prepared)
        if base in INTEGER_NUMERIC_HEADERS:
            return "" if prepared == "" else str(prepared)
        if base in NUMERIC_HEADERS:
            if prepared == "":
                return ""
            try:
                return f"{float(prepared):.2f}".replace(".", ",")
            except (TypeError, ValueError):
                return str(prepared)
        return str(prepared)

    def _build_item(self, header: str, value: Any, row_id: int, *, editable: bool = False, align_left: bool = False) -> QTableWidgetItem:
        item = SortableTableWidgetItem(self._display_text(header, value))
        flags = Qt.ItemIsEnabled | Qt.ItemIsSelectable
        if editable:
            flags |= Qt.ItemIsEditable
        item.setFlags(flags)
        item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter if align_left else Qt.AlignCenter)
        item.setData(Qt.UserRole, row_id)
        item.setData(Qt.UserRole + 1, self._sort_value(value))
        item.setData(Qt.UserRole + 2, value)
        return item

    def _populate_table(self, rows: list[TargetPriceCalculation]):
        self._updating_table = True
        self.table.setSortingEnabled(False)
        self._show_report_header()
        self.table.clearContents()
        self.table.setRowCount(len(rows))
        self._table_row_ids = []
        self._pending_changes.clear()
        self._pending_deletes.clear()
        self._deleted_row_snapshots.clear()

        left_headers = {"Supplier", "Our Product Name", "Brand", "Supplier (donor)", "Currency", "Currency (donor)"}
        for row_idx, calc in enumerate(rows):
            row_id = int(calc.id)
            self._table_row_ids.append(row_id)
            values = self._row_values(calc)
            for col_idx, value in enumerate(values):
                header = self.HEADERS[col_idx]
                self.table.setItem(
                    row_idx,
                    col_idx,
                    self._build_item(header, value, row_id, editable=False, align_left=header in left_headers),
                )

        self.table.resizeColumnsToContents()
        self.table.setSortingEnabled(True)
        self._updating_table = False

    def apply_pending_changes(self):
        if not self._pending_deletes:
            self.show_message("Нет изменений для применения")
            return
        try:
            with self.get_session() as session:
                ids = [int(row_id) for row_id in self._pending_deletes if row_id is not None]
                if ids:
                    session.query(TargetPriceCalculation).filter(TargetPriceCalculation.id.in_(ids)).delete(synchronize_session=False)
                session.commit()
            deleted = len(self._pending_deletes)
            self._pending_deletes.clear()
            self._deleted_row_snapshots.clear()
            self.build_report()
            self.show_message(f"Удалено строк: {deleted}")
        except Exception as e:
            self.show_error_message(f"Ошибка сохранения в базу данных: {e}")

    def _cell_value_for_export(self, row: int, col: int) -> Any:
        item = self.table.item(row, col)
        if item is None:
            return ""
        raw_value = item.data(Qt.UserRole + 2)
        if raw_value is not None:
            return raw_value
        return item.text()

    def _collect_export_rows(self) -> list[list[Any]]:
        rows: list[list[Any]] = []
        for row in range(self.table.rowCount()):
            item = self.table.item(row, 0)
            row_id = item.data(Qt.UserRole) if item else None
            if row_id in self._pending_deletes:
                continue
            values: list[Any] = []
            for col, header in enumerate(self.HEADERS):
                if header == "id":
                    continue
                values.append(self._cell_value_for_export(row, col))
            rows.append(values)
        return rows

    @staticmethod
    def _safe_filename(value: str) -> str:
        s = (value or "").strip()
        for ch in ["\\", "/", ":", "*", "?", '"', "<", ">", "|"]:
            s = s.replace(ch, "_")
        return s or "TargetPriceReport"

    def save_excel(self):
        try:
            if self.table.rowCount() == 0:
                self.show_message("Сначала сформируй отчет")
                return
            export_headers = [h for h in self.HEADERS if h != "id"]
            export_rows = self._collect_export_rows()
            timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            supplier = self._combo_value(self.ui.line_SupplName) or "all"
            default_name = f"Target price report_{self._safe_filename(supplier)}_{timestamp}.xlsx"
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "Сохранить Excel",
                str(BASE_DIR / default_name),
                "Excel files (*.xlsx)",
            )
            if not file_path:
                return
            if not file_path.lower().endswith(".xlsx"):
                file_path += ".xlsx"

            def done(output_path: object):
                path = Path(output_path)
                self.show_message("Файл сохранен")
                QDesktopServices.openUrl(path.as_uri())

            ok = start_excel_export(
                self,
                _export_target_price_history_file,
                kwargs={"headers": export_headers, "rows": export_rows, "output_path": file_path},
                on_finished=done,
                on_error=lambda text: self.show_error_message(str(text)),
                button=self.ui.btn_SaveExcel,
                busy_text="Формируется...",
                restore_text="Save Excel",
            )
            if ok:
                self.show_message("Excel файл формируется в фоновом режиме. Можно продолжать работать в программе.")
            else:
                self.show_message("Excel файл уже формируется. Можно продолжать работать в программе.")
        except Exception as e:
            self.show_error_message(str(e))

    def show_message(self, text: str):
        if hasattr(self.ui, "label_msg"):
            self.ui.label_msg.setText(text)
            self.ui.label_msg.setProperty("active", bool(text))
            self.ui.label_msg.style().unpolish(self.ui.label_msg)
            self.ui.label_msg.style().polish(self.ui.label_msg)
            self.ui.label_msg.setVisible(bool(text))

    def clear_message(self):
        self.show_message("")

    def show_error_message(self, text: str):
        self.show_message(text)
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
        copy_button.clicked.connect(lambda: QApplication.clipboard().setText(text))
        msg.exec_()
