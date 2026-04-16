from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

from PySide6.QtCore import QFile, Qt, QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QHBoxLayout,
    QHeaderView,
    QMenu,
    QMessageBox,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)
from PySide6.QtUiTools import QUiLoader

from app.db.db import SessionLocal
from app.db.models import (
    Product,
    TempIsImport,
    TempStockImport,
    TempSupplierOrdersImport,
)
from app.exports.product_stock_import_export import ProductStockImportExport
from app.services.product_matching import ProductMatchingService
from app.services.product_stock_run import ProductStockImportRun
from app.utils.batch import get_current_username
from app.utils.parsers import parse_loose_number
from app.utils.text import clean_multi_spaces
from app.ui.table_style import setup_data_table

BASE_DIR = Path(__file__).resolve().parents[2]
UI_PATH = BASE_DIR / "app" / "ui" / "windows" / "stock_supplier_orders.ui"


@dataclass(slots=True)
class ColumnDef:
    key: str
    header: str
    editable: bool = False
    kind: str = "text"  # text | product_combo | brand_combo | checkbox | indicator


MODE_STOCK = "stock"
MODE_ORDERS = "orders"
MODE_IS = "is"


COLUMN_DEFS: dict[str, list[ColumnDef]] = {
    MODE_STOCK: [
        ColumnDef("selected_product_id", "Our product", kind="product_combo"),
        ColumnDef("source_article", "Source article", editable=True),
        ColumnDef("source_sku", "SKU", editable=True),
        ColumnDef("source_product_name", "Source product", editable=True),
        ColumnDef("stock_qty", "Stock qty", editable=True),
        ColumnDef("markdown_qty", "Markdown qty", editable=True),
        ColumnDef("reserve_qty", "Reserve qty", editable=True),
        ColumnDef("has_lpc_warning", "LPC !", kind="indicator"),
        ColumnDef("lpc", "LPC", editable=True),
        ColumnDef("landed_cost", "Landed cost", editable=True),
        ColumnDef("distr_price", "Distr price", editable=True),
        ColumnDef("promo_price", "Promo price", editable=True),
        ColumnDef("new_product_name", "Product name (new)", editable=True),
        ColumnDef("new_brand", "Brand (new)", kind="brand_combo"),
        ColumnDef("new_pack", "Pack (new)", editable=True),
        ColumnDef("new_is_excise", "Excise (new)", kind="checkbox"),
    ],
    MODE_ORDERS: [
        ColumnDef("selected_product_id", "Our product", kind="product_combo"),
        ColumnDef("source_article", "Source article", editable=True),
        ColumnDef("source_product_name", "Source product", editable=True),
        ColumnDef("transit_qty", "Transit qty", editable=True),
        ColumnDef("order_qty", "Order qty", editable=True),
        ColumnDef("new_product_name", "Product name (new)", editable=True),
        ColumnDef("new_brand", "Brand (new)", kind="brand_combo"),
        ColumnDef("new_pack", "Pack (new)", editable=True),
        ColumnDef("new_is_excise", "Excise (new)", kind="checkbox"),
    ],
    MODE_IS: [
        ColumnDef("selected_product_id", "Our product", kind="product_combo"),
        ColumnDef("source_article", "Source article", editable=True),
        ColumnDef("source_product_name", "Source product", editable=True),
        ColumnDef("confirmed_qty", "Confirmed qty", editable=True),
        ColumnDef("remains_qty", "Remains qty", editable=True),
        ColumnDef("stock_qty", "Stock qty", editable=True),
        ColumnDef("new_product_name", "Product name (new)", editable=True),
        ColumnDef("new_brand", "Brand (new)", kind="brand_combo"),
        ColumnDef("new_pack", "Pack (new)", editable=True),
        ColumnDef("new_is_excise", "Excise (new)", kind="checkbox"),
    ],
}


NUMERIC_FIELDS = {
    "stock_qty", "markdown_qty", "reserve_qty", "lpc", "landed_cost", "distr_price", "promo_price",
    "transit_qty", "order_qty", "confirmed_qty", "remains_qty", "new_pack",
}


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


class ProductStockPage(QWidget):
    def __init__(self):
        super().__init__()
        self.ui = load_ui(UI_PATH)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.ui)

        self.table = self.ui.table
        self._updating_table = False
        self._batch_id = ""
        self._imported_by = get_current_username()
        self._current_file_path = ""
        self._mode = MODE_STOCK

        self.setup_ui()
        self.setup_connections()
        self.start_new_batch()
        self.refresh_filters()
        self.apply_mode(MODE_STOCK)
        self.show_message("")

    def get_session(self):
        return SessionLocal()

    def setup_ui(self):
        setup_data_table(self.table, sorting=False)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)

        for widget in (self.ui.line_RowsLoaded, self.ui.line_TotalQty, self.ui.line_RowsError):
            widget.setReadOnly(True)

    def setup_connections(self):
        self.ui.radio_Stock.toggled.connect(lambda checked: checked and self.apply_mode(MODE_STOCK))
        self.ui.radio_Orders.toggled.connect(lambda checked: checked and self.apply_mode(MODE_ORDERS))
        self.ui.radio_IS.toggled.connect(lambda checked: checked and self.apply_mode(MODE_IS))

        self.ui.line_FindProduct.textChanged.connect(self.refresh_product_combos)
        self.ui.cbo_FindBrand.currentTextChanged.connect(self.refresh_product_combos)

        self.ui.btn_Import.clicked.connect(self.import_file)
        self.ui.btn_Save.clicked.connect(self.save_all)
        self.ui.btn_Reset.clicked.connect(self.reset_all)
        self.table.customContextMenuRequested.connect(self.show_context_menu)
        self.table.itemChanged.connect(self.on_item_changed)

    def show_message(self, text: str):
        self.ui.label_msg.setText(text)

    def show_error_message(self, text: str):
        self.ui.label_msg.setText(text)

    def start_new_batch(self):
        self._batch_id = datetime.now().strftime("PS_%Y%m%d_%H%M%S_%f")
        self._imported_by = get_current_username()
        self._current_file_path = ""

    def apply_mode(self, mode: str):
        self._mode = mode
        titles = {
            MODE_STOCK: "Update stock",
            MODE_ORDERS: "Update supplier orders",
            MODE_IS: "Update IS",
        }
        self.ui.Title_label.setText(titles[mode])
        self.load_table()
        self.refresh_counters()

    def refresh_filters(self):
        with self.get_session() as session:
            brands = (
                session.query(Product.brand)
                .filter(Product.brand.isnot(None), Product.brand != "")
                .distinct()
                .order_by(Product.brand.asc())
                .all()
            )
        self.ui.cbo_FindBrand.blockSignals(True)
        self.ui.cbo_FindBrand.clear()
        self.ui.cbo_FindBrand.addItem("-")
        for row in brands:
            self.ui.cbo_FindBrand.addItem(row[0])
        self.ui.cbo_FindBrand.blockSignals(False)

    def _temp_model(self):
        return {
            MODE_STOCK: TempStockImport,
            MODE_ORDERS: TempSupplierOrdersImport,
            MODE_IS: TempIsImport,
        }[self._mode]

    def _product_match_mode(self) -> str:
        return MODE_IS if self._mode == MODE_IS else MODE_STOCK

    def _get_filtered_products(self):
        with self.get_session() as session:
            query = session.query(Product).filter(Product.name.isnot(None), Product.name != "")
            brand = self.ui.cbo_FindBrand.currentText().strip()
            find_text = self.ui.line_FindProduct.text().strip().lower()
            if brand and brand != "-":
                query = query.filter(Product.brand == brand)
            products = query.order_by(Product.name.asc()).all()
            if find_text:
                products = [p for p in products if find_text in (p.name or "").lower()]
            return products

    def _get_brand_values(self):
        with self.get_session() as session:
            rows = (
                session.query(Product.brand)
                .filter(Product.brand.isnot(None), Product.brand != "")
                .distinct()
                .order_by(Product.brand.asc())
                .all()
            )
        return [r[0] for r in rows]

    def _query_mode_rows(self, session):
        model = self._temp_model()
        return (
            session.query(model)
            .filter(model.batch_id == self._batch_id, model.imported_by == self._imported_by)
            .order_by(model.import_row_no.asc(), model.id.asc())
            .all()
        )

    def load_table(self):
        with self.get_session() as session:
            rows = self._query_mode_rows(session)
            data = [self._row_to_dict(row) for row in rows]
        self.display_table(data)
        self.refresh_counters()

    def _row_to_dict(self, row) -> dict[str, Any]:
        base = {
            "id": row.id,
            "selected_product_id": row.selected_product_id,
            "source_article": row.source_article,
            "source_product_name": row.source_product_name,
            "new_product_name": row.new_product_name,
            "new_brand": row.new_brand,
            "new_pack": row.new_pack,
            "new_is_excise": bool(row.new_is_excise) if row.new_is_excise is not None else False,
        }
        if isinstance(row, TempStockImport):
            base.update({
                "source_sku": row.source_sku,
                "stock_qty": row.stock_qty,
                "markdown_qty": row.markdown_qty,
                "reserve_qty": row.reserve_qty,
                "has_lpc_warning": bool(row.has_lpc_warning),
                "lpc": row.lpc,
                "landed_cost": row.landed_cost,
                "distr_price": row.distr_price,
                "promo_price": row.promo_price,
            })
        elif isinstance(row, TempSupplierOrdersImport):
            base.update({
                "transit_qty": row.transit_qty,
                "order_qty": row.order_qty,
            })
        elif isinstance(row, TempIsImport):
            base.update({
                "confirmed_qty": row.confirmed_qty,
                "remains_qty": row.remains_qty,
                "stock_qty": row.stock_qty,
            })
        return base

    def display_table(self, data: list[dict[str, Any]]):
        self._updating_table = True
        try:
            columns = COLUMN_DEFS[self._mode]
            self.table.blockSignals(True)
            self.table.clear()
            self.table.setColumnCount(len(columns))
            self.table.setHorizontalHeaderLabels([c.header for c in columns])
            self.table.setRowCount(len(data))
            brand_values = self._get_brand_values()

            for row_index, row_data in enumerate(data):
                row_id = row_data["id"]
                for col_index, col in enumerate(columns):
                    value = row_data.get(col.key)
                    if col.kind == "product_combo":
                        self.table.setCellWidget(row_index, col_index, self._build_product_combo(row_id, value))
                    elif col.kind == "brand_combo":
                        self.table.setCellWidget(row_index, col_index, self._build_brand_combo(row_id, value, brand_values))
                    elif col.kind == "checkbox":
                        self.table.setCellWidget(row_index, col_index, self._build_checkbox(row_id, col.key, bool(value)))
                    elif col.kind == "indicator":
                        self.table.setCellWidget(row_index, col_index, self._build_indicator(bool(value)))
                    else:
                        item = QTableWidgetItem("" if value is None else str(value))
                        item.setData(Qt.UserRole, row_id)
                        flags = Qt.ItemIsSelectable | Qt.ItemIsEnabled
                        if col.editable:
                            flags |= Qt.ItemIsEditable
                        item.setFlags(flags)
                        self.table.setItem(row_index, col_index, item)
        finally:
            self.table.blockSignals(False)
            self._updating_table = False

    def _build_indicator(self, checked: bool):
        checkbox = QCheckBox()
        checkbox.setChecked(checked)
        checkbox.setEnabled(False)
        container = QWidget()
        layout = QHBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(checkbox)
        layout.setAlignment(Qt.AlignCenter)
        return container

    def _build_checkbox(self, row_id: int, field_name: str, checked: bool):
        checkbox = QCheckBox()
        checkbox.setChecked(checked)
        checkbox.toggled.connect(lambda state, rid=row_id, fn=field_name: self.update_temp_field(rid, fn, bool(state), reload=False, clear_selected=True))
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
        combo.currentTextChanged.connect(
            lambda text, rid=row_id: self.update_temp_field(rid, "new_brand", clean_multi_spaces(text) or None, reload=False, clear_selected=True)
        )
        return combo

    def _build_product_combo(self, row_id: int, selected_product_id: int | None):
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
                if product is not None:
                    combo.addItem(product.name, product.id)
        idx = combo.findData(selected_product_id)
        if idx >= 0:
            combo.setCurrentIndex(idx)
        combo.currentIndexChanged.connect(
            lambda _, rid=row_id, c=combo: self.update_temp_field(rid, "selected_product_id", c.currentData(), reload=False)
        )
        return combo

    def on_item_changed(self, item: QTableWidgetItem):
        if self._updating_table:
            return
        row_id = item.data(Qt.UserRole)
        if row_id is None:
            return
        col = COLUMN_DEFS[self._mode][item.column()]
        if not col.editable:
            return
        value: Any = item.text().strip()
        if col.key in NUMERIC_FIELDS:
            value = parse_loose_number(value)
        elif value == "":
            value = None

        clear_selected = col.key in {"new_product_name", "new_pack"}
        self.update_temp_field(row_id, col.key, value, reload=False, clear_selected=clear_selected)

        if col.key in {"source_article", "source_product_name"}:
            self.try_auto_match_row(row_id)
        elif col.key in {"new_product_name", "new_brand", "new_pack", "new_is_excise"}:
            self.refresh_counters()

    def _get_row_by_id(self, session, row_id: int):
        return session.query(self._temp_model()).filter(self._temp_model().id == row_id).first()

    def update_temp_field(self, row_id: int, field_name: str, value: Any, reload: bool = False, clear_selected: bool = False):
        if self._updating_table:
            return
        try:
            with self.get_session() as session:
                row = self._get_row_by_id(session, row_id)
                if row is None:
                    return
                setattr(row, field_name, value)
                if clear_selected and value not in (None, ""):
                    row.selected_product_id = None
                if field_name == "selected_product_id" and value is not None:
                    row.new_product_name = None
                    row.new_brand = None
                    row.new_pack = None
                    row.new_is_excise = None
                session.commit()
        except Exception as e:
            self.show_error_message(str(e))
            return

        if reload:
            self.load_table()
        else:
            self.refresh_counters()

    def try_auto_match_row(self, row_id: int):
        try:
            with self.get_session() as session:
                row = self._get_row_by_id(session, row_id)
                if row is None or row.selected_product_id is not None:
                    return
                matcher = ProductMatchingService(session)
                if self._product_match_mode() == MODE_IS:
                    product = matcher.find_is_product(row.source_article, row.source_product_name)
                else:
                    product = matcher.find_stock_product(row.source_article, row.source_product_name)
                if product is not None:
                    row.selected_product_id = product.id
                    session.commit()
                    self.load_table()
                else:
                    self.refresh_counters()
        except Exception as e:
            self.show_error_message(str(e))

    def refresh_product_combos(self):
        self.load_table()

    def refresh_counters(self):
        with self.get_session() as session:
            model = self._temp_model()
            rows = self._query_mode_rows(session)
            row_count = len(rows)
            error_count = sum(1 for row in rows if row.selected_product_id is None)
            total_qty = 0
            for row in rows:
                if model is TempStockImport:
                    total_qty += float(row.stock_qty or 0) + float(row.markdown_qty or 0) + float(row.reserve_qty or 0)
                elif model is TempSupplierOrdersImport:
                    total_qty += float(row.transit_qty or 0) + float(row.order_qty or 0)
                else:
                    total_qty += float(row.confirmed_qty or 0) + float(row.remains_qty or 0) + float(row.stock_qty or 0)

        self.ui.line_RowsLoaded.setText(str(row_count))
        self.ui.line_RowsError.setText(str(error_count))
        self.ui.line_TotalQty.setText(f"{total_qty:.4f}".rstrip("0").rstrip("."))

    def import_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Выберите Excel файл", "", "Excel files (*.xls *.xlsx *.xlsm)")
        if not file_path:
            return
        try:
            self._current_file_path = file_path
            with self.get_session() as session:
                runner = ProductStockImportRun(session)
                self.start_new_batch()
                if self._mode == MODE_STOCK:
                    stats = runner.import_stock(file_path=file_path, imported_by=self._imported_by, batch_id=self._batch_id)
                elif self._mode == MODE_ORDERS:
                    stats = runner.import_supplier_orders(file_path=file_path, imported_by=self._imported_by, batch_id=self._batch_id)
                else:
                    stats = runner.import_is(file_path=file_path, imported_by=self._imported_by, batch_id=self._batch_id)
                session.commit()

            files = self.export_issue_files()
            self.load_table()
            self.show_message(self._build_import_message(stats, files))
        except Exception as e:
            self.show_error_message(str(e))

    def _build_import_message(self, stats: dict[str, Any], files: list[str]) -> str:
        base = f"Импорт завершен. Загружено: {stats.get('imported_count', 0)}. Автоподбор: {stats.get('matched_count', 0)}."
        if not files:
            return base
        return base + " Файлы: " + "; ".join(files)

    def _build_sibling_path(self, suffix: str) -> Path:
        source = Path(self._current_file_path)
        return source.parent / f"{suffix}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"

    def export_issue_files(self) -> list[str]:
        created: list[str] = []
        if not self._current_file_path:
            return created
        with self.get_session() as session:
            runner = ProductStockImportRun(session)
            if self._mode == MODE_STOCK:
                issue_path = runner.export_stock_product_issues(self._batch_id, self._imported_by, self._build_sibling_path("Stock_ProductIssues"))
                warn_path = runner.export_stock_lpc_warnings(self._batch_id, self._imported_by, self._build_sibling_path("Stock_LPCWarnings"))
                session.commit()
                if issue_path:
                    created.append(str(issue_path))
                if warn_path:
                    created.append(str(warn_path))
            elif self._mode == MODE_ORDERS:
                issue_path = runner.export_supplier_orders_product_issues(self._batch_id, self._imported_by, self._build_sibling_path("SupplierOrders_ProductIssues"))
                session.commit()
                if issue_path:
                    created.append(str(issue_path))
            else:
                issue_path = runner.export_is_product_issues(self._batch_id, self._imported_by, self._build_sibling_path("IS_ProductIssues"))
                session.commit()
                if issue_path:
                    created.append(str(issue_path))
        return created

    def save_all(self):
        try:
            with self.get_session() as session:
                rows = self._query_mode_rows(session)
                if not rows:
                    self.show_error_message("Нет данных для сохранения. Сначала импортируйте файл или добавьте строки.")
                    return
                matched_count = sum(1 for row in rows if row.selected_product_id is not None)
                if matched_count == 0:
                    self.show_error_message("В текущем импорте нет ни одной строки с SelectedProductID.")
                    return

                runner = ProductStockImportRun(session)
                if self._mode == MODE_STOCK:
                    stats = runner.save_stock(self._batch_id, self._imported_by)
                    msg = "Данные по остаткам успешно сохранены."
                elif self._mode == MODE_ORDERS:
                    stats = runner.save_supplier_orders(self._batch_id, self._imported_by)
                    msg = "Данные по заказам поставщиков успешно сохранены."
                else:
                    stats = runner.save_is(self._batch_id, self._imported_by)
                    msg = "Данные IS успешно сохранены."
                session.commit()

            self.start_new_batch()
            self.load_table()
            self.show_message(
                msg +
                f" Создано продуктов: {stats.get('created_products_count', 0)}. "
                f"Обновлено связей: {stats.get('product_articles_count', 0)}. "
                f"Сохранено строк: {stats.get('saved_count', 0)}."
            )
        except Exception as e:
            self.show_error_message(str(e))

    def reset_all(self):
        if QMessageBox.question(self, "Подтверждение", "Сбросить все данные текущего импорта?") != QMessageBox.Yes:
            return
        try:
            with self.get_session() as session:
                service = ProductStockImportRun(session).service
                if self._mode == MODE_STOCK:
                    service.delete_stock_rows(self._batch_id, self._imported_by)
                elif self._mode == MODE_ORDERS:
                    service.delete_supplier_orders_rows(self._batch_id, self._imported_by)
                else:
                    service.delete_is_rows(self._batch_id, self._imported_by)
                session.commit()
            self.ui.line_FindProduct.clear()
            self.ui.cbo_FindBrand.setCurrentIndex(0)
            self.start_new_batch()
            self.load_table()
            self.show_message("Форма очищена. Можно начинать новый импорт.")
        except Exception as e:
            self.show_error_message(str(e))

    def show_context_menu(self, position):
        menu = QMenu()
        copy_action = menu.addAction("Копировать")
        open_files_action = None
        if self._current_file_path:
            open_files_action = menu.addAction("Открыть папку файла")
        action = menu.exec(self.table.viewport().mapToGlobal(position))
        if action == copy_action:
            self.copy_cell_content()
        elif action == open_files_action:
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(Path(self._current_file_path).parent)))

    def copy_cell_content(self):
        selected_items = self.table.selectedItems()
        if not selected_items:
            return
        clipboard = QApplication.clipboard()
        if len(selected_items) == 1:
            clipboard.setText(selected_items[0].text())
        else:
            rows: dict[int, dict[int, str]] = {}
            for item in selected_items:
                rows.setdefault(item.row(), {})[item.column()] = item.text()
            text = "\n".join("\t".join(cols[c] for c in sorted(cols)) for _, cols in sorted(rows.items()))
            clipboard.setText(text)
        self.show_message("Скопировано")
