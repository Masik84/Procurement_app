from __future__ import annotations

from dataclasses import dataclass
from types import MethodType
from typing import Callable

from PySide6.QtCore import QObject, QTimer
from PySide6.QtWidgets import QComboBox

from app.db.models import Product
from app.utils.text import clean_multi_spaces


SUPPORTED_PAGES = {
    "SupplierPricesPage",
    "TargetPricesPage",
    "ProductSearchPage",
    "ProductArticlesPage",
    "CustomerCostsPage",
    "ProductStockPage",
    "OrderPlanningPage",
    "PriceHistoryPage",
}


@dataclass(frozen=True, slots=True)
class ProductSelectionRow:
    id: int
    name: str
    brand: str
    family: str
    pack: object
    qty_in_box: object
    abc_category: str
    is_excise: bool


class ProductSelectionFilter(QObject):
    """Shared controller for the top Product search + Brand selector.

    The static widgets are always defined in the page's .ui file.  This class
    only owns their data/behaviour and the filtering of product choices used by
    table cell editors.
    """

    DEFAULT_DEBOUNCE_MS = 250

    def __init__(self, page, *, result_limit: int | None = None) -> None:
        super().__init__(page)
        self.page = page
        self.ui = page.ui
        self.brand_combo = self.ui.cbo_FindBrand
        self.search_edit = self.ui.line_FindProduct
        self.result_limit = result_limit
        self._products: list[ProductSelectionRow] = []
        self._by_id: dict[int, ProductSelectionRow] = {}
        self._legacy_filter_slot: Callable | None = None
        self._refresh_callback: Callable[[], None] | None = None

        self._timer = QTimer(self)
        self._timer.setSingleShot(True)
        self._timer.setInterval(self.DEFAULT_DEBOUNCE_MS)
        self._timer.timeout.connect(self._refresh_open_product_editor)

        self._remember_and_disconnect_legacy_signals()
        self._install_page_adapters()
        self.reload()

        self.brand_combo.currentTextChanged.connect(self.schedule_refresh)
        self.search_edit.textChanged.connect(self.schedule_refresh)

        # Price History has two independent scopes: left report filters and
        # these top product-picking controls. Switching report mode resets only
        # the top picker here; the page itself owns/reset its left filters.
        if self.page.__class__.__name__ == "PriceHistoryPage" and hasattr(self.ui, "line_TableName"):
            self.ui.line_TableName.currentTextChanged.connect(self.reset)

    # ------------------------------------------------------------------
    # Product cache / top filter widgets
    # ------------------------------------------------------------------
    def reload(self) -> None:
        current_brand = clean_multi_spaces(self.brand_combo.currentText())

        with self.page.get_session() as session:
            rows = (
                session.query(
                    Product.id,
                    Product.name,
                    Product.brand,
                    Product.family,
                    Product.pack,
                    Product.qty_in_box,
                    Product.abc_category,
                    Product.is_excise,
                )
                .filter(Product.name.isnot(None), Product.name != "")
                .order_by(Product.name.asc(), Product.id.asc())
                .all()
            )

        products: list[ProductSelectionRow] = []
        for row in rows:
            try:
                product_id = int(row.id)
            except (TypeError, ValueError):
                continue
            name = str(row.name or "")
            if not name:
                continue
            products.append(
                ProductSelectionRow(
                    id=product_id,
                    name=name,
                    brand=str(row.brand or ""),
                    family=str(row.family or ""),
                    pack=row.pack,
                    qty_in_box=row.qty_in_box,
                    abc_category=str(row.abc_category or ""),
                    is_excise=bool(row.is_excise),
                )
            )

        self._products = products
        self._by_id = {product.id: product for product in products}

        brands = self.brand_names()
        self.brand_combo.blockSignals(True)
        try:
            self.brand_combo.clear()
            self.brand_combo.addItem("-")
            if brands:
                self.brand_combo.addItems(brands)
            index = self.brand_combo.findText(current_brand) if current_brand else -1
            self.brand_combo.setCurrentIndex(index if index >= 0 else 0)
        finally:
            self.brand_combo.blockSignals(False)

    def brand_names(self, *, refresh: bool = False) -> list[str]:
        if refresh:
            self.reload()
        return sorted(
            {clean_multi_spaces(product.brand) for product in self._products if clean_multi_spaces(product.brand)},
            key=str.casefold,
        )

    def filtered_products(self, *, refresh: bool = True) -> list[ProductSelectionRow]:
        # Keep product choices current after another page creates/renames a
        # product. The top filters are debounced, so this is one compact DB
        # read per actual editor refresh rather than one read per keystroke.
        if refresh:
            self.reload()
        brand_filter = clean_multi_spaces(self.brand_combo.currentText())
        text_filter = clean_multi_spaces(self.search_edit.text()).casefold()

        result: list[ProductSelectionRow] = []
        for product in self._products:
            if brand_filter and brand_filter != "-" and clean_multi_spaces(product.brand) != brand_filter:
                continue
            if text_filter and text_filter not in product.name.casefold():
                continue
            result.append(product)
            if self.result_limit is not None and len(result) >= self.result_limit:
                break
        return result

    def product_by_id(self, product_id: object) -> ProductSelectionRow | None:
        try:
            return self._by_id.get(int(product_id))
        except (TypeError, ValueError):
            return None

    def reset(self, *_args) -> None:
        self._timer.stop()
        self.search_edit.blockSignals(True)
        self.brand_combo.blockSignals(True)
        try:
            self.search_edit.clear()
            self.brand_combo.setCurrentIndex(0)
        finally:
            self.search_edit.blockSignals(False)
            self.brand_combo.blockSignals(False)
        self._refresh_open_product_editor()

    def schedule_refresh(self, *_args) -> None:
        self._timer.start()

    # ------------------------------------------------------------------
    # Integration with existing pages
    # ------------------------------------------------------------------
    def _remember_and_disconnect_legacy_signals(self) -> None:
        class_name = self.page.__class__.__name__

        if class_name == "CustomerCostsPage":
            slot = getattr(self.page, "schedule_product_combo_refresh", None)
            self._refresh_callback = getattr(self.page, "refresh_product_combos", None)
        elif class_name in {
            "SupplierPricesPage",
            "TargetPricesPage",
            "ProductSearchPage",
            "ProductArticlesPage",
            "ProductStockPage",
        }:
            slot = getattr(self.page, "refresh_current_product_combo", None)
            self._refresh_callback = slot
        elif class_name == "OrderPlanningPage":
            slot = getattr(self.page, "refresh_current_product_combo", None)
            self._refresh_callback = self._refresh_order_planning_combo
        elif class_name == "PriceHistoryPage":
            slot = None
            self._refresh_callback = self._refresh_price_history_combo
        else:
            slot = None

        self._legacy_filter_slot = slot
        if slot is None:
            return

        for signal in (self.brand_combo.currentTextChanged, self.search_edit.textChanged):
            try:
                signal.disconnect(slot)
            except (TypeError, RuntimeError):
                pass

    def _install_page_adapters(self) -> None:
        class_name = self.page.__class__.__name__

        # Pages that historically returned Product ORM objects only read scalar
        # product fields. ProductSelectionRow deliberately exposes the same
        # attributes and keeps the editor independent from a live DB session.
        if class_name in {
            "SupplierPricesPage",
            "TargetPricesPage",
            "ProductSearchPage",
            "PriceHistoryPage",
        }:
            self.page.get_filtered_products = MethodType(
                lambda _page: self.filtered_products(),
                self.page,
            )

        if class_name in {"SupplierPricesPage", "TargetPricesPage", "ProductSearchPage"}:
            self.page.get_brand_names = MethodType(
                lambda _page: self.brand_names(refresh=True),
                self.page,
            )
            self.page.load_find_brands = MethodType(
                lambda _page: self.reload(),
                self.page,
            )

        if class_name == "ProductArticlesPage":
            # Product Articles stores the selected product in the editor by
            # name rather than id. Keep that page-specific save behaviour,
            # but source both the names and Brand list from this shared filter.
            self.page._get_product_name_values = MethodType(
                lambda _page: [p.name for p in self.filtered_products()],
                self.page,
            )
            self.page.fill_in_prod_brand_list = MethodType(
                lambda _page: self.reload(),
                self.page,
            )

        if class_name == "CustomerCostsPage":
            self.page._get_filtered_products = MethodType(
                lambda _page: [(p.id, p.name) for p in self.filtered_products()],
                self.page,
            )
            self.page._get_brand_values = MethodType(
                lambda _page: self.brand_names(refresh=True),
                self.page,
            )
            self.page.refresh_filters = MethodType(
                lambda _page: self.reload(),
                self.page,
            )

        if class_name == "ProductStockPage":
            self.page._get_filtered_products = MethodType(
                lambda _page: self.filtered_products(),
                self.page,
            )
            self.page._get_brand_values = MethodType(
                lambda _page: self.brand_names(refresh=True),
                self.page,
            )
            self.page.refresh_filters = MethodType(
                lambda _page: self.reload(),
                self.page,
            )

        if class_name == "OrderPlanningPage":
            self.page.load_find_brands = MethodType(
                lambda _page: self.reload(),
                self.page,
            )
            self.page.build_product_combo = MethodType(
                lambda _page, selected_product_id, row_data=None: self._order_planning_build_product_combo(
                    selected_product_id, row_data
                ),
                self.page,
            )
            # Internal direct calls should use the shared controller too.
            self.page.refresh_current_product_combo = MethodType(
                lambda _page: self._refresh_order_planning_combo(),
                self.page,
            )

    def _refresh_open_product_editor(self) -> None:
        callback = self._refresh_callback
        if callback is not None:
            callback()

    # ------------------------------------------------------------------
    # Page-specific editor bridges. Data/filtering stays common above.
    # ------------------------------------------------------------------
    def _refresh_price_history_combo(self) -> None:
        table = getattr(self.page, "table", None)
        if table is None:
            return
        row = table.currentRow()
        if row < 0:
            return
        combo = table.cellWidget(row, 0)
        if not isinstance(combo, QComboBox):
            return

        current_id = combo.currentData()
        current_text = combo.currentText()
        products = self.filtered_products()

        combo.blockSignals(True)
        try:
            combo.clear()
            combo.addItem("", None)
            added_ids: set[int] = set()
            for product in products:
                combo.addItem(product.name, product.id)
                added_ids.add(product.id)

            selected = self.product_by_id(current_id)
            if selected is not None and selected.id not in added_ids:
                combo.addItem(selected.name, selected.id)

            index = combo.findData(current_id)
            if index >= 0:
                combo.setCurrentIndex(index)
            elif current_text:
                # Product editor in Price History is not intended for creating
                # a product, but preserve its current text while a filter is
                # being changed so the value does not visually disappear.
                combo.setCurrentText(current_text)
        finally:
            combo.blockSignals(False)

    def _order_planning_build_product_combo(self, selected_product_id: int | None, row_data: dict | None = None) -> QComboBox:
        combo = QComboBox()
        combo.setEditable(True)
        combo.setInsertPolicy(QComboBox.NoInsert)
        combo.setToolTip("Выберите продукт из базы или впишите новый Product Name вручную")
        combo.addItem("", None)

        added_ids: set[int] = set()
        for product in self.filtered_products():
            combo.addItem(product.name, product.id)
            added_ids.add(product.id)

        selected = self.product_by_id(selected_product_id)
        if selected is not None and selected.id not in added_ids:
            combo.addItem(selected.name, selected.id)

        index = combo.findData(selected_product_id)
        if index >= 0:
            combo.setCurrentIndex(index)
        else:
            row_data = row_data or {}
            combo.setCurrentText(clean_multi_spaces(row_data.get("product_name") or ""))
        return combo

    def _refresh_order_planning_combo(self) -> None:
        table = getattr(self.page, "table", None)
        if table is None:
            return
        table_row = table.currentRow()
        if table_row < 0:
            return

        if getattr(self.page, "_mode", None) == "check":
            product_col = self.page.CHECK_COLUMNS.index("product_name")
        else:
            product_col = self.page.CALC_COLUMNS.index("product_name")

        combo = table.cellWidget(table_row, product_col)
        if not isinstance(combo, QComboBox):
            return

        current_id = combo.currentData()
        current_text = combo.currentText()
        products = self.filtered_products()

        combo.blockSignals(True)
        try:
            combo.clear()
            combo.addItem("", None)
            added_ids: set[int] = set()
            for product in products:
                combo.addItem(product.name, product.id)
                added_ids.add(product.id)

            selected = self.product_by_id(current_id)
            if selected is not None and selected.id not in added_ids:
                combo.addItem(selected.name, selected.id)

            index = combo.findData(current_id)
            if index >= 0:
                combo.setCurrentIndex(index)
            else:
                combo.setCurrentText(current_text)
        finally:
            combo.blockSignals(False)


def attach_product_selection_filter(page) -> ProductSelectionFilter | None:
    """Attach the shared selector to a supported page, if its .ui has the fields."""
    if page.__class__.__name__ not in SUPPORTED_PAGES:
        return None
    ui = getattr(page, "ui", None)
    if ui is None or not hasattr(ui, "cbo_FindBrand") or not hasattr(ui, "line_FindProduct"):
        return None

    existing = getattr(page, "_product_selection_filter", None)
    if isinstance(existing, ProductSelectionFilter):
        return existing

    result_limit = 500 if page.__class__.__name__ == "CustomerCostsPage" else None
    controller = ProductSelectionFilter(page, result_limit=result_limit)
    page._product_selection_filter = controller
    return controller
