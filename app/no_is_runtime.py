from __future__ import annotations

import importlib
import sys
from decimal import Decimal
from typing import Any, Iterable, Mapping, Sequence


_FORBIDDEN_EXCEL_HEADERS = {"order is", "stock is"}
_FORBIDDEN_ORDER_STAGE_KEYS = {"CoralOrderQty", "CoralConfirmedQty"}

_INSTALLED = False
_EXCEL_GUARDS_INSTALLED = False
_ORIGINAL_IMPORT_MODULE = importlib.import_module


def _norm_header(value: object) -> str:
    return " ".join(str(value or "").replace("\n", " ").split()).casefold()


def _is_forbidden_header(value: object) -> bool:
    return _norm_header(value) in _FORBIDDEN_EXCEL_HEADERS


def _filtered_headers_and_indexes(headers: Sequence[object]) -> tuple[list[object], list[int]]:
    keep = [index for index, header in enumerate(headers) if not _is_forbidden_header(header)]
    return [headers[index] for index in keep], keep


def _filter_row_by_indexes(row: Sequence[Any], keep: Sequence[int]) -> list[Any]:
    return [row[index] if index < len(row) else "" for index in keep]


def _filter_dict_keys(row: Mapping[str, Any]) -> dict[str, Any]:
    return {
        key: value
        for key, value in row.items()
        if not _is_forbidden_header(key)
    }


def _install_excel_guards() -> None:
    """Install project-wide guards so Order IS / Stock IS cannot leak to Excel."""
    global _EXCEL_GUARDS_INSTALLED
    if _EXCEL_GUARDS_INSTALLED:
        return
    _EXCEL_GUARDS_INSTALLED = True

    # Shared fast COM writer.
    try:
        fast_writer = _ORIGINAL_IMPORT_MODULE("app.utils.excel_fast_writer")
        original_fast_writer = fast_writer.write_excel_table
        if not getattr(original_fast_writer, "_no_is_guard", False):
            def write_excel_table_no_is(
                ws,
                headers,
                rows,
                *,
                header_getter=None,
                value_getter=None,
                start_row=1,
                start_col=1,
                chunk_size=fast_writer.DEFAULT_CHUNK_SIZE,
            ):
                header_list = list(headers)
                visible_headers = [
                    header_getter(header) if header_getter is not None else header
                    for header in header_list
                ]
                keep = [
                    index
                    for index, visible in enumerate(visible_headers)
                    if not _is_forbidden_header(visible)
                ]
                if len(keep) == len(header_list):
                    return original_fast_writer(
                        ws,
                        header_list,
                        rows,
                        header_getter=header_getter,
                        value_getter=value_getter,
                        start_row=start_row,
                        start_col=start_col,
                        chunk_size=chunk_size,
                    )

                filtered_headers = [header_list[index] for index in keep]

                if value_getter is not None:
                    def filtered_getter(row, header, filtered_index):
                        original_index = keep[filtered_index]
                        return value_getter(row, header, original_index)

                    return original_fast_writer(
                        ws,
                        filtered_headers,
                        rows,
                        header_getter=header_getter,
                        value_getter=filtered_getter,
                        start_row=start_row,
                        start_col=start_col,
                        chunk_size=chunk_size,
                    )

                def filtered_rows():
                    for row in rows:
                        if isinstance(row, Mapping):
                            yield row
                        else:
                            yield _filter_row_by_indexes(row, keep)

                return original_fast_writer(
                    ws,
                    filtered_headers,
                    filtered_rows(),
                    header_getter=header_getter,
                    value_getter=None,
                    start_row=start_row,
                    start_col=start_col,
                    chunk_size=chunk_size,
                )

            write_excel_table_no_is._no_is_guard = True
            fast_writer.write_excel_table = write_excel_table_no_is
    except Exception:
        pass

    # Shared openpyxl writer.
    try:
        export_format = _ORIGINAL_IMPORT_MODULE("app.utils.excel_export_format")
        original_openpyxl_writer = export_format.write_openpyxl_dict_sheet
        if not getattr(original_openpyxl_writer, "_no_is_guard", False):
            def write_openpyxl_dict_sheet_no_is(ws, rows, *, widths=None):
                filtered = [_filter_dict_keys(row) for row in rows]
                filtered_widths = None
                if widths is not None:
                    filtered_widths = {
                        key: value for key, value in widths.items()
                        if not _is_forbidden_header(key)
                    }
                return original_openpyxl_writer(ws, filtered, widths=filtered_widths)

            write_openpyxl_dict_sheet_no_is._no_is_guard = True
            export_format.write_openpyxl_dict_sheet = write_openpyxl_dict_sheet_no_is

        integer_headers = getattr(export_format, "INTEGER_HEADERS", None)
        if isinstance(integer_headers, set):
            integer_headers.discard("order is")
            integer_headers.discard("stock is")
    except Exception:
        pass

    # CostCalc_ header source.
    try:
        format_rules = _ORIGINAL_IMPORT_MODULE("app.utils.excel_format_rules")
        original_cost_calc_headers = format_rules.cost_calc_headers
        if not getattr(original_cost_calc_headers, "_no_is_guard", False):
            def cost_calc_headers_no_is(*, quick_order_months, safe_stock_months):
                headers, standard_order_header = original_cost_calc_headers(
                    quick_order_months=quick_order_months,
                    safe_stock_months=safe_stock_months,
                )
                return [
                    header for header in headers if not _is_forbidden_header(header)
                ], standard_order_header

            cost_calc_headers_no_is._no_is_guard = True
            format_rules.cost_calc_headers = cost_calc_headers_no_is
    except Exception:
        pass

    # Shared Excel formatting maps.
    try:
        column_format = _ORIGINAL_IMPORT_MODULE("app.exports.excel_column_format")
        for attr in (
            "NUMERIC_HEADERS",
            "HEADER_WIDTHS",
            "NUMBER_FORMATS_LOCAL",
            "HEADER_FILL_COLORS",
            "HEADER_FONT_COLORS",
        ):
            value = getattr(column_format, attr, None)
            if isinstance(value, set):
                for item in list(value):
                    if _is_forbidden_header(item):
                        value.discard(item)
            elif isinstance(value, dict):
                for key in list(value):
                    if _is_forbidden_header(key):
                        value.pop(key, None)
    except Exception:
        pass

    try:
        output_headers = _ORIGINAL_IMPORT_MODULE("app.utils.output_headers")
        specs = getattr(output_headers, "HEADER_SPECS", None)
        if isinstance(specs, dict):
            for key in list(specs):
                if _is_forbidden_header(key):
                    specs.pop(key, None)
    except Exception:
        pass


def _patch_supplier_orders_importer(service_module) -> None:
    importer_cls = getattr(service_module, "SupplierOrdersImporter", None)
    if importer_cls is None:
        return

    original = importer_cls.read_excel
    if getattr(original, "_no_is_patch", False):
        return

    def read_excel_without_is(self, file_path):
        rows = original(self, file_path)
        for row in rows:
            normal_qty = float(row.get("order_qty") or 0)
            coral_order = float(row.get("is_order_qty") or 0)
            coral_confirmed = float(row.get("is_confirmed_order_qty") or 0)
            # CORAL is no longer separated into IS buckets:
            # order + confirmed are ordinary Purchase Order.
            row["order_qty"] = normal_qty + coral_order + coral_confirmed
            row["is_order_qty"] = 0
            row["is_confirmed_order_qty"] = 0
        return rows

    read_excel_without_is._no_is_patch = True
    importer_cls.read_excel = read_excel_without_is


def _patch_product_stock_service(service_module) -> None:
    service_cls = getattr(service_module, "ProductStockService", None)
    product_stock_cls = getattr(service_module, "ProductStock", None)
    if service_cls is None or product_stock_cls is None:
        return

    _patch_supplier_orders_importer(service_module)

    original = service_cls.save_supplier_orders_to_product_stock
    if getattr(original, "_no_is_patch", False):
        return

    def save_supplier_orders_preserve_legacy_is(self, batch_id: str, imported_by: str) -> int:
        # The DB schema is intentionally retained. Preserve historical IS values
        # while the legacy method updates Purchase Order.
        snapshot = {
            int(row.id): (
                row.is_order_qty,
                row.is_confirmed_order_qty,
                row.is_stock_qty,
                row.is_update_date,
            )
            for row in self.session.query(product_stock_cls).all()
            if row.id is not None
        }

        result = original(self, batch_id, imported_by)

        if snapshot:
            # The legacy method performs a bulk UPDATE with
            # synchronize_session=False. Some snapshot ORM objects therefore
            # still contain the old values in memory even though the DB row was
            # zeroed. Restore with explicit bulk UPDATEs so SQL is always sent.
            for row_id, old in snapshot.items():
                (
                    self.session.query(product_stock_cls)
                    .filter(product_stock_cls.id == row_id)
                    .update(
                        {
                            product_stock_cls.is_order_qty: old[0],
                            product_stock_cls.is_confirmed_order_qty: old[1],
                            product_stock_cls.is_stock_qty: old[2],
                            product_stock_cls.is_update_date: old[3],
                        },
                        synchronize_session=False,
                    )
                )
            self.session.flush()
            self.session.expire_all()

        return result

    save_supplier_orders_preserve_legacy_is._no_is_patch = True
    service_cls.save_supplier_orders_to_product_stock = save_supplier_orders_preserve_legacy_is


def _patch_product_stock_exporter(exporter_module) -> None:
    exporter_cls = getattr(exporter_module, "ProductStockExporter", None)
    temp_orders_cls = getattr(exporter_module, "TempSupplierOrdersImport", None)
    if exporter_cls is None or temp_orders_cls is None:
        return

    original = exporter_cls.export_supplier_orders_product_issues
    if getattr(original, "_no_is_patch", False):
        return

    def export_supplier_orders_product_issues_no_is(
        self,
        batch_id: str,
        imported_by: str,
        output_path,
    ):
        rows = (
            self.session.query(temp_orders_cls)
            .filter(
                temp_orders_cls.batch_id == batch_id,
                temp_orders_cls.imported_by == imported_by,
                temp_orders_cls.selected_product_id.is_(None),
            )
            .order_by(temp_orders_cls.import_row_no.asc())
            .all()
        )

        out = []
        for row in rows:
            out.append({
                "ImportRowNo": row.import_row_no,
                "SourceArticle": row.source_article,
                "SourceProductName": row.source_product_name,
                "OrderQty": row.order_qty,
                "Comment": (
                    "Не найден SelectedProductID. "
                    "Требуется сопоставление или создание нового продукта."
                ),
            })

        if not out:
            return None
        return self._write(out, output_path)

    export_supplier_orders_product_issues_no_is._no_is_patch = True
    exporter_cls.export_supplier_orders_product_issues = (
        export_supplier_orders_product_issues_no_is
    )


def _patch_product_stock_page(module) -> None:
    if getattr(module, "_NO_IS_PATCHED", False):
        return
    module._NO_IS_PATCHED = True
    _install_excel_guards()

    service_module = sys.modules.get("app.services.product_stock_service")
    if service_module is not None:
        _patch_product_stock_service(service_module)

    exporter_module = sys.modules.get("app.exports.product_stock_exporter")
    if exporter_module is not None:
        _patch_product_stock_exporter(exporter_module)

    orders_mode = getattr(module, "MODE_ORDERS", "orders")
    order_columns = getattr(module, "COLUMN_DEFS", {}).get(orders_mode, [])
    if order_columns:
        module.COLUMN_DEFS[orders_mode] = [
            col
            for col in order_columns
            if getattr(col, "key", "") not in {
                "is_order_qty",
                "is_confirmed_order_qty",
            }
        ]

    numeric_fields = getattr(module, "NUMERIC_FIELDS", None)
    if isinstance(numeric_fields, set):
        numeric_fields.discard("is_order_qty")
        numeric_fields.discard("is_confirmed_order_qty")
        numeric_fields.discard("confirmed_qty")
        numeric_fields.discard("remains_qty")

    page_cls = module.ProductStockPage

    original_init = page_cls.__init__
    if not getattr(original_init, "_no_is_patch", False):
        def init_no_is(self, *args, **kwargs):
            original_init(self, *args, **kwargs)
            radio = getattr(self.ui, "radio_IS", None)
            if radio is not None:
                try:
                    if radio.isChecked():
                        self.ui.radio_Stock.setChecked(True)
                        self.apply_mode(getattr(module, "MODE_STOCK", "stock"))
                    radio.setEnabled(False)
                    radio.hide()
                    radio.setParent(None)
                    radio.deleteLater()
                except Exception:
                    pass

        init_no_is._no_is_patch = True
        page_cls.__init__ = init_no_is

    original_apply_mode = page_cls.apply_mode
    if not getattr(original_apply_mode, "_no_is_patch", False):
        def apply_mode_no_is(self, mode):
            if mode == getattr(module, "MODE_IS", "is"):
                mode = getattr(module, "MODE_STOCK", "stock")
            return original_apply_mode(self, mode)

        apply_mode_no_is._no_is_patch = True
        page_cls.apply_mode = apply_mode_no_is

    original_collect = page_cls._collect_issue_sheets
    if not getattr(original_collect, "_no_is_patch", False):
        def collect_issue_sheets_no_is(self):
            sheets = original_collect(self)
            cleaned = []
            for sheet_name, rows in sheets:
                clean_rows = []
                for row in rows:
                    row = dict(row)
                    for key in _FORBIDDEN_ORDER_STAGE_KEYS:
                        row.pop(key, None)
                    for key in list(row):
                        if _is_forbidden_header(key):
                            row.pop(key, None)
                    clean_rows.append(row)
                cleaned.append((sheet_name, clean_rows))
            return cleaned

        collect_issue_sheets_no_is._no_is_patch = True
        page_cls._collect_issue_sheets = collect_issue_sheets_no_is


def _patch_order_planning_service(service_module) -> None:
    service_cls = getattr(service_module, "OrderPlanningService", None)
    if service_cls is None:
        return

    original = service_cls.build_display_rows
    if getattr(original, "_no_is_patch", False):
        return

    def build_display_rows_no_is(
        self,
        base_rows,
        quick_months=Decimal("0"),
        safe_months=Decimal("0"),
        brand_filter=None,
        family_filter=None,
        product_filter=None,
        vol_not_null=False,
        product_map=None,
    ):
        # Let legacy grouping/filtering run, but do not allow its old IS-aware
        # volume filter to discard rows before we recalculate.
        rows = original(
            self,
            base_rows,
            quick_months=quick_months,
            safe_months=safe_months,
            brand_filter=brand_filter,
            family_filter=family_filter,
            product_filter=product_filter,
            vol_not_null=False,
            product_map=product_map,
        )

        quick_m = self._to_decimal(quick_months)
        safe_m = self._to_decimal(safe_months)
        result = []

        for row in rows:
            avg_sales = self._to_decimal(row.get("avg_sales_month"))
            pack = self._to_decimal(row.get("pack"), Decimal("0"))
            stock_qty = self._to_decimal(row.get("stock"))
            transit_qty = self._to_decimal(row.get("transit"))
            order_qty = self._to_decimal(row.get("purchase_order"))

            free_st = stock_qty
            free_st_tr = stock_qty + transit_qty
            free_ord = stock_qty + transit_qty + order_qty

            safe_st_month = (
                self._round4(free_st / avg_sales)
                if avg_sales > 0 else Decimal("0")
            )
            safe_st_tr_month = (
                self._round4(free_st_tr / avg_sales)
                if avg_sales > 0 else Decimal("0")
            )
            safe_ord_month = (
                self._round4(free_ord / avg_sales)
                if avg_sales > 0 else Decimal("0")
            )

            std_l_raw = (safe_m - safe_ord_month) * avg_sales
            quick_l_raw = (quick_m - safe_st_tr_month) * avg_sales
            if std_l_raw < 0:
                std_l_raw = Decimal("0")
            if quick_l_raw < 0:
                quick_l_raw = Decimal("0")

            if pack > 0:
                std_pcs = self._ceil_decimal(std_l_raw / pack)
                quick_pcs = self._ceil_decimal(quick_l_raw / pack)
            else:
                std_pcs = Decimal("0")
                quick_pcs = Decimal("0")

            std_l = std_pcs * pack
            quick_l = quick_pcs * pack

            row = dict(row)
            row["free_stock_st"] = free_st
            row["free_stock_st_tr"] = free_st_tr
            row["free_stock_ord"] = free_ord
            row["safe_stock_st_month"] = safe_st_month
            row["safe_stock_st_tr_month"] = safe_st_tr_month
            row["safe_stock_ord_month"] = safe_ord_month
            row["quick_order_pcs"] = quick_pcs
            row["quick_order_l"] = self._round4(quick_l)
            row["std_order_pcs"] = std_pcs
            row["std_order_l"] = self._round4(std_l)
            row.pop("order_is", None)
            row.pop("stock_is", None)

            # Legacy IS-only products must not remain visible as current stock/order.
            if (
                avg_sales == 0
                and stock_qty == 0
                and transit_qty == 0
                and order_qty == 0
                and row.get("product_id")
                and not row.get("is_auto_matched")
            ):
                continue

            is_new_or_unchecked = (
                not bool(row.get("product_id"))
                or bool(row.get("is_auto_matched"))
            )
            if (
                vol_not_null
                and std_l <= 0
                and quick_l <= 0
                and not is_new_or_unchecked
            ):
                continue

            result.append(row)

        return result

    build_display_rows_no_is._no_is_patch = True
    service_cls.build_display_rows = build_display_rows_no_is


def _patch_order_planning_exporter(exporter_module) -> None:
    exporter_cls = getattr(exporter_module, "OrderPlanningExporter", None)
    if exporter_cls is None:
        return

    original = exporter_cls.build_export_data
    if getattr(original, "_no_is_patch", False):
        return

    def build_export_data_no_is(self, display_rows, supplier_price_age_months=3):
        headers, rows = original(
            self,
            display_rows,
            supplier_price_age_months=supplier_price_age_months,
        )
        filtered_headers, keep = _filtered_headers_and_indexes(headers)
        filtered_rows = [_filter_row_by_indexes(row, keep) for row in rows]
        return filtered_headers, filtered_rows

    build_export_data_no_is._no_is_patch = True
    exporter_cls.build_export_data = build_export_data_no_is


def _patch_order_planning_page(module) -> None:
    if getattr(module, "_NO_IS_PATCHED", False):
        return
    module._NO_IS_PATCHED = True
    _install_excel_guards()

    service_module = sys.modules.get("app.services.order_planning_service")
    if service_module is not None:
        _patch_order_planning_service(service_module)

    exporter_module = sys.modules.get("app.exports.order_planning_exporter")
    if exporter_module is not None:
        _patch_order_planning_exporter(exporter_module)

    page_cls = module.OrderPlanningPage
    paired = list(zip(page_cls.CALC_COLUMNS, page_cls.CALC_HEADERS))
    paired = [
        (key, header)
        for key, header in paired
        if key not in {"order_is", "stock_is"}
        and not _is_forbidden_header(header)
    ]
    page_cls.CALC_COLUMNS = [key for key, _ in paired]
    page_cls.CALC_HEADERS = [header for _, header in paired]
    page_cls.NUMERIC_COLUMNS = set(page_cls.NUMERIC_COLUMNS)
    page_cls.NUMERIC_COLUMNS.discard("order_is")
    page_cls.NUMERIC_COLUMNS.discard("stock_is")


def _patch_supplier_price_exporter(exporter_module) -> None:
    exporter_cls = getattr(exporter_module, "SupplierPriceExporter", None)
    if exporter_cls is None:
        return

    # This exporter imported cost_calc_headers by name, so patch its local symbol too.
    original_cost_headers = exporter_module.cost_calc_headers
    if not getattr(original_cost_headers, "_no_is_patch", False):
        def local_cost_calc_headers_no_is(*, quick_order_months, safe_stock_months):
            headers, standard_header = original_cost_headers(
                quick_order_months=quick_order_months,
                safe_stock_months=safe_stock_months,
            )
            return [
                header for header in headers if not _is_forbidden_header(header)
            ], standard_header

        local_cost_calc_headers_no_is._no_is_patch = True
        exporter_module.cost_calc_headers = local_cost_calc_headers_no_is

    original_calc = exporter_cls._calc_order_planning_export_values
    if not getattr(original_calc, "_no_is_patch", False):
        def calc_order_planning_export_values_no_is(
            self,
            *,
            product_id,
            stock,
            pack,
            quick_months,
            order_months,
        ):
            avg_sales_month = self._get_latest_avg_sales_month(product_id)
            if avg_sales_month is None:
                return {
                    "Ср.Продажи мес": None,
                    "к Быстрому заказу, л": Decimal("0"),
                    "к Заказу, л": Decimal("0"),
                }

            stock_qty = self._to_decimal(getattr(stock, "stock_qty", None))
            transit_qty = self._to_decimal(getattr(stock, "transit_qty", None))
            order_qty = self._to_decimal(getattr(stock, "order_qty", None))

            free_st_tr = stock_qty + transit_qty
            free_plus_ord = stock_qty + transit_qty + order_qty

            return {
                "Ср.Продажи мес": avg_sales_month,
                "к Быстрому заказу, л": self._calc_order_liters(
                    quick_months,
                    avg_sales_month,
                    free_st_tr,
                    pack,
                ),
                "к Заказу, л": self._calc_order_liters(
                    order_months,
                    avg_sales_month,
                    free_plus_ord,
                    pack,
                ),
            }

        calc_order_planning_export_values_no_is._no_is_patch = True
        exporter_cls._calc_order_planning_export_values = (
            calc_order_planning_export_values_no_is
        )

    original_build = exporter_cls.build_export_rows
    if not getattr(original_build, "_no_is_patch", False):
        def build_export_rows_no_is(self, *args, **kwargs):
            rows = original_build(self, *args, **kwargs)

            Product = getattr(exporter_module, "Product")
            ProductStock = getattr(exporter_module, "ProductStock")
            stock_by_name = {
                (product.name or "").strip(): stock
                for product, stock in (
                    self.session.query(Product, ProductStock)
                    .join(ProductStock, ProductStock.product_id == Product.id)
                    .all()
                )
            }

            for row in rows:
                stock = stock_by_name.get(str(row.get("Our Product Name") or "").strip())
                if stock is not None:
                    row["Transit"] = self._excel_value(
                        getattr(stock, "transit_qty", None)
                    )
                row.pop("Order IS", None)
                row.pop("Stock IS", None)

            return rows

        build_export_rows_no_is._no_is_patch = True
        exporter_cls.build_export_rows = build_export_rows_no_is


def _patch_supplier_prices_page(module) -> None:
    if getattr(module, "_NO_IS_PATCHED", False):
        return
    module._NO_IS_PATCHED = True
    _install_excel_guards()

    exporter_module = sys.modules.get("app.exports.supplier_price_exporter")
    if exporter_module is not None:
        _patch_supplier_price_exporter(exporter_module)


def _patch_price_reports_page(module) -> None:
    if getattr(module, "_NO_IS_PATCHED", False):
        return
    module._NO_IS_PATCHED = True
    _install_excel_guards()

    page_cls = module.PriceReportsPage

    original_product_headers = page_cls._build_product_headers
    original_product_row = page_cls._build_product_row
    original_supplier_headers = page_cls._build_supplier_headers
    original_supplier_row = page_cls._build_supplier_row

    def product_headers_no_is(self, supplier_count):
        return [
            header
            for header in original_product_headers(self, supplier_count)
            if not _is_forbidden_header(header)
        ]

    def product_row_no_is(self, product, stock, options, supplier_count):
        original_headers = original_product_headers(self, supplier_count)
        row = original_product_row(self, product, stock, options, supplier_count)
        _, keep = _filtered_headers_and_indexes(original_headers)
        return _filter_row_by_indexes(row, keep)

    def supplier_headers_no_is(self, show_prev, other_count):
        return [
            header
            for header in original_supplier_headers(self, show_prev, other_count)
            if not _is_forbidden_header(header)
        ]

    def supplier_row_no_is(
        self,
        product,
        stock,
        chosen,
        prev,
        alternatives,
        show_prev,
        other_count,
    ):
        original_headers = original_supplier_headers(self, show_prev, other_count)
        row = original_supplier_row(
            self,
            product,
            stock,
            chosen,
            prev,
            alternatives,
            show_prev,
            other_count,
        )
        _, keep = _filtered_headers_and_indexes(original_headers)
        return _filter_row_by_indexes(row, keep)

    page_cls._build_product_headers = product_headers_no_is
    page_cls._build_product_row = product_row_no_is
    page_cls._build_supplier_headers = supplier_headers_no_is
    page_cls._build_supplier_row = supplier_row_no_is

    original_loader = page_cls._load_order_plan_export_values_by_product_name
    if not getattr(original_loader, "_no_is_patch", False):
        def load_order_plan_export_values_no_is(
            self,
            *,
            quick_months,
            order_months,
        ):
            result = original_loader(
                self,
                quick_months=quick_months,
                order_months=order_months,
            )

            Product = getattr(module, "Product")
            ProductStock = getattr(module, "ProductStock")

            with self.get_session() as session:
                stocks = {
                    int(stock.product_id): stock
                    for stock in session.query(ProductStock).all()
                    if stock.product_id is not None
                }
                products = session.query(Product).all()

                for product in products:
                    name = (product.name or "").strip()
                    values = result.get(name)
                    if not values:
                        continue

                    avg_sales = values.get("Ср.Продажи мес")
                    if avg_sales in (None, ""):
                        continue

                    stock = stocks.get(int(product.id))
                    stock_qty = (
                        self._to_decimal(
                            getattr(stock, "stock_qty", None),
                            Decimal("0"),
                        )
                        or Decimal("0")
                    )
                    transit_qty = (
                        self._to_decimal(
                            getattr(stock, "transit_qty", None),
                            Decimal("0"),
                        )
                        or Decimal("0")
                    )
                    order_qty = (
                        self._to_decimal(
                            getattr(stock, "order_qty", None),
                            Decimal("0"),
                        )
                        or Decimal("0")
                    )

                    free_plus_ord = stock_qty + transit_qty + order_qty
                    values["к Заказу, л"] = self._calc_order_liters_for_export(
                        months=order_months,
                        avg_sales_month=avg_sales,
                        free_qty=free_plus_ord,
                        pack=getattr(product, "pack", None),
                    )

            return result

        load_order_plan_export_values_no_is._no_is_patch = True
        page_cls._load_order_plan_export_values_by_product_name = (
            load_order_plan_export_values_no_is
        )


def _after_import(name: str, module) -> None:
    try:
        if name == "app.page_functions.product_stock_page":
            _patch_product_stock_page(module)
        elif name == "app.page_functions.order_planning_page":
            _patch_order_planning_page(module)
        elif name == "app.page_functions.price_reports_page":
            _patch_price_reports_page(module)
        elif name == "app.page_functions.supplier_prices_page":
            _patch_supplier_prices_page(module)
    except Exception as exc:
        # Keep the application startable and make failures visible in the console.
        print(f"[no_is_runtime] patch failed for {name}: {exc}")


def _patched_import_module(name: str, package: str | None = None):
    module = _ORIGINAL_IMPORT_MODULE(name, package)
    _after_import(module.__name__, module)
    return module


def install_no_is_runtime() -> None:
    """Enable the no-IS runtime once, without touching the existing DB schema."""
    global _INSTALLED
    if _INSTALLED:
        return
    _INSTALLED = True
    importlib.import_module = _patched_import_module
