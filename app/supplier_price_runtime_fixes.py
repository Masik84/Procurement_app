from __future__ import annotations

"""Supplier Price runtime fixes.

Keeps the current no-IS CostCalc adapter compatible with the optimized
group-product export and speeds up Supplier Price import by matching rows before
staging insert instead of inserting, rereading and updating the same temp rows.
"""

import logging
import time

logger = logging.getLogger(__name__)

_installed = False


def _install_costcalc_no_is_compatibility() -> None:
    from decimal import Decimal

    from app.exports.supplier_price_exporter import SupplierPriceExporter

    current = SupplierPriceExporter._calc_order_planning_export_values
    if getattr(current, "_supplier_price_avg_sales_compat", False):
        return

    def calc_order_planning_export_values_compat(
        self,
        *,
        product_id,
        stock,
        pack,
        quick_months,
        order_months,
        avg_sales_month=None,
        **_kwargs,
    ):
        # The optimized exporter bulk-loads avg_sales_month.  Use that value
        # directly; fall back to the legacy lookup only for older callers.
        if avg_sales_month is None:
            avg_sales_month = self._get_latest_avg_sales_month(product_id)

        if avg_sales_month is None:
            return {
                "Ср.Продажи мес": None,
                "к Быстрому заказу, л": Decimal("0"),
                "к Заказу, л": Decimal("0"),
            }

        # Current application mode intentionally excludes IS quantities.
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

    calc_order_planning_export_values_compat._supplier_price_avg_sales_compat = True
    # no_is_runtime may revisit SupplierPriceExporter later when the Supplier
    # Price page is imported. Mark this replacement as an already-applied
    # no-IS implementation so the legacy wrapper cannot overwrite it again.
    calc_order_planning_export_values_compat._no_is_patch = True
    calc_order_planning_export_values_compat._supplier_price_previous = current
    SupplierPriceExporter._calc_order_planning_export_values = (
        calc_order_planning_export_values_compat
    )


def _install_fast_supplier_price_import() -> None:
    from PySide6.QtWidgets import QFileDialog

    from app.utils import page_background_integration as pbi

    current_attach = pbi._attach_supplier_prices
    if getattr(current_attach, "_supplier_price_import_bulk_fast", False):
        return

    def attach_supplier_prices_import_fast(page) -> None:
        # First keep all existing Supplier Price background/save optimizations.
        current_attach(page)

        # Reduce GUI repaint cost when hundreds of imported rows are rendered.
        if not getattr(page, "_supplier_price_display_updates_fast", False):
            page._supplier_price_display_updates_fast = True
            previous_display = page.display_rows

            def display_rows_fast(rows):
                table = page.table
                previous_blocked = table.blockSignals(True)
                previous_updates = table.updatesEnabled()
                table.setUpdatesEnabled(False)
                try:
                    previous_display(rows)
                finally:
                    table.blockSignals(previous_blocked)
                    table.setUpdatesEnabled(previous_updates)
                    if previous_updates:
                        table.viewport().update()

            page.display_rows = display_rows_fast

        old_import_slot = getattr(page, "_background_supplier_price_import_slot", None)
        if old_import_slot is not None:
            pbi._disconnect(page.ui.btn_Import.clicked, old_import_slot)

        def import_background_fast() -> None:
            try:
                supplier_id = page.ensure_supplier()
            except Exception as exc:
                page.show_error_message(str(exc))
                return

            file_path, _ = QFileDialog.getOpenFileName(
                page,
                "Выберите файл прайс-листа",
                "",
                "Excel files (*.xls *.xlsx)",
            )
            if not file_path:
                return

            batch_id = page.batch_id
            imported_by = page.imported_by
            import_date = page.get_price_date()

            def work(progress):
                from app.db.db import SessionLocal
                from app.db.models import SupplierPriceCalculation, TempPriceImport
                from app.imports.supplier_price_importer import SupplierPriceImporter
                from app.services.qty_in_box_service import normalize_qty_in_box
                from app.services.supplier_price_service import SupplierPriceService

                total_started = time.perf_counter()
                timings: dict[str, float] = {}

                try:
                    started = time.perf_counter()
                    progress("Прайс поставщика: читаю Excel...")
                    rows = SupplierPriceImporter().read_excel(file_path)
                    timings["excel"] = time.perf_counter() - started
                    progress(
                        f"Excel прочитан за {timings['excel']:.2f} с. "
                        f"Строк: {len(rows)}."
                    )

                    with SessionLocal() as session:
                        service = SupplierPriceService(session)
                        matcher = service.product_matching_service

                        # Build matching dictionaries once, before the row loop.
                        started = time.perf_counter()
                        progress("Прайс поставщика: загружаю справочники сопоставления...")
                        matcher._ensure_product_caches()
                        matcher._ensure_link_caches()
                        timings["references"] = time.perf_counter() - started
                        progress(
                            "Справочники сопоставления загружены за "
                            f"{timings['references']:.2f} с."
                        )

                        started = time.perf_counter()
                        progress("Прайс поставщика: выполняю автоматическое сопоставление...")
                        mappings: list[dict] = []
                        matched = 0
                        match_cache: dict[tuple[str, str], object] = {}

                        for row in rows:
                            cache_key = (
                                str(row.get("supplier_article") or ""),
                                str(row.get("product_name") or ""),
                            )
                            if cache_key in match_cache:
                                product = match_cache[cache_key]
                            else:
                                product = matcher.find_price_import_product(
                                    supplier_article=row.get("supplier_article"),
                                    supplier_product_name=row.get("product_name"),
                                )
                                match_cache[cache_key] = product

                            selected_product_id = None
                            new_qty_in_box = None
                            new_is_excise = False
                            if product is not None:
                                selected_product_id = int(product.id)
                                try:
                                    new_qty_in_box = normalize_qty_in_box(product.qty_in_box)
                                except ValueError:
                                    new_qty_in_box = None
                                new_is_excise = bool(product.is_excise)
                                matched += 1

                            mappings.append({
                                "supplier_article": row.get("supplier_article") or None,
                                "product_name": row.get("product_name") or None,
                                "price": row.get("price"),
                                "price_pack": row.get("price_pack"),
                                "price_box": row.get("price_box"),
                                "qty_pcs": row.get("qty_pcs"),
                                "qty_box": row.get("qty_box"),
                                "volume_l": row.get("volume_l"),
                                "supplier_id": int(supplier_id),
                                "import_date": import_date,
                                "batch_id": batch_id,
                                "imported_by": imported_by,
                                "import_row_no": row.get("import_row_no"),
                                "selected_product_id": selected_product_id,
                                "new_product_name": None,
                                "new_brand": None,
                                "new_pack": None,
                                "new_qty_in_box": new_qty_in_box,
                                "new_is_excise": new_is_excise,
                            })

                        timings["matching"] = time.perf_counter() - started
                        progress(
                            f"Сопоставление завершено за {timings['matching']:.2f} с. "
                            f"Сопоставлено: {matched}/{len(rows)}."
                        )

                        started = time.perf_counter()
                        progress("Прайс поставщика: сохраняю staging...")

                        # Replace the current batch without intermediate flushes.
                        session.query(SupplierPriceCalculation).filter(
                            SupplierPriceCalculation.batch_id == batch_id,
                            SupplierPriceCalculation.imported_by == imported_by,
                        ).delete(synchronize_session=False)
                        session.query(TempPriceImport).filter(
                            TempPriceImport.batch_id == batch_id,
                            TempPriceImport.imported_by == imported_by,
                        ).delete(synchronize_session=False)

                        if mappings:
                            # One executemany insert instead of INSERT -> SELECT ->
                            # UPDATE for almost every automatically matched row.
                            session.execute(TempPriceImport.__table__.insert(), mappings)

                        session.commit()
                        timings["staging"] = time.perf_counter() - started
                        progress(
                            f"Staging сохранён за {timings['staging']:.2f} с."
                        )

                    timings["background_total"] = time.perf_counter() - total_started
                    progress(
                        "Импорт и сопоставление завершены за "
                        f"{timings['background_total']:.2f} с."
                    )
                    return {
                        "count": len(rows),
                        "matched": matched,
                        "file_path": file_path,
                        "timings": timings,
                    }
                finally:
                    SessionLocal.remove()

            def done(payload):
                page.selected_file_path = payload["file_path"]

                display_started = time.perf_counter()
                page.load_table_rows()
                display_elapsed = time.perf_counter() - display_started

                timings = payload.get("timings") or {}
                background_elapsed = timings.get("background_total")
                total_text = (
                    f"; фон: {background_elapsed:.2f} с"
                    if isinstance(background_elapsed, (int, float))
                    else ""
                )
                page.show_message(
                    f"Данные импортированы: {payload['count']}; "
                    f"автоматически сопоставлено: {payload['matched']}"
                    f"{total_text}; таблица: {display_elapsed:.2f} с"
                )

            pbi._start_page_task(
                page,
                title="Импорт и сопоставление прайса поставщика",
                work=work,
                on_finished=done,
                read_resources={"products", "product_articles"},
                write_resources={f"temp_price_import:{imported_by}"},
                intro=(
                    "Прайс: запускаю ускоренное чтение и сопоставление в фоне. "
                    "Можно работать в других вкладках."
                ),
                use_progress=True,
            )

        page._background_supplier_price_import_slot = import_background_fast
        page.ui.btn_Import.clicked.connect(import_background_fast)

    attach_supplier_prices_import_fast._supplier_price_import_bulk_fast = True
    attach_supplier_prices_import_fast._supplier_price_import_previous = current_attach
    pbi._attach_supplier_prices = attach_supplier_prices_import_fast


def install_supplier_price_runtime_fixes() -> None:
    global _installed
    if _installed:
        return
    _installed = True

    _install_costcalc_no_is_compatibility()
    _install_fast_supplier_price_import()
