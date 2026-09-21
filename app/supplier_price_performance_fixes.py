from __future__ import annotations

"""Performance fixes for Supplier Price save/background flow.

The project already installs narrow runtime compatibility modules from
``app.__init__``.  Keep the same pattern here so the change is isolated and can
be copied into an existing checkout without rewriting the large shared
background integration module.

What is fixed:
* only genuinely edited Supplier Price rows are snapshotted/sent to the worker;
* edited temp rows are loaded with one ``IN (...)`` query instead of N queries;
* Supplier/Product/date fields for the current batch are updated in one SQL
  statement;
* ProductArticle links use the existing batch helper;
* Qty-in-Box pack defaults are cached per SQLAlchemy Session;
* cost-calculation reference data is preloaded before the row loop;
* DB save/calculation and CostCalc Excel export are two separate background
  tasks, so DB save no longer waits for Excel COM;
* save stages report elapsed time in the background task window.
"""

import copy
import logging
import time
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from typing import Any

logger = logging.getLogger(__name__)
_installed = False


def _install_service_optimizations() -> None:
    from app.db.models import PackType, TempPriceImport
    from app.services import supplier_price_service as service_module
    from app.services.qty_in_box_service import SINGLE_ITEM_PACK_MARKERS
    from app.services.supplier_price_service import SupplierPriceService
    from app.utils.parsers import parse_loose_number
    from app.utils.text import clean_multi_spaces

    # ------------------------------------------------------------------
    # Qty in Box: one PackType load per Session instead of one SELECT/row.
    # SupplierPriceService imported this helper directly, so patch that local
    # symbol only; other pages/services retain their original behaviour.
    # ------------------------------------------------------------------
    if not getattr(service_module.default_qty_in_box_for_pack, "_supplier_price_cached", False):

        def cached_default_qty_in_box_for_pack(session, pack: object) -> int | None:
            cache_key = "_supplier_price_default_qty_by_pack"
            cache = session.info.get(cache_key)
            if cache is None:
                cache = {}
                for pack_type in session.query(PackType).all():
                    volume = parse_loose_number(pack_type.volume)
                    if volume is None:
                        continue
                    pack_name = clean_multi_spaces(pack_type.name).upper()
                    cache[Decimal(volume)] = (
                        1 if any(marker in pack_name for marker in SINGLE_ITEM_PACK_MARKERS) else None
                    )
                session.info[cache_key] = cache

            pack_number = parse_loose_number(pack)
            if pack_number is None:
                return None
            return cache.get(Decimal(pack_number))

        cached_default_qty_in_box_for_pack._supplier_price_cached = True
        service_module.default_qty_in_box_for_pack = cached_default_qty_in_box_for_pack

    # ------------------------------------------------------------------
    # ProductArticle: reuse the batch implementation that is already the
    # common mechanism in ProductMatchingService.
    # ------------------------------------------------------------------
    original_articles = SupplierPriceService.create_or_update_product_articles
    if not getattr(original_articles, "_supplier_price_batch_fast", False):

        def create_or_update_product_articles_fast(self, batch_id: str, imported_by: str) -> int:
            rows = (
                self.session.query(TempPriceImport)
                .filter(
                    TempPriceImport.batch_id == batch_id,
                    TempPriceImport.imported_by == imported_by,
                    TempPriceImport.selected_product_id.isnot(None),
                )
                .order_by(TempPriceImport.import_row_no.asc(), TempPriceImport.id.asc())
                .all()
            )
            if not rows:
                return 0

            links: list[tuple[int, object, object]] = []
            for row in rows:
                product_id = int(row.selected_product_id)
                safe_name = clean_multi_spaces(row.product_name)
                tokens = self.product_matching_service.split_article_tokens(row.supplier_article)
                if tokens:
                    links.extend((product_id, token, safe_name or None) for token in tokens)
                elif safe_name:
                    links.append((product_id, None, safe_name))

            if links:
                self.product_matching_service.create_product_articles_if_missing_batch(links)
            self.session.flush()
            # Preserve the method's existing meaning: count processed temp rows,
            # not only newly inserted ProductArticle records.
            return len(rows)

        create_or_update_product_articles_fast._supplier_price_batch_fast = True
        create_or_update_product_articles_fast._supplier_price_original = original_articles
        SupplierPriceService.create_or_update_product_articles = create_or_update_product_articles_fast

    # ------------------------------------------------------------------
    # Cost calculation: warm Product/Supplier/FixedCosts/PackType/MarkingRate
    # and FX caches in a fixed number of queries before the calculation loop.
    # ------------------------------------------------------------------
    original_calc = SupplierPriceService.save_supplier_price_calculations
    if not getattr(original_calc, "_supplier_price_preload_fast", False):

        def save_supplier_price_calculations_fast(
            self,
            batch_id: str,
            imported_by: str,
            fx_rate,
            currency_code: str,
            rf_prices_include_vat: bool = False,
        ) -> int:
            refs = (
                self.session.query(
                    TempPriceImport.selected_product_id,
                    TempPriceImport.supplier_id,
                )
                .filter(
                    TempPriceImport.batch_id == batch_id,
                    TempPriceImport.imported_by == imported_by,
                    TempPriceImport.selected_product_id.isnot(None),
                    TempPriceImport.price.isnot(None),
                )
                .all()
            )
            self.currency_cost_service.preload_reference_data(
                product_ids={
                    int(product_id)
                    for product_id, _supplier_id in refs
                    if product_id is not None
                },
                supplier_ids={
                    int(supplier_id)
                    for _product_id, supplier_id in refs
                    if supplier_id is not None
                },
            )
            return original_calc(
                self,
                batch_id=batch_id,
                imported_by=imported_by,
                fx_rate=fx_rate,
                currency_code=currency_code,
                rf_prices_include_vat=rf_prices_include_vat,
            )

        save_supplier_price_calculations_fast._supplier_price_preload_fast = True
        save_supplier_price_calculations_fast._supplier_price_original = original_calc
        SupplierPriceService.save_supplier_price_calculations = save_supplier_price_calculations_fast


def _apply_pending_changes_bulk(
    *,
    session,
    TempPriceImport,
    pending_changes: dict[int, dict],
    pending_deletes: set[int],
    supplier_id: int,
    import_date,
    batch_id: str,
    imported_by: str,
) -> None:
    """Apply GUI edits using O(1) SQL round-trips with respect to row count."""
    from app.services.qty_in_box_service import normalize_qty_in_box
    from app.utils.parsers import parse_loose_number
    from app.utils.text import clean_multi_spaces

    batch_filter = (
        TempPriceImport.batch_id == batch_id,
        TempPriceImport.imported_by == imported_by,
    )

    # Supplier/date are properties of the whole imported batch.  Updating them
    # once is both faster and more correct when the user changed form fields
    # after Import but did not touch every table row.
    session.query(TempPriceImport).filter(*batch_filter).update(
        {
            TempPriceImport.supplier_id: supplier_id,
            TempPriceImport.import_date: import_date,
        },
        synchronize_session=False,
    )

    row_ids: list[int] = []
    for row_id in pending_changes:
        try:
            row_ids.append(int(row_id))
        except (TypeError, ValueError):
            continue

    rows_by_id = {}
    if row_ids:
        rows = (
            session.query(TempPriceImport)
            .filter(
                TempPriceImport.id.in_(row_ids),
                *batch_filter,
            )
            .all()
        )
        rows_by_id = {int(row.id): row for row in rows}

    numeric_fields = {
        "price", "price_pack", "price_box", "qty_pcs", "qty_box", "volume_l",
        "new_pack", "new_qty_in_box",
    }

    for raw_row_id, changes in pending_changes.items():
        try:
            row = rows_by_id.get(int(raw_row_id))
        except (TypeError, ValueError):
            row = None
        if row is None:
            continue

        for key, value in changes.items():
            if key == "selected_product_id":
                if value in (None, "", 0):
                    row.selected_product_id = None
                else:
                    try:
                        row.selected_product_id = int(value)
                    except (TypeError, ValueError):
                        continue
            elif key in numeric_fields:
                parsed = parse_loose_number(value)
                if key == "new_qty_in_box":
                    parsed = normalize_qty_in_box(parsed, field_name="Qty in Box (for new)")
                setattr(row, key, parsed if parsed is not None else None)
            else:
                setattr(row, key, value)

        if row.selected_product_id is not None:
            row.new_brand = None
            row.new_pack = None
        else:
            has_new_product_data = any([
                bool(clean_multi_spaces(row.new_product_name)),
                bool(clean_multi_spaces(row.new_brand)),
                row.new_pack is not None,
                row.new_qty_in_box is not None,
                bool(row.new_is_excise),
            ])
            if has_new_product_data and row.new_is_excise is None:
                row.new_is_excise = False

    if pending_deletes:
        session.query(TempPriceImport).filter(
            TempPriceImport.id.in_(pending_deletes),
            *batch_filter,
        ).delete(synchronize_session=False)

    # SessionLocal has autoflush=False; all service queries below must see edits.
    session.flush()


def _supplier_apply_changes_in_worker_fast(
    *,
    supplier_id: int,
    currency_code: str,
    rf_prices_include_vat: bool,
    save_history: bool,
    fx_rate,
    import_date,
    batch_id: str,
    imported_by: str,
    pending_changes: dict[int, dict],
    pending_deletes: set[int],
    selected_file_path: str,
    export_output_path: str = "",
    quick_order_months: int | None = None,
    safe_stock_months: int | None = None,
    supplier_price_age_months: int = 3,
    progress=None,
) -> dict:
    """Fast DB save worker.  CostCalc Excel is intentionally NOT generated here."""
    from app.db.db import SessionLocal
    from app.db.models import TempPriceImport
    from app.exports.supplier_price_exporter import SupplierPriceExporter
    from app.services.product_matching_service import MissingPackTypeError
    from app.services.supplier_price_service import SupplierPriceService

    report = progress or (lambda _text: None)
    base_dir = Path(__file__).resolve().parents[1]
    total_started = time.perf_counter()
    timings: dict[str, float] = {}

    def stage_done(key: str, started: float, caption: str) -> None:
        elapsed = time.perf_counter() - started
        timings[key] = elapsed
        report(f"{caption}: {elapsed:.2f} с")

    try:
        with SessionLocal() as session:
            service = SupplierPriceService(session)
            try:
                started = time.perf_counter()
                report("Прайс поставщика: сохраняю изменения строк...")
                _apply_pending_changes_bulk(
                    session=session,
                    TempPriceImport=TempPriceImport,
                    pending_changes=pending_changes,
                    pending_deletes=pending_deletes,
                    supplier_id=supplier_id,
                    import_date=import_date,
                    batch_id=batch_id,
                    imported_by=imported_by,
                )
                stage_done("rows", started, "Изменения строк сохранены")

                started = time.perf_counter()
                report("Прайс поставщика: проверяю и создаю новые продукты...")
                service.validate_new_products_before_save(batch_id, imported_by)
                service.create_products_from_temp(batch_id, imported_by)
                service.automatch_remaining_rows_from_current_batch(batch_id, imported_by)
                stage_done("products", started, "Продукты проверены/сопоставлены")

                started = time.perf_counter()
                report("Прайс поставщика: проверяю упаковки и связки артикулов...")
                qty_in_box_warnings = service.prepare_box_data_and_update_products(batch_id, imported_by)
                service.create_or_update_product_articles(batch_id, imported_by)
                service.fill_price_from_price_pack(batch_id, imported_by)
                stage_done("links", started, "Упаковки и связки обработаны")

                saved_prices_count = 0
                if save_history:
                    started = time.perf_counter()
                    report("Прайс поставщика: сохраняю историю цен...")
                    saved_prices_count = service.save_prices_to_history_and_current(
                        batch_id=batch_id,
                        imported_by=imported_by,
                        currency_code=currency_code,
                        rf_prices_include_vat=rf_prices_include_vat,
                    )
                    stage_done("history", started, "История цен сохранена")

                started = time.perf_counter()
                report("Прайс поставщика: рассчитываю себестоимость...")
                saved_calculations_count = service.save_supplier_price_calculations(
                    batch_id=batch_id,
                    imported_by=imported_by,
                    fx_rate=fx_rate,
                    currency_code=currency_code,
                    rf_prices_include_vat=rf_prices_include_vat,
                )
                stage_done("calculation", started, "Себестоимость рассчитана")

                started = time.perf_counter()
                session.commit()
                stage_done("commit", started, "Данные записаны в БД")
            except MissingPackTypeError as exc:
                session.rollback()
                return {"status": "missing_pack", "error": exc}

            warning_path = None
            warning_error = None
            if qty_in_box_warnings:
                try:
                    report("Прайс поставщика: формирую предупреждения Qty in Box...")
                    output_dir = (
                        Path(selected_file_path).resolve().parent
                        if selected_file_path
                        else base_dir
                    )
                    output_path = output_dir / f"Warning_Qty_in_Box_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
                    warning_path = SupplierPriceExporter(session).export_qty_in_box_warnings(
                        qty_in_box_warnings,
                        output_path,
                    )
                except Exception as exc:  # warning export must not roll back saved data
                    warning_error = str(exc)

            timings["total"] = time.perf_counter() - total_started
            report(f"Сохранение прайса завершено за {timings['total']:.2f} с.")
            return {
                "status": "success",
                "saved_prices_count": saved_prices_count,
                "saved_calculations_count": saved_calculations_count,
                "warning_path": str(warning_path) if warning_path else "",
                "warning_error": warning_error or "",
                "rf_prices_include_vat": rf_prices_include_vat,
                "supplier_id": supplier_id,
                "timings": timings,
                # Kept in the signature/snapshot for compatibility, but Excel
                # runs in a second task after this worker has finished.
                "export_requested": bool(export_output_path),
            }
    finally:
        SessionLocal.remove()


def _install_page_flow() -> None:
    from PySide6.QtCore import QUrl
    from PySide6.QtGui import QDesktopServices
    from PySide6.QtWidgets import QFileDialog

    from app.utils import page_background_integration as pbi
    from app.utils.text import clean_multi_spaces

    original_attach = pbi._attach_supplier_prices
    if getattr(original_attach, "_supplier_price_performance_fast", False):
        return

    def attach_supplier_prices_fast(page) -> None:
        # Patch instance-level GUI tracking BEFORE the original integration
        # stores/connects its callbacks.
        if not getattr(page, "_supplier_price_pending_tracking_fast", False):
            page._supplier_price_pending_tracking_fast = True

            original_display_rows = page.display_rows

            def display_rows_fast(rows):
                original_display_rows(rows)
                # The legacy display method creates an empty dict for every row.
                # Empty entries are not edits and must never be sent to DB save.
                page._pending_changes = {
                    row_id: changes
                    for row_id, changes in page._pending_changes.items()
                    if changes
                }

            page.display_rows = display_rows_fast

            def snapshot_changed_rows_only() -> None:
                """Snapshot final cell text only for rows already marked dirty."""
                if not page._pending_changes:
                    return
                pending_ids = set(page._pending_changes)
                for row, row_id in enumerate(page._table_row_ids):
                    if row >= page.table.rowCount():
                        break
                    if row_id not in pending_ids:
                        continue
                    changes = page._pending_changes[row_id]
                    for column, column_name in enumerate(page.columns):
                        # Combo/checkbox changes have dedicated handlers and are
                        # already present in _pending_changes.
                        if column_name in {"selected_product_id", "new_is_excise"}:
                            continue
                        item = page.table.item(row, column)
                        if item is None:
                            continue
                        value = clean_multi_spaces(item.text()).upper()
                        changes[column_name] = value or None

            page._snapshot_table_values = snapshot_changed_rows_only

        original_attach(page)

        # Replace only the Supplier Price Save slot installed by the original
        # background integration.  Import remains unchanged.
        old_save_slot = getattr(page, "_background_supplier_price_save_slot", None)
        if old_save_slot is not None:
            pbi._disconnect(page.ui.btn_Save.clicked, old_save_slot)

        def choose_costcalc_output_path() -> str:
            supplier_name = (
                clean_multi_spaces(page.ui.line_NewSupplier.text())
                or clean_multi_spaces(page.ui.cbo_SupplName.currentText())
                or "Supplier"
            )
            safe_supplier_name = supplier_name
            for ch in ['\\', '/', ':', '*', '?', '"', '<', '>', '|']:
                safe_supplier_name = safe_supplier_name.replace(ch, '_')
            default_name = f"CostCalc_{safe_supplier_name}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
            base_dir = Path(__file__).resolve().parents[1]
            file_path, _ = QFileDialog.getSaveFileName(
                page,
                "Сохранить Excel файл",
                str(base_dir / default_name),
                "Excel files (*.xlsx)",
            )
            if not file_path:
                return ""
            if not file_path.lower().endswith(".xlsx"):
                file_path += ".xlsx"
            return file_path

        def save_background_fast() -> None:
            from app.db.models import TempPriceImport
            from app.utils.pack_type_prompt import resolve_missing_pack_for_temp_rows

            page._commit_open_editors()
            has_batch_rows = page._has_rows_in_current_batch()
            if not page._pending_changes and not page._pending_deletes and not has_batch_rows:
                page.show_message("Нет изменений")
                return

            try:
                supplier_id = page.ensure_supplier(save_existing_changes=True)
                currency_code = clean_multi_spaces(page.ui.cbo_Currency.currentText()).upper()
                if not currency_code or currency_code == "-":
                    raise ValueError("Выбери валюту")

                rf_prices_include_vat = False
                if page.ui.cbo_SupplierRF.currentText() == "да":
                    rf_prices_include_vat = page.ask_rf_prices_include_vat()
                fx_rate = page.parse_decimal_field(page.ui.line_ExchangeRate, "Курс")

                export_output_path = ""
                quick_order_months = None
                safe_stock_months = None
                if page.ask_export_calculated_excel():
                    quick_order_months, safe_stock_months = page.ask_order_planning_months_for_export()
                    # Cancel in the period dialog means: save DB, do not export.
                    if safe_stock_months is not None:
                        export_output_path = choose_costcalc_output_path()

                export_config = {
                    "export_output_path": export_output_path,
                    "quick_order_months": quick_order_months,
                    "safe_stock_months": safe_stock_months,
                    "supplier_price_age_months": page.get_supplier_price_age_months(),
                }
            except Exception as exc:
                page.show_error_message(str(exc))
                return

            def build_snapshot() -> dict:
                # Keep only real edits.  This is intentionally repeated here in
                # case an older display path left empty placeholders behind.
                pending_changes = {
                    int(row_id): copy.deepcopy(changes)
                    for row_id, changes in page._pending_changes.items()
                    if changes
                }
                return {
                    "supplier_id": int(supplier_id),
                    "currency_code": currency_code,
                    "rf_prices_include_vat": bool(rf_prices_include_vat),
                    "save_history": page.ui.cbo_History.currentText() == "да",
                    "fx_rate": fx_rate,
                    "import_date": page.get_price_date(),
                    "batch_id": page.batch_id,
                    "imported_by": page.imported_by,
                    "pending_changes": pending_changes,
                    "pending_deletes": set(page._pending_deletes),
                    "selected_file_path": page.selected_file_path,
                    **export_config,
                }

            def finalize_saved(
                *,
                saved_prices_count: int,
                saved_calculations_count: int,
                warning_error: str = "",
                export_path: str = "",
                export_error: str = "",
            ) -> None:
                page.cleanup_current_batch(start_new_batch_after=True)
                page.reset_form_fields_after_successful_save()

                if export_path:
                    page.show_message("Данные сохранены, CostCalc Excel сохранён")
                elif saved_calculations_count == 0 and saved_prices_count == 0:
                    page.show_message(
                        "Данные сохранены. Строки без цены использованы только для продуктов и связок"
                    )
                else:
                    page.show_message("Данные сохранены")

                errors: list[str] = []
                if warning_error:
                    errors.append(f"Warning-файл не удалось сохранить: {warning_error}")
                if export_error:
                    errors.append(f"CostCalc Excel не удалось сохранить: {export_error}")
                if errors:
                    page.show_error_message("Данные в БД сохранены, но:\n" + "\n".join(errors))

            def launch_costcalc_export(snapshot: dict, save_payload: dict) -> bool:
                output_path = snapshot.get("export_output_path") or ""
                if not output_path:
                    return False

                def export_work(progress):
                    from app.db.db import SessionLocal
                    from app.exports.supplier_price_exporter import SupplierPriceExporter

                    try:
                        started = time.perf_counter()
                        progress("CostCalc: формирую Excel...")
                        with SessionLocal() as session:
                            try:
                                result_path = SupplierPriceExporter(session).export_calculated(
                                    batch_id=snapshot["batch_id"],
                                    imported_by=snapshot["imported_by"],
                                    supplier_id=snapshot["supplier_id"],
                                    output_path=output_path,
                                    source_file_path=snapshot.get("selected_file_path") or None,
                                    quick_order_months=snapshot.get("quick_order_months"),
                                    safe_stock_months=snapshot.get("safe_stock_months"),
                                    supplier_price_age_months=snapshot.get("supplier_price_age_months"),
                                )
                                elapsed = time.perf_counter() - started
                                progress(f"CostCalc Excel сформирован за {elapsed:.2f} с.")
                                return {"status": "success", "path": str(result_path)}
                            except Exception as exc:
                                logger.exception("CostCalc Excel export failed after DB save")
                                return {"status": "error", "error": str(exc)}
                    finally:
                        SessionLocal.remove()

                def export_done(export_payload):
                    export_path = ""
                    export_error = ""
                    if export_payload.get("status") == "success":
                        export_path = export_payload.get("path") or ""
                        if export_path:
                            QDesktopServices.openUrl(QUrl.fromLocalFile(str(export_path)))
                    else:
                        export_error = export_payload.get("error") or "Неизвестная ошибка формирования CostCalc Excel"

                    finalize_saved(
                        saved_prices_count=int(save_payload.get("saved_prices_count") or 0),
                        saved_calculations_count=int(save_payload.get("saved_calculations_count") or 0),
                        warning_error=save_payload.get("warning_error") or "",
                        export_path=export_path,
                        export_error=export_error,
                    )

                started = pbi._start_page_task(
                    page,
                    title="Формирование CostCalc Excel",
                    work=export_work,
                    on_finished=export_done,
                    read_resources={
                        "products", "product_articles", "supplier_prices",
                        "supplier_price_calculations", "product_stock", "product_uc3_history",
                        f"temp_price_import:{page.imported_by}",
                    },
                    write_resources=set(),
                    intro="Данные в БД сохранены. CostCalc Excel формируется отдельной фоновой задачей.",
                    use_progress=True,
                )
                return bool(started)

            def launch_save(snapshot: dict) -> None:
                def work(progress):
                    return _supplier_apply_changes_in_worker_fast(**snapshot, progress=progress)

                def done(payload):
                    if payload.get("status") == "missing_pack":
                        error = payload.get("error")
                        if resolve_missing_pack_for_temp_rows(
                            page,
                            error,
                            model=TempPriceImport,
                            batch_id=page.batch_id,
                            imported_by=page.imported_by,
                            pending_changes=page._pending_changes,
                        ):
                            retry_snapshot = build_snapshot()
                            launch_save(retry_snapshot)
                        return

                    page.rf_prices_include_vat = bool(payload.get("rf_prices_include_vat"))
                    page._pending_changes.clear()
                    page._pending_deletes.clear()

                    warning_path = payload.get("warning_path") or ""
                    if warning_path:
                        QDesktopServices.openUrl(QUrl.fromLocalFile(str(warning_path)))

                    saved_prices_count = int(payload.get("saved_prices_count") or 0)
                    saved_calculations_count = int(payload.get("saved_calculations_count") or 0)

                    if snapshot.get("export_output_path") and saved_calculations_count > 0:
                        if launch_costcalc_export(snapshot, payload):
                            return
                        finalize_saved(
                            saved_prices_count=saved_prices_count,
                            saved_calculations_count=saved_calculations_count,
                            warning_error=payload.get("warning_error") or "",
                            export_error="Не удалось запустить фоновую задачу CostCalc Excel",
                        )
                        return

                    finalize_saved(
                        saved_prices_count=saved_prices_count,
                        saved_calculations_count=saved_calculations_count,
                        warning_error=payload.get("warning_error") or "",
                    )

                pbi._start_page_task(
                    page,
                    title="Сохранение и расчёт прайса поставщика",
                    work=work,
                    on_finished=done,
                    read_resources={
                        "suppliers", "fixed_costs", "exchange_rates",
                        "product_stock", "product_uc3_history",
                    },
                    write_resources={
                        "products", "product_articles", "supplier_prices",
                        "supplier_price_calculations", f"temp_price_import:{page.imported_by}",
                    },
                    intro="Прайс: сохраняю и рассчитываю в фоне. Можно работать в других вкладках.",
                    use_progress=True,
                )

            launch_save(build_snapshot())

        page._background_supplier_price_save_slot = save_background_fast
        page.ui.btn_Save.clicked.connect(save_background_fast)

    attach_supplier_prices_fast._supplier_price_performance_fast = True
    attach_supplier_prices_fast._supplier_price_original = original_attach
    pbi._attach_supplier_prices = attach_supplier_prices_fast


def install_supplier_price_performance_fixes() -> None:
    global _installed
    if _installed:
        return
    _installed = True
    _install_service_optimizations()
    _install_page_flow()
