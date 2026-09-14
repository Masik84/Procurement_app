from __future__ import annotations

import copy
import logging
from datetime import datetime
from pathlib import Path
from types import MethodType
from typing import Any, Callable

from PySide6.QtCore import QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import QFileDialog

from app.utils.background_tasks import TaskConflictError, get_background_task_manager
from app.utils.text import clean_multi_spaces

logger = logging.getLogger(__name__)


def _disconnect(signal, slot) -> None:
    try:
        signal.disconnect(slot)
    except (TypeError, RuntimeError):
        pass


def _set_page_busy(page, busy: bool, message: str | None = None) -> None:
    ui = getattr(page, "ui", None)
    if ui is not None:
        ui.setEnabled(not busy)
    if message and hasattr(page, "show_message"):
        page.show_message(message)


def _show_background_error(page, text: str) -> None:
    _set_page_busy(page, False)
    if hasattr(page, "show_error_message"):
        page.show_error_message(text)


def _start_page_task(
    page,
    *,
    title: str,
    work: Callable[[], Any],
    on_finished: Callable[[Any], None],
    read_resources=(),
    write_resources=(),
    intro: str,
) -> bool:
    try:
        handle = get_background_task_manager().start_task(
            title,
            work,
            read_resources=read_resources,
            write_resources=write_resources,
            owner=page,
        )
    except TaskConflictError as exc:
        page.show_error_message(str(exc))
        return False
    except Exception as exc:
        page.show_error_message(str(exc))
        return False

    _set_page_busy(page, True, intro)
    page._procurement_background_handle = handle

    def finished(result):
        page._procurement_background_handle = None
        _set_page_busy(page, False)
        on_finished(result)

    def failed(message: str):
        page._procurement_background_handle = None
        _show_background_error(page, message)

    handle.finished.connect(finished)
    handle.failed.connect(failed)
    return True




def _attach_foreground_save_guard(page, *, slot_name: str, write_resources) -> None:
    marker = f"_resource_guard_{slot_name}"
    if getattr(page, marker, False):
        return
    setattr(page, marker, True)
    original = getattr(page, slot_name, None)
    button = getattr(getattr(page, "ui", None), "btn_Save", None)
    if not callable(original) or button is None:
        return
    _disconnect(button.clicked, original)

    def guarded():
        try:
            get_background_task_manager().assert_available(write_resources=set(write_resources))
        except TaskConflictError as exc:
            if hasattr(page, "show_error_message"):
                page.show_error_message(str(exc))
            return
        original()

    setattr(page, f"_{slot_name}_resource_guard_slot", guarded)
    button.clicked.connect(guarded)

# ---------------------------------------------------------------------------
# Product Mapping / Portfolio
# ---------------------------------------------------------------------------

def _attach_product_mapping(page) -> None:
    if getattr(page, "_background_integration_product_mapping", False):
        return
    page._background_integration_product_mapping = True

    original_import = page.import_portfolio
    original_save = page.save
    page._sync_import_portfolio = original_import
    page._sync_save_product_mapping = original_save

    _disconnect(page.ui.btn_Import.clicked, original_import)
    _disconnect(page.ui.btn_Save.clicked, original_save)

    def import_background():
        file_path, _ = QFileDialog.getOpenFileName(
            page,
            "Выберите файл Портфель",
            "",
            "Excel files (*.xlsx *.xlsm *.xls)",
        )
        if not file_path:
            return

        def work():
            from app.db.db import SessionLocal
            from app.imports.product_mapping_portfolio_importer import ProductMappingPortfolioImporter
            from app.services.product_mapping_service import ProductMappingService

            try:
                portfolio_df = ProductMappingPortfolioImporter().read_excel(file_path)
                with SessionLocal() as session:
                    service = ProductMappingService(session)
                    deleted = service.cleanup_stale_new_links(portfolio_df)
                    session.commit()
                    result = service.check_products(portfolio_df)
                return {
                    "portfolio_df": portfolio_df,
                    "result": result,
                    "loaded_rows": len(portfolio_df),
                    "deleted": deleted,
                }
            finally:
                SessionLocal.remove()

        def done(payload):
            page._portfolio_df = payload["portfolio_df"].copy()
            page._clear_applied_left_filters()
            page._apply_check_result(payload["result"], loaded_rows=payload["loaded_rows"])

        _start_page_task(
            page,
            title="Сопоставление продуктов: Портфель",
            work=work,
            on_finished=done,
            read_resources={"products", "product_articles", "pack_types"},
            write_resources={"sales_product_links"},
            intro="Портфель загружен. Сопоставление выполняется в фоне — можно работать в других вкладках.",
        )

    def save_guarded():
        try:
            get_background_task_manager().assert_available(
                read_resources={"pack_types"},
                write_resources={"products", "product_articles", "sales_product_links"},
            )
        except TaskConflictError as exc:
            page.show_error_message(str(exc))
            return
        original_save()

    page._background_import_portfolio_slot = import_background
    page._guarded_product_mapping_save_slot = save_guarded
    page.ui.btn_Import.clicked.connect(import_background)
    page.ui.btn_Save.clicked.connect(save_guarded)


# ---------------------------------------------------------------------------
# Supplier Price
# ---------------------------------------------------------------------------

def _supplier_apply_changes_in_worker(
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
) -> dict:
    from app.db.db import SessionLocal
    from app.db.models import TempPriceImport
    from app.exports.supplier_price_exporter import SupplierPriceExporter
    from app.services.product_matching_service import MissingPackTypeError
    from app.services.qty_in_box_service import normalize_qty_in_box
    from app.services.supplier_price_service import SupplierPriceService
    from app.utils.parsers import parse_loose_number

    base_dir = Path(__file__).resolve().parents[2]
    try:
        with SessionLocal() as session:
            service = SupplierPriceService(session)
            try:
                for row_id, changes in pending_changes.items():
                    row = session.query(TempPriceImport).filter(TempPriceImport.id == row_id).first()
                    if row is None:
                        continue
                    row.supplier_id = supplier_id
                    row.import_date = import_date
                    for key, value in changes.items():
                        if key == "selected_product_id":
                            if value in (None, "", 0):
                                row.selected_product_id = None
                            else:
                                try:
                                    row.selected_product_id = int(value)
                                except (TypeError, ValueError):
                                    continue
                        elif key in {
                            "price", "price_pack", "price_box", "qty_pcs", "qty_box", "volume_l",
                            "new_pack", "new_qty_in_box",
                        }:
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
                        TempPriceImport.batch_id == batch_id,
                        TempPriceImport.imported_by == imported_by,
                    ).delete(synchronize_session=False)

                # autoflush=False: dependent service queries must see the edits.
                session.flush()
                service.validate_new_products_before_save(batch_id, imported_by)
                service.create_products_from_temp(batch_id, imported_by)
                service.automatch_remaining_rows_from_current_batch(batch_id, imported_by)
                qty_in_box_warnings = service.prepare_box_data_and_update_products(batch_id, imported_by)
                service.create_or_update_product_articles(batch_id, imported_by)
                service.fill_price_from_price_pack(batch_id, imported_by)

                saved_prices_count = 0
                if save_history:
                    saved_prices_count = service.save_prices_to_history_and_current(
                        batch_id=batch_id,
                        imported_by=imported_by,
                        currency_code=currency_code,
                        rf_prices_include_vat=rf_prices_include_vat,
                    )

                saved_calculations_count = service.save_supplier_price_calculations(
                    batch_id=batch_id,
                    imported_by=imported_by,
                    fx_rate=fx_rate,
                    currency_code=currency_code,
                    rf_prices_include_vat=rf_prices_include_vat,
                )
                session.commit()
            except MissingPackTypeError as exc:
                session.rollback()
                return {"status": "missing_pack", "error": exc}

            warning_path = None
            warning_error = None
            if qty_in_box_warnings:
                try:
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

            return {
                "status": "success",
                "saved_prices_count": saved_prices_count,
                "saved_calculations_count": saved_calculations_count,
                "warning_path": str(warning_path) if warning_path else "",
                "warning_error": warning_error or "",
                "rf_prices_include_vat": rf_prices_include_vat,
                "supplier_id": supplier_id,
            }
    finally:
        SessionLocal.remove()


def _attach_supplier_prices(page) -> None:
    if getattr(page, "_background_integration_supplier_prices", False):
        return
    page._background_integration_supplier_prices = True

    original_import = page.import_file
    original_save = page.apply_pending_changes
    page._sync_supplier_price_import = original_import
    page._sync_supplier_price_save = original_save
    _disconnect(page.ui.btn_Import.clicked, original_import)
    _disconnect(page.ui.btn_Save.clicked, original_save)

    def import_background():
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

        def work():
            from app.db.db import SessionLocal
            from app.imports.supplier_price_importer import SupplierPriceImporter
            from app.services.supplier_price_service import SupplierPriceService

            try:
                rows = SupplierPriceImporter().read_excel(file_path)
                with SessionLocal() as session:
                    service = SupplierPriceService(session)
                    service.import_rows_to_temp(
                        supplier_id=supplier_id,
                        batch_id=batch_id,
                        imported_by=imported_by,
                        rows=rows,
                        import_date=import_date,
                        replace_existing_batch_rows=True,
                    )
                    matched = service.automatch_temp_rows(batch_id, imported_by)
                    session.commit()
                return {"count": len(rows), "matched": matched, "file_path": file_path}
            finally:
                SessionLocal.remove()

        def done(payload):
            page.selected_file_path = payload["file_path"]
            page.load_table_rows()
            page.show_message(
                f"Данные импортированы: {payload['count']}; автоматически сопоставлено: {payload['matched']}"
            )

        _start_page_task(
            page,
            title="Импорт и сопоставление прайса поставщика",
            work=work,
            on_finished=done,
            read_resources={"products", "product_articles"},
            write_resources={f"temp_price_import:{imported_by}"},
            intro="Прайс импортируется и сопоставляется в фоне — можно работать в других вкладках.",
        )

    def save_background():
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
        except Exception as exc:
            page.show_error_message(str(exc))
            return

        snapshot = {
            "supplier_id": int(supplier_id),
            "currency_code": currency_code,
            "rf_prices_include_vat": bool(rf_prices_include_vat),
            "save_history": page.ui.cbo_History.currentText() == "да",
            "fx_rate": fx_rate,
            "import_date": page.get_price_date(),
            "batch_id": page.batch_id,
            "imported_by": page.imported_by,
            "pending_changes": copy.deepcopy(page._pending_changes),
            "pending_deletes": set(page._pending_deletes),
            "selected_file_path": page.selected_file_path,
        }

        def work():
            return _supplier_apply_changes_in_worker(**snapshot)

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
                    # Retry from a fresh DB session after the GUI has created the
                    # missing PackType pair and updated pending rows.
                    save_background()
                return

            page.rf_prices_include_vat = bool(payload.get("rf_prices_include_vat"))
            page._pending_changes.clear()
            page._pending_deletes.clear()

            warning_path = payload.get("warning_path") or ""
            if warning_path:
                QDesktopServices.openUrl(QUrl.fromLocalFile(str(warning_path)))

            saved_prices_count = int(payload.get("saved_prices_count") or 0)
            saved_calculations_count = int(payload.get("saved_calculations_count") or 0)
            warning_error = payload.get("warning_error") or ""

            if saved_calculations_count > 0 and page.ask_export_calculated_excel():
                quick_order_months, safe_stock_months = page.ask_order_planning_months_for_export()
                try:
                    export_started = page.export_calculated_excel(
                        int(payload["supplier_id"]),
                        quick_order_months=quick_order_months,
                        safe_stock_months=safe_stock_months,
                        cleanup_after_success=True,
                    )
                    if export_started:
                        page.show_message("Данные сохранены, Excel файл формируется в фоновом режиме")
                        if warning_error:
                            page.show_error_message(
                                f"Данные сохранены, но Warning-файл не удалось сохранить:\n{warning_error}"
                            )
                        return
                except Exception as exc:
                    warning_error = warning_error or str(exc)

            page.cleanup_current_batch(start_new_batch_after=True)
            page.reset_form_fields_after_successful_save()
            if saved_calculations_count == 0 and saved_prices_count == 0:
                page.show_message("Данные сохранены. Строки без цены использованы только для продуктов и связок")
            else:
                page.show_message("Данные сохранены")
            if warning_error:
                page.show_error_message(
                    f"Данные сохранены, но дополнительный Excel-файл не удалось выгрузить:\n{warning_error}"
                )

        _start_page_task(
            page,
            title="Расчёт прайса поставщика",
            work=work,
            on_finished=done,
            read_resources={
                "suppliers", "fixed_costs", "exchange_rates", "product_stock", "product_uc3_history"
            },
            write_resources={
                "products", "product_articles", "supplier_prices", "supplier_price_calculations",
                f"temp_price_import:{page.imported_by}",
            },
            intro="Прайс поставщика рассчитывается в фоне — можно работать в других вкладках.",
        )

    page._background_supplier_price_import_slot = import_background
    page._background_supplier_price_save_slot = save_background
    page.ui.btn_Import.clicked.connect(import_background)
    page.ui.btn_Save.clicked.connect(save_background)


# ---------------------------------------------------------------------------
# Target uC3: matching priority + remembered manual aliases
# ---------------------------------------------------------------------------

def _attach_product_uc3(page) -> None:
    if getattr(page, "_background_integration_product_uc3", False):
        return
    page._background_integration_product_uc3 = True

    original_import = page.import_excel
    original_save = page.apply_pending_changes
    original_finish_product_edit = page.finish_product_edit
    page._sync_product_uc3_import = original_import
    page._sync_product_uc3_save = original_save
    page._sync_product_uc3_finish_product_edit = original_finish_product_edit
    _disconnect(page.ui.btn_Import.clicked, original_import)
    _disconnect(page.ui.btn_Save.clicked, original_save)

    def finish_product_edit_with_alias_flag(_page, row: int, row_key: str, combo):
        original_finish_product_edit(row, row_key, combo)
        pending = page._pending_changes.get(row_key)
        if pending is None:
            return
        if clean_multi_spaces(pending.get("file_product_name")):
            # Any explicit user selection/correction becomes the remembered
            # name alias for subsequent Target uC3 imports.
            pending["manual_alias_required"] = True
            pending["match_source"] = "manual"

    page.finish_product_edit = MethodType(finish_product_edit_with_alias_flag, page)

    def import_with_saved_aliases():
        from app.imports.product_uc3_importer import ProductUc3Importer
        from app.services.product_uc3_matching_service import ProductUc3MatchingService

        try:
            file_path, _ = QFileDialog.getOpenFileName(
                page,
                "Выберите файл Target uC3",
                "",
                "Excel files (*.xls *.xlsx)",
            )
            if not file_path:
                return
            rows = ProductUc3Importer().read_excel(file_path)
            if not rows:
                page.show_message("Нет строк для импорта")
                return

            preview = []
            missing = []
            page._pending_changes.clear()
            page._pending_deletes.clear()
            page._deleted_row_snapshots.clear()
            page._new_rows.clear()

            with page.get_session() as session:
                matcher = ProductUc3MatchingService(session)
                for source in rows:
                    key = f"new::{page._temp_row_id}"
                    page._temp_row_id -= 1
                    source_product_name = clean_multi_spaces(source.get("product_name"))
                    match = matcher.resolve(source_product_name)
                    product_id = int(match.product_id) if match else None
                    our_product_name = match.product_name if match else ""
                    manual_alias_required = match is None
                    if manual_alias_required:
                        missing.append(source_product_name)
                    values = {
                        "product_id": product_id,
                        "product_name": our_product_name,
                        "file_product_name": source_product_name,
                        "target_uc3": source.get("target_uc3"),
                        "walk_away_uc3": source.get("walk_away_uc3"),
                        "manual_alias_required": manual_alias_required,
                        "match_source": match.source if match else "manual",
                    }
                    page._new_rows.add(key)
                    page._pending_changes[key] = values
                    preview.append({"row_key": key, **values, "change_date": None})

            page._populate_table(preview)
            if missing:
                page.show_message(f"Импорт: {len(preview)}; не найдено: {len(missing)}.")
            else:
                page.show_message(f"Импорт: {len(preview)} строк.")
        except Exception as exc:
            page.show_error_message(str(exc))

    def save_with_manual_aliases():
        from app.db.models import ProductUc3History
        from app.services.product_uc3_matching_service import ProductUc3MatchingService
        from app.services.product_uc3_service import ProductUc3Service
        from app.utils.gui_table_actions import commit_active_table_item_editors

        try:
            get_background_task_manager().assert_available(
                write_resources={"product_uc3_history", "product_articles"},
            )
        except TaskConflictError as exc:
            page.show_error_message(str(exc))
            return

        commit_active_table_item_editors(page.table)
        try:
            with page.get_session() as session:
                service = ProductUc3Service(session)
                matcher = ProductUc3MatchingService(session)
                for key in list(page._pending_deletes):
                    if isinstance(key, str) and key.startswith("db::"):
                        row_id = int(key.split("::", 1)[1])
                        session.query(ProductUc3History).filter(
                            ProductUc3History.id == row_id
                        ).delete(synchronize_session=False)
                session.flush()

                created = 0
                unchanged = 0
                reassigned = 0
                aliases_saved = 0
                for table_row in range(page.table.rowCount()):
                    key = page._row_key_at(table_row)
                    if key is None or key in page._pending_deletes:
                        continue
                    if key not in page._new_rows and key not in page._pending_changes:
                        continue
                    values = page._row_display_values(table_row, key)
                    product_id = values.get("product_id")
                    if not product_id:
                        raise ValueError(f"Строка {table_row + 1}: выберите Product Name")
                    product_id = int(product_id)

                    changes = page._pending_changes.get(key, {})
                    if key.startswith("db::"):
                        row_id = int(key.split("::", 1)[1])
                        history_row = (
                            session.query(ProductUc3History)
                            .filter(ProductUc3History.id == row_id)
                            .first()
                        )
                        if history_row is None:
                            continue
                        if int(history_row.product_id) != product_id:
                            history_row.product_id = product_id
                            session.flush()
                            reassigned += 1
                        if not ({"target_uc3", "walk_away_uc3"} & set(changes)):
                            continue

                    _obj, was_created = service.save_values(
                        product_id=product_id,
                        target_uc3=values.get("target_uc3"),
                        walk_away_uc3=values.get("walk_away_uc3"),
                    )
                    if was_created:
                        created += 1
                    else:
                        unchanged += 1

                    if values.get("manual_alias_required") and clean_multi_spaces(values.get("file_product_name")):
                        matcher.save_manual_alias(
                            product_id=product_id,
                            source_name=values.get("file_product_name"),
                        )
                        aliases_saved += 1

                session.commit()

            page._pending_changes.clear()
            page._pending_deletes.clear()
            page._deleted_row_snapshots.clear()
            page._new_rows.clear()
            page.find_rows()
            parts = [f"Сохранено новых значений: {created}", f"без изменений: {unchanged}"]
            if reassigned:
                parts.append(f"изменён Product: {reassigned}")
            if aliases_saved:
                parts.append(f"сохранено связок названий: {aliases_saved}")
            page.show_message("; ".join(parts))
        except Exception as exc:
            page.show_error_message(str(exc))

    page._target_uc3_import_alias_slot = import_with_saved_aliases
    page._target_uc3_save_alias_slot = save_with_manual_aliases
    page.ui.btn_Import.clicked.connect(import_with_saved_aliases)
    page.ui.btn_Save.clicked.connect(save_with_manual_aliases)


# ---------------------------------------------------------------------------
# Public lazy-page hook
# ---------------------------------------------------------------------------

def attach_page_background_integration(page) -> None:
    """Attach only the integrations relevant to this concrete page instance."""
    if getattr(page, "_procurement_page_integration_attached", False):
        return
    page._procurement_page_integration_attached = True

    class_name = page.__class__.__name__
    if class_name == "ProductMappingPage":
        _attach_product_mapping(page)
    elif class_name == "SupplierPricesPage":
        _attach_supplier_prices(page)
    elif class_name == "ProductUc3Page":
        _attach_product_uc3(page)
    elif class_name == "ProductsPage":
        _attach_foreground_save_guard(
            page, slot_name="apply_pending_changes", write_resources={"products"}
        )
    elif class_name == "ProductArticlesPage":
        _attach_foreground_save_guard(
            page, slot_name="apply_pending_changes", write_resources={"product_articles"}
        )
