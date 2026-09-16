from __future__ import annotations

import copy
import logging
from datetime import datetime
from pathlib import Path
from types import MethodType
from typing import Any, Callable

from PySide6.QtCore import QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import QAbstractButton, QFileDialog

from app.utils.background_tasks import TaskConflictError, get_background_task_manager
from app.utils.text import clean_multi_spaces

logger = logging.getLogger(__name__)


def _disconnect(signal, slot) -> None:
    try:
        signal.disconnect(slot)
    except (TypeError, RuntimeError):
        pass


def _ensure_tab_close_guard(page) -> None:
    """Do not detach a page while one of its background tasks is still active."""
    window = page.window()
    ui = getattr(window, "ui", None)
    tab_widget = getattr(ui, "tabWidget", None) if ui is not None else None
    original_close = getattr(window, "close_tab", None)
    if tab_widget is None or not callable(original_close):
        return
    if getattr(window, "_procurement_tab_close_guard_installed", False):
        return

    window._procurement_tab_close_guard_installed = True
    window._procurement_original_close_tab = original_close
    _disconnect(tab_widget.tabCloseRequested, original_close)

    def guarded_close(index: int) -> None:
        owner = tab_widget.widget(index) if 0 <= index < tab_widget.count() else None
        manager = get_background_task_manager()
        if owner is not None and manager.has_active_tasks(owner):
            try:
                from app.utils.message_dialogs import show_warning

                show_warning(
                    window,
                    "На этой вкладке ещё выполняется фоновая операция. "
                    "Дождитесь её завершения, после этого вкладку можно закрыть.",
                    title="Фоновая операция",
                )
            except Exception:
                logger.exception("Не удалось показать предупреждение о фоновой операции")
            return
        original_close(index)

    window._procurement_guarded_close_tab = guarded_close
    tab_widget.tabCloseRequested.connect(guarded_close)


def _set_page_busy(page, busy: bool, message: str | None = None) -> None:
    """Disable only controls of the page that owns the running task.

    This mirrors Daily-Report--new-: the current page cannot start another
    import/save/change-mode action while its worker is running, but the main
    tab widget and every other page stay fully usable.  Previous Procurement
    revisions either disabled the whole page root or left every button active;
    both behaviours were wrong.
    """
    page._procurement_background_busy = bool(busy)
    ui = getattr(page, "ui", None)

    if busy:
        if ui is not None and not getattr(page, "_procurement_busy_widgets", None):
            saved: list[tuple[object, bool]] = []
            seen: set[int] = set()

            # All page buttons (Import/Save/Reset/Add/Search/mode radio buttons
            # etc.) are unsafe to press while the page's own data task is active.
            try:
                buttons = ui.findChildren(QAbstractButton)
            except Exception:
                buttons = []
            for widget in buttons:
                marker = id(widget)
                if marker in seen:
                    continue
                seen.add(marker)
                try:
                    saved.append((widget, bool(widget.isEnabled())))
                    widget.setEnabled(False)
                except RuntimeError:
                    pass

            # Editing the table while the worker is replacing/saving its temp
            # rows can create a race.  Block only that table, not the whole page.
            table = getattr(page, "table", None)
            if table is not None and id(table) not in seen:
                seen.add(id(table))
                try:
                    saved.append((table, bool(table.isEnabled())))
                    table.setEnabled(False)
                except RuntimeError:
                    pass

            # Daily Report also freezes mode/source selectors that determine what
            # the running task means.  Keep the list narrow and page-local.
            critical_names = {
                "SupplierPricesPage": (
                    "cbo_SupplName", "line_NewSupplier", "cbo_Currency", "date_Price",
                    "cbo_SupplierRF", "cbo_viaNovo", "cbo_Customs", "cbo_Marking",
                    "cbo_History", "cbo_Rating",
                ),
                "ProductStockPage": (),  # mode selectors are radio buttons above
                "CustomerCostsPage": ("spb_SuppPriceAge",),
                "TargetPricesPage": (),
                "OrderPlanningPage": (),
            }.get(page.__class__.__name__, ())
            for name in critical_names:
                widget = getattr(ui, name, None)
                if widget is None or id(widget) in seen:
                    continue
                seen.add(id(widget))
                try:
                    saved.append((widget, bool(widget.isEnabled())))
                    widget.setEnabled(False)
                except RuntimeError:
                    pass

            page._procurement_busy_widgets = saved
    else:
        saved = list(getattr(page, "_procurement_busy_widgets", []) or [])
        page._procurement_busy_widgets = []
        for widget, was_enabled in saved:
            try:
                widget.setEnabled(bool(was_enabled))
            except RuntimeError:
                pass

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
    use_progress: bool = False,
) -> bool:
    def finished(result):
        page._procurement_background_handle = None
        _set_page_busy(page, False)
        on_finished(result)

    def failed(message: str):
        page._procurement_background_handle = None
        _show_background_error(page, message)

    try:
        manager = get_background_task_manager()
        # A page keeps a single active task handle.  Do not allow a second task
        # from the same page to overwrite that handle while keeping the whole
        # UI enabled.  Tasks from other pages are still allowed whenever their
        # declared resources do not conflict.
        if manager.has_active_tasks(page):
            if hasattr(page, "show_message"):
                page.show_message(
                    "На этой вкладке уже выполняется фоновая операция. "
                    "Можно продолжать работать в других вкладках."
                )
            return False

        handle = manager.start_task(
            title,
            work,
            read_resources=read_resources,
            write_resources=write_resources,
            owner=page,
            on_finished=finished,
            on_failed=failed,
            use_progress=use_progress,
            intro=intro,
        )
    except TaskConflictError as exc:
        page.show_error_message(str(exc))
        return False
    except Exception as exc:
        page.show_error_message(str(exc))
        return False

    page._procurement_background_handle = handle
    if use_progress and hasattr(page, "show_message"):
        handle.progress.connect(
            lambda text: page.show_message(" ".join(str(text).split())[:240])
        )
    _set_page_busy(page, True, intro)
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
    original_save_excel = page.save_excel
    page._sync_import_portfolio = original_import
    page._sync_save_product_mapping = original_save
    page._sync_save_excel_product_mapping = original_save_excel

    _disconnect(page.ui.btn_Import.clicked, original_import)
    _disconnect(page.ui.btn_Save.clicked, original_save)
    _disconnect(page.ui.btn_SaveExcel.clicked, original_save_excel)

    def import_background():
        file_path, _ = QFileDialog.getOpenFileName(
            page,
            "Выберите файл Портфель",
            "",
            "Excel files (*.xlsx *.xlsm *.xls)",
        )
        if not file_path:
            return

        def work(progress):
            from app.db.db import SessionLocal
            from app.imports.product_mapping_portfolio_importer import ProductMappingPortfolioImporter
            from app.services.product_mapping_service import ProductMappingService

            try:
                progress("Сопоставление продуктов: читаю Портфель...")
                portfolio_df = ProductMappingPortfolioImporter().read_excel(file_path)
                progress(f"Сопоставление продуктов: прочитано строк: {len(portfolio_df)}")
                with SessionLocal() as session:
                    service = ProductMappingService(session)
                    progress("Сопоставление продуктов: очищаю устаревшие NEW-связки...")
                    deleted = service.cleanup_stale_new_links(portfolio_df)
                    session.commit()
                    progress("Сопоставление продуктов: выполняю сопоставление...")
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
            intro="Портфель: запускаю чтение и сопоставление в фоне. Можно работать в других вкладках.",
            use_progress=True,
        )

    def save_background():
        from app.services.product_matching_service import MissingPackTypeError
        from app.utils.pack_type_prompt import ask_pack_type
        from app.utils.parsers import parse_loose_number

        page._commit_open_product_editor()
        pending_deletes = {
            clean_multi_spaces(code)
            for code in page._pending_deletes
            if clean_multi_spaces(code)
        }
        base_rows = page._all_rows if page._mode == "check" else page._rows
        rows = [
            copy.deepcopy(row)
            for row in base_rows
            if clean_multi_spaces(row.get("sales_code")) not in pending_deletes
            and not bool(row.get("is_ignored"))
        ]
        if not rows and not pending_deletes:
            page.show_message("Нет строк для сохранения")
            return
        mode = page._mode
        portfolio_df = page._portfolio_df.copy() if page._portfolio_df is not None else None

        def launch(rows_snapshot):
            def work(progress):
                from app.db.db import SessionLocal
                from app.services.product_mapping_service import ProductMappingService

                try:
                    with SessionLocal() as session:
                        service = ProductMappingService(session)
                        try:
                            progress("Сопоставление продуктов: сохраняю удаления...")
                            deleted = service.delete_links(pending_deletes)
                            progress("Сопоставление продуктов: сохраняю выбранные связки...")
                            count = service.save_rows(rows_snapshot) if rows_snapshot else 0
                            session.commit()
                            result = None
                            if mode == "check" and portfolio_df is not None:
                                progress("Сопоставление продуктов: перепроверяю текущий Портфель...")
                                result = service.check_products(portfolio_df)
                            return {"status": "success", "count": count, "deleted": deleted, "result": result}
                        except MissingPackTypeError as exc:
                            session.rollback()
                            return {"status": "missing_pack", "error": exc}
                finally:
                    SessionLocal.remove()

            def done(payload):
                if payload.get("status") == "missing_pack":
                    exc = payload.get("error")
                    selected_pack = ask_pack_type(page, exc)
                    if selected_pack is None:
                        return
                    requested_pack = parse_loose_number(exc.requested_pack)
                    for row in page._all_rows:
                        if row.get("product_id"):
                            continue
                        pack = row.get("new_pack") if row.get("new_pack") not in (None, "") else row.get("sales_pack")
                        if parse_loose_number(pack) == requested_pack:
                            row["new_pack"] = selected_pack
                            row["sales_pack"] = selected_pack
                    page.apply_filters()
                    refreshed = [
                        copy.deepcopy(row)
                        for row in (page._all_rows if page._mode == "check" else page._rows)
                        if clean_multi_spaces(row.get("sales_code")) not in pending_deletes
                        and not bool(row.get("is_ignored"))
                    ]
                    launch(refreshed)
                    return

                page._pending_deletes.clear()
                page._deleted_row_snapshots.clear()
                if mode == "check" and payload.get("result") is not None:
                    page._apply_check_result(payload["result"])
                else:
                    page.search_saved()
                count = int(payload.get("count") or 0)
                deleted = int(payload.get("deleted") or 0)
                if deleted:
                    page.show_message(f"Сопоставления сохранены: {count}; удалено: {deleted}")
                else:
                    page.show_message(f"Сопоставления сохранены: {count}")

            _start_page_task(
                page,
                title="Сопоставление продуктов: сохранение",
                work=work,
                on_finished=done,
                read_resources={"pack_types"},
                write_resources={"products", "product_articles", "sales_product_links"},
                intro="Сопоставление продуктов сохраняется в фоне. Можно работать в других вкладках.",
                use_progress=True,
            )

        launch(rows)

    def save_excel_background():
        if not page._rows:
            page.show_message("Нет строк для выгрузки")
            return
        base_dir = Path(__file__).resolve().parents[2]
        default = f"Product_Mapping_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
        path, _ = QFileDialog.getSaveFileName(page, "Сохранить Excel", str(base_dir / default), "Excel files (*.xlsx)")
        if not path:
            return
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"
        data = [
            {header: row.get(field) for field, header in zip(page.COLUMNS, page.HEADERS)}
            for row in page._rows
        ]
        from app.workers.excel_export_worker import start_excel_export

        def do_export():
            import pandas as pd
            pd.DataFrame(data).to_excel(path, index=False)
            return path

        def done(output_path):
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(output_path)))
            page.show_message("Excel файл сохранен")

        if not start_excel_export(page, do_export, on_finished=done, on_error=page.show_error_message, button=page.ui.btn_SaveExcel):
            page.show_message("Excel файл уже формируется")

    page._background_import_portfolio_slot = import_background
    page._background_save_product_mapping_slot = save_background
    page._background_save_excel_product_mapping_slot = save_excel_background
    page.ui.btn_Import.clicked.connect(import_background)
    page.ui.btn_Save.clicked.connect(save_background)
    page.ui.btn_SaveExcel.clicked.connect(save_excel_background)


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
    export_output_path: str = "",
    quick_order_months: int | None = None,
    safe_stock_months: int | None = None,
    supplier_price_age_months: int = 3,
    progress: Callable[[str], None] | None = None,
) -> dict:
    from app.db.db import SessionLocal
    from app.db.models import TempPriceImport
    from app.exports.supplier_price_exporter import SupplierPriceExporter
    from app.services.product_matching_service import MissingPackTypeError
    from app.services.qty_in_box_service import normalize_qty_in_box
    from app.services.supplier_price_service import SupplierPriceService
    from app.utils.parsers import parse_loose_number

    report = progress or (lambda _text: None)
    base_dir = Path(__file__).resolve().parents[2]
    try:
        with SessionLocal() as session:
            service = SupplierPriceService(session)
            try:
                report("Прайс поставщика: сохраняю изменения строк...")
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

                # SessionLocal uses autoflush=False. All subsequent service queries
                # must see the edits made above.
                session.flush()

                report("Прайс поставщика: проверяю новые продукты...")
                service.validate_new_products_before_save(batch_id, imported_by)
                service.create_products_from_temp(batch_id, imported_by)

                report("Прайс поставщика: сопоставляю продукты и упаковки...")
                service.automatch_remaining_rows_from_current_batch(batch_id, imported_by)
                qty_in_box_warnings = service.prepare_box_data_and_update_products(batch_id, imported_by)
                service.create_or_update_product_articles(batch_id, imported_by)
                service.fill_price_from_price_pack(batch_id, imported_by)

                saved_prices_count = 0
                if save_history:
                    report("Прайс поставщика: сохраняю историю цен...")
                    saved_prices_count = service.save_prices_to_history_and_current(
                        batch_id=batch_id,
                        imported_by=imported_by,
                        currency_code=currency_code,
                        rf_prices_include_vat=rf_prices_include_vat,
                    )

                report("Прайс поставщика: рассчитываю себестоимость...")
                saved_calculations_count = service.save_supplier_price_calculations(
                    batch_id=batch_id,
                    imported_by=imported_by,
                    fx_rate=fx_rate,
                    currency_code=currency_code,
                    rf_prices_include_vat=rf_prices_include_vat,
                )
                session.commit()
                report("Прайс поставщика: данные сохранены в БД.")
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

            calculated_export_path = None
            calculated_export_error = None
            if export_output_path and saved_calculations_count > 0:
                try:
                    report("Прайс поставщика: формирую CostCalc Excel...")
                    calculated_export_path = SupplierPriceExporter(session).export_calculated(
                        batch_id=batch_id,
                        imported_by=imported_by,
                        supplier_id=supplier_id,
                        output_path=export_output_path,
                        source_file_path=selected_file_path or None,
                        quick_order_months=quick_order_months,
                        safe_stock_months=safe_stock_months,
                        supplier_price_age_months=supplier_price_age_months,
                    )
                    report("Прайс поставщика: CostCalc Excel сформирован.")
                except Exception as exc:
                    calculated_export_error = str(exc)

            return {
                "status": "success",
                "saved_prices_count": saved_prices_count,
                "saved_calculations_count": saved_calculations_count,
                "warning_path": str(warning_path) if warning_path else "",
                "warning_error": warning_error or "",
                "export_path": str(calculated_export_path) if calculated_export_path else "",
                "export_error": calculated_export_error or "",
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

        def work(progress):
            from app.db.db import SessionLocal
            from app.imports.supplier_price_importer import SupplierPriceImporter
            from app.services.supplier_price_service import SupplierPriceService

            try:
                progress("Прайс поставщика: читаю Excel...")
                logger.info("BACKGROUND SUPPLIER PRICE | READ EXCEL | file=%s", file_path)
                rows = SupplierPriceImporter().read_excel(file_path)
                progress(f"Прайс поставщика: прочитано строк: {len(rows)}. Сохраняю staging...")
                logger.info("BACKGROUND SUPPLIER PRICE | EXCEL ROWS | %s", len(rows))
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
                    progress("Прайс поставщика: выполняю автоматическое сопоставление...")
                    matched = service.automatch_temp_rows(batch_id, imported_by)
                    session.commit()
                    logger.info(
                        "BACKGROUND SUPPLIER PRICE | COMMIT | rows=%s | matched=%s",
                        len(rows),
                        matched,
                    )
                progress("Прайс поставщика: импорт завершён.")
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
            intro="Прайс: запускаю чтение и сопоставление в фоне. Можно работать в других вкладках.",
            use_progress=True,
        )

    def _choose_costcalc_output_path() -> str:
        supplier_name = (
            clean_multi_spaces(page.ui.line_NewSupplier.text())
            or clean_multi_spaces(page.ui.cbo_SupplName.currentText())
            or "Supplier"
        )
        safe_supplier_name = supplier_name
        for ch in ['\\', '/', ':', '*', '?', '"', '<', '>', '|']:
            safe_supplier_name = safe_supplier_name.replace(ch, '_')
        default_name = f"CostCalc_{safe_supplier_name}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
        base_dir = Path(__file__).resolve().parents[2]
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

    def save_background():
        from app.db.models import TempPriceImport
        from app.utils.pack_type_prompt import resolve_missing_pack_for_temp_rows

        # GUI PRE-FLIGHT. All questions and file dialogs must happen BEFORE the
        # background task starts. A worker must never wait for hidden GUI input.
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
                # Cancel in the period dialog means: save data, but do not export.
                if safe_stock_months is not None:
                    export_output_path = _choose_costcalc_output_path()

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
            return {
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
                **export_config,
            }

        def launch(snapshot: dict) -> None:
            def work(progress):
                return _supplier_apply_changes_in_worker(**snapshot, progress=progress)

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
                        # Keep the user's already selected period/path; only refresh
                        # the mutable table snapshot and retry in a fresh DB session.
                        retry_snapshot = build_snapshot()
                        launch(retry_snapshot)
                    return

                page.rf_prices_include_vat = bool(payload.get("rf_prices_include_vat"))
                page._pending_changes.clear()
                page._pending_deletes.clear()

                warning_path = payload.get("warning_path") or ""
                if warning_path:
                    QDesktopServices.openUrl(QUrl.fromLocalFile(str(warning_path)))

                export_path = payload.get("export_path") or ""
                if export_path:
                    QDesktopServices.openUrl(QUrl.fromLocalFile(str(export_path)))

                saved_prices_count = int(payload.get("saved_prices_count") or 0)
                saved_calculations_count = int(payload.get("saved_calculations_count") or 0)
                warning_error = payload.get("warning_error") or ""
                export_error = payload.get("export_error") or ""

                page.cleanup_current_batch(start_new_batch_after=True)
                page.reset_form_fields_after_successful_save()

                if export_path:
                    page.show_message("Данные сохранены, CostCalc Excel сохранён")
                elif saved_calculations_count == 0 and saved_prices_count == 0:
                    page.show_message("Данные сохранены. Строки без цены использованы только для продуктов и связок")
                else:
                    page.show_message("Данные сохранены")

                errors = []
                if warning_error:
                    errors.append(f"Warning-файл не удалось сохранить: {warning_error}")
                if export_error:
                    errors.append(f"CostCalc Excel не удалось сохранить: {export_error}")
                if errors:
                    page.show_error_message("Данные в БД сохранены, но:\n" + "\n".join(errors))

            _start_page_task(
                page,
                title="Сохранение и расчёт прайса поставщика",
                work=work,
                on_finished=done,
                read_resources={
                    "suppliers", "fixed_costs", "exchange_rates", "product_stock", "product_uc3_history"
                },
                write_resources={
                    "products", "product_articles", "supplier_prices", "supplier_price_calculations",
                    f"temp_price_import:{page.imported_by}",
                },
                intro="Прайс: сохраняю и рассчитываю в фоне. Можно работать в других вкладках.",
                use_progress=True,
            )

        launch(build_snapshot())

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
    original_find_rows = page.find_rows
    original_reset_all = page.reset_all
    original_on_current_changed = page.on_current_changed
    page._sync_product_uc3_import = original_import
    page._sync_product_uc3_save = original_save
    page._sync_product_uc3_finish_product_edit = original_finish_product_edit
    page._sync_product_uc3_find_rows = original_find_rows
    page._sync_product_uc3_reset_all = original_reset_all
    _disconnect(page.ui.btn_Import.clicked, original_import)
    _disconnect(page.ui.btn_Save.clicked, original_save)
    _disconnect(page.ui.btn_Search.clicked, original_find_rows)
    _disconnect(page.ui.btn_ResetAll.clicked, original_reset_all)
    _disconnect(page.ui.chb_CalcToday.stateChanged, original_on_current_changed)

    def _set_import_preview_mode(enabled: bool) -> None:
        # Product Name from file is an import-preview helper only.  Keep the
        # physical column because the legacy row/index code uses fixed column
        # numbers, but make it completely invisible outside Excel import.
        if page.table.columnCount() > 1:
            page.table.setColumnHidden(1, not enabled)

    def _clear_target_table() -> None:
        page._updating_table = True
        try:
            page.table.setSortingEnabled(False)
            page.table.clearContents()
            page.table.setRowCount(0)
        finally:
            page.table.setSortingEnabled(True)
            page._updating_table = False
        _set_import_preview_mode(False)

    def find_rows_only_by_search(_page=None):
        _set_import_preview_mode(False)
        original_find_rows()
        _set_import_preview_mode(False)

    def reset_without_search():
        # The original Reset correctly resets all filter widgets, but it also
        # calls self.find_rows().  Suppress that reload: Target uC3 must query
        # the DB only after an explicit Search click.
        page._target_uc3_suppress_find = True
        try:
            original_reset_all()
        finally:
            page._target_uc3_suppress_find = False
        page._pending_changes.clear()
        page._pending_deletes.clear()
        page._deleted_row_snapshots.clear()
        page._new_rows.clear()
        _clear_target_table()
        page.show_message("Фильтры и поля сброшены")

    def controlled_find_rows(_page=None):
        if getattr(page, "_target_uc3_suppress_find", False):
            _clear_target_table()
            return
        find_rows_only_by_search()

    page.find_rows = MethodType(controlled_find_rows, page)
    page._target_uc3_search_slot = find_rows_only_by_search
    page._target_uc3_reset_slot = reset_without_search
    page.ui.btn_Search.clicked.connect(find_rows_only_by_search)
    page.ui.btn_ResetAll.clicked.connect(reset_without_search)
    # Switching Current/history only changes the search criteria. It must not
    # trigger a database query by itself.
    page.ui.chb_CalcToday.stateChanged.connect(lambda _state: None)

    # __init__ in the legacy page loads DB rows immediately.  The integration
    # is attached on first Show, so remove those rows before the page becomes
    # usable.  From this point onward Search is the only normal DB loader.
    _clear_target_table()

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
        file_path, _ = QFileDialog.getOpenFileName(
            page,
            "Выберите файл Target uC3",
            "",
            "Excel files (*.xls *.xlsx)",
        )
        if not file_path:
            return

        def work():
            from app.db.db import SessionLocal
            from app.imports.product_uc3_importer import ProductUc3Importer
            from app.services.product_uc3_matching_service import ProductUc3MatchingService

            try:
                rows = ProductUc3Importer().read_excel(file_path)
                prepared = []
                with SessionLocal() as session:
                    matcher = ProductUc3MatchingService(session)
                    for source in rows:
                        source_product_name = clean_multi_spaces(source.get("product_name"))
                        match = matcher.resolve(source_product_name)
                        prepared.append({
                            "product_id": int(match.product_id) if match else None,
                            "product_name": match.product_name if match else "",
                            "file_product_name": source_product_name,
                            "target_uc3": source.get("target_uc3"),
                            "walk_away_uc3": source.get("walk_away_uc3"),
                            "manual_alias_required": match is None,
                            "match_source": match.source if match else "manual",
                        })
                return prepared
            finally:
                SessionLocal.remove()

        def done(prepared):
            if not prepared:
                page.show_message("Нет строк для импорта")
                return

            preview = []
            missing = 0
            page._pending_changes.clear()
            page._pending_deletes.clear()
            page._deleted_row_snapshots.clear()
            page._new_rows.clear()

            for values in prepared:
                key = f"new::{page._temp_row_id}"
                page._temp_row_id -= 1
                values = dict(values)
                if values.get("manual_alias_required"):
                    missing += 1
                page._new_rows.add(key)
                page._pending_changes[key] = values
                preview.append({"row_key": key, **values, "change_date": None})

            page._populate_table(preview)
            _set_import_preview_mode(True)
            if missing:
                page.show_message(f"Импорт: {len(preview)}; не найдено: {missing}.")
            else:
                page.show_message(f"Импорт: {len(preview)} строк.")

        _start_page_task(
            page,
            title="Target uC3: импорт и сопоставление",
            work=work,
            on_finished=done,
            read_resources={"products", "product_articles"},
            intro="Target uC3 сопоставляется в фоне — можно работать в других вкладках.",
        )

    def save_with_manual_aliases():
        from app.utils.gui_table_actions import commit_active_table_item_editors

        commit_active_table_item_editors(page.table)
        rows_snapshot = []
        for table_row in range(page.table.rowCount()):
            key = page._row_key_at(table_row)
            if key is None or key in page._pending_deletes:
                continue
            if key not in page._new_rows and key not in page._pending_changes:
                continue
            values = copy.deepcopy(page._row_display_values(table_row, key))
            product_id = values.get("product_id")
            if not product_id:
                page.show_error_message(f"Строка {table_row + 1}: выберите Product Name")
                return
            rows_snapshot.append({
                "key": str(key),
                "values": values,
                "changes": copy.deepcopy(page._pending_changes.get(key, {})),
            })
        pending_deletes = set(page._pending_deletes)
        if not rows_snapshot and not pending_deletes:
            page.show_message("Нет изменений")
            return

        def work(progress):
            from app.db.db import SessionLocal
            from app.db.models import ProductUc3History
            from app.services.product_uc3_matching_service import ProductUc3MatchingService
            from app.services.product_uc3_service import ProductUc3Service

            try:
                with SessionLocal() as session:
                    service = ProductUc3Service(session)
                    matcher = ProductUc3MatchingService(session)
                    progress("Target uC3: удаляю отмеченные строки...")
                    for key in pending_deletes:
                        if isinstance(key, str) and key.startswith("db::"):
                            row_id = int(key.split("::", 1)[1])
                            session.query(ProductUc3History).filter(
                                ProductUc3History.id == row_id
                            ).delete(synchronize_session=False)
                    session.flush()

                    created = unchanged = reassigned = aliases_saved = 0
                    total = max(1, len(rows_snapshot))
                    for idx, item in enumerate(rows_snapshot, start=1):
                        values = item["values"]
                        changes = item["changes"]
                        key = item["key"]
                        product_id = int(values["product_id"])
                        if idx == 1 or idx == total or idx % 25 == 0:
                            progress(f"Target uC3: сохраняю строки {idx}/{total}...")

                        if key.startswith("db::"):
                            row_id = int(key.split("::", 1)[1])
                            history_row = session.query(ProductUc3History).filter(
                                ProductUc3History.id == row_id
                            ).first()
                            if history_row is None:
                                continue
                            if int(history_row.product_id) != product_id:
                                history_row.product_id = product_id
                                session.flush()
                                reassigned += 1
                            if not ({"target_uc3", "walk_away_uc3"} & set(changes)):
                                if values.get("manual_alias_required") and clean_multi_spaces(values.get("file_product_name")):
                                    matcher.save_manual_alias(product_id=product_id, source_name=values.get("file_product_name"))
                                    aliases_saved += 1
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
                    return {
                        "created": created,
                        "unchanged": unchanged,
                        "reassigned": reassigned,
                        "aliases_saved": aliases_saved,
                    }
            finally:
                SessionLocal.remove()

        def done(payload):
            page._pending_changes.clear()
            page._pending_deletes.clear()
            page._deleted_row_snapshots.clear()
            page._new_rows.clear()
            _clear_target_table()
            parts = [
                f"Сохранено новых значений: {payload['created']}",
                f"без изменений: {payload['unchanged']}",
            ]
            if payload.get("reassigned"):
                parts.append(f"изменён Product: {payload['reassigned']}")
            if payload.get("aliases_saved"):
                parts.append(f"сохранено связок названий: {payload['aliases_saved']}")
            parts.append("Для загрузки данных из БД нажмите Search")
            page.show_message("; ".join(parts))

        _start_page_task(
            page,
            title="Target uC3: сохранение",
            work=work,
            on_finished=done,
            write_resources={"product_uc3_history", "product_articles"},
            intro="Target uC3 сохраняется в фоне. Можно работать в других вкладках.",
            use_progress=True,
        )

    page._target_uc3_import_alias_slot = import_with_saved_aliases
    page._target_uc3_save_alias_slot = save_with_manual_aliases
    page.ui.btn_Import.clicked.connect(import_with_saved_aliases)
    page.ui.btn_Save.clicked.connect(save_with_manual_aliases)



# ---------------------------------------------------------------------------
# Stock / Supplier Orders / IS
# ---------------------------------------------------------------------------

def _ensure_product_stock_primary_key_alias() -> None:
    """Compatibility for the legacy no-IS runtime wrapper.

    ProductStock is keyed by ``product_id`` and intentionally has no separate
    ``id`` column. An older runtime compatibility wrapper addressed this model
    through ``.id`` while preserving legacy IS fields during Supplier Orders
    save. Registering an ORM synonym keeps that wrapper compatible without a DB
    migration and without creating a fake database column.
    """
    from sqlalchemy.orm import synonym
    from app.db.models import ProductStock

    mapper = ProductStock.__mapper__
    if "id" not in mapper.attrs:
        mapper.add_property("id", synonym("product_id"))


def _attach_product_stock(page) -> None:
    if getattr(page, "_background_integration_product_stock", False):
        return
    page._background_integration_product_stock = True
    _ensure_product_stock_primary_key_alias()

    original_import = page.import_file
    original_save = page.save_all
    page._sync_product_stock_import = original_import
    page._sync_product_stock_save = original_save
    _disconnect(page.ui.btn_Import.clicked, original_import)
    _disconnect(page.ui.btn_Save.clicked, original_save)

    def import_background():
        file_path, _ = QFileDialog.getOpenFileName(
            page,
            "Выберите Excel файл",
            "",
            "Excel files (*.xls *.xlsx *.xlsm)",
        )
        if not file_path:
            return

        page.start_new_batch()
        page._current_file_path = file_path
        mode = page._mode
        batch_id = page._batch_id
        imported_by = page._imported_by

        def work(progress):
            from app.db.db import SessionLocal
            from app.services.product_stock_service import ProductStockService

            progress("Остатки/Заказы/IS: читаю Excel...")
            logger.info(
                "BACKGROUND STOCK IMPORT | READ EXCEL | mode=%s | file=%s",
                mode,
                file_path,
            )
            try:
                with SessionLocal() as session:
                    runner = ProductStockService(session)

                    # Do not call runner.import_* here. Those legacy convenience
                    # methods run cleanup_old_for_all(), which touches EVERY temp
                    # table in one transaction. That hidden write can make an
                    # unrelated Supplier Price import wait on the stock/order task.
                    # The page already performs daily stale-temp cleanup when it
                    # opens; for this batch we only replace its own rows.
                    if mode == "stock":
                        rows = runner.stock_importer.read_excel(file_path)
                        logger.info("BACKGROUND STOCK IMPORT | EXCEL ROWS | %s", len(rows))
                        imported_count = runner.import_stock_rows(
                            rows=rows,
                            batch_id=batch_id,
                            imported_by=imported_by,
                            replace_existing=True,
                        )
                        matched_count = runner.automatch_stock_rows(batch_id, imported_by)
                    elif mode == "orders":
                        rows = runner.supplier_orders_importer.read_excel(file_path)
                        logger.info("BACKGROUND ORDERS IMPORT | EXCEL ROWS | %s", len(rows))
                        imported_count = runner.import_supplier_orders_rows(
                            rows=rows,
                            batch_id=batch_id,
                            imported_by=imported_by,
                            replace_existing=True,
                        )
                        matched_count = runner.automatch_supplier_orders_rows(batch_id, imported_by)
                    else:
                        rows = runner.is_importer.read_excel(file_path)
                        logger.info("BACKGROUND IS IMPORT | EXCEL ROWS | %s", len(rows))
                        imported_count = runner.import_is_rows(
                            rows=rows,
                            batch_id=batch_id,
                            imported_by=imported_by,
                            replace_existing=True,
                        )
                        matched_count = runner.automatch_is_rows(batch_id, imported_by)

                    progress("Остатки/Заказы/IS: сохраняю staging и сопоставление...")
                    session.commit()
                    logger.info(
                        "BACKGROUND STOCK IMPORT | COMMIT | mode=%s | imported=%s | matched=%s",
                        mode,
                        imported_count,
                        matched_count,
                    )
                    return {
                        "batch_id": batch_id,
                        "imported_count": imported_count,
                        "matched_count": matched_count,
                    }
            finally:
                SessionLocal.remove()

        def done(result):
            page.load_table()
            page.offer_save_issue_file()
            imported = int(result.get("imported_count") or 0)
            matched = int(result.get("matched_count") or 0)
            page.show_message(f"Импорт завершен. Строк: {imported}; сопоставлено: {matched}.")

        _start_page_task(
            page,
            title={
                "stock": "Импорт остатков",
                "orders": "Импорт заказов поставщиков",
                "is": "Импорт IS",
            }.get(mode, "Импорт остатков/заказов"),
            work=work,
            on_finished=done,
            read_resources={"products", "product_articles"},
            write_resources={f"temp_product_stock:{imported_by}"},
            intro="Импорт и сопоставление выполняются в фоне — можно работать в других вкладках.",
            use_progress=True,
        )

    def save_background():
        from app.db.models import TempIsImport, TempStockImport, TempSupplierOrdersImport
        from app.services.product_matching_service import MissingPackTypeError
        from app.services.product_stock_service import ProductStockService
        from app.utils.pack_type_prompt import resolve_missing_pack_for_temp_rows

        page._commit_open_editors()
        mode = page._mode
        batch_id = page._batch_id
        imported_by = page._imported_by
        pending_deletes = set(getattr(page, "_pending_deletes", set()) or set())
        model_by_mode = {
            "stock": TempStockImport,
            "orders": TempSupplierOrdersImport,
            "is": TempIsImport,
        }
        temp_model = model_by_mode[mode]

        def work(progress):
            from app.db.db import SessionLocal

            try:
                with SessionLocal() as session:
                    runner = ProductStockService(session)
                    try:
                        progress("Остатки/Заказы/IS: проверяю строки перед сохранением...")
                        if pending_deletes:
                            session.query(temp_model).filter(
                                temp_model.id.in_(pending_deletes),
                                temp_model.batch_id == batch_id,
                                temp_model.imported_by == imported_by,
                            ).delete(synchronize_session=False)
                            session.flush()

                        rows = (
                            session.query(temp_model)
                            .filter(
                                temp_model.batch_id == batch_id,
                                temp_model.imported_by == imported_by,
                            )
                            .all()
                        )
                        if not rows:
                            raise ValueError(
                                "Нет данных для сохранения. Сначала импортируйте файл или добавьте строки."
                            )
                        matched_count = sum(
                            1 for row in rows if row.selected_product_id is not None
                        )
                        if matched_count == 0:
                            can_auto_create = (
                                mode in {"stock", "orders"}
                                and any(ProductStockService.can_create_new_product(row) for row in rows)
                            )
                            if not can_auto_create:
                                raise ValueError(
                                    "В текущем импорте нет ни одной строки с SelectedProductID."
                                )

                        progress("Остатки/Заказы/IS: записываю данные в БД...")
                        if mode == "stock":
                            stats = runner.save_stock(batch_id, imported_by)
                            message = "Данные по остаткам успешно сохранены."
                        elif mode == "orders":
                            stats = runner.save_supplier_orders(batch_id, imported_by)
                            message = "Данные по заказам поставщиков успешно сохранены."
                        else:
                            stats = runner.save_is(batch_id, imported_by)
                            message = "Данные IS успешно сохранены."
                        session.commit()
                        return {"status": "success", "stats": stats, "message": message}
                    except MissingPackTypeError as exc:
                        session.rollback()
                        return {"status": "missing_pack", "error": exc}
            finally:
                SessionLocal.remove()

        def done(payload):
            if payload.get("status") == "missing_pack":
                error = payload.get("error")
                if resolve_missing_pack_for_temp_rows(
                    page,
                    error,
                    model=temp_model,
                    batch_id=batch_id,
                    imported_by=imported_by,
                ):
                    page.load_table()
                    save_background()
                return

            pending = getattr(page, "_pending_deletes", None)
            if isinstance(pending, set):
                pending.clear()
            snapshots = getattr(page, "_deleted_row_snapshots", None)
            if isinstance(snapshots, list):
                snapshots.clear()
            page.start_new_batch()
            page.clear_table()
            page.refresh_counters()
            page.show_message(payload.get("message") or "Данные сохранены.")

        _start_page_task(
            page,
            title={
                "stock": "Сохранение остатков",
                "orders": "Сохранение заказов поставщиков",
                "is": "Сохранение IS",
            }.get(mode, "Сохранение остатков/заказов"),
            work=work,
            on_finished=done,
            read_resources={"pack_types"},
            write_resources={
                "products",
                "product_articles",
                "product_stock",
                f"temp_product_stock:{imported_by}",
            },
            intro="Сохранение выполняется в фоне — можно работать в других вкладках.",
            use_progress=True,
        )

    page._background_product_stock_import_slot = import_background
    page._background_product_stock_save_slot = save_background
    page.ui.btn_Import.clicked.connect(import_background)
    page.ui.btn_Save.clicked.connect(save_background)


# ---------------------------------------------------------------------------
# Product Search
# ---------------------------------------------------------------------------

def _attach_product_search(page) -> None:
    if getattr(page, "_background_integration_product_search", False):
        return
    page._background_integration_product_search = True

    original_import = page.import_file
    original_save = page.apply_pending_changes
    _disconnect(page.ui.btn_Import.clicked, original_import)
    _disconnect(page.ui.btn_Save.clicked, original_save)

    def import_background():
        file_path, _ = QFileDialog.getOpenFileName(
            page,
            "Выберите файл для поиска продуктов",
            "",
            "Excel files (*.xls *.xlsx)",
        )
        if not file_path:
            return
        batch_id = page.batch_id
        imported_by = page.imported_by

        def work(progress):
            from app.db.db import SessionLocal
            from app.imports.product_search_importer import ProductSearchImporter
            from app.services.product_search_service import ProductSearchService

            try:
                progress("Поиск продуктов: читаю Excel...")
                rows = ProductSearchImporter().read_excel(file_path)
                progress(f"Поиск продуктов: прочитано строк: {len(rows)}. Загружаю staging...")
                with SessionLocal() as session:
                    service = ProductSearchService(session)
                    service.import_rows_to_temp(
                        batch_id=batch_id,
                        imported_by=imported_by,
                        rows=rows,
                        import_date=datetime.now(),
                        replace_existing_batch_rows=True,
                    )
                    progress("Поиск продуктов: выполняю автоматическое сопоставление...")
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
            title="Поиск продуктов: импорт и сопоставление",
            work=work,
            on_finished=done,
            read_resources={"products", "product_articles"},
            write_resources={f"temp_product_search:{imported_by}"},
            intro="Поиск продуктов: импорт и сопоставление выполняются в фоне.",
            use_progress=True,
        )

    def save_background():
        from app.utils.message_dialogs import ask_yes_no

        page._commit_open_editors()
        if not page._pending_changes and not page._pending_deletes:
            page.show_message("Нет изменений")
            return

        save_to_excel = ask_yes_no(
            page,
            "Сохранить в Excel?",
            title="Сохранение",
            default_yes=False,
        )
        save_path = ""
        if save_to_excel:
            base_dir = Path(__file__).resolve().parents[2]
            save_path, _ = QFileDialog.getSaveFileName(
                page,
                "Сохранить итоговый файл",
                str(base_dir / f"ProductSearch_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"),
                "Excel files (*.xlsx)",
            )
            if save_path and not save_path.lower().endswith(".xlsx"):
                save_path += ".xlsx"
            if not save_path:
                save_to_excel = False

        batch_id = page.batch_id
        imported_by = page.imported_by
        pending_changes = copy.deepcopy(page._pending_changes)
        pending_deletes = set(page._pending_deletes)

        def launch(changes_snapshot, deletes_snapshot):
            def work(progress):
                from sqlalchemy.orm import joinedload
                from app.db.db import SessionLocal
                from app.db.models import TempProductSearchImport
                from app.exports.product_search_exporter import ProductSearchExporter
                from app.services.product_matching_service import MissingPackTypeError
                from app.services.product_search_service import ProductSearchService
                from app.services.qty_in_box_service import normalize_qty_in_box
                from app.utils.parsers import parse_loose_number

                try:
                    with SessionLocal() as session:
                        service = ProductSearchService(session)
                        try:
                            progress("Поиск продуктов: сохраняю изменения staging...")
                            for row_id, changes in changes_snapshot.items():
                                row = session.query(TempProductSearchImport).filter(
                                    TempProductSearchImport.id == row_id
                                ).first()
                                if row is None:
                                    continue
                                for key, value in changes.items():
                                    if key in {"new_pack", "new_qty_in_box"}:
                                        parsed = parse_loose_number(value)
                                        if key == "new_qty_in_box":
                                            parsed = normalize_qty_in_box(parsed, field_name="Qty in Box (for new)")
                                        setattr(row, key, parsed if parsed is not None else None)
                                    else:
                                        setattr(row, key, value)
                                has_new_product_data = any([
                                    bool(clean_multi_spaces(row.new_product_name)),
                                    bool(clean_multi_spaces(row.new_brand)),
                                    row.new_pack is not None,
                                    row.new_qty_in_box is not None,
                                    bool(row.new_is_excise),
                                ])
                                if row.selected_product_id is None and has_new_product_data and row.new_is_excise is None:
                                    row.new_is_excise = False
                                if row.selected_product_id is not None:
                                    row.new_product_name = None
                                    row.new_brand = None
                                    row.new_pack = None
                                    row.new_qty_in_box = None
                                    row.new_is_excise = None
                            session.flush()
                            if deletes_snapshot:
                                session.query(TempProductSearchImport).filter(
                                    TempProductSearchImport.id.in_(deletes_snapshot),
                                    TempProductSearchImport.batch_id == batch_id,
                                    TempProductSearchImport.imported_by == imported_by,
                                ).delete(synchronize_session=False)
                                session.flush()

                            progress("Поиск продуктов: проверяю и создаю новые продукты...")
                            service.validate_new_products_before_save(batch_id, imported_by)
                            service.create_products_from_temp(batch_id, imported_by)
                            service.create_or_update_product_articles(batch_id, imported_by)
                            session.flush()

                            export_rows = []
                            if save_to_excel:
                                progress("Поиск продуктов: подготавливаю Excel...")
                                rows_for_export = session.query(TempProductSearchImport).options(
                                    joinedload(TempProductSearchImport.selected_product)
                                ).filter(
                                    TempProductSearchImport.batch_id == batch_id,
                                    TempProductSearchImport.imported_by == imported_by,
                                ).order_by(
                                    TempProductSearchImport.import_row_no.asc(),
                                    TempProductSearchImport.id.asc(),
                                ).all()
                                for row in rows_for_export:
                                    product = row.selected_product
                                    export_rows.append({
                                        "source_article": row.source_article or "",
                                        "source_product_name": row.source_product_name or "",
                                        "product_name": product.name if product else (row.new_product_name or ""),
                                        "brand": product.brand if product else (row.new_brand or ""),
                                        "pack": product.pack if product else row.new_pack,
                                        "qty_in_box": product.qty_in_box if product else row.new_qty_in_box,
                                        "abc_category": (product.abc_category or "-") if product else "-",
                                        "is_excise": product.is_excise if product else row.new_is_excise,
                                    })

                            service.delete_temp_rows_for_user(imported_by)
                            session.commit()

                            output = ""
                            export_error = ""
                            if save_to_excel and save_path and export_rows:
                                progress("Поиск продуктов: формирую итоговый Excel...")
                                try:
                                    output = ProductSearchExporter().export_result(save_path, export_rows)
                                except Exception as exc:
                                    export_error = str(exc)
                            return {"status": "success", "output": str(output or ""), "export_error": export_error}
                        except MissingPackTypeError as exc:
                            session.rollback()
                            return {"status": "missing_pack", "error": exc}
                finally:
                    SessionLocal.remove()

            def done(payload):
                if payload.get("status") == "missing_pack":
                    from app.db.models import TempProductSearchImport
                    from app.utils.pack_type_prompt import resolve_missing_pack_for_temp_rows
                    if resolve_missing_pack_for_temp_rows(
                        page,
                        payload.get("error"),
                        model=TempProductSearchImport,
                        batch_id=batch_id,
                        imported_by=imported_by,
                        pending_changes=page._pending_changes,
                    ):
                        launch(copy.deepcopy(page._pending_changes), set(page._pending_deletes))
                    return

                page._pending_changes.clear()
                page._pending_deletes.clear()
                page.load_find_brands()
                page.load_table_rows()
                output = payload.get("output") or ""
                export_error = payload.get("export_error") or ""
                if output:
                    QDesktopServices.openUrl(QUrl.fromLocalFile(str(output)))
                    page.show_message("Данные сохранены, Excel файл сохранён")
                elif export_error:
                    page.show_error_message("Данные в БД сохранены, но Excel не удалось выгрузить:\n" + export_error)
                elif save_to_excel:
                    page.show_message("Данные сохранены в БД. Нет данных для выгрузки")
                else:
                    page.show_message("Данные сохранены в БД")

            _start_page_task(
                page,
                title="Поиск продуктов: сохранение",
                work=work,
                on_finished=done,
                read_resources={"pack_types"},
                write_resources={"products", "product_articles", f"temp_product_search:{imported_by}"},
                intro="Поиск продуктов: сохраняю изменения в фоне.",
                use_progress=True,
            )

        launch(pending_changes, pending_deletes)

    page._background_product_search_import_slot = import_background
    page._background_product_search_save_slot = save_background
    page.ui.btn_Import.clicked.connect(import_background)
    page.ui.btn_Save.clicked.connect(save_background)


# ---------------------------------------------------------------------------
# Customer Cost import (calculation/save already uses start_excel_export)
# ---------------------------------------------------------------------------

def _attach_customer_costs(page) -> None:
    if getattr(page, "_background_integration_customer_costs", False):
        return
    page._background_integration_customer_costs = True
    original_import = page.import_file
    original_calc = page.calculate_costs
    original_save = page.save_all
    _disconnect(page.ui.btn_Import.clicked, original_import)
    _disconnect(page.ui.btn_CalcCost.clicked, original_calc)
    _disconnect(page.ui.btn_Save.clicked, original_save)

    def import_background():
        file_path, _ = QFileDialog.getOpenFileName(page, "Выберите файл", "", "Excel files (*.xls *.xlsx)")
        if not file_path:
            return
        page.start_new_batch()
        page._current_file_path = file_path
        batch_id = page._batch_id
        imported_by = page._imported_by

        def work(progress):
            from app.db.db import SessionLocal
            from app.imports.customer_cost_importer import CustomerCostImporter
            from app.services.customer_cost_service import CustomerCostService
            try:
                progress("Стоимость клиенту: читаю Excel...")
                rows = CustomerCostImporter().read_excel(file_path)
                progress(f"Стоимость клиенту: прочитано строк: {len(rows)}. Загружаю staging...")
                with SessionLocal() as session:
                    service = CustomerCostService(session)
                    service.import_rows(rows=rows, batch_id=batch_id, imported_by=imported_by)
                    progress("Стоимость клиенту: выполняю автоматическое сопоставление...")
                    matched = service.automatch_temp_rows(batch_id=batch_id, imported_by=imported_by)
                    session.commit()
                return {"count": len(rows), "matched": matched}
            finally:
                SessionLocal.remove()

        def done(payload):
            page.load_table()
            page.show_message(f"Данные импортированы: {payload['count']}; автоматически сопоставлено: {payload['matched']}")

        _start_page_task(
            page, title="Расчет стоимости клиенту: импорт и сопоставление",
            work=work, on_finished=done,
            read_resources={"products", "product_articles", "supplier_prices"},
            write_resources={f"temp_customer_cost:{imported_by}"},
            intro="Стоимость клиенту: импорт и сопоставление выполняются в фоне.",
            use_progress=True,
        )

    def _delete_pending(session, pending_deletes, batch_id, imported_by):
        if not pending_deletes:
            return
        from app.db.models import TempCustomerCostImport, TempCustomerCostOption
        session.query(TempCustomerCostOption).filter(
            TempCustomerCostOption.temp_import_id.in_(pending_deletes),
            TempCustomerCostOption.batch_id == batch_id,
            TempCustomerCostOption.imported_by == imported_by,
        ).delete(synchronize_session=False)
        session.query(TempCustomerCostImport).filter(
            TempCustomerCostImport.id.in_(pending_deletes),
            TempCustomerCostImport.batch_id == batch_id,
            TempCustomerCostImport.imported_by == imported_by,
        ).delete(synchronize_session=False)
        session.flush()

    def calculate_background():
        page._commit_open_editors()
        if not page._ensure_known_pack_types():
            return
        default_name = f"CustCostCalc_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
        base_dir = Path(__file__).resolve().parents[2]
        save_path, _ = QFileDialog.getSaveFileName(page, "Сохранить расчет", str(base_dir / default_name), "Excel files (*.xlsx)")
        if not save_path:
            return
        batch_id = page._batch_id
        imported_by = page._imported_by
        manual_prices = copy.deepcopy(page._collect_manual_prices())
        supplier_price_age_months = page.get_supplier_price_age_months()
        pending_deletes = set(getattr(page, "_pending_deletes", set()) or set())

        def work(progress):
            from app.db.db import SessionLocal
            from app.exports.customer_cost_exporter import CustomerCostExporter
            from app.services.customer_cost_service import CustomerCostService
            try:
                with SessionLocal() as session:
                    progress("Стоимость клиенту: применяю изменения строк...")
                    _delete_pending(session, pending_deletes, batch_id, imported_by)
                    service = CustomerCostService(session)
                    progress("Стоимость клиенту: выполняю расчет...")
                    service.run_calculation(batch_id, imported_by, manual_prices=manual_prices, supplier_price_age_months=supplier_price_age_months)
                    progress("Стоимость клиенту: формирую Excel...")
                    output = CustomerCostExporter(session).export_calculated(batch_id, imported_by, save_path)
                    session.commit()
                    return str(output)
            finally:
                SessionLocal.remove()

        def done(output_path):
            page._pending_deletes.clear()
            page._deleted_row_snapshots.clear()
            page.load_table()
            page.refresh_filters()
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(output_path)))
            page.show_message("Расчет выполнен")

        _start_page_task(
            page, title="Расчет стоимости клиенту",
            work=work, on_finished=done,
            read_resources={"products", "supplier_prices", "product_stock"},
            write_resources={"customer_price_calculations", f"temp_customer_cost:{imported_by}"},
            intro="Стоимость клиенту: расчет выполняется в фоне.", use_progress=True,
        )

    def save_background():
        page._commit_open_editors()
        if not page._ensure_known_pack_types():
            return
        base_dir = Path(__file__).resolve().parents[2]
        folder = QFileDialog.getExistingDirectory(page, "Папка для файлов менеджеров", str(base_dir))
        if not folder:
            return
        batch_id = page._batch_id
        imported_by = page._imported_by
        manual_prices = copy.deepcopy(page._collect_manual_prices())
        supplier_price_age_months = page.get_supplier_price_age_months()
        pending_deletes = set(getattr(page, "_pending_deletes", set()) or set())

        def work(progress):
            from app.db.db import SessionLocal
            from app.exports.customer_cost_exporter import CustomerCostExporter
            from app.services.customer_cost_service import CustomerCostService
            try:
                with SessionLocal() as session:
                    progress("Стоимость клиенту: применяю изменения строк...")
                    _delete_pending(session, pending_deletes, batch_id, imported_by)
                    service = CustomerCostService(session)
                    progress("Стоимость клиенту: выполняю расчет...")
                    service.run_calculation(batch_id, imported_by, manual_prices=manual_prices, supplier_price_age_months=supplier_price_age_months)
                    progress("Стоимость клиенту: сохраняю расчеты...")
                    service.save_calculations(batch_id, imported_by)
                    progress("Стоимость клиенту: формирую файлы менеджеров...")
                    output = CustomerCostExporter(session).export_kam_files(batch_id, imported_by, folder)
                    service.delete_temp_rows_for_user(imported_by)
                    session.commit()
                    return output
            finally:
                SessionLocal.remove()

        def done(_output):
            page._pending_deletes.clear()
            page._deleted_row_snapshots.clear()
            page.start_new_batch()
            page._current_file_path = ""
            page.show_message("Данные сохранены")

        _start_page_task(
            page, title="Стоимость клиенту: сохранение",
            work=work, on_finished=done,
            read_resources={"products", "supplier_prices", "product_stock"},
            write_resources={"customer_price_calculations", f"temp_customer_cost:{imported_by}"},
            intro="Стоимость клиенту: сохранение выполняется в фоне.", use_progress=True,
        )

    page._background_customer_cost_import_slot = import_background
    page._background_customer_cost_calc_slot = calculate_background
    page._background_customer_cost_save_slot = save_background
    page.ui.btn_Import.clicked.connect(import_background)
    page.ui.btn_CalcCost.clicked.connect(calculate_background)
    page.ui.btn_Save.clicked.connect(save_background)


# ---------------------------------------------------------------------------
# Target Price import (calculation/export already uses start_excel_export)
# ---------------------------------------------------------------------------

def _attach_target_prices(page) -> None:
    if getattr(page, "_background_integration_target_prices", False):
        return
    page._background_integration_target_prices = True
    original_import = page.import_file
    original_calc = page.calculate_costs
    original_save = page.save_all
    _disconnect(page.ui.btn_Import.clicked, original_import)
    _disconnect(page.ui.btn_CalcCost.clicked, original_calc)
    _disconnect(page.ui.btn_Save.clicked, original_save)

    def import_background():
        try:
            supplier_id = page.ensure_supplier()
        except Exception as exc:
            page.show_error_message(str(exc))
            return
        file_path, _ = QFileDialog.getOpenFileName(
            page,
            "Выберите файл",
            str(Path(__file__).resolve().parents[2]),
            "Excel files (*.xls *.xlsx)",
        )
        if not file_path:
            return
        page.start_new_batch()
        batch_id = page.batch_id
        imported_by = page.imported_by

        def work(progress):
            from app.db.db import SessionLocal
            from app.imports.target_price_importer import TargetPriceImporter
            from app.services.target_price_service import TargetPriceService
            try:
                progress("Target Price: читаю Excel...")
                rows = TargetPriceImporter().read_excel(file_path)
                progress(f"Target Price: прочитано строк: {len(rows)}. Загружаю staging...")
                with SessionLocal() as session:
                    service = TargetPriceService(session)
                    service.import_rows(rows=rows, batch_id=batch_id, imported_by=imported_by, supplier_id=supplier_id)
                    progress("Target Price: выполняю автоматическое сопоставление...")
                    matched = service.automatch_temp_rows(batch_id, imported_by)
                    session.commit()
                return {"count": len(rows), "matched": matched}
            finally:
                SessionLocal.remove()

        def done(payload):
            page._showing_options = False
            page.load_table()
            page.show_message(f"Данные импортированы: {payload['count']}; автоматически сопоставлено: {payload['matched']}")

        _start_page_task(
            page, title="Target Price: импорт и сопоставление",
            work=work, on_finished=done,
            read_resources={"products", "product_articles"},
            write_resources={f"temp_target_price:{imported_by}"},
            intro="Target Price: импорт и сопоставление выполняются в фоне.", use_progress=True,
        )

    def _delete_pending(session, pending_deletes, batch_id, imported_by):
        if not pending_deletes:
            return
        from app.db.models import TempTargetPriceImport, TempTargetPriceOption
        session.query(TempTargetPriceOption).filter(
            TempTargetPriceOption.temp_import_id.in_(pending_deletes),
            TempTargetPriceOption.batch_id == batch_id,
            TempTargetPriceOption.imported_by == imported_by,
        ).delete(synchronize_session=False)
        session.query(TempTargetPriceImport).filter(
            TempTargetPriceImport.id.in_(pending_deletes),
            TempTargetPriceImport.batch_id == batch_id,
            TempTargetPriceImport.imported_by == imported_by,
        ).delete(synchronize_session=False)
        session.flush()

    def calculate_background():
        page._commit_open_editors()
        if not page._ensure_known_pack_types():
            return
        try:
            supplier_id = page.ensure_supplier()
            supplier_name = clean_multi_spaces(page.ui.cbo_SupplName.currentText()) or clean_multi_spaces(page.ui.line_NewSupplier.text()) or "NoName"
            safe_supplier_name = "".join(ch if ch not in r'<>:\"/\\|?*' else "_" for ch in supplier_name)
            default = f"TargetPriceCalc_{safe_supplier_name}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
            manual_full_costs = copy.deepcopy(page.collect_manual_full_costs_from_table())
            supplier_price_age_months = page.get_supplier_price_age_months()
            save_path, _ = QFileDialog.getSaveFileName(page, "Сохранить расчет", str(Path(__file__).resolve().parents[2] / default), "Excel files (*.xlsx)")
            if not save_path:
                return
            batch_id = page.batch_id
            imported_by = page.imported_by
            pending_deletes = set(getattr(page, "_pending_deletes", set()) or set())
        except Exception as exc:
            page.show_error_message(str(exc))
            return

        def work(progress):
            from app.db.db import SessionLocal
            from app.exports.target_price_exporter import TargetPriceExporter
            from app.services.target_price_service import TargetPriceService
            try:
                with SessionLocal() as session:
                    progress("Target Price: применяю изменения строк...")
                    _delete_pending(session, pending_deletes, batch_id, imported_by)
                    service = TargetPriceService(session)
                    progress("Target Price: выполняю расчет...")
                    service.run_calculation(batch_id, imported_by, manual_full_costs=manual_full_costs, supplier_price_age_months=supplier_price_age_months)
                    progress("Target Price: формирую Excel...")
                    output = TargetPriceExporter(session).export_calculated(batch_id, imported_by, save_path)
                    session.commit()
                    return str(output)
            finally:
                SessionLocal.remove()

        def done(output_path):
            page._pending_deletes.clear()
            page._deleted_row_snapshots.clear()
            page._showing_options = True
            page.load_table()
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(output_path)))
            page.show_message("Расчет выполнен")

        _start_page_task(
            page, title="Target Price: расчет",
            work=work, on_finished=done,
            read_resources={"products", "supplier_prices", "product_stock", "product_uc3_history"},
            write_resources={f"temp_target_price:{imported_by}"},
            intro="Target Price: расчет выполняется в фоне.", use_progress=True,
        )

    def save_background():
        page._commit_open_editors()
        if not page._ensure_known_pack_types():
            return
        try:
            if not page._showing_options:
                page.show_error_message("Сначала запустите расчет")
                return
            supplier_id = page.ensure_supplier()
            supplier_name = clean_multi_spaces(page.ui.cbo_SupplName.currentText()) or clean_multi_spaces(page.ui.line_NewSupplier.text()) or "NoName"
            safe_supplier_name = "".join(ch if ch not in r'<>:\"/\\|?*' else "_" for ch in supplier_name)
            default = f"TargetPrice_{safe_supplier_name}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
            manual_full_costs = copy.deepcopy(page.collect_manual_full_costs_from_table())
            save_path, _ = QFileDialog.getSaveFileName(page, "Сохранить target price", str(Path(__file__).resolve().parents[2] / default), "Excel files (*.xlsx)")
            if not save_path:
                return
            currency = clean_multi_spaces(page.ui.cbo_Currency.currentText()).upper()
            if currency == "-":
                currency = ""
            batch_id = page.batch_id
            imported_by = page.imported_by
            fx_rate = page.parse_decimal_field(page.ui.line_ExchangeRate, "Курс")
            transport = page.parse_decimal_field(page.ui.line_Transport, "Транспорт")
            reexport = page.parse_percent_field(page.ui.line_Reexport, "Реэкспорт")
            insurance = page.parse_percent_field(page.ui.line_Insurance, "Insurance %")
            fx_markup = page.parse_percent_field(page.ui.line_FXMarkup, "FX markup %")
            fx_markup_abs = page.parse_decimal_field(page.ui.line_FXMarkupAbs, "FX markup abs")
            has_customs = page.ui.cbo_Customs.currentText() == "да"
            via_novo = page.ui.cbo_viaNovo.currentText() == "через Ново"
            pending_deletes = set(getattr(page, "_pending_deletes", set()) or set())
        except Exception as exc:
            page.show_error_message(str(exc))
            return

        def work(progress):
            from app.db.db import SessionLocal
            from app.exports.target_price_exporter import TargetPriceExporter
            from app.services.target_price_service import TargetPriceService
            try:
                with SessionLocal() as session:
                    progress("Target Price: применяю изменения строк...")
                    _delete_pending(session, pending_deletes, batch_id, imported_by)
                    service = TargetPriceService(session)
                    progress("Target Price: сохраняю расчеты...")
                    service.save_target_calculations(
                        batch_id=batch_id,
                        imported_by=imported_by,
                        target_supplier_id=supplier_id,
                        currency_code=currency,
                        fx_rate=fx_rate,
                        transport=transport,
                        reexport=reexport,
                        insurance=insurance,
                        fx_markup=fx_markup,
                        fx_markup_abs=fx_markup_abs,
                        has_customs=has_customs,
                        via_novo=via_novo,
                        manual_full_costs=manual_full_costs,
                    )
                    progress("Target Price: формирую итоговый Excel...")
                    output = TargetPriceExporter(session).export_final(batch_id, imported_by, save_path)
                    service.delete_temp_rows_for_user(imported_by)
                    session.commit()
                    return str(output)
            finally:
                SessionLocal.remove()

        def done(output_path):
            page._pending_deletes.clear()
            page._deleted_row_snapshots.clear()
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(output_path)))
            page.start_new_batch()
            page.show_message("Target price сохранен")

        _start_page_task(
            page, title="Target Price: сохранение",
            work=work, on_finished=done,
            read_resources={"products", "supplier_prices", "product_stock", "product_uc3_history"},
            write_resources={"target_price_calculations", f"temp_target_price:{imported_by}"},
            intro="Target Price: сохранение выполняется в фоне.", use_progress=True,
        )

    page._background_target_price_import_slot = import_background
    page._background_target_price_calc_slot = calculate_background
    page._background_target_price_save_slot = save_background
    page.ui.btn_Import.clicked.connect(import_background)
    page.ui.btn_CalcCost.clicked.connect(calculate_background)
    page.ui.btn_Save.clicked.connect(save_background)


# ---------------------------------------------------------------------------
# Order Planning calculation
# ---------------------------------------------------------------------------

def _attach_order_planning(page) -> None:
    if getattr(page, "_background_integration_order_planning", False):
        return
    page._background_integration_order_planning = True
    original_calculate = page.calculate
    original_save = page.save
    _disconnect(page.ui.btn_Calculate.clicked, original_calculate)
    _disconnect(page.ui.btn_Save.clicked, original_save)

    def calculate_background():
        if page._today_dates_selected():
            page.show_error_message("Выбери период расчета. Он должен быть не менее месяца")
            return
        period_from, period_to = page._selected_period()
        if period_to < period_from:
            page.show_error_message("Дата 'Продажи по' не может быть меньше даты 'Продажи с'")
            return
        if (period_to - period_from).days + 1 < 30:
            page.show_error_message("Выбери период расчета. Он должен быть не менее месяца")
            return

        def work(progress):
            from app.db.db import SessionLocal
            from app.services.order_planning_service import OrderPlanningService
            try:
                progress("Планирование закупок: читаю продажи и остатки...")
                with SessionLocal() as session:
                    result = OrderPlanningService(session).calculate(period_from, period_to)
                progress(f"Планирование закупок: расчет завершён. Строк: {len(result.rows)}")
                return result
            finally:
                SessionLocal.remove()

        def done(result):
            page._base_rows = result.rows
            page._period_from = result.period_from
            page._period_to = result.period_to
            page._set_period_label()
            page.recalculate_current_rows()
            msg = f"Расчет выполнен. Строк: {len(page._rows)}."
            if result.unmatched_count:
                msg += f" Не сопоставлено продуктов: {result.unmatched_count}. Справочники → Сопоставление продуктов."
            page.show_message(msg)

        _start_page_task(
            page, title="Планирование закупок: расчет",
            work=work, on_finished=done,
            read_resources={"products", "product_stock", "sales_product_links"},
            intro="Планирование закупок: расчет выполняется в фоне.", use_progress=True,
        )

    def save_background():
        from app.services.product_matching_service import MissingPackTypeError
        from app.utils.pack_type_prompt import ask_pack_type
        from app.utils.parsers import parse_loose_number

        page._commit_open_product_editors()
        mode = page._mode
        rows_snapshot = [copy.deepcopy(row) for row in page._rows]

        if mode != "check":
            if not page._period_from or not page._period_to:
                page.show_error_message("Сначала сделай расчет")
                return
            missing = [row for row in rows_snapshot if row.get("sales_code") and not row.get("product_id")]
            if missing:
                page.show_error_message(
                    f"Не сопоставлено продуктов: {len(missing)}. Открой Справочники → Сопоставление продуктов."
                )
                return

        def launch(rows_for_save):
            period_from = page._period_from
            period_to = page._period_to

            def work(progress):
                from app.db.db import SessionLocal
                from app.services.order_planning_service import OrderPlanningService
                try:
                    with SessionLocal() as session:
                        service = OrderPlanningService(session)
                        try:
                            if mode == "check":
                                progress("Планирование закупок: сохраняю сопоставления...")
                                count = service.save_product_links(rows_for_save)
                                session.commit()
                                progress("Планирование закупок: перепроверяю продукты...")
                                result = service.check_products()
                                return {"status": "success", "mode": "check", "count": count, "result": result}
                            progress("Планирование закупок: сохраняю расчет...")
                            count = service.save_calculation(rows_for_save, period_from, period_to)
                            session.commit()
                            return {"status": "success", "mode": "calc", "count": count}
                        except MissingPackTypeError as exc:
                            session.rollback()
                            return {"status": "missing_pack", "error": exc}
                finally:
                    SessionLocal.remove()

            def done(payload):
                if payload.get("status") == "missing_pack":
                    exc = payload.get("error")
                    selected_pack = ask_pack_type(page, exc)
                    if selected_pack is None:
                        return
                    requested_pack = parse_loose_number(exc.requested_pack)
                    for row in page._rows:
                        if row.get("product_id"):
                            continue
                        row_pack = row.get("new_pack")
                        if row_pack in (None, ""):
                            row_pack = row.get("sales_pack") if row.get("sales_pack") not in (None, "") else row.get("pack")
                        if parse_loose_number(row_pack) == requested_pack:
                            row["new_pack"] = selected_pack
                            row["sales_pack"] = selected_pack
                            row["pack"] = selected_pack
                    page._base_rows = [dict(row) for row in page._rows]
                    page.display_rows(page._rows, page._mode)
                    launch([copy.deepcopy(row) for row in page._rows])
                    return

                if payload.get("mode") == "check":
                    result = payload["result"]
                    page.display_rows(result.rows, "check")
                    if result.rows:
                        page.show_message(f"Сопоставления сохранены: {payload['count']}. Осталось изменений: {len(result.rows)}")
                    else:
                        page.show_message(f"Сопоставления сохранены: {payload['count']}. Изменений не найдено")
                else:
                    page.show_message(f"Расчет сохранен. Продуктов: {payload['count']}")

            _start_page_task(
                page, title="Планирование закупок: сохранение",
                work=work, on_finished=done,
                read_resources={"pack_types"} if mode == "check" else set(),
                write_resources={"order_planning_calculations", "products", "sales_product_links"} if mode == "check" else {"order_planning_calculations"},
                intro="Планирование закупок: сохранение выполняется в фоне.", use_progress=True,
            )

        launch(rows_snapshot)

    page._background_order_planning_calculate_slot = calculate_background
    page._background_order_planning_save_slot = save_background
    page.ui.btn_Calculate.clicked.connect(calculate_background)
    page.ui.btn_Save.clicked.connect(save_background)


# ---------------------------------------------------------------------------
# Legacy report Excel exporters -> central background manager
# ---------------------------------------------------------------------------

def _attach_legacy_report_export(page) -> None:
    if getattr(page, "_background_integration_report_export", False):
        return
    page._background_integration_report_export = True

    class_name = page.__class__.__name__
    if class_name == "PriceReportsPage":
        import app.page_functions.price_reports_page as report_module
        from app.workers.excel_export_worker import start_excel_export

        def start_managed_export(
            *, headers, rows, output_path, report_mode, quick_order_months, safe_stock_months
        ):
            page.ui.btn_ExportExcel.setEnabled(False)
            page.ui.btn_ExportExcel.setText("Формируется...")
            page.show_message("Excel файл формируется в фоновом режиме. Можно продолжать работать в программе.")
            start_excel_export(
                page,
                report_module._export_price_report_file,
                kwargs={
                    "headers": headers,
                    "rows": rows,
                    "output_path": output_path,
                    "report_mode": report_mode,
                    "quick_order_months": quick_order_months,
                    "safe_stock_months": safe_stock_months,
                },
                on_finished=page._on_excel_export_finished,
                on_error=page._on_excel_export_error,
                button=page.ui.btn_ExportExcel,
                busy_text="Формируется...",
                restore_text=page._export_button_text or "Export Excel",
            )

        page._start_excel_export = start_managed_export

    elif class_name == "CustomerCostsReportsPage":
        import app.page_functions.customer_costs_reports_page as report_module
        from app.workers.excel_export_worker import start_excel_export

        def start_managed_export(*, headers, rows, output_path):
            page.ui.btn_ExportExcel.setEnabled(False)
            page.ui.btn_ExportExcel.setText("Формируется...")
            page.show_message("Excel файл формируется в фоновом режиме. Можно продолжать работать в программе.")
            start_excel_export(
                page,
                report_module._export_customer_cost_report_file,
                kwargs={"headers": headers, "rows": rows, "output_path": output_path},
                on_finished=page._on_excel_export_finished,
                on_error=page._on_excel_export_error,
                button=page.ui.btn_ExportExcel,
                busy_text="Формируется...",
                restore_text=page._export_button_text or "Export Excel",
            )

        page._start_excel_export = start_managed_export


# ---------------------------------------------------------------------------
# Public lazy-page hook
# ---------------------------------------------------------------------------

def attach_page_background_integration(page) -> None:
    """Attach only the integrations relevant to this concrete page instance."""
    if getattr(page, "_procurement_page_integration_attached", False):
        return
    page._procurement_page_integration_attached = True
    _ensure_tab_close_guard(page)

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
    elif class_name == "ProductStockPage":
        _attach_product_stock(page)
    elif class_name == "ProductSearchPage":
        _attach_product_search(page)
    elif class_name == "CustomerCostsPage":
        _attach_customer_costs(page)
    elif class_name == "TargetPricesPage":
        _attach_target_prices(page)
    elif class_name == "OrderPlanningPage":
        _attach_order_planning(page)
    elif class_name in {"PriceReportsPage", "CustomerCostsReportsPage"}:
        _attach_legacy_report_export(page)
