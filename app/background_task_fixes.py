from __future__ import annotations

"""Small runtime compatibility fixes for background tasks.

This module intentionally does not replace BackgroundTaskManager or its
worker/thread lifecycle.

Responsibilities:
1) keep informational messages of an active background operation in the task
   log instead of opening modal information dialogs;
2) make the no-IS Order Planning exporter wrapper transparent to the restored
   3/5-month export arguments;
3) keep Supplier Price save progress concise;
4) prevent the same progress message from being added twice.
"""

import logging
from typing import Any

from PySide6.QtCore import QTimer

logger = logging.getLogger(__name__)
_installed = False


def _install_order_planning_exporter_no_is_compat() -> None:
    import sys
    from app import no_is_runtime as no_is

    def apply_compat(exporter_module) -> None:
        exporter_cls = getattr(exporter_module, "OrderPlanningExporter", None)
        if exporter_cls is None:
            return

        current = exporter_cls.build_export_data
        if getattr(current, "_order_months_compat", False):
            return

        original = current
        if getattr(current, "_no_is_patch", False):
            try:
                closure_values = {
                    name: cell.cell_contents
                    for name, cell in zip(
                        current.__code__.co_freevars,
                        current.__closure__ or (),
                    )
                }
                captured = closure_values.get("original")
                if callable(captured):
                    original = captured
            except Exception:
                logger.exception("Не удалось раскрыть legacy no-IS wrapper Order Planning")
                return

        def build_export_data_no_is(
            self,
            display_rows,
            supplier_price_age_months: int = 3,
            quick_order_months: int = 3,
            safe_stock_months: int = 5,
        ):
            headers, rows = original(
                self,
                display_rows,
                supplier_price_age_months=supplier_price_age_months,
                quick_order_months=quick_order_months,
                safe_stock_months=safe_stock_months,
            )
            filtered_headers, keep = no_is._filtered_headers_and_indexes(headers)
            filtered_rows = [no_is._filter_row_by_indexes(row, keep) for row in rows]
            return filtered_headers, filtered_rows

        build_export_data_no_is._no_is_patch = True
        build_export_data_no_is._order_months_compat = True
        exporter_cls.build_export_data = build_export_data_no_is

    def patch_order_planning_exporter_compat(exporter_module) -> None:
        apply_compat(exporter_module)

    patch_order_planning_exporter_compat._order_months_compat = True
    no_is._patch_order_planning_exporter = patch_order_planning_exporter_compat

    already_loaded = sys.modules.get("app.exports.order_planning_exporter")
    if already_loaded is not None:
        apply_compat(already_loaded)


def _install_manager_progress_dedupe() -> None:
    """Ignore only an exact consecutive duplicate in the central task log."""
    from app.utils.background_tasks import BackgroundTaskManager

    current = BackgroundTaskManager._on_progress
    if getattr(current, "_procurement_progress_dedupe", False):
        return

    original = current

    def on_progress_deduped(self, task_id: str, text: str) -> None:
        record = self.record(task_id)
        message = str(text or "").strip()
        if record is not None and message:
            log = getattr(record, "log", None)
            if isinstance(log, list) and log and log[-1] == message:
                return
        original(self, task_id, text)

    on_progress_deduped._procurement_progress_dedupe = True
    on_progress_deduped._procurement_original = original
    BackgroundTaskManager._on_progress = on_progress_deduped


def _install_background_info_routing() -> None:
    from app.utils import page_background_integration as pbi

    original_start = pbi._start_page_task
    if getattr(original_start, "_background_info_routing", False):
        return

    def start_page_task_without_modal_info(
        page,
        *,
        title: str,
        work,
        on_finished,
        read_resources=(),
        write_resources=(),
        intro: str,
        use_progress: bool = False,
    ) -> bool:
        manager = pbi.get_background_task_manager()
        state = getattr(page, "_procurement_background_message_router_state", None)

        if state is None:
            original_show = getattr(page, "show_message", None)
            had_instance_attr = "show_message" in getattr(page, "__dict__", {})
            previous_instance_value = getattr(page, "__dict__", {}).get("show_message")
            state = {
                "original_show": original_show,
                "had_instance_attr": had_instance_attr,
                "previous_instance_value": previous_instance_value,
                "restored": False,
            }
            setattr(page, "_procurement_background_message_router_state", state)

            def latest_owner_record():
                tasks = getattr(manager, "_tasks", {})
                for record in reversed(list(tasks.values())):
                    if getattr(record, "owner", None) is page and getattr(record, "status", "") in {
                        "running", "finished", "failed"
                    }:
                        return record
                return None

            def background_only_message(text: Any = "", *_args, **_kwargs) -> None:
                message = str(text or "").strip()
                if not message:
                    return
                record = latest_owner_record()
                if record is None:
                    return

                # Progress from the worker is already stored by BackgroundTaskManager.
                # Do not mirror it back into the same log while the task is running.
                if getattr(record, "status", "") == "running":
                    return

                log = getattr(record, "log", None)
                if isinstance(log, list):
                    if message in log:
                        return
                    if getattr(record, "status", "") == "finished" and log and log[-1] == "Готово.":
                        log.insert(len(log) - 1, message)
                    else:
                        log.append(message)
                manager.task_progress.emit(record.task_id, message)

            state["router"] = background_only_message
            if callable(original_show):
                setattr(page, "show_message", background_only_message)

        def restore_show_message_if_idle() -> None:
            current = getattr(page, "_procurement_background_message_router_state", None)
            if current is not state or state.get("restored"):
                return
            if manager.has_active_tasks(page):
                return

            state["restored"] = True
            original_show = state.get("original_show")
            try:
                if state.get("had_instance_attr"):
                    setattr(page, "show_message", state.get("previous_instance_value"))
                else:
                    delattr(page, "show_message")
            except (AttributeError, RuntimeError):
                if callable(original_show):
                    try:
                        setattr(page, "show_message", original_show)
                    except (AttributeError, RuntimeError):
                        pass
            try:
                delattr(page, "_procurement_background_message_router_state")
            except (AttributeError, RuntimeError):
                pass

        work_to_start = work
        if use_progress and title == "Сохранение и расчёт прайса поставщика":
            original_work = work

            def concise_supplier_price_work(progress):
                def concise_progress(text: Any = "") -> None:
                    message = str(text or "").strip()
                    if message.startswith("Сохранение прайса завершено"):
                        progress(message)

                return original_work(concise_progress)

            work_to_start = concise_supplier_price_work

        try:
            started = original_start(
                page,
                title=title,
                work=work_to_start,
                on_finished=on_finished,
                read_resources=read_resources,
                write_resources=write_resources,
                intro=intro,
                use_progress=use_progress,
            )
        except Exception:
            restore_show_message_if_idle()
            raise

        if not started:
            restore_show_message_if_idle()
            return False

        handle = getattr(page, "_procurement_background_handle", None)
        if handle is None:
            restore_show_message_if_idle()
            return True

        # Do not disconnect handle.progress here. Other code may legitimately
        # subscribe to that signal. The router above simply ignores the duplicate
        # page.show_message mirror while the task is active.
        handle.finished.connect(lambda _result: QTimer.singleShot(0, restore_show_message_if_idle))
        handle.failed.connect(lambda _message: QTimer.singleShot(0, restore_show_message_if_idle))
        return True

    start_page_task_without_modal_info._background_info_routing = True
    start_page_task_without_modal_info._background_original = original_start
    pbi._start_page_task = start_page_task_without_modal_info


def install_background_task_fixes() -> None:
    global _installed
    if _installed:
        return
    _install_order_planning_exporter_no_is_compat()
    _install_manager_progress_dedupe()
    _install_background_info_routing()
    _installed = True
