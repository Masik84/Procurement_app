from __future__ import annotations

from typing import Any, Callable

from PySide6.QtCore import QObject, Slot

from app.utils.background_tasks import TaskConflictError, get_background_task_manager


class _ExcelExportCallbackProxy(QObject):
    def __init__(
        self,
        *,
        on_finished: Callable[[Any], None] | None,
        on_error: Callable[[str], None] | None,
        finish_ui: Callable[[], None],
        parent: QObject | None = None,
    ) -> None:
        super().__init__(parent)
        self._on_finished = on_finished
        self._on_error = on_error
        self._finish_ui = finish_ui

    @Slot(object)
    def handle_finished(self, result: Any) -> None:
        self._finish_ui()
        if self._on_finished is not None:
            self._on_finished(result)

    @Slot(str)
    def handle_error(self, error_text: str) -> None:
        self._finish_ui()
        if self._on_error is not None:
            self._on_error(error_text)


def start_excel_export(
    owner: QObject,
    export_func: Callable[..., Any],
    *,
    on_finished: Callable[[Any], None] | None = None,
    on_error: Callable[[str], None] | None = None,
    button: Any | None = None,
    busy_text: str = "Формируется...",
    restore_text: str | None = None,
    args: tuple[Any, ...] = (),
    kwargs: dict[str, Any] | None = None,
) -> bool:
    """Run an Excel export through the application-wide background manager.

    The public API stays compatible with the previous dedicated-QThread helper,
    but all exports now participate in safe application shutdown and central
    thread ownership.
    """

    if getattr(owner, "_excel_export_task_handle", None) is not None:
        return False

    kwargs = kwargs or {}
    if button is not None:
        if restore_text is None:
            restore_text = button.text()
        button.setEnabled(False)
        button.setText(busy_text)

    def finish_ui() -> None:
        if button is not None:
            button.setEnabled(True)
            button.setText(restore_text or "Export Excel")
        setattr(owner, "_excel_export_task_handle", None)
        setattr(owner, "_excel_export_callback_proxy", None)
        # Old pages initialise/check these attributes. Keep them coherent even
        # though the dedicated worker/thread implementation is no longer used.
        setattr(owner, "_excel_export_thread", None)
        setattr(owner, "_excel_export_worker", None)

    proxy = _ExcelExportCallbackProxy(
        on_finished=on_finished,
        on_error=on_error,
        finish_ui=finish_ui,
        parent=owner,
    )
    setattr(owner, "_excel_export_callback_proxy", proxy)

    def work():
        return export_func(*args, **kwargs)

    owner_class = owner.__class__.__name__
    read_resources = set()
    if owner_class == "SupplierPricesPage":
        read_resources = {
            "products",
            "supplier_price_calculations",
            "product_stock",
            "product_uc3_history",
        }

    try:
        handle = get_background_task_manager().start_task(
            "Формирование Excel",
            work,
            read_resources=read_resources,
            owner=owner,
        )
    except TaskConflictError:
        finish_ui()
        return False
    except Exception as exc:
        finish_ui()
        if on_error is not None:
            on_error(str(exc))
        return False

    setattr(owner, "_excel_export_task_handle", handle)
    handle.finished.connect(proxy.handle_finished)
    handle.failed.connect(proxy.handle_error)
    return True
