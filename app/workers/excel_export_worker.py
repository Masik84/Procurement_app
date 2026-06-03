from __future__ import annotations

import traceback
from typing import Any, Callable

from PySide6.QtCore import QObject, QThread, Signal, Slot, Qt


class ExcelExportWorker(QObject):
    """Run a long Excel export in a background QThread.

    The worker never touches Qt widgets. UI callbacks are executed through a
    QObject proxy that lives in the owner's/main thread.
    """

    finished = Signal(object)
    error = Signal(str)

    def __init__(self, export_func: Callable[..., Any], *args: Any, **kwargs: Any) -> None:
        super().__init__()
        self._export_func = export_func
        self._args = args
        self._kwargs = kwargs

    @Slot()
    def run(self) -> None:
        try:
            result = self._export_func(*self._args, **self._kwargs)
            self.finished.emit(result)
        except Exception as exc:
            details = traceback.format_exc()
            print(details)
            message = str(exc).strip() or exc.__class__.__name__
            self.error.emit(message)


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
    """Start ``export_func`` in a QThread and keep references on ``owner``."""

    if getattr(owner, "_excel_export_thread", None) is not None:
        return False

    kwargs = kwargs or {}
    thread = QThread(owner)
    worker = ExcelExportWorker(export_func, *args, **kwargs)
    worker.moveToThread(thread)

    if button is not None:
        if restore_text is None:
            restore_text = button.text()
        button.setEnabled(False)
        button.setText(busy_text)

    def finish_ui() -> None:
        if button is not None:
            button.setEnabled(True)
            button.setText(restore_text or "Export Excel")

    proxy = _ExcelExportCallbackProxy(
        on_finished=on_finished,
        on_error=on_error,
        finish_ui=finish_ui,
        parent=owner,
    )

    def clear_refs() -> None:
        setattr(owner, "_excel_export_thread", None)
        setattr(owner, "_excel_export_worker", None)
        setattr(owner, "_excel_export_callback_proxy", None)

    setattr(owner, "_excel_export_thread", thread)
    setattr(owner, "_excel_export_worker", worker)
    setattr(owner, "_excel_export_callback_proxy", proxy)

    thread.started.connect(worker.run)
    worker.finished.connect(proxy.handle_finished, Qt.QueuedConnection)
    worker.error.connect(proxy.handle_error, Qt.QueuedConnection)
    worker.finished.connect(thread.quit)
    worker.error.connect(thread.quit)
    worker.finished.connect(worker.deleteLater)
    worker.error.connect(worker.deleteLater)
    thread.finished.connect(thread.deleteLater)
    thread.finished.connect(clear_refs)
    thread.start()
    return True
