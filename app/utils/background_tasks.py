from __future__ import annotations

import logging
import traceback
import uuid
from dataclasses import dataclass, field
from typing import Any, Callable, Iterable

from PySide6.QtCore import QEvent, QObject, QThread, Qt, Signal, Slot
from PySide6.QtWidgets import QApplication, QMainWindow, QMessageBox, QWidget

logger = logging.getLogger(__name__)

TaskCallable = Callable[[], Any]


class TaskConflictError(RuntimeError):
    """Raised when a new task would conflict with an active background task."""


class _TaskWorker(QObject):
    finished = Signal(object)
    failed = Signal(str, str)

    def __init__(self, func: TaskCallable) -> None:
        super().__init__()
        self._func = func

    @Slot()
    def run(self) -> None:
        try:
            self.finished.emit(self._func())
        except Exception as exc:  # pragma: no cover - GUI boundary
            details = traceback.format_exc()
            logger.exception("Background task failed")
            self.failed.emit(str(exc).strip() or exc.__class__.__name__, details)


class BackgroundTaskHandle(QObject):
    finished = Signal(object)
    failed = Signal(str)

    def __init__(self, task_id: str, title: str, parent: QObject | None = None) -> None:
        super().__init__(parent)
        self.task_id = task_id
        self.title = title


@dataclass(slots=True)
class _TaskRecord:
    task_id: str
    title: str
    thread: QThread
    worker: _TaskWorker
    handle: BackgroundTaskHandle
    read_resources: set[str] = field(default_factory=set)
    write_resources: set[str] = field(default_factory=set)
    owner: QObject | None = None
    status: str = "running"


class BackgroundTaskManager(QObject):
    """One application-wide owner for long-running QThreads.

    Tasks declare the data they read and write. Read/read combinations may run
    together; any overlap that includes a writer is blocked. Qt widgets are
    never touched from worker threads: callers receive results through the
    handle's signals in the GUI thread.
    """

    def __init__(self, app: QApplication) -> None:
        super().__init__(app)
        self._app = app
        self._tasks: dict[str, _TaskRecord] = {}
        self._shutting_down = False
        app.installEventFilter(self)
        app.aboutToQuit.connect(self._final_shutdown)

    @staticmethod
    def _clean_resources(values: Iterable[str] | None) -> set[str]:
        return {str(value).strip() for value in (values or ()) if str(value).strip()}

    def active_count(self) -> int:
        return sum(record.status == "running" for record in self._tasks.values())

    def active_titles(self) -> list[str]:
        return [record.title for record in self._tasks.values() if record.status == "running"]

    def _conflicting_record(
        self,
        *,
        read_resources: set[str],
        write_resources: set[str],
    ) -> _TaskRecord | None:
        for record in self._tasks.values():
            if record.status != "running":
                continue
            # New writes conflict with every active read/write of the same data.
            if write_resources & (record.read_resources | record.write_resources):
                return record
            # New reads conflict only with active writers.
            if read_resources & record.write_resources:
                return record
        return None

    def assert_available(
        self,
        *,
        read_resources: Iterable[str] | None = None,
        write_resources: Iterable[str] | None = None,
    ) -> None:
        reads = self._clean_resources(read_resources)
        writes = self._clean_resources(write_resources)
        conflict = self._conflicting_record(read_resources=reads, write_resources=writes)
        if conflict is None:
            return
        raise TaskConflictError(
            f"Сейчас выполняется «{conflict.title}». "
            "Эти операции используют одни и те же данные. Дождитесь завершения текущей операции."
        )

    def start_task(
        self,
        title: str,
        func: TaskCallable,
        *,
        read_resources: Iterable[str] | None = None,
        write_resources: Iterable[str] | None = None,
        owner: QObject | None = None,
    ) -> BackgroundTaskHandle:
        if self._shutting_down:
            raise RuntimeError("Приложение завершает работу; новую фоновую операцию запустить нельзя.")

        reads = self._clean_resources(read_resources)
        writes = self._clean_resources(write_resources)
        self.assert_available(read_resources=reads, write_resources=writes)

        task_id = uuid.uuid4().hex
        thread_parent = self
        thread = QThread(thread_parent)
        worker = _TaskWorker(func)
        worker.moveToThread(thread)
        handle = BackgroundTaskHandle(task_id, title, self)

        record = _TaskRecord(
            task_id=task_id,
            title=str(title),
            thread=thread,
            worker=worker,
            handle=handle,
            read_resources=reads,
            write_resources=writes,
            owner=owner,
        )
        self._tasks[task_id] = record

        thread.started.connect(worker.run)
        worker.finished.connect(
            lambda result, tid=task_id: self._on_finished(tid, result),
            Qt.QueuedConnection,
        )
        worker.failed.connect(
            lambda message, details, tid=task_id: self._on_failed(tid, message, details),
            Qt.QueuedConnection,
        )
        worker.finished.connect(thread.quit, Qt.DirectConnection)
        worker.failed.connect(thread.quit, Qt.DirectConnection)
        worker.finished.connect(worker.deleteLater)
        worker.failed.connect(worker.deleteLater)
        thread.finished.connect(lambda tid=task_id: self._on_thread_finished(tid))
        thread.finished.connect(thread.deleteLater)
        thread.start()
        return handle

    @Slot(str, object)
    def _on_finished(self, task_id: str, result: object) -> None:
        record = self._tasks.get(task_id)
        if record is None:
            return
        record.status = "finished"
        record.handle.finished.emit(result)

    @Slot(str, str, str)
    def _on_failed(self, task_id: str, message: str, details: str) -> None:
        record = self._tasks.get(task_id)
        if record is None:
            return
        record.status = "failed"
        if details:
            logger.error("Background task %s failed:\n%s", record.title, details)
        record.handle.failed.emit(message)

    def _on_thread_finished(self, task_id: str) -> None:
        self._tasks.pop(task_id, None)

    def shutdown(self, timeout_ms: int = 15000) -> bool:
        """Wait for active workers; never terminate a Python/Qt thread forcibly."""
        self._shutting_down = True
        records = [record for record in self._tasks.values() if record.thread.isRunning()]
        if not records:
            return True

        # Workers are DB / Excel functions that finish cooperatively. Calling
        # QThread.terminate() here would risk an open DB transaction or native
        # Qt crash, so shutdown only waits for safe completion.
        remaining = max(int(timeout_ms), 0)
        for index, record in enumerate(records):
            if not record.thread.isRunning():
                continue
            threads_left = max(len(records) - index, 1)
            wait_for = max(remaining // threads_left, 1) if remaining else 0
            if wait_for <= 0 or not record.thread.wait(wait_for):
                self._shutting_down = False
                return False
            remaining = max(remaining - wait_for, 0)
        return True

    @Slot()
    def _final_shutdown(self) -> None:
        # Normal window close is intercepted below, so this is only a fallback.
        self.shutdown(timeout_ms=5000)

    def eventFilter(self, watched: QObject, event) -> bool:  # noqa: N802 - Qt API
        if event.type() != QEvent.Close or self.active_count() == 0:
            return super().eventFilter(watched, event)
        if not isinstance(watched, QMainWindow) or not watched.isWindow():
            return super().eventFilter(watched, event)

        # Let a short task finish immediately. If it is still running, consume
        # the close event so QApplication cannot be destroyed under a live
        # worker thread (a known Windows/PySide source of 0xC0000005).
        if self.shutdown(timeout_ms=15000):
            return super().eventFilter(watched, event)

        titles = "\n".join(f"• {title}" for title in self.active_titles())
        QMessageBox.warning(
            watched if isinstance(watched, QWidget) else None,
            "Procurement App — завершение работы",
            "Фоновая операция ещё выполняется. Программа пока не будет закрыта, "
            "чтобы не повредить данные.\n\n"
            + (titles or "")
            + "\n\nПодождите несколько секунд и нажмите закрыть ещё раз.",
        )
        self._shutting_down = False
        return True


def get_background_task_manager(parent: QWidget | None = None) -> BackgroundTaskManager:
    app = QApplication.instance()
    if app is None:
        raise RuntimeError("QApplication ещё не создан")
    manager = getattr(app, "_procurement_background_task_manager", None)
    if not isinstance(manager, BackgroundTaskManager):
        manager = BackgroundTaskManager(app)
        setattr(app, "_procurement_background_task_manager", manager)
    return manager
