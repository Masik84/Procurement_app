from __future__ import annotations

import logging
import traceback
import uuid
from dataclasses import dataclass, field
from typing import Any, Callable, Iterable

from PySide6.QtCore import QEvent, QObject, QThread, QTimer, Qt, Signal, Slot
from PySide6.QtWidgets import (
    QApplication, QDialog, QHBoxLayout, QLabel, QListWidget, QListWidgetItem,
    QMainWindow, QMessageBox, QPlainTextEdit, QPushButton, QSplitter,
    QVBoxLayout, QWidget,
)

logger = logging.getLogger(__name__)

TaskCallable = Callable[..., Any]


class TaskConflictError(RuntimeError):
    """Raised when a new task would conflict with an active background task."""


class _TaskWorker(QObject):
    progress = Signal(str)
    finished = Signal(object)
    failed = Signal(str, str)

    def __init__(self, func: TaskCallable, title: str, *, use_progress: bool = False) -> None:
        super().__init__()
        self._func = func
        self._title = str(title)
        self._use_progress = bool(use_progress)

    @Slot()
    def run(self) -> None:
        logger.info("BACKGROUND | START | %s", self._title)
        try:
            if self._use_progress:
                result = self._func(lambda text: self.progress.emit(str(text)))
            else:
                result = self._func()
        except Exception as exc:  # pragma: no cover - GUI boundary
            details = traceback.format_exc()
            logger.exception("BACKGROUND | FAIL | %s", self._title)
            self.failed.emit(str(exc).strip() or exc.__class__.__name__, details)
            return
        logger.info("BACKGROUND | FINISH | %s", self._title)
        self.finished.emit(result)


class BackgroundTaskHandle(QObject):
    progress = Signal(str)
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
    thread: QThread | None
    worker: _TaskWorker | None
    handle: BackgroundTaskHandle
    read_resources: set[str] = field(default_factory=set)
    write_resources: set[str] = field(default_factory=set)
    owner: QObject | None = None
    status: str = "running"
    log: list[str] = field(default_factory=list)


class BackgroundTaskCenterDialog(QDialog):
    """Non-modal task monitor mirroring the working Daily-Report pattern."""

    def __init__(self, manager: "BackgroundTaskManager", parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.manager = manager
        self.setWindowTitle("Фоновые задачи")
        self.setModal(False)
        self.resize(760, 420)

        root = QVBoxLayout(self)
        self.summary = QLabel("Нет активных задач")
        self.summary.setWordWrap(True)
        root.addWidget(self.summary)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        self.task_list = QListWidget()
        self.task_list.setMinimumWidth(260)
        self.log_edit = QPlainTextEdit()
        self.log_edit.setReadOnly(True)
        self.log_edit.setLineWrapMode(QPlainTextEdit.LineWrapMode.NoWrap)
        splitter.addWidget(self.task_list)
        splitter.addWidget(self.log_edit)
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        root.addWidget(splitter, 1)

        buttons = QHBoxLayout()
        buttons.addStretch(1)
        self.minimize_button = QPushButton("Свернуть")
        self.hide_button = QPushButton("Скрыть")
        buttons.addWidget(self.minimize_button)
        buttons.addWidget(self.hide_button)
        root.addLayout(buttons)
        self.minimize_button.clicked.connect(self.showMinimized)
        self.hide_button.clicked.connect(self.hide)
        self.task_list.currentItemChanged.connect(self._show_selected)

        manager.task_started.connect(self.on_task_started)
        manager.task_progress.connect(self.on_task_progress)
        manager.task_done.connect(self.on_task_done)

    def _item(self, task_id: str):
        for i in range(self.task_list.count()):
            item = self.task_list.item(i)
            if item.data(Qt.ItemDataRole.UserRole) == task_id:
                return item
        return None

    @staticmethod
    def _label(record: _TaskRecord) -> str:
        prefix = {"running": "●", "finished": "✓", "failed": "✗"}.get(record.status, "•")
        return f"{prefix} {record.title}"

    def _refresh_summary(self) -> None:
        active = self.manager.active_count()
        self.summary.setText(
            f"Выполняется фоновых задач: {active}. Можно работать в других разделах."
            if active else "Фоновые задачи завершены"
        )
        self.setWindowTitle(f"Фоновые задачи ({active})" if active else "Фоновые задачи")

    @Slot(str, str)
    def on_task_started(self, task_id: str, _title: str) -> None:
        record = self.manager.record(task_id)
        if record is None:
            return
        item = QListWidgetItem(self._label(record))
        item.setData(Qt.ItemDataRole.UserRole, task_id)
        self.task_list.insertItem(0, item)
        self.task_list.setCurrentItem(item)
        self._refresh_summary()
        was_minimized = self.isMinimized()
        if not self.isVisible():
            self.show()
        if not was_minimized:
            self.raise_()
            self.activateWindow()

    @Slot(str, str)
    def on_task_progress(self, task_id: str, _text: str) -> None:
        record = self.manager.record(task_id)
        item = self._item(task_id)
        if record is not None and item is not None:
            item.setText(self._label(record))
        current = self.task_list.currentItem()
        if current is not None and current.data(Qt.ItemDataRole.UserRole) == task_id:
            self._render(task_id)
        self._refresh_summary()

    @Slot(str, bool)
    def on_task_done(self, task_id: str, _ok: bool) -> None:
        self.on_task_progress(task_id, "")

    def _show_selected(self, current, _previous) -> None:
        if current is None:
            self.log_edit.clear()
            return
        self._render(current.data(Qt.ItemDataRole.UserRole))

    def _render(self, task_id: str) -> None:
        record = self.manager.record(task_id)
        if record is None:
            self.log_edit.clear()
            return
        self.log_edit.setPlainText("\n".join(record.log))
        bar = self.log_edit.verticalScrollBar()
        bar.setValue(bar.maximum())

    def closeEvent(self, event) -> None:  # noqa: N802
        self.hide()
        event.accept()


class BackgroundTaskManager(QObject):
    task_started = Signal(str, str)
    task_progress = Signal(str, str)
    task_done = Signal(str, bool)
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
        self._task_center: BackgroundTaskCenterDialog | None = None
        app.installEventFilter(self)
        app.aboutToQuit.connect(self._final_shutdown)

    @staticmethod
    def _clean_resources(values: Iterable[str] | None) -> set[str]:
        return {str(value).strip() for value in (values or ()) if str(value).strip()}

    def active_count(self) -> int:
        return sum(record.status == "running" for record in self._tasks.values())

    def active_titles(self) -> list[str]:
        return [record.title for record in self._tasks.values() if record.status == "running"]

    def record(self, task_id: str) -> _TaskRecord | None:
        return self._tasks.get(task_id)

    def ensure_task_center(self, parent: QWidget | None = None) -> BackgroundTaskCenterDialog:
        if self._task_center is None:
            owner = parent.window() if parent is not None else None
            self._task_center = BackgroundTaskCenterDialog(self, owner)
        return self._task_center

    def show_task_center(self, parent: QWidget | None = None) -> None:
        dialog = self.ensure_task_center(parent)
        if dialog.isMinimized():
            return
        dialog.show()
        dialog.raise_()
        dialog.activateWindow()

    def has_active_tasks(self, owner: QObject | None = None) -> bool:
        if owner is None:
            return self.active_count() > 0
        return any(
            record.status == "running" and record.owner is owner
            for record in self._tasks.values()
        )

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
        on_finished: Callable[[Any], None] | None = None,
        on_failed: Callable[[str], None] | None = None,
        use_progress: bool = False,
        intro: str = "",
    ) -> BackgroundTaskHandle:
        if self._shutting_down:
            raise RuntimeError("Приложение завершает работу; новую фоновую операцию запустить нельзя.")

        reads = self._clean_resources(read_resources)
        writes = self._clean_resources(write_resources)
        self.assert_available(read_resources=reads, write_resources=writes)

        task_id = uuid.uuid4().hex
        thread_parent = self
        thread = QThread(thread_parent)
        worker = _TaskWorker(func, title, use_progress=use_progress)
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
        if intro:
            record.log.append(str(intro))
        self._tasks[task_id] = record

        thread.started.connect(worker.run)
        worker.progress.connect(
            lambda text, tid=task_id: self._on_progress(tid, text),
            Qt.QueuedConnection,
        )
        worker.finished.connect(
            lambda result, tid=task_id: self._on_finished(tid, result),
            Qt.QueuedConnection,
        )
        worker.failed.connect(
            lambda message, details, tid=task_id: self._on_failed(tid, message, details),
            Qt.QueuedConnection,
        )
        # Same safe ordering as Daily-Report--new-.  Do not force a DirectConnection
        # from the worker thread into the QThread wrapper; let Qt queue quit in
        # the normal order after result/failure delivery has been scheduled.
        worker.finished.connect(thread.quit)
        worker.failed.connect(lambda _message, _details, target=thread: target.quit())
        worker.finished.connect(worker.deleteLater)
        worker.failed.connect(worker.deleteLater)
        thread.finished.connect(lambda tid=task_id: self._on_thread_finished(tid))
        thread.finished.connect(thread.deleteLater)

        # Register caller callbacks before starting the worker.  This removes
        # the old race without deferring QThread.start() through the GUI event
        # loop.  Even an instant worker cannot finish invisibly: worker signals
        # are queued back to this manager in the GUI thread.
        if on_finished is not None:
            handle.finished.connect(on_finished)
        if on_failed is not None:
            handle.failed.connect(on_failed)

        logger.info(
            "BACKGROUND | SCHEDULED | %s | task_id=%s | read=%s | write=%s",
            title,
            task_id,
            sorted(reads),
            sorted(writes),
        )
        self.ensure_task_center(owner if isinstance(owner, QWidget) else None)
        self.task_started.emit(task_id, str(title))
        if intro:
            self.task_progress.emit(task_id, str(intro))
        self.show_task_center(owner if isinstance(owner, QWidget) else None)
        # Same lifecycle as Daily-Report--new-: start on the next GUI event-loop
        # turn so every page callback is connected before even a very small task
        # can finish.  Crucially, the task record is NOT removed when QThread
        # stops; _on_finished/_on_failed must still be able to deliver the result
        # to the GUI page.
        QTimer.singleShot(0, thread.start)
        return handle


    @Slot(str, str)
    def _on_progress(self, task_id: str, text: str) -> None:
        record = self._tasks.get(task_id)
        if record is None or record.status != "running":
            return
        message = str(text or "").strip()
        if not message:
            return
        logger.info("BACKGROUND | PROGRESS | %s | %s", record.title, message)
        record.log.append(message)
        record.handle.progress.emit(message)
        self.task_progress.emit(task_id, message)

    @Slot(str, object)
    def _on_finished(self, task_id: str, result: object) -> None:
        record = self._tasks.get(task_id)
        if record is None:
            return
        record.status = "finished"
        record.log.append("Готово.")
        record.handle.finished.emit(result)
        self.task_done.emit(task_id, True)

    @Slot(str, str, str)
    def _on_failed(self, task_id: str, message: str, details: str) -> None:
        record = self._tasks.get(task_id)
        if record is None:
            return
        record.status = "failed"
        if details:
            logger.error("Background task %s failed:\n%s", record.title, details)
        record.log.append(f"Ошибка: {message}")
        if details:
            record.log.append(details.rstrip())
        record.handle.failed.emit(message)
        self.task_done.emit(task_id, False)

    def _on_thread_finished(self, task_id: str) -> None:
        # Do NOT pop the task here. QThread.finished may be delivered before the
        # queued worker.finished/worker.failed callback. Removing the record at
        # this point loses the result and leaves the page waiting forever.
        # Daily-Report--new- keeps the record and only clears native Qt refs.
        record = self._tasks.get(task_id)
        if record is None:
            return
        record.worker = None
        record.thread = None
        logger.info(
            "BACKGROUND | THREAD STOPPED | %s | task_id=%s | status=%s",
            record.title,
            task_id,
            record.status,
        )

    def shutdown(self, timeout_ms: int = 15000) -> bool:
        """Wait for live workers before QApplication/native Qt teardown."""
        self._shutting_down = True
        live_threads: list[QThread] = []
        for record in self._tasks.values():
            thread = record.thread
            if thread is None:
                continue
            try:
                running = thread.isRunning()
            except RuntimeError:
                record.thread = None
                record.worker = None
                continue
            if not running:
                record.thread = None
                record.worker = None
                continue
            thread.requestInterruption()
            thread.quit()
            live_threads.append(thread)

        if not live_threads:
            return True

        per_thread = max(250, int(timeout_ms / max(1, len(live_threads))))
        all_stopped = True
        for thread in live_threads:
            try:
                if not thread.wait(per_thread):
                    all_stopped = False
            except RuntimeError:
                continue

        app = QApplication.instance()
        if app is not None:
            app.processEvents()

        if not all_stopped:
            self._shutting_down = False
        return all_stopped

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
    manager = getattr(app, "_background_task_manager", None)
    if not isinstance(manager, BackgroundTaskManager):
        manager = getattr(app, "_procurement_background_task_manager", None)
    if not isinstance(manager, BackgroundTaskManager):
        manager = BackgroundTaskManager(app)
    # Keep both names during the transition so older Procurement helpers and the
    # Daily-Report-style shutdown path always resolve the same manager.
    setattr(app, "_background_task_manager", manager)
    setattr(app, "_procurement_background_task_manager", manager)
    return manager
