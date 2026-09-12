from __future__ import annotations

from enum import Enum

from PySide6.QtCore import Qt
from PySide6.QtGui import QGuiApplication
from PySide6.QtWidgets import (
    QApplication,
    QDialog,
    QHBoxLayout,
    QLabel,
    QPlainTextEdit,
    QPushButton,
    QSizePolicy,
    QStyle,
    QVBoxLayout,
    QWidget,
)


LONG_TEXT_LIMIT = 420
SUMMARY_TEXT_LIMIT = 280
DETAILS_MAX_WIDTH = 760
DETAILS_MAX_HEIGHT = 560
COMPACT_MAX_WIDTH = 540


class MessageKind(str, Enum):
    ERROR = "error"
    WARNING = "warning"
    INFO = "info"
    QUESTION = "question"


_ICON_BY_KIND = {
    MessageKind.ERROR: QStyle.StandardPixmap.SP_MessageBoxCritical,
    MessageKind.WARNING: QStyle.StandardPixmap.SP_MessageBoxWarning,
    MessageKind.INFO: QStyle.StandardPixmap.SP_MessageBoxInformation,
    MessageKind.QUESTION: QStyle.StandardPixmap.SP_MessageBoxQuestion,
}


_DEFAULT_TITLE = {
    MessageKind.ERROR: "Ошибка",
    MessageKind.WARNING: "Предупреждение",
    MessageKind.INFO: "Сообщение",
    MessageKind.QUESTION: "Подтверждение",
}


def _clean_text(text: object) -> str:
    value = str(text or "").strip()
    return value or "Нет текста сообщения."


def _is_long_message(text: str) -> bool:
    if len(text) > LONG_TEXT_LIMIT:
        return True
    lines = text.splitlines()
    if len(lines) > 5:
        return True
    return any(len(line) > SUMMARY_TEXT_LIMIT for line in lines)


def _summary_for(text: str) -> str:
    """Return a compact, useful first view without DB/traceback payloads."""
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    if not lines:
        return "Нет текста сообщения."

    # Database/traceback errors usually start with the useful human-readable
    # exception and then continue with LINE/SQL/parameters.  Keep the useful
    # first part in the compact view and put the complete payload behind Show.
    stop_prefixes = (
        "LINE ",
        "[SQL:",
        "[parameters:",
        "[parameters ",
        "(Background on this error",
        "Traceback ",
    )
    useful: list[str] = []
    for line in lines:
        if useful and line.startswith(stop_prefixes):
            break
        if line == "^":
            break
        useful.append(line)
        if len(" ".join(useful)) >= SUMMARY_TEXT_LIMIT:
            break
        # One normal exception line is usually the best compact summary.
        if len(useful) >= 2:
            break

    summary = " ".join(useful) if useful else lines[0]
    summary = " ".join(summary.split())
    if len(summary) <= SUMMARY_TEXT_LIMIT:
        return summary

    cut = summary[: SUMMARY_TEXT_LIMIT - 1].rstrip()
    if " " in cut:
        cut = cut.rsplit(" ", 1)[0]
    return f"{cut}…"


class AppMessageDialog(QDialog):
    """Single application-wide popup for info, warnings, errors and questions.

    Long messages always open compact.  The complete text is available through
    Show in a bounded, scrollable text area and can be copied in full.
    """

    def __init__(
        self,
        parent: QWidget | None,
        *,
        title: str,
        text: object,
        kind: MessageKind,
        question: bool = False,
        default_yes: bool = False,
    ) -> None:
        super().__init__(parent)
        self.setModal(True)
        self.setWindowTitle(title or _DEFAULT_TITLE[kind])
        self.setWindowFlag(Qt.WindowContextHelpButtonHint, False)

        self._full_text = _clean_text(text)
        self._has_details = _is_long_message(self._full_text)
        self._question = question
        self._default_yes = default_yes
        self._expanded = False

        self._build_ui(kind)
        self._apply_compact_geometry()

    def _build_ui(self, kind: MessageKind) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(18, 16, 18, 14)
        root.setSpacing(12)

        message_row = QHBoxLayout()
        message_row.setSpacing(14)

        icon_label = QLabel(self)
        icon = QApplication.style().standardIcon(_ICON_BY_KIND[kind])
        icon_label.setPixmap(icon.pixmap(32, 32))
        icon_label.setFixedSize(36, 36)
        icon_label.setAlignment(Qt.AlignTop | Qt.AlignHCenter)
        message_row.addWidget(icon_label, 0, Qt.AlignTop)

        summary = QLabel(self)
        summary.setObjectName("appMessageSummary")
        summary.setText(_summary_for(self._full_text) if self._has_details else self._full_text)
        summary.setWordWrap(True)
        summary.setTextInteractionFlags(Qt.TextSelectableByMouse | Qt.TextSelectableByKeyboard)
        summary.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        message_row.addWidget(summary, 1)
        root.addLayout(message_row)

        self._details = QPlainTextEdit(self)
        self._details.setObjectName("appMessageDetails")
        self._details.setReadOnly(True)
        self._details.setPlainText(self._full_text)
        self._details.setLineWrapMode(QPlainTextEdit.LineWrapMode.NoWrap)
        self._details.setMinimumHeight(220)
        self._details.setMaximumHeight(360)
        self._details.setVisible(False)
        root.addWidget(self._details, 1)

        buttons = QHBoxLayout()
        buttons.setSpacing(8)

        copy_button = QPushButton("Copy", self)
        copy_button.setToolTip("Скопировать полный текст сообщения")
        copy_button.clicked.connect(self._copy_full_text)
        buttons.addWidget(copy_button)

        self._show_button: QPushButton | None = None
        if self._has_details:
            self._show_button = QPushButton("Show", self)
            self._show_button.setToolTip("Показать полный текст")
            self._show_button.clicked.connect(self._toggle_details)
            buttons.addWidget(self._show_button)

        buttons.addStretch(1)

        if self._question:
            yes_button = QPushButton("Да", self)
            no_button = QPushButton("Нет", self)
            yes_button.clicked.connect(self.accept)
            no_button.clicked.connect(self.reject)
            if self._default_yes:
                yes_button.setDefault(True)
                yes_button.setFocus()
            else:
                no_button.setDefault(True)
                no_button.setFocus()
            buttons.addWidget(yes_button)
            buttons.addWidget(no_button)
        else:
            ok_button = QPushButton("OK", self)
            ok_button.clicked.connect(self.accept)
            ok_button.setDefault(True)
            ok_button.setFocus()
            buttons.addWidget(ok_button)

        root.addLayout(buttons)

    def _available_geometry(self):
        screen = self.screen()
        if screen is None and self.parentWidget() is not None:
            screen = self.parentWidget().screen()
        if screen is None:
            screen = QGuiApplication.primaryScreen()
        return screen.availableGeometry() if screen is not None else None

    def _apply_compact_geometry(self) -> None:
        available = self._available_geometry()
        if available is None:
            self.setMaximumWidth(COMPACT_MAX_WIDTH)
            self.adjustSize()
            return

        width = min(COMPACT_MAX_WIDTH, max(360, available.width() - 80))
        self.setMaximumWidth(width)
        self.resize(width, self.sizeHint().height())
        self._center_in_available_geometry(available)

    def _apply_expanded_geometry(self) -> None:
        available = self._available_geometry()
        if available is None:
            self.setMaximumSize(DETAILS_MAX_WIDTH, DETAILS_MAX_HEIGHT)
            self.resize(DETAILS_MAX_WIDTH, DETAILS_MAX_HEIGHT)
            return

        width = min(DETAILS_MAX_WIDTH, max(420, int(available.width() * 0.80)))
        height = min(DETAILS_MAX_HEIGHT, max(340, int(available.height() * 0.75)))
        width = min(width, available.width() - 40)
        height = min(height, available.height() - 40)
        self.setMaximumSize(width, height)
        self.resize(width, height)
        self._center_in_available_geometry(available)

    def _center_in_available_geometry(self, available) -> None:
        frame = self.frameGeometry()
        frame.moveCenter(available.center())
        self.move(frame.topLeft())

    def _copy_full_text(self) -> None:
        QGuiApplication.clipboard().setText(self._full_text)

    def _toggle_details(self) -> None:
        if self._show_button is None:
            return
        self._expanded = not self._expanded
        self._details.setVisible(self._expanded)
        self._show_button.setText("Hide" if self._expanded else "Show")
        self._show_button.setToolTip("Скрыть полный текст" if self._expanded else "Показать полный текст")

        if self._expanded:
            self._apply_expanded_geometry()
            self._details.setFocus()
        else:
            # Drop expanded size limits before recalculating compact size.
            self.setMaximumSize(16777215, 16777215)
            self._apply_compact_geometry()


def show_message(
    parent: QWidget | None,
    text: object,
    *,
    title: str | None = None,
    kind: MessageKind = MessageKind.INFO,
) -> None:
    dialog = AppMessageDialog(
        parent,
        title=title or _DEFAULT_TITLE[kind],
        text=text,
        kind=kind,
    )
    dialog.exec()


def show_error(parent: QWidget | None, text: object, *, title: str = "Ошибка") -> None:
    show_message(parent, text, title=title, kind=MessageKind.ERROR)


def show_warning(parent: QWidget | None, text: object, *, title: str = "Предупреждение") -> None:
    show_message(parent, text, title=title, kind=MessageKind.WARNING)


def show_info(parent: QWidget | None, text: object, *, title: str = "Сообщение") -> None:
    show_message(parent, text, title=title, kind=MessageKind.INFO)


def ask_yes_no(
    parent: QWidget | None,
    text: object,
    *,
    title: str = "Подтверждение",
    default_yes: bool = False,
) -> bool:
    dialog = AppMessageDialog(
        parent,
        title=title,
        text=text,
        kind=MessageKind.QUESTION,
        question=True,
        default_yes=default_yes,
    )
    return dialog.exec() == QDialog.Accepted
