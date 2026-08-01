from __future__ import annotations

import math
import re

from PySide6.QtCore import QSize, Qt
from PySide6.QtGui import QFont, QFontMetricsF, QPainter
from PySide6.QtWidgets import (
    QHeaderView,
    QProxyStyle,
    QStyle,
    QStyleOption,
    QStyleOptionHeader,
    QTableView,
)


# Header text is changed only while Qt paints the GUI. The model/header values
# themselves stay untouched, so DB fields, internal names and Excel exports keep
# their original spelling without line breaks.
GUI_HEADER_MAX_LINE_LENGTH = 14
GUI_HEADER_MIN_WRAP_LENGTH = 13
GUI_HEADER_MAX_LINES = 3
GUI_HEADER_HORIZONTAL_PADDING = 8
GUI_HEADER_VERTICAL_PADDING = 6

_SPACE_RE = re.compile(r"\s+")


def format_gui_header(header: object) -> str:
    """Return a compact GUI-only header with balanced line breaks."""
    text = str(header or "").strip()
    if not text or "\n" in text:
        return text

    text = _SPACE_RE.sub(" ", text)
    words = text.split(" ")
    if len(words) < 2 or len(text) < GUI_HEADER_MIN_WRAP_LENGTH:
        return text

    two_lines = _balanced_lines(words, 2)
    if max(map(len, two_lines)) <= GUI_HEADER_MAX_LINE_LENGTH:
        return "\n".join(two_lines)

    line_count = min(GUI_HEADER_MAX_LINES, len(words))
    return "\n".join(_balanced_lines(words, line_count))


def _balanced_lines(words: list[str], line_count: int) -> list[str]:
    """Split contiguous words into visually balanced lines."""
    if line_count <= 1 or len(words) <= 1:
        return [" ".join(words)]

    line_count = min(line_count, len(words))
    best_lines: list[str] | None = None
    best_score: tuple[int, int, int] | None = None

    def search(start: int, remaining: int, current: list[str]) -> None:
        nonlocal best_lines, best_score
        if remaining == 1:
            candidate = current + [" ".join(words[start:])]
            lengths = [len(line) for line in candidate]
            score = (
                max(lengths),
                max(lengths) - min(lengths),
                sum(length * length for length in lengths),
            )
            if best_score is None or score < best_score:
                best_score = score
                best_lines = candidate
            return

        last_start = len(words) - remaining + 1
        for end in range(start + 1, last_start + 1):
            search(
                end,
                remaining - 1,
                current + [" ".join(words[start:end])],
            )

    search(0, line_count, [])
    return best_lines or [" ".join(words)]


def gui_header_line_count(header: object) -> int:
    formatted = format_gui_header(header)
    return max(1, formatted.count("\n") + 1)


def calculate_gui_header_base_height(
    table: QTableView,
    *,
    base_font_point_size: float,
    minimum_height: int = 24,
) -> int:
    """Calculate the unscaled height required by the longest GUI header."""
    model = table.model()
    if model is None:
        return minimum_height

    columns = model.columnCount()
    if columns <= 0:
        return minimum_height

    max_lines = 1
    for column in range(columns):
        value = model.headerData(
            column,
            Qt.Orientation.Horizontal,
            Qt.ItemDataRole.DisplayRole,
        )
        max_lines = max(max_lines, gui_header_line_count(value))

    font = QFont(table.horizontalHeader().font())
    if base_font_point_size > 0:
        font.setPointSizeF(float(base_font_point_size))
    metrics = QFontMetricsF(font)

    required = math.ceil(
        max_lines * metrics.lineSpacing() + GUI_HEADER_VERTICAL_PADDING * 2
    )
    return max(int(minimum_height), required)


class _GuiHeaderProxyStyle(QProxyStyle):
    """Draw line breaks in horizontal GUI headers without changing model data."""

    def drawControl(
        self,
        element: QStyle.ControlElement,
        option: QStyleOption,
        painter: QPainter,
        widget=None,
    ) -> None:
        if (
            element == QStyle.ControlElement.CE_HeaderLabel
            and isinstance(option, QStyleOptionHeader)
            and option.orientation == Qt.Orientation.Horizontal
        ):
            wrapped = QStyleOptionHeader(option)
            wrapped.text = format_gui_header(option.text)
            wrapped.textAlignment = Qt.AlignmentFlag.AlignCenter
            super().drawControl(element, wrapped, painter, widget)
            return

        super().drawControl(element, option, painter, widget)

    def sizeFromContents(
        self,
        contents_type: QStyle.ContentsType,
        option: QStyleOption,
        contents_size: QSize,
        widget=None,
    ) -> QSize:
        if (
            contents_type == QStyle.ContentsType.CT_HeaderSection
            and isinstance(option, QStyleOptionHeader)
            and option.orientation == Qt.Orientation.Horizontal
        ):
            wrapped = QStyleOptionHeader(option)
            wrapped.text = format_gui_header(option.text)
            size = super().sizeFromContents(
                contents_type,
                wrapped,
                contents_size,
                widget,
            )
            lines = wrapped.text.splitlines() or [""]
            metrics = wrapped.fontMetrics
            text_width = max(metrics.horizontalAdvance(line) for line in lines)
            text_height = len(lines) * metrics.lineSpacing()
            size.setWidth(
                max(size.width(), text_width + GUI_HEADER_HORIZONTAL_PADDING * 2)
            )
            size.setHeight(
                max(size.height(), text_height + GUI_HEADER_VERTICAL_PADDING * 2)
            )
            return size

        return super().sizeFromContents(
            contents_type,
            option,
            contents_size,
            widget,
        )


def install_gui_header_style(table: QTableView) -> None:
    """Install the GUI-only wrapping style once for a table header."""
    header: QHeaderView = table.horizontalHeader()
    if header.property("gui_header_wrap_style_installed"):
        return

    proxy = _GuiHeaderProxyStyle()
    proxy.setParent(header)
    header.setStyle(proxy)
    header.setDefaultAlignment(Qt.AlignmentFlag.AlignCenter)
    header.setProperty("gui_header_wrap_style_installed", True)
    # Keep a Python reference as well; this avoids premature wrapper cleanup in
    # some PySide6 versions even though the QObject parent is already assigned.
    header._gui_header_wrap_proxy = proxy
