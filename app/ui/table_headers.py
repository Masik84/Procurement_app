from __future__ import annotations

import math
import re

from PySide6.QtCore import QSize, Qt
from PySide6.QtGui import QFont, QFontMetricsF, QPainter, QPalette
from PySide6.QtWidgets import (
    QHeaderView,
    QProxyStyle,
    QStyle,
    QStyleOption,
    QStyleOptionHeader,
    QTableView,
)


# Единый источник ТОЛЬКО для отображения заголовков в GUI.
# Настоящие/логические названия колонок в таблицах, БД, импорте и Excel
# остаются без \n.
#
# Подход повторяет идею Daily-Report--new-:
# явные переносы для известных длинных заголовков + компактный fallback.
GUI_HEADER_LABELS: dict[str, str] = {
    "Material number": "Material\nnumber",
    "Supplier Product Name": "Supplier Product\nName",
    "Selected Product": "Selected\nProduct",
    "Supplier article": "Supplier\narticle",
    "Customer Product Name": "Customer Product\nName",

    "Price, pack": "Price,\npack",
    "Price, box": "Price,\nbox",
    "Qty, pcs": "Qty,\npcs",
    "Qty, box": "Qty,\nbox",
    "Volume, L": "Volume,\nL",
    "Qty in Box": "Qty\nin Box",

    "Product Family": "Product\nFamily",
    "Product name (for new)": "Product name\n(for new)",
    "Brand (for new)": "Brand\n(for new)",
    "Pack (for new)": "Pack\n(for new)",
    "Qty in Box (for new)": "Qty\nin Box\n(for new)",
    "Excise duty (for new)": "Excise duty\n(for new)",

    "Cost Novo with VAT": "Cost Novo\nwith VAT",
    "Full Cost Msk": "Full Cost\nMsk",
    "Target Price, pack": "Target Price,\npack",
}

GUI_HEADER_MAX_LINE_LENGTH = 14
GUI_HEADER_MIN_WRAP_LENGTH = 15
GUI_HEADER_MAX_LINES = 3
GUI_HEADER_HORIZONTAL_PADDING = 8
GUI_HEADER_VERTICAL_PADDING = 5

_SPACE_RE = re.compile(r"\s+")


def format_gui_header(header: object) -> str:
    """Вернуть компактную GUI-подпись с реальными переносами строк."""
    text = str(header or "").strip()
    if not text:
        return text

    if "\n" in text:
        return text

    text = _SPACE_RE.sub(" ", text)

    explicit = GUI_HEADER_LABELS.get(text)
    if explicit is not None:
        return explicit

    words = text.split(" ")
    if len(words) < 2 or len(text) < GUI_HEADER_MIN_WRAP_LENGTH:
        return text

    two_lines = _balanced_lines(words, 2)
    if max(map(len, two_lines)) <= GUI_HEADER_MAX_LINE_LENGTH:
        return "\n".join(two_lines)

    line_count = min(GUI_HEADER_MAX_LINES, len(words))
    return "\n".join(_balanced_lines(words, line_count))


def _balanced_lines(words: list[str], line_count: int) -> list[str]:
    """Разбить последовательность слов на визуально сбалансированные строки."""
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
    """Высота шапки по максимальному числу строк GUI-заголовка."""
    model = table.model()
    if model is None or model.columnCount() <= 0:
        return minimum_height

    max_lines = 1
    for column in range(model.columnCount()):
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
    """Рисует настоящий многострочный текст заголовка.

    Важно: модель таблицы сохраняет исходный однострочный header text.
    Поэтому существующий код Procurement, который читает
    horizontalHeaderItem(...).text(), продолжает получать прежнее имя колонки.
    """

    def drawControl(
        self,
        element: QStyle.ControlElement,
        option: QStyleOption,
        painter: QPainter,
        widget=None,
    ) -> None:
        if (
            element == QStyle.ControlElement.CE_Header
            and isinstance(option, QStyleOptionHeader)
            and option.orientation == Qt.Orientation.Horizontal
        ):
            display_text = format_gui_header(option.text)

            if "\n" not in display_text:
                super().drawControl(element, option, painter, widget)
                return

            # Сначала базовый стиль рисует фон/границы/стрелку сортировки,
            # но БЕЗ текста.
            background_option = QStyleOptionHeader(option)
            background_option.text = ""
            super().drawControl(element, background_option, painter, widget)

            # Затем текст рисуем сами. В отличие от прежней версии это не
            # зависит от того, вызывает ли системный стиль CE_HeaderLabel.
            label_rect = self.subElementRect(
                QStyle.SubElement.SE_HeaderLabel,
                option,
                widget,
            )

            painter.save()
            try:
                if widget is not None:
                    painter.setFont(widget.font())
                painter.setPen(
                    option.palette.color(QPalette.ColorRole.ButtonText)
                )
                painter.drawText(
                    label_rect,
                    Qt.AlignmentFlag.AlignCenter
                    | Qt.TextFlag.TextWordWrap,
                    display_text,
                )
            finally:
                painter.restore()
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
            display_text = format_gui_header(option.text)
            lines = display_text.splitlines() or [""]

            # Daily-Report--new- считает ширину по самой длинной строке
            # заголовка. Критично НЕ брать max() с width от системного стиля,
            # иначе Windows снова раздувает колонку по исходной длинной строке.
            blank_option = QStyleOptionHeader(option)
            blank_option.text = ""
            base_size = super().sizeFromContents(
                contents_type,
                blank_option,
                QSize(0, 0),
                widget,
            )

            metrics = option.fontMetrics
            text_width = max(
                metrics.horizontalAdvance(line)
                for line in lines
            )
            text_height = len(lines) * metrics.lineSpacing()

            width = (
                base_size.width()
                + text_width
                + GUI_HEADER_HORIZONTAL_PADDING * 2
            )
            height = max(
                base_size.height(),
                text_height + GUI_HEADER_VERTICAL_PADDING * 2,
            )

            return QSize(max(24, width), max(24, height))

        return super().sizeFromContents(
            contents_type,
            option,
            contents_size,
            widget,
        )


def install_gui_header_style(table: QTableView) -> None:
    """Подключить многострочную GUI-шапку к таблице один раз."""
    header: QHeaderView = table.horizontalHeader()
    if header.property("gui_header_wrap_style_installed"):
        return

    proxy = _GuiHeaderProxyStyle()
    proxy.setParent(header)
    header.setStyle(proxy)
    header.setDefaultAlignment(Qt.AlignmentFlag.AlignCenter)
    header.setProperty("gui_header_wrap_style_installed", True)

    # PySide6: держим Python-ссылку, хотя QObject-parent уже назначен.
    header._gui_header_wrap_proxy = proxy
