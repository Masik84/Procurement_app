from __future__ import annotations

import math
import re

from PySide6.QtCore import QSize, Qt
from PySide6.QtGui import QFont, QFontMetricsF, QPainter
from PySide6.QtWidgets import (
    QHeaderView,
    QStyle,
    QStyleOptionHeader,
    QTableView,
    QTableWidget,
)


# Kept for compatibility with page code that may already use the role.  The
# safe implementation below does NOT replace QTableWidgetItem objects and does
# not override their text()/setText() methods.
HEADER_SOURCE_ROLE = Qt.UserRole + 705


# GUI-only captions. DB/Excel/logical column names stay unchanged in the model.
GUI_HEADER_LABELS: dict[str, str] = {
    "Material number": "Material\nnumber",
    "Supplier Product Name": "Supplier\nProduct Name",
    "Selected Product": "Selected\nProduct",
    "Supplier article": "Supplier\narticle",
    "Supplier Article": "Supplier\nArticle",
    "Customer Product Name": "Customer\nProduct Name",
    "Product Family": "Product\nFamily",
    "Product name (variant)": "Product name\n(variant)",
    "Base currency": "Base \ncurrency",

    "Price, pack": "Price,\npack",
    "Price, box": "Price,\nbox",
    "Qty, pcs": "Qty,\npcs",
    "Qty, box": "Qty,\nbox",
    "Volume, L": "Volume,\nL",
    "Qty in Box": "Qty\nin Box",

    "Product name (for new)": "Product\nname\n(for new)",
    "Brand (for new)": "Brand\n(for new)",
    "Pack (for new)": "Pack\n(for new)",
    "Qty in Box (for new)": "Qty\nin Box\n(for new)",
    "Excise duty (for new)": "Excise\nduty\n(for new)",

    "Cost Novo with VAT": "Cost Novo\nwith VAT",
    "Cost Novo with VAT (prev)": "Cost Novo\nwith VAT (prev)",
    "Full Cost Msk": "Full Cost\nMsk",
    "Full Cost Msk (prev)": "Full Cost\nMsk (prev)",
    "Target Price, pack": "Target Price,\npack",
    "Target Price (Pack)": "Target Price\n(Pack)",
    "Manual Full Cost": "Manual\nFull Cost",
    "Manual Supplier": "Manual\nSupplier",
    "Supplier name": "Supplier\nname",
    "Price date": "Price\ndate",
    "last update": "last\nupdate",
    "last update (prev)": "last update\n(prev)",
    "FX rate Best1": "FX rate\nBest1",
    "FX rate Best2": "FX rate\nBest2",
    "Currency Best1": "Currency\nBest1",
    "Currency Best2": "Currency\nBest2",
    "Best full Price, L": "Best full\nPrice, L",
    "Best full Price, L 2": "Best full\nPrice, L 2",
    "Our Product Name": "Our Product\nName",
    "Volume to take": "Volume\nto take",
    "Purchase Order": "Purchase\nOrder",
    "Reserve E-Comm": "Reserve\nE-Comm",
}

GUI_HEADER_MAX_LINE_LENGTH = 14
GUI_HEADER_MIN_WRAP_LENGTH = 15
GUI_HEADER_MAX_LINES = 3
GUI_HEADER_HORIZONTAL_PADDING = 28
GUI_CELL_HORIZONTAL_PADDING = 20
GUI_COLUMN_MIN_WIDTH = 55
GUI_COLUMN_MAX_WIDTH = 300
AUTOSIZE_FULL_SCAN_ROWS = 300
AUTOSIZE_SAMPLE_ROWS = 120

_SPACE_RE = re.compile(r"\\s+")


def header_display_name(header: object) -> str:
    """Return a GUI-only caption with real line breaks."""
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
            search(end, remaining - 1, current + [" ".join(words[start:end])])

    search(0, line_count, [])
    return best_lines or [" ".join(words)]


class MultilineHeaderView(QHeaderView):
    """Paint wrapped captions without modifying model/header item data.

    This is deliberately a normal Qt header subclass.  No Qt method is replaced
    on a live QTableWidget instance and no QTableWidgetItem virtual method is
    overridden, which avoids the PySide lifetime problems that caused the
    native 0xC0000005 startup crash.
    """

    def paintSection(self, painter: QPainter, rect, logical_index: int) -> None:
        if not rect.isValid():
            return

        option = QStyleOptionHeader()
        self.initStyleOptionForIndex(option, logical_index)
        option.rect = rect
        option.text = header_display_name(option.text)
        option.textAlignment = Qt.AlignmentFlag.AlignCenter
        self.style().drawControl(
            QStyle.ControlElement.CE_Header,
            option,
            painter,
            self,
        )

    def sectionSizeFromContents(self, logical_index: int) -> QSize:
        base = super().sectionSizeFromContents(logical_index)
        model = self.model()
        if model is None:
            return base

        source = model.headerData(
            logical_index,
            Qt.Orientation.Horizontal,
            Qt.ItemDataRole.DisplayRole,
        )
        visible = header_display_name(source)
        lines = visible.splitlines() or [""]
        if len(lines) <= 1:
            return base

        metrics = self.fontMetrics()
        width = max(metrics.horizontalAdvance(line) for line in lines)
        width += GUI_HEADER_HORIZONTAL_PADDING
        height = len(lines) * metrics.lineSpacing() + 10

        return QSize(
            max(GUI_COLUMN_MIN_WIDTH, min(int(width), GUI_COLUMN_MAX_WIDTH)),
            max(base.height(), int(height)),
        )


def install_gui_table_headers(table: QTableView) -> None:
    """Install safe wrapped rendering on Procurement QTableWidget headers."""
    if not isinstance(table, QTableWidget):
        return

    old_header = table.horizontalHeader()
    if isinstance(old_header, MultilineHeaderView):
        return

    # Read properties before setHorizontalHeader(); Qt owns/deletes the old
    # header when it is replaced.
    sections_clickable = old_header.sectionsClickable()
    sections_movable = old_header.sectionsMovable()
    highlight_sections = old_header.highlightSections()
    stretch_last = old_header.stretchLastSection()
    default_alignment = old_header.defaultAlignment()
    minimum_section_size = old_header.minimumSectionSize()
    default_section_size = old_header.defaultSectionSize()
    sort_indicator_shown = old_header.isSortIndicatorShown()
    sort_section = old_header.sortIndicatorSection()
    sort_order = old_header.sortIndicatorOrder()

    header = MultilineHeaderView(Qt.Orientation.Horizontal, table)
    header.setSectionsClickable(sections_clickable)
    header.setSectionsMovable(sections_movable)
    header.setHighlightSections(highlight_sections)
    header.setStretchLastSection(stretch_last)
    header.setDefaultAlignment(default_alignment)
    header.setMinimumSectionSize(minimum_section_size)
    header.setDefaultSectionSize(default_section_size)
    header.setSortIndicatorShown(sort_indicator_shown)
    if sort_section >= 0:
        header.setSortIndicator(sort_section, sort_order)

    table.setHorizontalHeader(header)
    table.setProperty("procurement_multiline_headers_installed", True)


def table_header_name(table: QTableWidget, column: int) -> str:
    """Return the logical header name. Model/header item text is never wrapped."""
    item = table.horizontalHeaderItem(column)
    if item is None:
        return ""
    source = item.data(HEADER_SOURCE_ROLE)
    if source is not None:
        return str(source)
    return " ".join(str(item.text()).splitlines()).strip()


def table_header_names(table: QTableWidget) -> list[str]:
    return [table_header_name(table, column) for column in range(table.columnCount())]


def _source_header_text(table: QTableView, column: int) -> str:
    model = table.model()
    if model is None:
        return ""
    value = model.headerData(
        column,
        Qt.Orientation.Horizontal,
        Qt.ItemDataRole.DisplayRole,
    )
    return str(value or "")


def calculate_gui_header_base_height(
    table: QTableView,
    *,
    base_font_point_size: float,
    minimum_height: int = 24,
) -> int:
    model = table.model()
    if model is None or model.columnCount() <= 0:
        return minimum_height

    max_lines = 1
    for column in range(model.columnCount()):
        visible = header_display_name(_source_header_text(table, column))
        max_lines = max(max_lines, len(visible.splitlines()) or 1)

    font = QFont(table.horizontalHeader().font())
    if base_font_point_size > 0:
        font.setPointSizeF(float(base_font_point_size))
    metrics = QFontMetricsF(font)
    required = math.ceil(max_lines * metrics.lineSpacing() + 10)
    return max(int(minimum_height), required)


def _sample_row_indexes(row_count: int) -> list[int]:
    if row_count <= AUTOSIZE_FULL_SCAN_ROWS:
        return list(range(row_count))
    if row_count <= 0:
        return []

    first_count = min(60, row_count)
    indexes = list(range(first_count))
    remaining = AUTOSIZE_SAMPLE_ROWS - first_count
    if remaining <= 0:
        return indexes

    step = max(1, row_count // remaining)
    indexes.extend(range(first_count, row_count, step))
    return sorted(set(indexes[:AUTOSIZE_SAMPLE_ROWS]))


def resize_columns_for_multiline_headers(table: QTableView) -> None:
    """Compact wrapped-header columns without modifying header data."""
    model = table.model()
    if model is None:
        return

    header_metrics = table.horizontalHeader().fontMetrics()
    body_metrics = table.fontMetrics()
    sample_rows = _sample_row_indexes(model.rowCount())

    for column in range(model.columnCount()):
        visible = header_display_name(_source_header_text(table, column))
        lines = visible.splitlines() or [""]
        if len(lines) <= 1:
            continue

        width = max(header_metrics.horizontalAdvance(line) for line in lines)
        width += GUI_HEADER_HORIZONTAL_PADDING

        for row in sample_rows:
            index = model.index(row, column)
            value = model.data(index, Qt.ItemDataRole.DisplayRole)
            if value is not None:
                cell_text = str(value)
                if len(cell_text) > 80:
                    cell_text = cell_text[:80]
                width = max(
                    width,
                    body_metrics.horizontalAdvance(cell_text)
                    + GUI_CELL_HORIZONTAL_PADDING,
                )

            if isinstance(table, QTableWidget):
                widget = table.cellWidget(row, column)
                if widget is not None:
                    width = max(width, widget.sizeHint().width() + 10)

        width = max(GUI_COLUMN_MIN_WIDTH, min(int(width), GUI_COLUMN_MAX_WIDTH))
        table.horizontalHeader().resizeSection(column, width)
