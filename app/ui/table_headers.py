from __future__ import annotations

import math
import re

from PySide6.QtCore import Qt
from PySide6.QtGui import QFont, QFontMetricsF
from PySide6.QtWidgets import QTableWidget, QTableWidgetItem, QTableView


# The original/logical header name is stored separately from the text shown in
# the GUI. This mirrors Daily-Report--new-: business logic keeps the original
# name, while the user sees compact multi-line captions.
HEADER_SOURCE_ROLE = Qt.UserRole + 705


# Explicit GUI captions for headers where a deliberate line layout is clearer
# than automatic balancing. These values affect GUI tables only; Excel/DB names
# remain unchanged.
GUI_HEADER_LABELS: dict[str, str] = {
    "Material number": "Material\nnumber",
    "Supplier Product Name": "Supplier\nProduct Name",
    "Selected Product": "Selected\nProduct",
    "Supplier article": "Supplier\narticle",
    "Supplier Article": "Supplier\nArticle",
    "Customer Product Name": "Customer\nProduct Name",
    "Product Family": "Product\nFamily",
    "Product name (variant)": "Product name\n(variant)",
    "Категория ABC": "Категория \nABC",
    "Excise duty": "Excise \nduty",
    
    "Base currency": "Base \ncurrency",
    "FX markup %": "FX markup \n%",
    "FX markup abs": "FX markup \nabs",
    "Import duty": "Import \nduty",
    "Rating calc": "Rating \ncalc",
    "Marks for us": "Marks \nfor us",

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

# Long unlisted names are wrapped automatically so the same behaviour works in
# every current/future window without adding page-specific code.
GUI_HEADER_MAX_LINE_LENGTH = 14
GUI_HEADER_MIN_WRAP_LENGTH = 15
GUI_HEADER_MAX_LINES = 3
GUI_HEADER_HORIZONTAL_PADDING = 28
GUI_CELL_HORIZONTAL_PADDING = 20
GUI_COLUMN_MIN_WIDTH = 55
GUI_COLUMN_MAX_WIDTH = 300
AUTOSIZE_FULL_SCAN_ROWS = 300
AUTOSIZE_SAMPLE_ROWS = 120

_SPACE_RE = re.compile(r"\s+")


def header_display_name(header: object) -> str:
    """Return the GUI caption for a logical header name."""
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


class GuiHeaderItem(QTableWidgetItem):
    """Header item with separate logical and visible text.

    Qt paints DisplayRole, which contains the real multi-line caption. Python
    page code historically calls item.text(); keeping that call returning the
    logical source name avoids breaking existing save/filter/edit logic.
    """

    def __init__(self, source_name: object):
        source = str(source_name or "")
        super().__init__()
        QTableWidgetItem.setData(self, HEADER_SOURCE_ROLE, source)
        QTableWidgetItem.setData(
            self,
            Qt.ItemDataRole.DisplayRole,
            header_display_name(source),
        )
        self.setTextAlignment(Qt.AlignmentFlag.AlignCenter)

    def text(self) -> str:
        source = self.data(HEADER_SOURCE_ROLE)
        if source is not None:
            return str(source)
        return QTableWidgetItem.text(self).replace("\n", " ")

    def setText(self, text: str) -> None:
        source = str(text or "")
        QTableWidgetItem.setData(self, HEADER_SOURCE_ROLE, source)
        QTableWidgetItem.setData(
            self,
            Qt.ItemDataRole.DisplayRole,
            header_display_name(source),
        )


def _copy_header_roles(source_item: QTableWidgetItem, target_item: QTableWidgetItem) -> None:
    """Preserve optional styling/metadata from explicitly supplied header items."""
    roles = (
        Qt.ItemDataRole.DecorationRole,
        Qt.ItemDataRole.FontRole,
        Qt.ItemDataRole.TextAlignmentRole,
        Qt.ItemDataRole.BackgroundRole,
        Qt.ItemDataRole.ForegroundRole,
        Qt.ItemDataRole.CheckStateRole,
        Qt.ItemDataRole.SizeHintRole,
        Qt.ItemDataRole.ToolTipRole,
        Qt.ItemDataRole.StatusTipRole,
        Qt.ItemDataRole.WhatsThisRole,
    )
    for role in roles:
        value = source_item.data(role)
        if value is not None:
            target_item.setData(role, value)


def _logical_name_from_item(item: QTableWidgetItem) -> str:
    source = item.data(HEADER_SOURCE_ROLE)
    if source is not None:
        return str(source)

    # For a normal QTableWidgetItem, text() is the source text. For an already
    # multi-line item supplied by a page, normalise only the line separators.
    return " ".join(str(QTableWidgetItem.text(item)).splitlines()).strip()


def install_gui_table_headers(table: QTableWidget) -> None:
    """Install Daily-Report-style multi-line headers on one table.

    Every active Procurement table calls setup_data_table(), so installing the
    wrappers there applies the behaviour to every application window and to
    dynamically created headers as well.
    """
    if table.property("procurement_multiline_headers_installed"):
        return

    original_set_labels = table.setHorizontalHeaderLabels
    original_set_item = table.setHorizontalHeaderItem

    table._procurement_original_set_horizontal_header_labels = original_set_labels
    table._procurement_original_set_horizontal_header_item = original_set_item

    def set_horizontal_header_labels(labels) -> None:
        logical_names = [str(value or "") for value in labels]
        # Let QTableWidget do its normal bookkeeping first, then replace every
        # generated header item with our source/display-aware item.
        original_set_labels(logical_names)
        for column, logical_name in enumerate(logical_names):
            current = table.horizontalHeaderItem(column)
            new_item = GuiHeaderItem(logical_name)
            if current is not None:
                _copy_header_roles(current, new_item)
            original_set_item(column, new_item)

    def set_horizontal_header_item(column: int, item: QTableWidgetItem | None) -> None:
        if item is None:
            original_set_item(column, item)
            return
        if isinstance(item, GuiHeaderItem):
            original_set_item(column, item)
            return

        logical_name = _logical_name_from_item(item)
        new_item = GuiHeaderItem(logical_name)
        _copy_header_roles(item, new_item)
        original_set_item(column, new_item)

    table.setHorizontalHeaderLabels = set_horizontal_header_labels
    table.setHorizontalHeaderItem = set_horizontal_header_item
    table.setProperty("procurement_multiline_headers_installed", True)

    # Convert headers that were already defined in the .ui file before
    # setup_data_table() was called.
    for column in range(table.columnCount()):
        current = table.horizontalHeaderItem(column)
        if current is None or isinstance(current, GuiHeaderItem):
            continue
        logical_name = _logical_name_from_item(current)
        new_item = GuiHeaderItem(logical_name)
        _copy_header_roles(current, new_item)
        original_set_item(column, new_item)


def table_header_name(table: QTableWidget, column: int) -> str:
    """Return the logical header name without GUI line breaks."""
    item = table.horizontalHeaderItem(column)
    if item is None:
        return ""
    source = item.data(HEADER_SOURCE_ROLE)
    if source is not None:
        return str(source)
    return " ".join(str(item.text()).splitlines()).strip()


def table_header_names(table: QTableWidget) -> list[str]:
    return [table_header_name(table, column) for column in range(table.columnCount())]


def _header_display_text(table: QTableView, column: int) -> str:
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
    """Calculate header height from the actual visible line count."""
    model = table.model()
    if model is None or model.columnCount() <= 0:
        return minimum_height

    max_lines = 1
    for column in range(model.columnCount()):
        visible = _header_display_text(table, column)
        if "\n" not in visible:
            visible = header_display_name(visible)
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
    """Size multi-line columns by the longest visible header line + cell data.

    Daily-Report--new- deliberately measures each header line separately. This
    prevents a caption such as ``Qty\nin Box\n(for new)`` from making the whole
    column as wide as ``Qty in Box (for new)``.
    """
    model = table.model()
    if model is None:
        return

    header_metrics = table.horizontalHeader().fontMetrics()
    body_metrics = table.fontMetrics()
    row_count = model.rowCount()
    sample_rows = _sample_row_indexes(row_count)

    for column in range(model.columnCount()):
        header_text = _header_display_text(table, column)
        lines = header_text.splitlines() or [""]
        if len(lines) <= 1:
            continue

        width = max(
            header_metrics.horizontalAdvance(line)
            for line in lines
        ) + GUI_HEADER_HORIZONTAL_PADDING

        for row in sample_rows:
            index = model.index(row, column)
            value = model.data(index, Qt.ItemDataRole.DisplayRole)
            if value is not None:
                cell_text = str(value)
                if len(cell_text) > 80:
                    cell_text = cell_text[:80]
                width = max(
                    width,
                    body_metrics.horizontalAdvance(cell_text) + GUI_CELL_HORIZONTAL_PADDING,
                )

            if isinstance(table, QTableWidget):
                widget = table.cellWidget(row, column)
                if widget is not None:
                    width = max(width, widget.sizeHint().width() + 10)

        width = max(GUI_COLUMN_MIN_WIDTH, min(int(width), GUI_COLUMN_MAX_WIDTH))
        # resizeSection bypasses Procurement's setColumnWidth wrapper; while the
        # scale manager is performing autosize it also avoids feeding scaled
        # dimensions back as new base widths.
        table.horizontalHeader().resizeSection(column, width)
