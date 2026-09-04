from __future__ import annotations

import math

from PySide6.QtCore import QObject, QSignalBlocker, QTimer, Qt
from PySide6.QtGui import QFont, QFontMetricsF
from PySide6.QtWidgets import QTableView, QTableWidget, QTableWidgetItem


# Same approach as Daily-Report--new-:
# - keep the standard QHeaderView from the .ui file;
# - put real \n into the QTableWidget header item;
# - keep the logical/original name in a separate role.
# No custom QHeaderView, paintSection, QProxyStyle or live-method monkey patching.
HEADER_SOURCE_ROLE = Qt.UserRole + 705


GUI_HEADER_LABELS: dict[str, str] = {
    "Material number": "Material\nnumber",
    "Supplier Product Name": "Supplier Product\nName",
    "Selected Product": "Selected\nProduct",
    "Supplier article": "Supplier\narticle",
    "Supplier Article": "Supplier\nArticle",
    "Customer Product Name": "Customer Product\nName",
    "Product Family": "Product\nFamily",
    "Product name (variant)": "Product name\n(variant)",
    "Base currency": "Base\ncurrency",
    "Mark for us": "Mark for\nus",

    "Price, pack": "Price,\npack",
    "Price, box": "Price,\nbox",
    "Qty, pcs": "Qty,\npcs",
    "Qty, box": "Qty,\nbox",
    "Volume, L": "Volume,\nL",
    "Qty in Box": "Qty\nin Box",

    "Product name (for new)": "Product name\n(for new)",
    "Brand (for new)": "Brand\n(for new)",
    "Pack (for new)": "Pack\n(for new)",
    "Qty in Box (for new)": "Qty\nin Box\n(for new)",
    "Excise duty (for new)": "Excise duty\n(for new)",

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
    "Transport cost per L": "Transport\ncost per L",
    "Re-export %": "Re-export\n%",
    "Insurance %": "Insurance\n%",
    "FX markup %": "FX markup\n%",
    "FX markup abs": "FX markup\nabs",
    "Agent fee": "Agent\nfee",
}

GUI_HEADER_HORIZONTAL_PADDING = 14
GUI_HEADER_SORT_INDICATOR_EXTRA = 16
GUI_COLUMN_MIN_WIDTH = 55


def header_display_name(column_name: object) -> str:
    """Return only the visible GUI header caption."""
    source = str(column_name or "").strip()
    if not source:
        return source
    if "\n" in source:
        return source

    explicit = GUI_HEADER_LABELS.get(source)
    if explicit is not None:
        return explicit

    # Same general rule as Daily-Report--new-: a header of exactly two words
    # is shown in two lines. Longer captions are wrapped only by the explicit
    # map above, so the GUI remains predictable.
    parts = source.split()
    if len(parts) == 2:
        return "\n".join(parts)
    return source


def set_table_header(table: QTableWidget, column: int, column_name: object) -> None:
    """Set one visible multiline header and preserve its logical name."""
    source = str(column_name or "")
    item = QTableWidgetItem(header_display_name(source))
    item.setData(HEADER_SOURCE_ROLE, source)
    item.setTextAlignment(Qt.AlignCenter)
    table.setHorizontalHeaderItem(column, item)


def set_table_headers(table: QTableWidget, columns) -> None:
    """Set all table headers using the Daily-Report--new- scheme."""
    for column, column_name in enumerate(columns):
        set_table_header(table, column, column_name)


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


class _HeaderItemSync(QObject):
    """Convert ordinary setHorizontalHeaderLabels() calls to safe header items.

    Existing Procurement pages already call setHorizontalHeaderLabels() in many
    places. Replacing those live Qt methods caused the previous PySide crashes.
    Instead we listen to the model's normal header-change signals and, after the
    change completes, rewrite only the header items. The standard QHeaderView
    itself is never replaced, so all .ui/QSS styling remains untouched.
    """

    def __init__(self, table: QTableWidget):
        super().__init__(table)
        self._table = table
        self._scheduled = False
        self._applying = False

        model = table.model()
        model.headerDataChanged.connect(self.schedule)
        model.modelReset.connect(self.schedule)
        model.columnsInserted.connect(self.schedule)
        model.columnsRemoved.connect(self.schedule)
        model.rowsInserted.connect(self.schedule)
        model.rowsRemoved.connect(self.schedule)
        model.dataChanged.connect(self.schedule)
        self.schedule()

    def schedule(self, *_args) -> None:
        if self._scheduled or self._applying:
            return
        self._scheduled = True
        QTimer.singleShot(0, self.apply)

    def apply(self) -> None:
        self._scheduled = False
        if self._applying:
            return

        table = self._table
        if table is None:
            return

        self._applying = True
        try:
            model = table.model()
            blocker = QSignalBlocker(model)
            try:
                for column in range(table.columnCount()):
                    item = table.horizontalHeaderItem(column)
                    if item is None:
                        raw = model.headerData(
                            column,
                            Qt.Horizontal,
                            Qt.DisplayRole,
                        )
                        if raw is None:
                            continue
                        source = str(raw)
                    else:
                        stored = item.data(HEADER_SOURCE_ROLE)
                        source = (
                            str(stored)
                            if stored is not None
                            else " ".join(str(item.text()).splitlines()).strip()
                        )

                    visible = header_display_name(source)
                    if item is None:
                        item = QTableWidgetItem(visible)
                        table.setHorizontalHeaderItem(column, item)
                    elif item.text() != visible:
                        item.setText(visible)

                    item.setData(HEADER_SOURCE_ROLE, source)
                    item.setTextAlignment(Qt.AlignCenter)
            finally:
                del blocker

            header = table.horizontalHeader()
            header.setDefaultAlignment(Qt.AlignCenter)
            header.updateGeometry()
            header.viewport().update()
            resize_columns_for_multiline_headers(table)
        finally:
            self._applying = False


def install_gui_table_headers(table: QTableView) -> None:
    """Enable multiline header items while keeping the original QHeaderView."""
    if not isinstance(table, QTableWidget):
        return
    if table.property("procurement_multiline_header_items_installed"):
        return

    sync = _HeaderItemSync(table)
    table._procurement_header_item_sync = sync
    table.setProperty("procurement_multiline_header_items_installed", True)


def _source_header_text(table: QTableView, column: int) -> str:
    if isinstance(table, QTableWidget):
        return table_header_name(table, column)
    model = table.model()
    if model is None:
        return ""
    value = model.headerData(column, Qt.Horizontal, Qt.DisplayRole)
    return str(value or "")


def calculate_gui_header_base_height(
    table: QTableView,
    *,
    base_font_point_size: float,
    minimum_height: int = 24,
) -> int:
    """Calculate height from the same visible text used by header items."""
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
    required = math.ceil(max_lines * metrics.lineSpacing() + 6)
    return max(int(minimum_height), required)


def resize_columns_for_multiline_headers(table: QTableView) -> None:
    """Size every column to the larger of header text or cell contents.

    Qt first calculates the natural width from the actual cells/widgets.  We
    then enforce a minimum based on the longest visible line of the GUI header.
    This keeps short-data columns wide enough to show their complete captions
    while still allowing long cell values to determine a larger width.
    """
    model = table.model()
    if model is None:
        return

    header = table.horizontalHeader()
    header_metrics = header.fontMetrics()

    for column in range(model.columnCount()):
        if table.isColumnHidden(column):
            continue

        # Let Qt calculate the width required by the current cell contents,
        # delegates and cell widgets.  Unlike the previous implementation this
        # is done for ALL columns, not only those whose caption contains \n.
        table.resizeColumnToContents(column)
        content_width = header.sectionSize(column)

        visible = header_display_name(_source_header_text(table, column))
        lines = visible.splitlines() or [""]
        header_width = max(header_metrics.horizontalAdvance(line) for line in lines)
        header_width += GUI_HEADER_HORIZONTAL_PADDING

        # Sorting arrows occupy part of a header section. Reserve space so the
        # final letters of the caption do not disappear under the indicator.
        if header.isSortIndicatorShown():
            header_width += GUI_HEADER_SORT_INDICATOR_EXTRA

        final_width = max(GUI_COLUMN_MIN_WIDTH, int(content_width), int(header_width))
        header.resizeSection(column, final_width)
