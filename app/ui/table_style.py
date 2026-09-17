from decimal import Decimal, InvalidOperation
from datetime import date, datetime

from PySide6.QtCore import Qt
from PySide6.QtGui import QBrush, QColor
from PySide6.QtWidgets import QHeaderView, QStyledItemDelegate, QTableWidget, QTableWidgetItem, QWidget

from app.ui.table_headers import (
    header_display_name,
    install_gui_table_headers,
    resize_columns_for_multiline_headers,
    table_header_name,
    table_header_names,
)
from app.ui.table_scale import register_table_for_scaling
from app.utils.excel_format_rules import ABC_VALUE_FILLS, normalize_rule_header
from app.utils.table_sort import numeric_id_value


class NumericTableWidgetItem(QTableWidgetItem):
    """QTableWidget item whose integer text is compared numerically."""

    def __lt__(self, other) -> bool:
        if isinstance(other, QTableWidgetItem):
            left = numeric_id_value(self.data(Qt.ItemDataRole.DisplayRole))
            right = numeric_id_value(other.data(Qt.ItemDataRole.DisplayRole))
            if left is not None and right is not None:
                return left < right
        return super().__lt__(other)


_INTEGER_HEADER_NAMES = frozenset({
    "id",
    "qty in box",
    "qty in box (for new)",
    "кол-во в упак",
    "кол во в упак",
    "кол_во_в_упак",
    "qty, pcs",
    "qty pcs",
    "qty, box",
    "qty box",
})

_TEXT_HEADER_MARKERS = (
    "article",
    "артикул",
    "code",
    "код",
    "number",
    "номер",
    "name",
    "наименование",
    "brand",
    "family",
    "supplier",
    "поставщик",
    "customer",
    "клиент",
    "currency",
    "валюта",
    "status",
    "статус",
    "comment",
    "комментар",
    "document",
    "документ",
    "source",
    "источник",
    "type",
    "тип",
    "batch",
    "user",
    "login",
)

_DATE_HEADER_MARKERS = ("date", "дата")


def _normalise_header_name(value: object) -> str:
    return " ".join(str(value or "").replace("_", " ").split()).strip().casefold()


def _looks_textual_header(header_name: str) -> bool:
    normalized = _normalise_header_name(header_name)
    if any(marker in normalized for marker in _DATE_HEADER_MARKERS):
        return False
    return any(marker in normalized for marker in _TEXT_HEADER_MARKERS)


def _is_integer_header(header_name: str) -> bool:
    return _normalise_header_name(header_name) in _INTEGER_HEADER_NAMES


def _is_date_header(header_name: str) -> bool:
    normalized = _normalise_header_name(header_name)
    return any(marker in normalized for marker in _DATE_HEADER_MARKERS)


def _is_abc_header(header_name: str) -> bool:
    return normalize_rule_header(header_name) == "Категория ABC"


def _format_date_table_value(value) -> str | None:
    if isinstance(value, datetime):
        if value.hour or value.minute or value.second or value.microsecond:
            return value.strftime("%d.%m.%Y %H:%M:%S")
        return value.strftime("%d.%m.%Y")
    if isinstance(value, date):
        return value.strftime("%d.%m.%Y")
    return None


def format_gui_table_value(header_name: str, value) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "Да" if value else "Нет"

    date_text = _format_date_table_value(value)
    if date_text is not None:
        return date_text

    text = str(value).strip()
    if text.casefold() in {"nan", "none", "nat"}:
        return ""
    if _is_date_header(header_name):
        return text
    if _is_integer_header(header_name):
        return format_integer_table_value(value)
    if _looks_textual_header(header_name):
        return text

    number = _table_decimal(value)
    if number is not None:
        return format_decimal_table_value(number)
    if text.endswith("%"):
        return text.replace(".", ",")
    return text


def gui_table_alignment(header_name: str, value=None) -> Qt.AlignmentFlag:
    if _is_abc_header(header_name):
        return Qt.AlignCenter | Qt.AlignVCenter
    if _is_integer_header(header_name) or _is_date_header(header_name):
        return Qt.AlignCenter | Qt.AlignVCenter
    if isinstance(value, (bool, int, float, Decimal)):
        return Qt.AlignCenter | Qt.AlignVCenter
    if not _looks_textual_header(header_name) and _table_decimal(value) is not None:
        return Qt.AlignCenter | Qt.AlignVCenter
    return Qt.AlignLeft | Qt.AlignVCenter


def set_table_alignment_overrides(table: QTableWidget, overrides: dict[str, object] | None = None) -> None:
    """Set opt-in per-header alignment rules without changing other tables."""
    normalized = {
        _normalise_header_name(header_name): alignment
        for header_name, alignment in (overrides or {}).items()
        if str(header_name or "").strip()
    }
    table._procurement_alignment_overrides = normalized
    table.viewport().update()


class GlobalTableDisplayDelegate(QStyledItemDelegate):
    """Delegate that applies the same display rules to every QTableWidget."""

    def __init__(self, table: QTableWidget):
        super().__init__(table)
        self._table = table

    def initStyleOption(self, option, index) -> None:
        super().initStyleOption(option, index)
        table = self._table
        if table is None:
            return
        header_name = table_header_name(table, index.column())
        raw_value = index.data(Qt.ItemDataRole.DisplayRole)
        option.text = format_gui_table_value(header_name, raw_value)

        alignment_overrides = getattr(table, "_procurement_alignment_overrides", None) or {}
        override = alignment_overrides.get(_normalise_header_name(header_name))
        option.displayAlignment = (
            override if override is not None else gui_table_alignment(header_name, raw_value)
        )

        if _is_abc_header(header_name):
            color = ABC_VALUE_FILLS.get(str(raw_value or "").strip().upper())
            if color is not None:
                option.backgroundBrush = QBrush(QColor(*color))

    def displayText(self, value, locale) -> str:
        return "" if value is None else str(value)


def install_global_table_display_rules(table: QTableWidget) -> None:
    if not isinstance(table, QTableWidget):
        return
    if table.property("procurement_global_table_display_rules"):
        return
    delegate = GlobalTableDisplayDelegate(table)
    table.setItemDelegate(delegate)
    table._procurement_global_display_delegate = delegate
    table.setProperty("procurement_global_table_display_rules", True)


def apply_global_table_display_rules(root: QWidget) -> None:
    if isinstance(root, QTableWidget):
        install_global_table_display_rules(root)
    for table in root.findChildren(QTableWidget):
        install_global_table_display_rules(table)


def setup_data_table(table: QTableWidget, *, sorting: bool = True) -> None:
    install_gui_table_headers(table)
    install_global_table_display_rules(table)

    table.setSelectionBehavior(QTableWidget.SelectItems)
    table.setEditTriggers(QTableWidget.DoubleClicked | QTableWidget.EditKeyPressed)
    table.setAlternatingRowColors(True)

    header = table.horizontalHeader()
    header.setSectionResizeMode(QHeaderView.Interactive)
    header.setStretchLastSection(False)

    table.verticalHeader().setVisible(False)
    table.verticalHeader().setSectionResizeMode(QHeaderView.Fixed)
    table.verticalHeader().setDefaultSectionSize(22)
    table.verticalHeader().setMinimumSectionSize(22)
    table.verticalHeader().setMaximumSectionSize(22)
    table.setSortingEnabled(sorting)
    table.setWordWrap(False)
    table.setTextElideMode(Qt.TextElideMode.ElideRight)
    table.setTabKeyNavigation(True)
    table.setCornerButtonEnabled(False)

    table.resizeColumnsToContents()
    resize_columns_for_multiline_headers(table)
    register_table_for_scaling(table)


GUI_DECIMAL_FIELDS = frozenset({"pack", "sales_pack", "new_pack"})
GUI_INTEGER_FIELDS = frozenset({
    "qty_in_box",
    "sales_qty_in_box",
    "product_qty_in_box",
    "new_qty_in_box",
})


def _table_decimal(value) -> Decimal | None:
    if value is None or isinstance(value, bool):
        return None

    text = str(value).strip()
    if not text or text.casefold() in {"nan", "none"}:
        return None

    normalized = text.replace("\xa0", "").replace(" ", "").replace(",", ".")
    try:
        number = Decimal(normalized)
    except (InvalidOperation, ValueError, TypeError):
        return None
    return number if number.is_finite() else None


def format_decimal_table_value(value, *, max_decimals: int = 4) -> str:
    if value is None:
        return ""

    number = _table_decimal(value)
    if number is None:
        return str(value).replace(".", ",")

    max_decimals = max(0, int(max_decimals))
    text = f"{number:.{max_decimals}f}"
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    if text in {"-0", "+0"}:
        text = "0"
    return text.replace(".", ",")


def format_integer_table_value(value) -> str:
    if value is None:
        return ""

    number = _table_decimal(value)
    if number is None:
        return str(value)
    if number == number.to_integral_value():
        return str(int(number))
    return format_decimal_table_value(number)


def format_table_field_value(field_name: str, value) -> str:
    field = str(field_name or "").strip()
    if field in GUI_INTEGER_FIELDS:
        return format_integer_table_value(value)
    if field in GUI_DECIMAL_FIELDS:
        return format_decimal_table_value(value)
    return format_table_value(value)


def format_table_value(value) -> str:
    if value is None:
        return ""
    if isinstance(value, Decimal):
        return str(value).replace(".", ",")
    text = str(value)
    if "." in text and any(ch.isdigit() for ch in text):
        return text.replace(".", ",")
    return text


def build_table_item(
    value,
    *,
    editable: bool = True,
    align_left: bool = False,
    user_data=None,
    numeric_sort: bool = False,
):
    item_class = NumericTableWidgetItem if numeric_sort else QTableWidgetItem
    item = item_class(format_table_value(value))

    flags = Qt.ItemIsEnabled | Qt.ItemIsSelectable
    if editable:
        flags |= Qt.ItemIsEditable
    item.setFlags(flags)

    item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter if align_left else Qt.AlignCenter)

    if user_data is not None:
        item.setData(Qt.UserRole, user_data)

    return item
