from decimal import Decimal

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QHeaderView, QTableWidget, QTableWidgetItem


def setup_data_table(table: QTableWidget, *, sorting: bool = True) -> None:
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
):
    item = QTableWidgetItem(format_table_value(value))

    flags = Qt.ItemIsEnabled | Qt.ItemIsSelectable
    if editable:
        flags |= Qt.ItemIsEditable
    item.setFlags(flags)

    item.setTextAlignment(
        Qt.AlignLeft | Qt.AlignVCenter if align_left else Qt.AlignCenter
    )

    if user_data is not None:
        item.setData(Qt.UserRole, user_data)

    return item