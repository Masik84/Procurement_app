from PySide6.QtCore import Qt
from PySide6.QtWidgets import QHeaderView, QTableWidget


def setup_data_table(table: QTableWidget, *, sorting: bool = True) -> None:
    table.setSelectionBehavior(QTableWidget.SelectItems)
    table.setEditTriggers(QTableWidget.DoubleClicked | QTableWidget.EditKeyPressed)
    table.setAlternatingRowColors(True)

    header = table.horizontalHeader()
    header.setSectionResizeMode(QHeaderView.Interactive)
    header.setStretchLastSection(False)

    table.verticalHeader().setVisible(False)
    table.setSortingEnabled(sorting)
    table.setWordWrap(False)
    table.setTextElideMode(Qt.TextElideMode.ElideRight)
    table.setTabKeyNavigation(True)
    table.setCornerButtonEnabled(False)

    table.resizeColumnsToContents()