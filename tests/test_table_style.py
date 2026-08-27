from __future__ import annotations

import os
import unittest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication, QTableWidget

from app.ui.table_style import build_table_item


class NumericTableWidgetItemTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication([])

    def test_id_items_sort_in_numeric_order_in_qtablewidget(self) -> None:
        table = QTableWidget(0, 1)
        table.setHorizontalHeaderLabels(["id"])
        for row, value in enumerate(["1", "10", "100", "1000", "2", "3", "11"]):
            table.insertRow(row)
            table.setItem(row, 0, build_table_item(value, numeric_sort=True))

        table.sortItems(0, Qt.SortOrder.AscendingOrder)

        self.assertEqual(
            [table.item(row, 0).text() for row in range(table.rowCount())],
            ["1", "2", "3", "10", "11", "100", "1000"],
        )


if __name__ == "__main__":
    unittest.main()
