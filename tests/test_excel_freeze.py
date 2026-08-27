from __future__ import annotations

import unittest

from app.utils.excel_freeze import apply_freeze_panes


class _Cell:
    def __init__(self, row: int, column: int, ws: "_Worksheet") -> None:
        self.Row = row
        self.Column = column
        self._ws = ws

    def Select(self) -> None:
        self._ws.selected = (self.Row, self.Column)


class _Window:
    def __init__(self) -> None:
        self.FreezePanes = True
        self.SplitRow = 0
        self.SplitColumn = 0
        self.ScrollRow = 9
        self.ScrollColumn = 7
        self.Zoom = 100


class _Worksheet:
    def __init__(self) -> None:
        self.window = _Window()
        self.Application = type("Application", (), {"ActiveWindow": self.window})()
        self.activated = False
        self.selected: tuple[int, int] | None = None

    def Activate(self) -> None:
        self.activated = True

    def Range(self, address: str) -> _Cell:
        cells = {"A1": (1, 1), "N2": (2, 14)}
        row, column = cells[address]
        return _Cell(row, column, self)


class ApplyFreezePanesTests(unittest.TestCase):
    def test_freeze_cell_resets_the_frozen_pane_to_column_a(self) -> None:
        ws = _Worksheet()

        apply_freeze_panes(ws, freeze_cell="N2", zoom=85)

        self.assertTrue(ws.activated)
        self.assertTrue(ws.window.FreezePanes)
        self.assertEqual(ws.window.SplitRow, 1)
        self.assertEqual(ws.window.SplitColumn, 13)
        self.assertEqual(ws.window.ScrollRow, 1)
        self.assertEqual(ws.window.ScrollColumn, 1)
        self.assertEqual(ws.window.Zoom, 85)
        self.assertEqual(ws.selected, (1, 1))

    def test_explicit_split_is_applied_without_a_freeze_cell(self) -> None:
        ws = _Worksheet()

        apply_freeze_panes(ws, split_row=1, split_column=5)

        self.assertEqual(ws.window.SplitRow, 1)
        self.assertEqual(ws.window.SplitColumn, 5)
        self.assertEqual(ws.window.ScrollColumn, 1)


if __name__ == "__main__":
    unittest.main()
