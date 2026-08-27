from __future__ import annotations

from typing import Any


def apply_freeze_panes(
    ws: Any,
    *,
    freeze_cell: str | None = None,
    split_row: int | None = None,
    split_column: int | None = None,
    zoom: int = 85,
) -> None:
    """Freeze an Excel worksheet and save it at the top-left scroll position.

    Win32 Excel persists the scroll position of every pane separately.  Merely
    selecting the freeze cell can therefore produce a workbook whose frozen
    pane starts at a later column even though columns A onward are frozen.
    """
    try:
        ws.Activate()
        window = ws.Application.ActiveWindow

        if freeze_cell is not None:
            cell = ws.Range(freeze_cell)
            split_row = max(int(cell.Row) - 1, 0)
            split_column = max(int(cell.Column) - 1, 0)

        row = max(int(split_row or 0), 0)
        column = max(int(split_column or 0), 0)

        window.FreezePanes = False
        window.SplitRow = row
        window.SplitColumn = column
        window.ScrollRow = 1
        window.ScrollColumn = 1
        window.Zoom = zoom
        window.FreezePanes = True

        # Activate the frozen top-left pane before resetting its own scroll.
        # ActiveWindow.ScrollColumn applies to the currently active pane only.
        ws.Range("A1").Select()
        window.ScrollRow = 1
        window.ScrollColumn = 1
    except Exception:
        # Formatting must not prevent the data export from being saved.
        pass
