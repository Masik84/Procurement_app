from __future__ import annotations

import logging
from pathlib import Path

import pythoncom
import win32com.client as win32

from app.utils.excel_fast_writer import write_excel_table
from app.utils.excel_format_rules import FORMATS, save_workbook_xlsx, set_number_format_safe


logger = logging.getLogger(__name__)


class ProductUc3Exporter:
    HEADERS = ["Product Name", "Target uC3", "Walk-Away uC3"]

    def _create_excel_app(self):
        pythoncom.CoInitialize()
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        return excel

    @staticmethod
    def _prepare_path(file_path: str | Path) -> Path:
        path = Path(file_path)
        if path.suffix.lower() != ".xlsx":
            path = path.with_suffix(".xlsx")
        path.parent.mkdir(parents=True, exist_ok=True)
        return path.resolve()

    def _format_sheet(self, ws, *, history: bool = False):
        ws.Cells.Font.Name = "Aptos Narrow"
        ws.Cells.Font.Size = 11
        last_col = "D" if history else "C"
        header = ws.Range(f"A1:{last_col}1")
        header.Font.Name = "Aptos Narrow"
        header.Font.Size = 11
        header.Font.Bold = True
        header.WrapText = True
        header.HorizontalAlignment = -4108
        header.VerticalAlignment = -4160
        header.Interior.Color = 0xCDCDCD
        ws.Rows(1).RowHeight = 32
        ws.Columns("A:A").ColumnWidth = 42
        ws.Columns("B:C").ColumnWidth = 16
        set_number_format_safe(ws.Columns("B:C"), FORMATS.INTEGER)
        if history:
            ws.Columns("D:D").ColumnWidth = 14
            set_number_format_safe(ws.Columns("D:D"), FORMATS.DATE)
        ws.Range(f"A1:{last_col}1").AutoFilter(1)

    def export_template(self, file_path: str | Path) -> Path:
        target = self._prepare_path(file_path)
        excel = wb = None
        try:
            if target.exists():
                target.unlink()
            excel = self._create_excel_app()
            wb = excel.Workbooks.Add()
            ws = wb.Worksheets(1)
            ws.Name = "Target uC3"
            for idx, header in enumerate(self.HEADERS, 1):
                ws.Cells(1, idx).Value = header
            self._format_sheet(ws)
            save_workbook_xlsx(wb, target)
            return target
        finally:
            try:
                if wb is not None:
                    wb.Close(SaveChanges=False)
            except Exception:
                logger.exception("Ошибка закрытия Excel")
            try:
                if excel is not None:
                    excel.Quit()
            except Exception:
                logger.exception("Ошибка закрытия Excel")
            pythoncom.CoUninitialize()

    def export_rows(self, rows: list[dict], file_path: str | Path) -> Path:
        target = self._prepare_path(file_path)
        excel = wb = None
        try:
            if target.exists():
                target.unlink()
            excel = self._create_excel_app()
            wb = excel.Workbooks.Add()
            ws = wb.Worksheets(1)
            ws.Name = "Target uC3"
            headers = self.HEADERS + ["Change date"]
            write_excel_table(
                ws,
                headers,
                rows,
                value_getter=lambda row, header, _col: row.get(header, ""),
            )
            self._format_sheet(ws, history=True)
            save_workbook_xlsx(wb, target)
            return target
        finally:
            try:
                if wb is not None:
                    wb.Close(SaveChanges=False)
            except Exception:
                logger.exception("Ошибка закрытия Excel")
            try:
                if excel is not None:
                    excel.Quit()
            except Exception:
                logger.exception("Ошибка закрытия Excel")
            pythoncom.CoUninitialize()
