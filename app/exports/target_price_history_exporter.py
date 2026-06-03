from __future__ import annotations

from datetime import date, datetime, time
from pathlib import Path
from typing import Sequence

import pythoncom
import pywintypes
import win32com.client as win32

from app.exports.excel_column_format import (
    DATE_HEADERS,
    apply_standard_worksheet_format,
    excel_value_by_header,
    normalize_header,
)


class TargetPriceHistoryExporter:
    """Excel export for TargetPriceHistoryPage.

    Visual Excel rules — colors, widths, alignments, number formats and zoom —
    are stored in app.exports.excel_column_format and are applied from there.
    """

    @staticmethod
    def _safe_sheet_title(value: str) -> str:
        title = (value or "Target prices")[:31]
        for ch in ["\\", "/", ":", "*", "?", "[", "]"]:
            title = title.replace(ch, "_")
        return title or "Target prices"

    @staticmethod
    def _create_excel_app():
        pythoncom.CoInitialize()
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        return excel

    @staticmethod
    def _to_com_date(value: object) -> object:
        if isinstance(value, datetime):
            return pywintypes.Time(datetime.combine(value.date(), time.min))
        if isinstance(value, date):
            return pywintypes.Time(datetime.combine(value, time.min))
        return value

    def _excel_value(self, header: str, value: object) -> object:
        base = normalize_header(header)
        result = excel_value_by_header(header, value)
        if base in DATE_HEADERS:
            return self._to_com_date(result)
        return result

    def _write_table(self, ws, headers: Sequence[str], rows: Sequence[Sequence[object]]) -> None:
        for col_index, header in enumerate(headers, start=1):
            ws.Cells(1, col_index).Value = str(header)

        for row_index, row_values in enumerate(rows, start=2):
            for col_index, header in enumerate(headers, start=1):
                value = row_values[col_index - 1] if col_index - 1 < len(row_values) else ""
                ws.Cells(row_index, col_index).Value = self._excel_value(str(header), value)

    def export_report(self, *, headers: Sequence[str], rows: Sequence[Sequence[object]], output_path: str | Path) -> Path:
        output_path = Path(output_path)
        if output_path.suffix.lower() != ".xlsx":
            output_path = output_path.with_suffix(".xlsx")
        output_path.parent.mkdir(parents=True, exist_ok=True)
        target_path = output_path.resolve()

        if target_path.exists():
            try:
                target_path.unlink()
            except PermissionError:
                raise PermissionError(
                    f"Не удается перезаписать файл:\n{target_path}\n\n"
                    f"Скорее всего, он открыт в Excel. Закрой файл и попробуй снова."
                )

        excel = None
        wb = None
        try:
            excel = self._create_excel_app()
            wb = excel.Workbooks.Add()
            ws = wb.Worksheets(1)
            ws.Name = self._safe_sheet_title("Target prices")

            header_list = [str(h) for h in headers]
            self._write_table(ws, header_list, rows)
            apply_standard_worksheet_format(ws, header_list, freeze_cell="F2", zoom=85)

            wb.SaveAs(str(target_path))
            return target_path
        finally:
            try:
                if wb is not None:
                    wb.Close(SaveChanges=False)
            except Exception:
                pass
            try:
                if excel is not None:
                    excel.Quit()
            except Exception:
                pass
            pythoncom.CoUninitialize()
