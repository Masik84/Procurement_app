from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import Any, Sequence

import pythoncom
import win32com.client as win32

from app.exports.excel_column_format import (
    DEFAULT_FONT_NAME,
    DEFAULT_FONT_SIZE,
    DEFAULT_HEADER_FILL,
    DEFAULT_HEADER_FONT,
    HEADER_FILL_COLORS,
    TEXT_LEFT_HEADERS,
    excel_value_by_header,
    normalize_header,
    number_format_for_header,
    width_for_header,
)


class CustomerCostReportExporter:
    """Excel export for CustomerCostsReportsPage."""

    def __init__(self) -> None:
        self._xl_center = -4108
        self._xl_left = -4131
        self._xl_right = -4152
        self._xl_vcenter = -4160

    @staticmethod
    def _rgb(r: int, g: int, b: int) -> int:
        return r + g * 256 + b * 65536

    @staticmethod
    def _excel_column_letter(col_num: int) -> str:
        result = ""
        while col_num > 0:
            col_num, remainder = divmod(col_num - 1, 26)
            result = chr(65 + remainder) + result
        return result

    @staticmethod
    def _excel_value(value: object) -> Any:
        if value is None:
            return ""
        if isinstance(value, Decimal):
            return float(value)
        if isinstance(value, (datetime, date)):
            return value
        if isinstance(value, bool):
            return "Да" if value else "Нет"
        return value

    @staticmethod
    def _safe_filename(value: str) -> str:
        s = (value or "").strip()
        for ch in ["\\", "/", ":", "*", "?", '"', "<", ">", "|"]:
            s = s.replace(ch, "_")
        return s or "CustomerCostReport"

    def _create_excel_app(self):
        pythoncom.CoInitialize()
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        return excel

    def _set_number_format_safe(self, target, format_en: str, format_local: str | None = None) -> None:
        try:
            target.NumberFormat = format_en
        except Exception:
            if format_local:
                target.NumberFormatLocal = format_local
            else:
                raise

    def _apply_header_common(self, ws, headers_count: int) -> None:
        ws.Cells.Font.Name = DEFAULT_FONT_NAME
        ws.Cells.Font.Size = DEFAULT_FONT_SIZE
        last_col = self._excel_column_letter(headers_count)
        header_range = ws.Range(f"A1:{last_col}1")
        header_range.Font.Name = DEFAULT_FONT_NAME
        header_range.Font.Size = DEFAULT_FONT_SIZE
        header_range.Font.Bold = True
        header_range.WrapText = True
        header_range.HorizontalAlignment = self._xl_center
        header_range.VerticalAlignment = self._xl_vcenter
        ws.Rows(1).RowHeight = 45

    def _write_table(self, ws, headers: Sequence[str], rows: Sequence[Sequence[object]]) -> None:
        for col_index, header in enumerate(headers, start=1):
            ws.Cells(1, col_index).Value = header
        for row_index, row in enumerate(rows, start=2):
            for col_index, _ in enumerate(headers, start=1):
                value = row[col_index - 1] if col_index - 1 < len(row) else ""
                ws.Cells(row_index, col_index).Value = excel_value_by_header(headers[col_index - 1], value)

    @staticmethod
    def _header_map(headers: Sequence[str]) -> dict[str, int]:
        return {str(header): idx + 1 for idx, header in enumerate(headers)}

    def _col(self, header_map: dict[str, int], header: str) -> str | None:
        idx = header_map.get(header)
        if not idx:
            return None
        return self._excel_column_letter(idx)

    def _set_width_by_header(self, ws, header_map: dict[str, int], header: str, width: float) -> None:
        letter = self._col(header_map, header)
        if letter:
            ws.Columns(f"{letter}:{letter}").ColumnWidth = width

    def _format_columns_by_headers(self, ws, header_map: dict[str, int], headers: Sequence[str], format_local: str) -> None:
        for header in headers:
            letter = self._col(header_map, header)
            if letter:
                try:
                    ws.Columns(f"{letter}:{letter}").NumberFormatLocal = format_local
                except Exception:
                    self._set_number_format_safe(ws.Columns(f"{letter}:{letter}"), "General", format_local)

    def _format_report(self, ws, headers: Sequence[str], rows_count: int) -> None:
        header_map = self._header_map(headers)
        last_col = self._excel_column_letter(len(headers))
        header_range = ws.Range(f"A1:{last_col}1")
        header_range.Interior.Color = DEFAULT_HEADER_FILL
        header_range.Font.Color = DEFAULT_HEADER_FONT

        for header in headers:
            base = normalize_header(header)
            letter = self._col(header_map, header)
            if not letter:
                continue

            if base in HEADER_FILL_COLORS:
                ws.Range(f"{letter}1").Interior.Color = HEADER_FILL_COLORS[base]

            fmt = number_format_for_header(header)
            if fmt:
                try:
                    ws.Columns(f"{letter}:{letter}").NumberFormatLocal = fmt
                except Exception:
                    self._set_number_format_safe(ws.Columns(f"{letter}:{letter}"), "General", fmt)

            ws.Columns(f"{letter}:{letter}").ColumnWidth = width_for_header(header, 10.0)
            if base in TEXT_LEFT_HEADERS:
                ws.Columns(f"{letter}:{letter}").HorizontalAlignment = self._xl_left
            else:
                ws.Columns(f"{letter}:{letter}").HorizontalAlignment = self._xl_center

        if rows_count > 0:
            data_range = ws.Range(f"A2:{last_col}{rows_count + 1}")
            data_range.VerticalAlignment = self._xl_vcenter

        ws.Range(f"A1:{last_col}1").AutoFilter(1)
        ws.Application.ActiveWindow.SplitRow = 1
        ws.Application.ActiveWindow.SplitColumn = 5
        ws.Application.ActiveWindow.FreezePanes = True
        ws.Application.ActiveWindow.Zoom = 85

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
            ws.Name = "Sheet1"
            self._write_table(ws, headers, rows)
            self._apply_header_common(ws, len(headers))
            self._format_report(ws, headers, len(rows))
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
