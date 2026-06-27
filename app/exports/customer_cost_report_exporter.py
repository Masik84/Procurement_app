from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import Any, Sequence

import pythoncom
import win32com.client as win32

from app.utils.excel_fast_writer import write_excel_table


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
            if value == 0:
                return ""
            return float(value)
        if isinstance(value, (datetime, date)):
            return value
        if isinstance(value, bool):
            return "Да" if value else "Нет"
        if isinstance(value, (int, float)) and value == 0:
            return ""
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
        ws.Cells.Font.Name = "Aptos Narrow"
        ws.Cells.Font.Size = 11
        last_col = self._excel_column_letter(headers_count)
        header_range = ws.Range(f"A1:{last_col}1")
        header_range.Font.Name = "Aptos Narrow"
        header_range.Font.Size = 11
        header_range.Font.Bold = True
        header_range.WrapText = True
        header_range.HorizontalAlignment = self._xl_center
        header_range.VerticalAlignment = self._xl_vcenter
        ws.Rows(1).RowHeight = 45

    def _write_table(self, ws, headers: Sequence[str], rows: Sequence[Sequence[object]]) -> None:
        write_excel_table(
            ws,
            headers,
            rows,
            value_getter=lambda row, _header, col_index: self._excel_value(
                row[col_index] if col_index < len(row) else "",
            ),
        )

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
                self._set_number_format_safe(ws.Columns(f"{letter}:{letter}"), "General", format_local)

    def _format_report(self, ws, headers: Sequence[str], rows_count: int) -> None:
        header_map = self._header_map(headers)
        last_col = self._excel_column_letter(len(headers))
        ws.Range(f"A1:{last_col}1").Interior.Color = self._rgb(205, 205, 205)
        ws.Range(f"A1:{last_col}1").Font.Color = self._rgb(0, 0, 0)

        for header in ("Supplier", "Supplier Price, L", "Currency", "FX rate"):
            letter = self._col(header_map, header)
            if letter:
                ws.Range(f"{letter}1").Interior.Color = self._rgb(146, 208, 80)

        for header in ("Cost Novo with VAT", "Full Cost Msk"):
            letter = self._col(header_map, header)
            if letter:
                ws.Range(f"{letter}1").Interior.Color = self._rgb(0, 176, 240)

        cost_headers = [
            "Supplier Price, L", "Cost Novo with VAT", "Full Cost Msk", "FX markup", "Transport",
            "Re-export", "Agent fee", "Bank fee", "Customs fee", "Additional customs",
            "Storage", "Move Novo", "Move Msk", "Marking",
        ]
        self._format_columns_by_headers(ws, header_map, cost_headers, '# ##0,00;[Red]-# ##0,00;"-"')
        self._format_columns_by_headers(ws, header_map, ["FX rate"], "# ##0")
        self._format_columns_by_headers(ws, header_map, ["Qty, pcs", "Volume, L"], '# ##0,00;[Red]-# ##0,00;"-"')
        self._format_columns_by_headers(ws, header_map, ["Дата", "Price date"], "ДД.ММ.ГГ;@")

        widths = {
            "Дата": 11.0,
            "Менеджер": 18.0,
            "Клиент": 22.0,
            "Customer Product Name": 31.14,
            "Our Product Name": 31.14,
            "Pack": 8.43,
            "Qty, pcs": 10.0,
            "Volume, L": 10.0,
            "Supplier": 16.14,
            "Supplier Article": 14.0,
            "Supplier Price, L": 10.0,
            "Currency": 8.14,
            "FX rate": 7.29,
            "Cost Novo with VAT": 10.0,
            "Full Cost Msk": 10.0,
            "Comments": 30.0,
        }
        for header, width in widths.items():
            self._set_width_by_header(ws, header_map, header, width)
        for header in headers:
            if header not in widths:
                self._set_width_by_header(ws, header_map, header, 10.0)

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
