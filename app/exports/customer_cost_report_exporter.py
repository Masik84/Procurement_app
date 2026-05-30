from __future__ import annotations

from pathlib import Path
from typing import Sequence

import pythoncom
import win32com.client as win32

from app.utils.excel_export_format import (
    COST_HEADER_COLOR,
    SUPPLIER_HEADER_COLOR,
    XL_VCENTER,
    apply_base_table_style,
    apply_column_formats,
    color_headers,
    excel_column_letter,
    set_widths_by_headers,
    write_table,
)


class CustomerCostReportExporter:
    """Excel export for CustomerCostsReportsPage."""

    def _create_excel_app(self):
        pythoncom.CoInitialize()
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        return excel

    @staticmethod
    def _safe_filename(value: str) -> str:
        s = (value or "").strip()
        for ch in ["\\", "/", ":", "*", "?", '"', "<", ">", "|"]:
            s = s.replace(ch, "_")
        return s or "CustomerCostReport"

    def _format_report(self, ws, headers: Sequence[str], rows_count: int) -> None:
        last_col = excel_column_letter(len(headers))

        color_headers(
            ws,
            headers,
            {
                "Supplier": SUPPLIER_HEADER_COLOR,
                "Supplier Price, L": SUPPLIER_HEADER_COLOR,
                "Currency": SUPPLIER_HEADER_COLOR,
                "FX rate": SUPPLIER_HEADER_COLOR,
                "Cost Novo withVAT": COST_HEADER_COLOR,
                "Full Cost Msk": COST_HEADER_COLOR,
            },
        )
        apply_column_formats(ws, headers)
        set_widths_by_headers(
            ws,
            headers,
            {
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
                "Supplier Price, L": 12.0,
                "Currency": 8.14,
                "FX rate": 8.0,
                "Cost Novo withVAT": 12.0,
                "Full Cost Msk": 12.0,
                "Comments": 30.0,
            },
        )

        if rows_count > 0:
            data_range = ws.Range(f"A2:{last_col}{rows_count + 1}")
            data_range.VerticalAlignment = XL_VCENTER
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
            write_table(ws, headers, rows)
            apply_base_table_style(ws, len(headers))
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
