from __future__ import annotations

from datetime import datetime
from decimal import Decimal
from pathlib import Path
from typing import Any, Sequence

import pythoncom
import win32com.client as win32


class PriceReportExporter:
    """Excel export for PriceReportsPage.

    The GUI page only prepares headers/rows and asks the user for a file path.
    All Excel-specific formatting lives here, by analogy with SupplierPriceExporter.
    """

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
        if isinstance(value, (int, float)) and value == 0:
            return ""
        return value

    @staticmethod
    def _excel_value_keep_zero(value: object) -> Any:
        if value is None:
            return ""
        if isinstance(value, Decimal):
            return float(value)
        return value

    @staticmethod
    def _base_order_header(header: object) -> str:
        text = str(header or "")
        if text.startswith("к Быстрому заказу, л"):
            return "к Быстрому заказу, л"
        if text.startswith("к Заказу, л"):
            return "к Заказу, л"
        return text

    @classmethod
    def _is_order_plan_export_header(cls, header: object) -> bool:
        return cls._base_order_header(header) in {"к Быстрому заказу, л", "к Заказу, л"}

    @staticmethod
    def _safe_filename(value: str) -> str:
        s = (value or "").strip()
        for ch in ["\\", "/", ":", "*", "?", '"', "<", ">", "|"]:
            s = s.replace(ch, "_")
        return s or "Supplier"

    def _create_excel_app(self):
        pythoncom.CoInitialize()
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        return excel

    def _apply_font_to_all_cells(self, ws) -> None:
        ws.Cells.Font.Name = "Aptos Narrow"
        ws.Cells.Font.Size = 11

    def _apply_header_common(self, ws, headers_count: int) -> None:
        self._apply_font_to_all_cells(ws)
        last_col = self._excel_column_letter(headers_count)
        header_range = ws.Range(f"A1:{last_col}1")
        header_range.Font.Name = "Aptos Narrow"
        header_range.Font.Size = 11
        header_range.Font.Bold = True
        header_range.WrapText = True
        header_range.HorizontalAlignment = self._xl_center
        header_range.VerticalAlignment = self._xl_vcenter
        ws.Rows(1).EntireRow.AutoFit()

    def _set_number_format_safe(self, target, format_en: str, format_local: str | None = None) -> None:
        try:
            target.NumberFormat = format_en
        except Exception:
            if format_local:
                target.NumberFormatLocal = format_local
            else:
                raise

    def _write_table(self, ws, headers: Sequence[str], rows: Sequence[Sequence[object]]) -> None:
        for col_index, header in enumerate(headers, start=1):
            ws.Cells(1, col_index).Value = header

        for row_index, row in enumerate(rows, start=2):
            for col_index, header in enumerate(headers, start=1):
                value = row[col_index - 1] if col_index - 1 < len(row) else ""
                if self._is_order_plan_export_header(header):
                    # Для колонок заказа 0 — это значение, а не пустая ячейка.
                    # Формат Excel сам покажет ноль как "-".
                    ws.Cells(row_index, col_index).Value = self._excel_value_keep_zero(value)
                else:
                    ws.Cells(row_index, col_index).Value = self._excel_value(value)

    def _header_map(self, headers: Sequence[str]) -> dict[str, int]:
        return {str(header): idx + 1 for idx, header in enumerate(headers)}

    def _col(self, header_map: dict[str, int], header: str) -> str | None:
        idx = header_map.get(header)
        if not idx:
            return None
        return self._excel_column_letter(idx)

    def _range_by_headers(self, ws, header_map: dict[str, int], first_header: str, last_header: str):
        first = header_map.get(first_header)
        last = header_map.get(last_header)
        if not first or not last:
            return None
        return ws.Range(f"{self._excel_column_letter(first)}1:{self._excel_column_letter(last)}1")

    def _color_header_range(self, ws, header_map: dict[str, int], first_header: str, last_header: str, color: int, font_color: int | None = None) -> None:
        rng = self._range_by_headers(ws, header_map, first_header, last_header)
        if rng is None:
            return
        rng.Interior.Color = color
        if font_color is not None:
            rng.Font.Color = font_color

    def _format_columns_by_headers(self, ws, header_map: dict[str, int], headers: Sequence[str], format_local: str) -> None:
        for header in headers:
            letter = self._col(header_map, header)
            if letter:
                self._set_number_format_safe(ws.Columns(f"{letter}:{letter}"), "General", format_local)

    def _set_width_by_header(self, ws, header_map: dict[str, int], header: str, width: float) -> None:
        letter = self._col(header_map, header)
        if letter:
            ws.Columns(f"{letter}:{letter}").ColumnWidth = width

    def _color_order_plan_headers(self, ws, header_map: dict[str, int]) -> None:
        order_headers = [h for h in header_map if h == "Ср.Продажи мес" or h.startswith("к Быстрому заказу, л") or h.startswith("к Заказу, л")]
        for header in order_headers:
            letter = self._col(header_map, header)
            if letter:
                ws.Range(f"{letter}1").Interior.Color = self._rgb(160, 43, 147)
                ws.Range(f"{letter}1").Font.Color = self._rgb(255, 255, 255)

    def _format_stock_and_order_plan_columns(self, ws, header_map: dict[str, int]) -> None:
        headers = [
            h for h in header_map
            if h in {"Stock", "Transit", "Purchase Order", "Order IS", "Stock IS", "Reserve cust", "Reserve E-Comm", "Damaged", "Ср.Продажи мес"}
            or h.startswith("к Быстрому заказу, л") or h.startswith("к Заказу, л")
        ]
        self._format_columns_by_headers(ws, header_map, headers, '# ##0;[Red]-# ##0;"-"')

    def _color_repeating_supplier_headers(self, ws, header_map: dict[str, int], start_idx: int) -> None:
        idx = start_idx
        while f"Cost Novo with VAT_{idx}" in header_map:
            self._color_header_range(ws, header_map, f"Cost Novo with VAT_{idx}", f"Full Cost Msk_{idx}", self._rgb(0, 176, 240))
            self._color_header_range(ws, header_map, f"Supplier_{idx}", f"Currency_{idx}", self._rgb(146, 208, 80))
            idx += 1

    def _format_fx_headers(self, ws, header_map: dict[str, int]) -> None:
        for header in header_map:
            if header == "FX rate" or header.startswith("FX rate_") or header in {"FX rate Best1", "FX rate Best2"}:
                letter = self._col(header_map, header)
                if letter:
                    self._set_number_format_safe(ws.Columns(f"{letter}:{letter}"), "General", "# ##0")
                    ws.Columns(f"{letter}:{letter}").ColumnWidth = 7.29

    def _set_common_data_alignment(self, ws, headers_count: int, rows_count: int) -> None:
        if rows_count <= 0:
            return
        last_col = self._excel_column_letter(headers_count)
        data_range = ws.Range(f"A2:{last_col}{rows_count + 1}")
        data_range.VerticalAlignment = self._xl_vcenter

    def _apply_supplier_price_like_widths_until_damaged(self, ws, header_map: dict[str, int]) -> None:
        # Ширины взяты из SupplierPriceExporter.export_calculated для блока до Damaged.
        for header in ("Supplier Product Name", "Our Product Name", "Product Name"):
            self._set_width_by_header(ws, header_map, header, 31.14)
        self._set_width_by_header(ws, header_map, "Qty, pcs", 10.00)
        self._set_width_by_header(ws, header_map, "Volume, L", 10.00)
        self._set_width_by_header(ws, header_map, "Best Suppl", 16.14)
        self._set_width_by_header(ws, header_map, "Best Suppl 2", 16.14)
        self._set_width_by_header(ws, header_map, "last update", 11.00)
        self._set_width_by_header(ws, header_map, "last update (prev)", 11.00)
        self._set_width_by_header(ws, header_map, "last update Best1", 9.43)
        self._set_width_by_header(ws, header_map, "last update Best2", 9.43)
        for header in header_map:
            if header in {"Stock", "Transit", "Purchase Order", "Order IS", "Stock IS", "Reserve cust", "Reserve E-Comm", "Damaged", "Ср.Продажи мес"} or header.startswith("к Быстрому заказу, л") or header.startswith("к Заказу, л"):
                self._set_width_by_header(ws, header_map, header, 8.14)

    def _apply_repeating_supplier_widths(self, ws, header_map: dict[str, int], start_idx: int) -> None:
        idx = start_idx
        while f"Cost Novo with VAT_{idx}" in header_map:
            self._set_width_by_header(ws, header_map, f"Cost Novo with VAT_{idx}", 9.00)
            self._set_width_by_header(ws, header_map, f"Full Cost Msk_{idx}", 9.00)
            self._set_width_by_header(ws, header_map, f"Supplier_{idx}", 16.14)
            self._set_width_by_header(ws, header_map, f"last update_{idx}", 9.43)
            self._set_width_by_header(ws, header_map, f"FX rate_{idx}", 7.29)
            self._set_width_by_header(ws, header_map, f"Currency_{idx}", 8.14)
            idx += 1

    def _format_product_report(self, ws, headers: Sequence[str], rows_count: int) -> None:
        header_map = self._header_map(headers)
        ws.Rows(1).RowHeight = 45

        self._color_header_range(ws, header_map, "Brand", "Pack", self._rgb(205, 205, 205))
        self._color_header_range(ws, header_map, "Дистр цена", "curr Landed cost", self._rgb(192, 0, 0), self._rgb(255, 255, 255))
        self._color_header_range(ws, header_map, "Stock", "Purchase Order", self._rgb(33, 92, 152), self._rgb(255, 255, 255))
        self._color_header_range(ws, header_map, "Order IS", "Stock IS", self._rgb(192, 0, 0), self._rgb(255, 255, 255))
        self._color_header_range(ws, header_map, "Reserve cust", "Damaged", self._rgb(33, 92, 152), self._rgb(255, 255, 255))
        self._color_order_plan_headers(ws, header_map)

        # self._format_columns_by_headers(ws, header_map, [], "# ##0,00_ ;[Red]-# ##0,00_ ;'-'")
        self._format_columns_by_headers(ws, header_map, ["Дистр цена", "Промо цена", "curr LPC", "curr Landed cost"], "# ##0 ₽")
        self._format_stock_and_order_plan_columns(ws, header_map)

        idx = 1
        while f"Cost Novo with VAT_{idx}" in header_map:
            self._format_columns_by_headers(ws, header_map, [f"Cost Novo with VAT_{idx}", f"Full Cost Msk_{idx}"], "# ##0 ₽")
            self._format_columns_by_headers(ws, header_map, [f"last update_{idx}"], "ДД.ММ.ГГ;@")
            self._format_columns_by_headers(ws, header_map, [f"FX rate_{idx}"], "# ##0")
            idx += 1

        self._set_width_by_header(ws, header_map, "Brand", 13.00)
        self._set_width_by_header(ws, header_map, "Product Name", 31.14)
        self._set_width_by_header(ws, header_map, "Pack", 8.43)
        self._set_width_by_header(ws, header_map, "Дистр цена", 8.43)
        self._set_width_by_header(ws, header_map, "Промо цена", 8.43)
        self._set_width_by_header(ws, header_map, "curr LPC", 8.43)
        self._set_width_by_header(ws, header_map, "curr Landed cost", 8.43)
        self._apply_supplier_price_like_widths_until_damaged(ws, header_map)
        self._apply_repeating_supplier_widths(ws, header_map, 1)
        self._color_repeating_supplier_headers(ws, header_map, 1)
        self._color_repeating_supplier_headers(ws, header_map, 3)
        self._format_fx_headers(ws, header_map)

        ws.Range(f"A1:{self._excel_column_letter(len(headers))}1").AutoFilter(1)
        self._freeze(ws, split_column=7)

    def _format_supplier_report(self, ws, headers: Sequence[str], rows_count: int) -> None:
        header_map = self._header_map(headers)
        ws.Rows(1).RowHeight = 60

        first_gray_end = "Full Cost Msk (prev)" if "Full Cost Msk (prev)" in header_map else "Full Cost Msk"
        self._color_header_range(ws, header_map, "Our Product Name", first_gray_end, self._rgb(205, 205, 205))
        self._color_header_range(ws, header_map, "Дистр цена", "curr Landed cost", self._rgb(192, 0, 0), self._rgb(255, 255, 255))
        self._color_header_range(ws, header_map, "Best Suppl", "Currency Best1", self._rgb(0, 176, 240))
        self._color_header_range(ws, header_map, "Best Suppl 2", "Currency Best2", self._rgb(146, 208, 80))
        self._color_header_range(ws, header_map, "Stock", "Purchase Order", self._rgb(33, 92, 152), self._rgb(255, 255, 255))
        self._color_header_range(ws, header_map, "Order IS", "Stock IS", self._rgb(192, 0, 0), self._rgb(255, 255, 255))
        self._color_header_range(ws, header_map, "Reserve cust", "Damaged", self._rgb(33, 92, 152), self._rgb(255, 255, 255))
        self._color_order_plan_headers(ws, header_map)

        self._format_columns_by_headers(ws, header_map, ["last update", "last update (prev)", "last update Best1", "last update Best2"], "ДД.ММ.ГГ;@")
        self._format_columns_by_headers(ws, header_map, ["FX rate", "FX rate Best1", "FX rate Best2"], "# ##0")
        self._format_columns_by_headers(ws, header_map, ["Price, L", "Price, pack", "Price, L (prev)"], "# ##0,00_ ;[Red]-# ##0,00_ ;'-'")
        self._format_columns_by_headers(ws, header_map, ["Дистр цена", "Промо цена", "Cost Novo with VAT", "Full Cost Msk", 
                                                                                            "Cost Novo with VAT (prev)", "Full Cost Msk (prev)", "curr LPC", 
                                                                                            "curr Landed cost", "Best full Price, L", "Best full Price, L 2"], "# ##0 ₽")
        self._format_stock_and_order_plan_columns(ws, header_map)

        idx = 3
        while f"Cost Novo with VAT_{idx}" in header_map:
            self._format_columns_by_headers(ws, header_map, [f"Cost Novo with VAT_{idx}", f"Full Cost Msk_{idx}"], "# ##0 ₽")
            self._format_columns_by_headers(ws, header_map, [f"last update_{idx}"], "ДД.ММ.ГГ;@")
            self._format_columns_by_headers(ws, header_map, [f"FX rate_{idx}"], "# ##0")
            idx += 1

        self._set_width_by_header(ws, header_map, "Our Product Name", 31.14)
        self._set_width_by_header(ws, header_map, "Pack", 8.43)
        for header in ("Price, L", "Price, pack"):
            self._set_width_by_header(ws, header_map, header, 9.14 if header == "Price, L" else 13.00)
        for header in ("Currency", "Cost Novo with VAT", "Full Cost Msk", "Price, L (prev)", "Cost Novo with VAT (prev)", "Full Cost Msk (prev)", "Currency Best1", "Currency Best2"):
            self._set_width_by_header(ws, header_map, header, 8.71)
        self._set_width_by_header(ws, header_map, "last update (prev)", 11.00)
        for header in ("FX rate", "FX rate Best1", "FX rate Best2"):
            self._set_width_by_header(ws, header_map, header, 7.29)
        self._set_width_by_header(ws, header_map, "Дистр цена", 8.43)
        self._set_width_by_header(ws, header_map, "Промо цена", 8.43)
        self._set_width_by_header(ws, header_map, "curr LPC", 8.43)
        self._set_width_by_header(ws, header_map, "curr Landed cost", 8.43)
        self._set_width_by_header(ws, header_map, "Best full Price, L", 8.43)
        self._set_width_by_header(ws, header_map, "Best full Price, L 2", 8.43)
        self._apply_supplier_price_like_widths_until_damaged(ws, header_map)
        self._apply_repeating_supplier_widths(ws, header_map, 3)
        self._format_fx_headers(ws, header_map)

        ws.Range(f"A1:{self._excel_column_letter(len(headers))}1").AutoFilter(1)
        freeze_col = max((header_map.get("Дистр цена") or 1) - 1, 1)
        self._freeze(ws, split_column=freeze_col)

    def _freeze(self, ws, split_column: int) -> None:
        try:
            ws.Activate()
            window = ws.Application.ActiveWindow
            window.FreezePanes = False
            window.SplitRow = 1
            window.SplitColumn = split_column
            window.ScrollRow = 1
            window.ScrollColumn = 1
            window.Zoom = 85
            window.FreezePanes = True
            ws.Range("A1").Select()
            window.ScrollRow = 1
            window.ScrollColumn = 1
        except Exception:
            pass


    @staticmethod
    def _headers_with_order_months(headers: Sequence[str], quick_order_months: int | None, safe_stock_months: int | None) -> list[str]:
        result: list[str] = []
        for header in headers:
            h = str(header)
            if h == "к Быстрому заказу, л" and quick_order_months is not None:
                result.append(f"к Быстрому заказу, л ({quick_order_months} м)")
            elif h == "к Заказу, л" and safe_stock_months is not None:
                result.append(f"к Заказу, л ({safe_stock_months} м)")
            else:
                result.append(h)
        return result

    def export_report(
        self,
        *,
        headers: Sequence[str],
        rows: Sequence[Sequence[object]],
        output_path: str | Path,
        report_mode: str,
        quick_order_months: int | None = None,
        safe_stock_months: int | None = None,
    ) -> Path:
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

            headers = self._headers_with_order_months(headers, quick_order_months, safe_stock_months)
            self._write_table(ws, headers, rows)
            self._apply_header_common(ws, len(headers))
            self._set_common_data_alignment(ws, len(headers), len(rows))

            if report_mode == "supplier":
                self._format_supplier_report(ws, headers, len(rows))
            else:
                self._format_product_report(ws, headers, len(rows))

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
