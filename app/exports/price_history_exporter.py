
from __future__ import annotations

from pathlib import Path

import pythoncom
import win32com.client as win32

from app.utils.excel_fast_writer import write_excel_table


class PriceHistoryExporter:
    def __init__(self):
        self._xl_center = -4108
        self._xl_vcenter = -4160

    def _create_excel_app(self):
        pythoncom.CoInitialize()
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        return excel

    def _apply_base_font(self, ws):
        ws.Cells.Font.Name = "Aptos Narrow"
        ws.Cells.Font.Size = 11

    def _apply_header_range(self, ws, start_col: str, end_col: str, fill_color: int | None = None, white_font: bool = False):
        rng = ws.Range(f"{start_col}1:{end_col}1")
        rng.Font.Name = "Aptos Narrow"
        rng.Font.Size = 11
        rng.Font.Bold = True
        rng.WrapText = True
        rng.HorizontalAlignment = self._xl_center
        rng.VerticalAlignment = self._xl_vcenter
        if fill_color is not None:
            rng.Interior.Color = fill_color
        if white_font:
            rng.Font.Color = 0xFFFFFF

    @staticmethod
    def _excel_invariant_number_format(format_code: str | None) -> str | None:
        if not format_code:
            return format_code
        return (
            str(format_code)
            .replace("ДД", "dd")
            .replace("ММ", "mm")
            .replace("ГГ", "yy")
            .replace(",0000", ".0000")
            .replace(",00", ".00")
            .replace(",0", ".0")
        )

    def _set_number_format_safe(self, target, format_local: str, format_en: str | None = None) -> None:
        candidates = [
            ("NumberFormatLocal", format_local),
            ("NumberFormat", format_en or self._excel_invariant_number_format(format_local) or format_local),
            ("NumberFormat", format_local),
            ("NumberFormat", "General"),
        ]
        for attr, fmt in candidates:
            try:
                setattr(target, attr, fmt)
                return
            except Exception:
                pass

    def export_template(self, file_path: str):
        excel = None
        wb = None
        try:
            excel = self._create_excel_app()
            wb = excel.Workbooks.Add()
            ws = wb.Worksheets(1)
            ws.Name = "Sheet1"

            headers = [
                "Supplier name",
                "Article",
                "Supplier Product Name",
                "Our Product Name",
                "Price date",
                "Price",
                "Currency",
            ]
            for col_index, header in enumerate(headers, start=1):
                ws.Cells(1, col_index).Value = header

            self._apply_base_font(ws)
            self._apply_header_range(ws, "A", "G", 0xCDCDCD, white_font=False)

            ws.Rows(1).RowHeight = 45
            ws.Columns("A:A").ColumnWidth = 24
            ws.Columns("B:B").ColumnWidth = 18
            ws.Columns("C:D").ColumnWidth = 32
            ws.Columns("E:E").ColumnWidth = 12
            ws.Columns("F:F").ColumnWidth = 12
            ws.Columns("G:G").ColumnWidth = 12

            wb.SaveAs(str(Path(file_path).resolve()))
        except PermissionError:
            raise
        except Exception as e:
            raise Exception(f"Ошибка при создании шаблона Excel: {e}")
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

    def export_rows(self, rows, file_path: str, report_type: str = "history"):
        excel = None
        wb = None
        try:
            excel = self._create_excel_app()
            wb = excel.Workbooks.Add()
            ws = wb.Worksheets(1)
            ws.Name = "Sheet1"

            headers = ["Product name", "Supplier name", "Price date", "Price", "Currency"]

            def value_for_header(row, header, _col_index):
                if header == "Product name":
                    return row.get("product_name", "")
                if header == "Supplier name":
                    return row.get("supplier_name", "")
                if header == "Price date":
                    return row.get("price_date") or ""
                if header == "Price":
                    price = row.get("price")
                    if price is None:
                        return ""
                    try:
                        return float(price)
                    except Exception:
                        return ""
                if header == "Currency":
                    return row.get("currency", "")
                return ""

            write_excel_table(ws, headers, rows, value_getter=value_for_header)

            self._apply_base_font(ws)
            self._apply_header_range(ws, "A", "E", 0xCDCDCD, white_font=False)

            if report_type == "current":
                ws.Rows(1).RowHeight = 45
            else:
                ws.Rows(1).RowHeight = 60

            ws.Range("A1:E1").AutoFilter(1)

            ws.Columns("A:A").ColumnWidth = 31.8
            ws.Columns("B:B").ColumnWidth = 24
            ws.Columns("C:C").ColumnWidth = 13
            ws.Columns("D:D").ColumnWidth = 13
            ws.Columns("E:E").ColumnWidth = 13

            if rows:
                last_row = len(rows) + 1
                self._set_number_format_safe(ws.Range(f"C2:C{last_row}"), "ДД.ММ.ГГ;@", "dd.mm.yy;@")
                self._set_number_format_safe(ws.Range(f"D2:D{last_row}"), "0,00", "0.00")

            wb.SaveAs(str(Path(file_path).resolve()))
        except PermissionError:
            raise
        except Exception as e:
            raise Exception(f"Ошибка при сохранении Excel: {e}")
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
    
    
    
    
# Backward-compatible alias.
PriceHistoryExcelExporter = PriceHistoryExporter
