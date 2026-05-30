
from __future__ import annotations

from pathlib import Path

import pythoncom
import win32com.client as win32

from app.utils.excel_export_format import write_and_format_table


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
                "Supplier Product name",
                "Our Product Name",
                "Price date",
                "Price",
                "Currency",
            ]
            write_and_format_table(ws, headers, [], widths={
                "Supplier name": 24, "Article": 18, "Supplier Product name": 32,
                "Our Product Name": 32, "Price date": 12, "Price": 12, "Currency": 12,
            }, apply_filter=False)

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
            table_rows = [
                [
                    row.get("product_name", ""),
                    row.get("supplier_name", ""),
                    row.get("price_date"),
                    row.get("price"),
                    row.get("currency", ""),
                ]
                for row in rows
            ]
            write_and_format_table(ws, headers, table_rows, widths={
                "Product name": 31.8,
                "Supplier name": 24,
                "Price date": 13,
                "Price": 13,
                "Currency": 13,
            }, apply_filter=True)
            ws.Rows(1).RowHeight = 45 if report_type == "current" else 60

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
