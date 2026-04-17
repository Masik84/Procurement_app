from __future__ import annotations

from pathlib import Path

import pythoncom
import win32com.client as win32


class ProductSearchExcelExporter:
    def __init__(self):
        self._xl_center = -4108
        self._xl_vcenter = -4160

    def _excel_column_letter(self, col_num: int) -> str:
        result = ""
        while col_num > 0:
            col_num, remainder = divmod(col_num - 1, 26)
            result = chr(65 + remainder) + result
        return result

    def _safe_value(self, value):
        if value is None:
            return ""
        return value

    def _create_excel_app(self):
        pythoncom.CoInitialize()
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        return excel

    def _apply_base_style(self, ws, headers_count: int, apply_filter: bool = True):
        ws.Cells.Font.Name = "Aptos Narrow"
        ws.Cells.Font.Size = 11

        if headers_count <= 0:
            return

        last_col = self._excel_column_letter(headers_count)
        header_range = ws.Range(f"A1:{last_col}1")

        header_range.Font.Name = "Aptos Narrow"
        header_range.Font.Size = 11
        header_range.Font.Bold = True
        header_range.Interior.Color = 0xCDCDCD
        header_range.WrapText = True
        header_range.HorizontalAlignment = self._xl_center
        header_range.VerticalAlignment = self._xl_vcenter

        ws.Rows(1).EntireRow.AutoFit()

        if apply_filter:
            try:
                ws.Range(f"A1:{last_col}1").AutoFilter(1)
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

            headers = ["Article", "Product name"]

            for col_index, header in enumerate(headers, start=1):
                ws.Cells(1, col_index).Value = header

            self._apply_base_style(ws, len(headers), apply_filter=False)

            ws.Columns("A:A").NumberFormat = "@"
            ws.Columns("A:B").ColumnWidth = 28

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

    def export_result(self, file_path: str, rows: list[dict]):
        excel = None
        wb = None

        try:
            excel = self._create_excel_app()
            wb = excel.Workbooks.Add()
            ws = wb.Worksheets(1)
            ws.Name = "Sheet1"

            headers = [
                "Article",
                "Supplier Product name",
                "Product name",
                "Brand",
                "Pack",
                "Excise duty",
            ]

            for col_index, header in enumerate(headers, start=1):
                ws.Cells(1, col_index).Value = header

            row_num = 2
            for row in rows:
                ws.Cells(row_num, 1).Value = self._safe_value(row.get("source_article"))
                ws.Cells(row_num, 2).Value = self._safe_value(row.get("source_product_name"))
                ws.Cells(row_num, 3).Value = self._safe_value(row.get("product_name"))
                ws.Cells(row_num, 4).Value = self._safe_value(row.get("brand"))
                ws.Cells(row_num, 5).Value = self._safe_value(row.get("pack"))
                ws.Cells(row_num, 6).Value = self._safe_value(row.get("is_excise"))
                row_num += 1

            self._apply_base_style(ws, len(headers), apply_filter=True)

            ws.Columns("A:A").NumberFormat = "@"
            ws.Columns("A:A").ColumnWidth = 18
            ws.Columns("B:D").ColumnWidth = 30
            ws.Columns("E:E").ColumnWidth = 12
            ws.Columns("F:F").ColumnWidth = 14

            wb.SaveAs(str(Path(file_path).resolve()))

        except PermissionError:
            raise
        except Exception as e:
            raise Exception(f"Ошибка при создании итогового Excel: {e}")
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
