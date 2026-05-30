from __future__ import annotations

from pathlib import Path

import pythoncom
import win32com.client as win32

from app.utils.excel_export_format import write_and_format_table


class ProductSearchExporter:
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

            write_and_format_table(ws, headers, [], widths={"Article": 28, "Product name": 28}, apply_filter=False)

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

            table_rows = [[
                row.get("source_article"), row.get("source_product_name"), row.get("product_name"),
                row.get("brand"), row.get("pack"), row.get("is_excise"),
            ] for row in rows]
            write_and_format_table(ws, headers, table_rows, widths={
                "Article": 18, "Supplier Product name": 30, "Product name": 30,
                "Brand": 30, "Pack": 12, "Excise duty": 14,
            }, apply_filter=True)

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

# Backward-compatible alias.
ProductSearchExcelExporter = ProductSearchExporter
