from __future__ import annotations

from pathlib import Path

import pythoncom
import win32com.client as win32
from app.utils.output_headers import standardize_output_header, apply_header_style_and_formats
from app.utils.excel_format_rules import FORMATS, set_number_format_safe


class ProductArticleExporter:
    def __init__(self):
        self._xl_center = -4108
        self._xl_vcenter = -4160
        self._xl_openxml_workbook = 51  # .xlsx

    def _excel_column_letter(self, col_num: int) -> str:
        result = ""
        while col_num > 0:
            col_num, remainder = divmod(col_num - 1, 26)
            result = chr(65 + remainder) + result
        return result

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

    def _normalize_save_path(self, file_path: str) -> str:
        path = Path(file_path).expanduser()
        if path.suffix.lower() != ".xlsx":
            path = path.with_suffix(".xlsx")

        path.parent.mkdir(parents=True, exist_ok=True)
        return str(path.resolve())

    def export_template(self, file_path: str):
        excel = None
        wb = None

        try:
            save_path = self._normalize_save_path(file_path)

            excel = self._create_excel_app()
            wb = excel.Workbooks.Add()
            ws = wb.Worksheets(1)
            ws.Name = "Sheet1"

            headers = ["Product name", "Article", "Product name (variant)"]
            for col_index, header in enumerate(headers, start=1):
                ws.Cells(1, col_index).Value = standardize_output_header(header)

            self._apply_base_style(ws, len(headers), apply_filter=False)

            ws.Columns("A:A").ColumnWidth = 34
            ws.Columns("B:B").ColumnWidth = 22
            ws.Columns("C:C").ColumnWidth = 34

            set_number_format_safe(ws.Columns("B:B"), FORMATS.TEXT)

            wb.SaveAs(Filename=save_path, FileFormat=self._xl_openxml_workbook)

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

            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

# Backward-compatible alias.
ProductArticleExcelExporter = ProductArticleExporter
