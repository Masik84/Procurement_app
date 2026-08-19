from __future__ import annotations

from pathlib import Path

import pythoncom
import win32com.client as win32

from app.exports.excel_column_format import apply_standard_worksheet_format, excel_value_by_header
from app.utils.excel_fast_writer import write_excel_table
from app.utils.excel_format_rules import FORMATS, set_number_format_safe, save_workbook_xlsx


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

    def _normalize_save_path(self, file_path: str) -> Path:
        path = Path(file_path).expanduser()
        if path.suffix.lower() != ".xlsx":
            path = path.with_suffix(".xlsx")
        path.parent.mkdir(parents=True, exist_ok=True)
        target_path = path.resolve()

        if target_path.exists():
            try:
                target_path.unlink()
            except PermissionError:
                raise PermissionError(
                    f"Не удается перезаписать файл:\n{target_path}\n\n"
                    "Скорее всего, он открыт в Excel. Закрой файл и попробуй снова."
                )

        return target_path

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

            set_number_format_safe(ws.Columns("A:A"), FORMATS.TEXT)
            ws.Columns("A:B").ColumnWidth = 28

            save_workbook_xlsx(wb, file_path)

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
            target_path = self._normalize_save_path(file_path)

            excel = self._create_excel_app()
            wb = excel.Workbooks.Add()
            ws = wb.Worksheets(1)
            ws.Name = "Sheet1"

            headers = [
                "Article",
                "Supplier Product Name",
                "Product name",
                "Brand",
                "Pack",
                "Категория ABC",
                "Excise duty",
            ]

            def value_for_header(row, header, _col_index):
                values_by_header = {
                    "Article": row.get("source_article"),
                    "Supplier Product Name": row.get("source_product_name"),
                    "Product name": row.get("product_name"),
                    "Brand": row.get("brand"),
                    "Pack": row.get("pack"),
                    "Категория ABC": row.get("abc_category") or "-",
                    "Excise duty": row.get("is_excise"),
                }
                return excel_value_by_header(str(header), self._safe_value(values_by_header.get(header)))

            write_excel_table(ws, headers, rows, value_getter=value_for_header)

            apply_standard_worksheet_format(ws, headers, freeze_cell="C2", zoom=85)

            save_workbook_xlsx(wb, target_path)
            return target_path

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
