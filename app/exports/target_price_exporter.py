from __future__ import annotations

from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Any

import pythoncom
import win32com.client as win32
from sqlalchemy.orm import Session

from app.db.models import Product, TargetPriceCalculation, TempTargetPriceImport, TempTargetPriceOption


class TargetPriceExporter:
    def __init__(self, session: Session):
        self.session = session
        self._xl_center = -4108
        self._xl_vcenter = -4160

    @staticmethod
    def _to_decimal(value: object) -> Decimal:
        if value is None:
            return Decimal("0")
        if isinstance(value, Decimal):
            return value
        return Decimal(str(value))

    @staticmethod
    def _excel_value_or_blank(value: object):
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
    def _round_fx_rate(value: object):
        if value is None or value == "":
            return ""
        try:
            return int(TargetPriceExporter._to_decimal(value).to_integral_value(rounding=ROUND_HALF_UP))
        except Exception:
            return ""

    @staticmethod
    def _excel_column_letter(col_num: int) -> str:
        result = ""
        while col_num > 0:
            col_num, remainder = divmod(col_num - 1, 26)
            result = chr(65 + remainder) + result
        return result

    @staticmethod
    def _rgb(r: int, g: int, b: int) -> int:
        return r + g * 256 + b * 65536

    def _create_excel_app(self):
        pythoncom.CoInitialize()
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        return excel

    def _apply_font_to_all_cells(self, ws):
        ws.Cells.Font.Name = "Aptos Narrow"
        ws.Cells.Font.Size = 11

    def _apply_header_common(self, ws, headers_count: int):
        self._apply_font_to_all_cells(ws)
        last_col = self._excel_column_letter(headers_count)
        header_range = ws.Range(f"A1:{last_col}1")
        header_range.Font.Name = "Aptos Narrow"
        header_range.Font.Size = 11
        header_range.Font.Bold = True
        header_range.WrapText = True
        header_range.HorizontalAlignment = self._xl_center
        header_range.VerticalAlignment = self._vcenter if hasattr(self, "_vcenter") else self._xl_vcenter
        ws.Rows(1).EntireRow.AutoFit()

    def _set_number_format_safe(self, target, format_code: str):
        try:
            target.NumberFormat = format_code
        except Exception:
            try:
                target.NumberFormatLocal = format_code
            except Exception:
                pass

    def _save_workbook(self, file_path: str | Path, sheet_name: str, headers: list[str], rows: list[dict], formatter):
        file_path = Path(file_path)
        if file_path.suffix.lower() != ".xlsx":
            file_path = file_path.with_suffix(".xlsx")
        file_path.parent.mkdir(parents=True, exist_ok=True)

        excel = None
        wb = None
        try:
            target_path = file_path.resolve()
            if target_path.exists():
                try:
                    target_path.unlink()
                except PermissionError:
                    raise PermissionError(
                        f"Не удается перезаписать файл:\n{target_path}\n\n"
                        f"Скорее всего, он открыт в Excel. Закрой файл и попробуй снова."
                    )
            excel = self._create_excel_app()
            wb = excel.Workbooks.Add()
            ws = wb.Worksheets(1)
            ws.Name = sheet_name
            for col_index, header in enumerate(headers, start=1):
                ws.Cells(1, col_index).Value = header
            for row_num, row in enumerate(rows, start=2):
                for col_index, header in enumerate(headers, start=1):
                    ws.Cells(row_num, col_index).Value = self._excel_value_or_blank(row.get(header))
            self._apply_header_common(ws, len(headers))
            formatter(ws, headers)
            last_col = self._excel_column_letter(len(headers))
            ws.Range(f"A1:{last_col}1").AutoFilter(1)
            try:
                excel.ActiveWindow.Zoom = 85
                ws.Activate()
                ws.Range("D2").Select()
                excel.ActiveWindow.FreezePanes = True
            except Exception:
                pass
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

    def export_template(self, file_path: str | Path) -> Path:
        headers = ["Material number", "Material"]
        def fmt(ws, headers):
            ws.Range("A1:B1").Interior.Color = self._rgb(205, 205, 205)
            self._set_number_format_safe(ws.Columns("A:A"), "@")
            self._set_number_format_safe(ws.Columns("B:B"), "@")
            ws.Columns("A:A").ColumnWidth = 18
            ws.Columns("B:B").ColumnWidth = 31.14
        return self._save_workbook(file_path, "Sheet1", headers, [], fmt)

    def _collect_calculated_rows(self, batch_id: str, imported_by: str) -> tuple[list[dict], int]:
        rows = self.session.query(TempTargetPriceImport).filter(
            TempTargetPriceImport.batch_id == batch_id,
            TempTargetPriceImport.imported_by == imported_by,
        ).order_by(TempTargetPriceImport.import_row_no.asc(), TempTargetPriceImport.id.asc()).all()
        out: list[dict] = []
        max_opt = 0
        for row in rows:
            product = self.session.query(Product).filter(Product.id == row.selected_product_id).first() if row.selected_product_id else None
            selected = self.session.query(TempTargetPriceOption).filter(TempTargetPriceOption.id == row.selected_option_id).first() if row.selected_option_id else None
            options = self.session.query(TempTargetPriceOption).filter(
                TempTargetPriceOption.batch_id == batch_id,
                TempTargetPriceOption.imported_by == imported_by,
                TempTargetPriceOption.temp_import_id == row.id,
            ).order_by(TempTargetPriceOption.opt_rank.asc(), TempTargetPriceOption.cost_novo_wvat.asc(), TempTargetPriceOption.id.asc()).all()
            max_opt = max(max_opt, len(options))
            data = {
                "Supplier Article": row.supplier_article,
                "Supplier Product Name": row.product_name,
                "Our Product Name": product.name if product else "",
                "Pack Price, L": selected.supplier_price if selected else None,
                "Price (Pack)": (selected.supplier_price * product.pack) if selected and product and product.pack else None,
                "Currency": selected.currency_code if selected else None,
                "FX rate": self._round_fx_rate(selected.fx_rate_used) if selected else None,
                "Дистр цена": None,
                "Промо цена": None,
                "curr LPC": None,
                "curr Landed cost": selected.full_cost_msk if selected else None,
            }
            for i, opt in enumerate(options, start=1):
                data[f"Cost Novo withVAT_{i}"] = opt.cost_novo_wvat
                data[f"Full Cost Msk_{i}"] = opt.full_cost_msk
                data[f"Supplier_{i}"] = opt.supplier_name
                data[f"last update_{i}"] = opt.price_date_used
                data[f"Currency_{i}"] = opt.currency_code
            out.append(data)
        for data in out:
            for i in range(1, max_opt + 1):
                data.setdefault(f"Cost Novo withVAT_{i}", None)
                data.setdefault(f"Full Cost Msk_{i}", None)
                data.setdefault(f"Supplier_{i}", None)
                data.setdefault(f"last update_{i}", None)
                data.setdefault(f"Currency_{i}", None)
        return out, max_opt

    def export_calculated(self, batch_id: str, imported_by: str, file_path: str | Path) -> Path:
        rows, max_opt = self._collect_calculated_rows(batch_id, imported_by)
        base_headers = [
            "Supplier Article", "Supplier Product Name", "Our Product Name", "Pack Price, L",
            "Price (Pack)", "Currency", "FX rate", "Дистр цена", "Промо цена", "curr LPC", "curr Landed cost",
        ]
        dynamic_headers: list[str] = []
        for i in range(1, max_opt + 1):
            dynamic_headers.extend([f"Cost Novo withVAT_{i}", f"Full Cost Msk_{i}", f"Supplier_{i}", f"last update_{i}", f"Currency_{i}"])
        headers = base_headers + dynamic_headers
        def fmt(ws, headers):
            last_base_col = self._excel_column_letter(len(base_headers))
            ws.Range(f"A1:{last_base_col}1").Interior.Color = self._rgb(205, 205, 205)
            start = len(base_headers) + 1
            for _ in range(1, max_opt + 1):
                c1 = self._excel_column_letter(start)
                c2 = self._excel_column_letter(start + 1)
                c3 = self._excel_column_letter(start + 2)
                c5 = self._excel_column_letter(start + 4)
                ws.Range(f"{c1}1:{c2}1").Interior.Color = self._rgb(0, 176, 240)
                ws.Range(f"{c3}1:{c5}1").Interior.Color = self._rgb(146, 208, 80)
                start += 5
            self._set_number_format_safe(ws.Columns("A:A"), "@")
            self._set_number_format_safe(ws.Columns("D:E"), '# ##0,0000')
            self._set_number_format_safe(ws.Columns("G:G"), '# ##0')
            self._set_number_format_safe(ws.Columns("H:K"), '# ##0 ₽')
            ws.Columns("A:A").ColumnWidth = 18
            ws.Columns("B:C").ColumnWidth = 31.14
            ws.Columns("D:K").ColumnWidth = 11
            start = len(base_headers) + 1
            for _ in range(1, max_opt + 1):
                c1 = self._excel_column_letter(start)
                c2 = self._excel_column_letter(start + 1)
                c3 = self._excel_column_letter(start + 2)
                c4 = self._excel_column_letter(start + 3)
                c5 = self._excel_column_letter(start + 4)
                self._set_number_format_safe(ws.Columns(f"{c1}:{c2}"), '# ##0 ₽')
                self._set_number_format_safe(ws.Columns(f"{c3}:{c3}"), "@")
                self._set_number_format_safe(ws.Columns(f"{c4}:{c4}"), "ДД.ММ.ГГ;@")
                self._set_number_format_safe(ws.Columns(f"{c5}:{c5}"), "@")
                ws.Columns(f"{c1}:{c2}").ColumnWidth = 10
                ws.Columns(f"{c3}:{c3}").ColumnWidth = 16
                ws.Columns(f"{c4}:{c4}").ColumnWidth = 11
                ws.Columns(f"{c5}:{c5}").ColumnWidth = 9
                start += 5
        return self._save_workbook(file_path, "Calculated", headers, rows, fmt)

    def _collect_final_rows(self, batch_id: str, imported_by: str) -> list[dict]:
        rows = self.session.query(TargetPriceCalculation).filter(
            TargetPriceCalculation.batch_id == batch_id,
            TargetPriceCalculation.imported_by == imported_by,
        ).order_by(TargetPriceCalculation.import_row_no.asc(), TargetPriceCalculation.id.asc()).all()
        out: list[dict] = []
        for row in rows:
            product = self.session.query(Product).filter(Product.id == row.product_id).first()
            donor = row.donor_supplier.name if row.donor_supplier else ""
            out.append({
                "Supplier Article": row.supplier_article,
                "Supplier Product Name": row.supplier_product_name,
                "Our Product Name": product.name if product else "",
                "Pack": product.pack if product else None,
                "Target Price, L": row.target_price_l,
                "Target Price (Pack)": row.target_price_pack,
                "Currency": row.currency_code,
                "FX rate": self._round_fx_rate(row.fx_rate_used),
                "fin.Supplier for calc": donor,
            })
        return out

    def export_final(self, batch_id: str, imported_by: str, file_path: str | Path) -> Path:
        headers = [
            "Supplier Article", "Supplier Product Name", "Our Product Name", "Pack",
            "Target Price, L", "Target Price (Pack)", "Currency", "FX rate", "fin.Supplier for calc",
        ]
        rows = self._collect_final_rows(batch_id, imported_by)
        def fmt(ws, headers):
            ws.Range("A1:I1").Interior.Color = self._rgb(205, 205, 205)
            ws.Range("E1:F1").Interior.Color = self._rgb(0, 176, 240)
            self._set_number_format_safe(ws.Columns("A:A"), "@")
            self._set_number_format_safe(ws.Columns("D:D"), "General")
            self._set_number_format_safe(ws.Columns("E:F"), '# ##0,0000')
            self._set_number_format_safe(ws.Columns("H:H"), '# ##0')
            ws.Columns("A:A").ColumnWidth = 18
            ws.Columns("B:C").ColumnWidth = 31.14
            ws.Columns("D:F").ColumnWidth = 12
            ws.Columns("G:H").ColumnWidth = 9
            ws.Columns("I:I").ColumnWidth = 18
        return self._save_workbook(file_path, "Target price", headers, rows, fmt)
