from __future__ import annotations

from datetime import datetime
from decimal import Decimal
from pathlib import Path
from typing import Any

import pythoncom
import win32com.client as win32
from sqlalchemy.orm import Session

from app.db.models import TempCustomerCostImport, TempCustomerCostOption


class CustomerCostExporter:
    def __init__(self, session: Session):
        self.session = session

        self._xl_center = -4108
        self._xl_vcenter = -4160

    @staticmethod
    def _safe_name(value: str) -> str:
        s = (value or "").strip() or "NoName"
        for ch in ['\\', '/', ':', '*', '?', '"', '<', '>', '|']:
            s = s.replace(ch, '_')
        return s

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
        header_range.VerticalAlignment = self._xl_vcenter
        ws.Rows(1).EntireRow.AutoFit()

    def _collect_export_rows(self, batch_id: str, imported_by: str) -> tuple[list[dict], int]:
        rows = self.session.query(TempCustomerCostImport).filter(
            TempCustomerCostImport.batch_id == batch_id,
            TempCustomerCostImport.imported_by == imported_by,
        ).order_by(
            TempCustomerCostImport.import_row_no.asc(),
            TempCustomerCostImport.id.asc(),
        ).all()

        out_rows: list[dict] = []
        max_opt = 0

        for row in rows:
            options = self.session.query(TempCustomerCostOption).filter(
                TempCustomerCostOption.batch_id == batch_id,
                TempCustomerCostOption.imported_by == imported_by,
                TempCustomerCostOption.temp_import_id == row.id,
            ).order_by(
                TempCustomerCostOption.opt_rank.asc(),
                TempCustomerCostOption.full_cost_msk.asc(),
                TempCustomerCostOption.id.asc(),
            ).all()

            max_opt = max(max_opt, len(options))

            data = {
                "Дата": row.request_date,
                "Менеджер": row.manager_name,
                "Клиент": row.customer_name,
                "Код продукта": row.supplier_article,
                "Название продукта": row.product_name,
                "Фасовка": row.pack,
                "Количество": row.qty_pcs,
                "Объем л": row.volume_l,
                "Вид закупки": row.purchase_type,
                "Условия оплаты": row.payment_terms,
                "Комментарии": row.comments,
            }

            for i, opt in enumerate(options, start=1):
                data[f"Cost Novo withVAT_{i}"] = opt.cost_novo_wvat
                data[f"Full Cost Msk_{i}"] = opt.full_cost_msk
                data[f"Supplier_{i}"] = opt.supplier_name
                data[f"last update_{i}"] = opt.price_date_used
                data[f"Currency_{i}"] = opt.currency_code

            out_rows.append(data)

        for row in out_rows:
            for i in range(1, max_opt + 1):
                row.setdefault(f"Cost Novo withVAT_{i}", None)
                row.setdefault(f"Full Cost Msk_{i}", None)
                row.setdefault(f"Supplier_{i}", None)
                row.setdefault(f"last update_{i}", None)
                row.setdefault(f"Currency_{i}", None)

        return out_rows, max_opt

    def export_calculated(self, batch_id: str, imported_by: str, file_path: str | Path) -> Path:
        file_path = Path(file_path)
        if file_path.suffix.lower() != ".xlsx":
            file_path = file_path.with_suffix(".xlsx")
        file_path.parent.mkdir(parents=True, exist_ok=True)

        rows, max_opt = self._collect_export_rows(batch_id=batch_id, imported_by=imported_by)

        base_headers = [
            "Дата",
            "Менеджер",
            "Клиент",
            "Код продукта",
            "Название продукта",
            "Фасовка",
            "Количество",
            "Объем л",
            "Вид закупки",
            "Условия оплаты",
            "Комментарии",
        ]

        dynamic_headers: list[str] = []
        for i in range(1, max_opt + 1):
            dynamic_headers.extend([
                f"Cost Novo withVAT_{i}",
                f"Full Cost Msk_{i}",
                f"Supplier_{i}",
                f"last update_{i}",
                f"Currency_{i}",
            ])

        headers = base_headers + dynamic_headers

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
            ws.Name = "Calculated"

            for col_index, header in enumerate(headers, start=1):
                ws.Cells(1, col_index).Value = header

            row_num = 2
            for row in rows:
                for col_index, header in enumerate(headers, start=1):
                    ws.Cells(row_num, col_index).Value = self._excel_value_or_blank(row.get(header))
                row_num += 1

            self._apply_header_common(ws, len(headers))

            if len(base_headers) > 0:
                last_base_col = self._excel_column_letter(len(base_headers))
                ws.Range(f"A1:{last_base_col}1").Interior.Color = self._rgb(205, 205, 205)

            start_col = len(base_headers) + 1
            for _ in range(1, max_opt + 1):
                c1 = self._excel_column_letter(start_col)
                c2 = self._excel_column_letter(start_col + 1)
                c3 = self._excel_column_letter(start_col + 2)
                c4 = self._excel_column_letter(start_col + 3)
                c5 = self._excel_column_letter(start_col + 4)

                ws.Range(f"{c1}1:{c2}1").Interior.Color = self._rgb(0, 176, 240)
                ws.Range(f"{c3}1:{c5}1").Interior.Color = self._rgb(146, 208, 80)

                start_col += 5

            ws.Columns("A:A").NumberFormatLocal = "ДД.ММ.ГГ;@"
            ws.Columns("B:C").NumberFormatLocal = "@"
            ws.Columns("D:D").NumberFormatLocal = "@"

            start_col = len(base_headers) + 1
            for _ in range(1, max_opt + 1):
                c1 = self._excel_column_letter(start_col)
                c2 = self._excel_column_letter(start_col + 1)
                c3 = self._excel_column_letter(start_col + 2)
                c4 = self._excel_column_letter(start_col + 3)
                c5 = self._excel_column_letter(start_col + 4)

                ws.Columns(f"{c1}:{c2}").NumberFormatLocal = '# ##0 ₽'
                ws.Columns(f"{c3}:{c3}").NumberFormatLocal = "@"
                ws.Columns(f"{c4}:{c4}").NumberFormatLocal = "ДД.ММ.ГГ;@"

                start_col += 5

            ws.Columns("A:A").ColumnWidth = 11.00
            ws.Columns("B:B").ColumnWidth = 16.14
            ws.Columns("C:C").ColumnWidth = 18.14
            ws.Columns("D:D").ColumnWidth = 16.00
            ws.Columns("E:E").ColumnWidth = 31.14
            ws.Columns("F:H").ColumnWidth = 10.50
            ws.Columns("I:J").ColumnWidth = 18.00
            ws.Columns("K:K").ColumnWidth = 24.00

            start_col = len(base_headers) + 1
            for _ in range(1, max_opt + 1):
                c1 = self._excel_column_letter(start_col)
                c2 = self._excel_column_letter(start_col + 1)
                c3 = self._excel_column_letter(start_col + 2)
                c4 = self._excel_column_letter(start_col + 3)
                c5 = self._excel_column_letter(start_col + 4)

                ws.Columns(f"{c1}:{c1}").ColumnWidth = 11.50
                ws.Columns(f"{c2}:{c2}").ColumnWidth = 11.50
                ws.Columns(f"{c3}:{c3}").ColumnWidth = 16.14
                ws.Columns(f"{c4}:{c4}").ColumnWidth = 10.14
                ws.Columns(f"{c5}:{c5}").ColumnWidth = 9.14

                start_col += 5

            last_col = self._excel_column_letter(len(headers))
            ws.Range(f"A1:{last_col}1").AutoFilter(1)

            try:
                excel.ActiveWindow.Zoom = 90
                ws.Activate()
                freeze_cell = self._excel_column_letter(len(base_headers) + 1) + "2"
                ws.Range(freeze_cell).Select()
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

    def _collect_kam_rows(self, batch_id: str, imported_by: str, manager_name: str, customer_name: str) -> list[dict]:
        rows = self.session.query(TempCustomerCostImport).filter(
            TempCustomerCostImport.batch_id == batch_id,
            TempCustomerCostImport.imported_by == imported_by,
            TempCustomerCostImport.manager_name == manager_name,
            TempCustomerCostImport.customer_name == customer_name,
        ).order_by(
            TempCustomerCostImport.import_row_no.asc(),
            TempCustomerCostImport.id.asc(),
        ).all()

        out_rows: list[dict] = []
        for row in rows:
            option = None
            if row.selected_option_id is not None:
                option = self.session.query(TempCustomerCostOption).filter(
                    TempCustomerCostOption.id == row.selected_option_id
                ).first()

            out_rows.append({
                "Дата": row.request_date,
                "Менеджер": row.manager_name,
                "Клиент": row.customer_name,
                "Код продукта": row.supplier_article,
                "Название продукта": row.product_name,
                "Фасовка": row.pack,
                "Количество": row.qty_pcs,
                "Объем л": row.volume_l,
                "Вид закупки": row.purchase_type,
                "Условия оплаты": row.payment_terms,
                "Комментарии": row.comments,
                "Кост руб л с НДС": option.cost_novo_wvat if option else None,
                "Поставщик": option.supplier_name if option else None,
                "Валюта": option.currency_code if option else None,
                "Курс": option.fx_rate_used if option else None,
            })

        return out_rows

    def _export_one_kam_file(
        self,
        batch_id: str,
        imported_by: str,
        manager_name: str,
        customer_name: str,
        file_path: str | Path,
    ) -> Path:
        file_path = Path(file_path)
        if file_path.suffix.lower() != ".xlsx":
            file_path = file_path.with_suffix(".xlsx")
        file_path.parent.mkdir(parents=True, exist_ok=True)

        rows = self._collect_kam_rows(batch_id, imported_by, manager_name, customer_name)

        headers = [
            "Дата",
            "Менеджер",
            "Клиент",
            "Код продукта",
            "Название продукта",
            "Фасовка",
            "Количество",
            "Объем л",
            "Вид закупки",
            "Условия оплаты",
            "Комментарии",
            "Кост руб л с НДС",
            "Поставщик",
            "Валюта",
            "Курс",
        ]

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
            ws.Name = "KAM"

            for col_index, header in enumerate(headers, start=1):
                ws.Cells(1, col_index).Value = header

            row_num = 2
            for row in rows:
                for col_index, header in enumerate(headers, start=1):
                    ws.Cells(row_num, col_index).Value = self._excel_value_or_blank(row.get(header))
                row_num += 1

            self._apply_header_common(ws, len(headers))

            # A:K серый
            ws.Range("A1:K1").Interior.Color = self._rgb(205, 205, 205)

            # L голубой
            ws.Range("L1:L1").Interior.Color = self._rgb(0, 176, 240)

            # M:O зеленый
            ws.Range("M1:O1").Interior.Color = self._rgb(146, 208, 80)

            # Форматы
            ws.Columns("A:A").NumberFormatLocal = "ДД.ММ.ГГ;@"
            ws.Columns("B:C").NumberFormatLocal = "@"
            ws.Columns("F:F").NumberFormatLocal = "General"
            ws.Columns("G:H").NumberFormatLocal = '# ##0;[Red]-# ##0;"-"'
            ws.Columns("L:L").NumberFormatLocal = '# ##0 ₽'

            # Ширины по образцу
            ws.Columns("A:A").ColumnWidth = 9.14
            ws.Columns("B:B").ColumnWidth = 13.00
            ws.Columns("C:C").ColumnWidth = 19.43
            ws.Columns("D:D").ColumnWidth = 12.57
            ws.Columns("E:E").ColumnWidth = 35.00
            ws.Columns("F:F").ColumnWidth = 9.14
            ws.Columns("G:G").ColumnWidth = 9.14
            ws.Columns("H:H").ColumnWidth = 13.00
            ws.Columns("I:I").ColumnWidth = 9.14
            ws.Columns("J:J").ColumnWidth = 13.00
            ws.Columns("K:K").ColumnWidth = 13.00
            ws.Columns("L:L").ColumnWidth = 10.86
            ws.Columns("M:M").ColumnWidth = 16.71
            ws.Columns("N:N").ColumnWidth = 7.86
            ws.Columns("O:O").ColumnWidth = 8.71

            ws.Range("A1:O1").AutoFilter(1)

            try:
                excel.ActiveWindow.Zoom = 90
                ws.Activate()
                ws.Range("L2").Select()
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

    def export_kam_files(self, batch_id: str, imported_by: str, folder_path: str | Path) -> list[Path]:
        folder_path = Path(folder_path)
        folder_path.mkdir(parents=True, exist_ok=True)

        groups = self.session.query(
            TempCustomerCostImport.manager_name,
            TempCustomerCostImport.customer_name,
        ).filter(
            TempCustomerCostImport.batch_id == batch_id,
            TempCustomerCostImport.imported_by == imported_by,
        ).distinct().all()

        result: list[Path] = []
        for manager_name, customer_name in groups:
            file_path = folder_path / f"{self._safe_name(manager_name)}_{self._safe_name(customer_name)}_KAM.xlsx"
            output_path = self._export_one_kam_file(
                batch_id=batch_id,
                imported_by=imported_by,
                manager_name=manager_name,
                customer_name=customer_name,
                file_path=file_path,
            )
            result.append(output_path)

        return result
# Backward-compatible alias.
CustomerCostExport = CustomerCostExporter
