from __future__ import annotations

from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Any

import pythoncom
import win32com.client as win32
from sqlalchemy.orm import Session

from app.utils.excel_fast_writer import write_excel_table
from app.utils.excel_format_rules import FORMATS, set_number_format_safe

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
    def _round_fx_rate(value: object):
        if value is None or value == "":
            return ""
        try:
            return int(CustomerCostExporter._to_decimal(value).to_integral_value(rounding=ROUND_HALF_UP))
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
        header_range.VerticalAlignment = self._xl_vcenter
        ws.Rows(1).EntireRow.AutoFit()

    def export_template(self, file_path: str | Path) -> Path:
        file_path = Path(file_path)
        if file_path.suffix.lower() != ".xlsx":
            file_path = file_path.with_suffix(".xlsx")
        file_path.parent.mkdir(parents=True, exist_ok=True)

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

            self._apply_header_common(ws, len(headers))
            self._apply_kam_layout(ws, headers)

            try:
                excel.ActiveWindow.Zoom = 85
                ws.Activate()
                cost_col = self._excel_column_letter(headers.index("Кост руб л с НДС") + 1)
                ws.Range(f"{cost_col}2").Select()
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

    def _apply_kam_layout(self, ws, headers: list[str]):
        header_map = {str(header): idx + 1 for idx, header in enumerate(headers)}

        def col(header: str) -> str | None:
            idx = header_map.get(header)
            return self._excel_column_letter(idx) if idx else None

        def color_range(first_header: str, last_header: str, color: int) -> None:
            first = header_map.get(first_header)
            last = header_map.get(last_header)
            if first and last:
                ws.Range(
                    f"{self._excel_column_letter(first)}1:{self._excel_column_letter(last)}1"
                ).Interior.Color = color

        color_range("Дата", "Комментарии", self._rgb(205, 205, 205))
        color_range("Кост руб л с НДС", "Кост руб л с НДС", self._rgb(0, 176, 240))
        color_range("Поставщик", "Курс", self._rgb(146, 208, 80))

        formats = {
            "Дата": FORMATS.DATE,
            "Менеджер": "@",
            "Клиент": "@",
            "Код продукта": "@",
            "Категория ABC": "@",
            "Количество": FORMATS.INTEGER,
            "Объем л": FORMATS.INTEGER,
            "Кост руб л с НДС": FORMATS.MONEY_RUB_SIMPLE,
        }
        for header, fmt in formats.items():
            letter = col(header)
            if letter:
                set_number_format_safe(ws.Columns(f"{letter}:{letter}"), fmt)

        widths = {
            "Дата": 8.00,
            "Менеджер": 13.00,
            "Клиент": 13.00,
            "Код продукта": 12.57,
            "Название продукта": 35.00,
            "Фасовка": 7.29,
            "Категория ABC": 12.00,
            "Количество": 7.29,
            "Объем л": 7.29,
            "Вид закупки": 12.00,
            "Условия оплаты": 10.00,
            "Комментарии": 10.00,
            "Кост руб л с НДС": 10.86,
            "Поставщик": 16.71,
            "Валюта": 7.86,
            "Курс": 8.71,
        }
        for header, width in widths.items():
            letter = col(header)
            if letter:
                ws.Columns(f"{letter}:{letter}").ColumnWidth = width

        last_col = self._excel_column_letter(len(headers))
        ws.Range(f"A1:{last_col}1").AutoFilter(1)

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

            product = row.selected_product
            data = {
                "Дата": row.request_date,
                "Менеджер": row.manager_name,
                "Клиент": row.customer_name,
                "Код продукта": row.supplier_article,
                "Название продукта": row.product_name,
                "Фасовка": row.pack,
                "Категория ABC": (product.abc_category or "-") if product else "-",
                "Количество": row.qty_pcs,
                "Объем л": row.volume_l,
                "Вид закупки": row.purchase_type,
                "Условия оплаты": row.payment_terms,
                "Комментарии": row.comments,
            }

            for i, opt in enumerate(options, start=1):
                data[f"Cost Novo with VAT_{i}"] = opt.cost_novo_wvat
                data[f"Full Cost Msk_{i}"] = opt.full_cost_msk
                data[f"Supplier_{i}"] = opt.supplier_name
                data[f"last update_{i}"] = opt.price_date_used
                data[f"FX rate_{i}"] = self._round_fx_rate(opt.fx_rate_used)
                data[f"Currency_{i}"] = opt.currency_code

            out_rows.append(data)

        for row in out_rows:
            for i in range(1, max_opt + 1):
                row.setdefault(f"Cost Novo with VAT_{i}", None)
                row.setdefault(f"Full Cost Msk_{i}", None)
                row.setdefault(f"Supplier_{i}", None)
                row.setdefault(f"last update_{i}", None)
                row.setdefault(f"FX rate_{i}", None)
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
            "Категория ABC",
            "Количество",
            "Объем л",
            "Вид закупки",
            "Условия оплаты",
            "Комментарии",
        ]

        dynamic_headers: list[str] = []
        for i in range(1, max_opt + 1):
            dynamic_headers.extend([
                f"Cost Novo with VAT_{i}",
                f"Full Cost Msk_{i}",
                f"Supplier_{i}",
                f"last update_{i}",
                f"FX rate_{i}",
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

            write_excel_table(
                ws,
                headers,
                rows,
                value_getter=lambda row, header, _col_index: self._excel_value_or_blank(row.get(header)),
            )

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
                c6 = self._excel_column_letter(start_col + 5)

                ws.Range(f"{c1}1:{c2}1").Interior.Color = self._rgb(0, 176, 240)
                ws.Range(f"{c3}1:{c6}1").Interior.Color = self._rgb(146, 208, 80)

                start_col += 6

            base_header_map = {str(header): idx + 1 for idx, header in enumerate(base_headers)}
            for header, fmt in {
                "Дата": FORMATS.DATE,
                "Менеджер": "@",
                "Клиент": "@",
                "Код продукта": "@",
                "Категория ABC": "@",
                "Количество": FORMATS.INTEGER,
                "Объем л": FORMATS.INTEGER,
            }.items():
                idx = base_header_map.get(header)
                if idx:
                    letter = self._excel_column_letter(idx)
                    set_number_format_safe(ws.Columns(f"{letter}:{letter}"), fmt)

            start_col = len(base_headers) + 1
            for _ in range(1, max_opt + 1):
                c1 = self._excel_column_letter(start_col)
                c2 = self._excel_column_letter(start_col + 1)
                c3 = self._excel_column_letter(start_col + 2)
                c4 = self._excel_column_letter(start_col + 3)
                c5 = self._excel_column_letter(start_col + 4)
                c6 = self._excel_column_letter(start_col + 5)

                set_number_format_safe(ws.Columns(f"{c1}:{c2}"), FORMATS.MONEY_RUB_SIMPLE)
                set_number_format_safe(ws.Columns(f"{c3}:{c3}"), FORMATS.TEXT)
                set_number_format_safe(ws.Columns(f"{c4}:{c4}"), FORMATS.DATE)
                set_number_format_safe(ws.Columns(f"{c5}:{c5}"), FORMATS.FX_INTEGER)
                set_number_format_safe(ws.Columns(f"{c6}:{c6}"), FORMATS.TEXT)

                start_col += 6

            base_widths = {
                "Дата": 11.00,
                "Менеджер": 16.14,
                "Клиент": 18.14,
                "Код продукта": 16.00,
                "Название продукта": 31.14,
                "Фасовка": 10.50,
                "Категория ABC": 12.00,
                "Количество": 10.50,
                "Объем л": 10.50,
                "Вид закупки": 18.00,
                "Условия оплаты": 18.00,
                "Комментарии": 24.00,
            }
            for header, width in base_widths.items():
                idx = base_header_map.get(header)
                if idx:
                    letter = self._excel_column_letter(idx)
                    ws.Columns(f"{letter}:{letter}").ColumnWidth = width

            start_col = len(base_headers) + 1
            for _ in range(1, max_opt + 1):
                c1 = self._excel_column_letter(start_col)
                c2 = self._excel_column_letter(start_col + 1)
                c3 = self._excel_column_letter(start_col + 2)
                c4 = self._excel_column_letter(start_col + 3)
                c5 = self._excel_column_letter(start_col + 4)
                c6 = self._excel_column_letter(start_col + 5)

                ws.Columns(f"{c1}:{c1}").ColumnWidth = 9.0
                ws.Columns(f"{c2}:{c2}").ColumnWidth = 9.0
                ws.Columns(f"{c3}:{c3}").ColumnWidth = 16.14
                ws.Columns(f"{c4}:{c4}").ColumnWidth = 10.14
                ws.Columns(f"{c5}:{c5}").ColumnWidth = 7.29
                ws.Columns(f"{c6}:{c6}").ColumnWidth = 9.14

                start_col += 6

            last_col = self._excel_column_letter(len(headers))
            ws.Range(f"A1:{last_col}1").AutoFilter(1)

            try:
                excel.ActiveWindow.Zoom = 85
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

            product = row.selected_product
            out_rows.append({
                "Дата": row.request_date,
                "Менеджер": row.manager_name,
                "Клиент": row.customer_name,
                "Код продукта": row.supplier_article,
                "Название продукта": row.product_name,
                "Фасовка": row.pack,
                "Категория ABC": (product.abc_category or "-") if product else "-",
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
            "Категория ABC",
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

            write_excel_table(
                ws,
                headers,
                rows,
                value_getter=lambda row, header, _col_index: self._excel_value_or_blank(row.get(header)),
            )

            self._apply_header_common(ws, len(headers))

            self._apply_kam_layout(ws, headers)

            try:
                excel.ActiveWindow.Zoom = 85
                ws.Activate()
                cost_col = self._excel_column_letter(headers.index("Кост руб л с НДС") + 1)
                ws.Range(f"{cost_col}2").Select()
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
