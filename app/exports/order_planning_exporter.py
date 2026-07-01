from __future__ import annotations

from datetime import datetime
from decimal import Decimal
from pathlib import Path
from typing import Any, Sequence

import pythoncom
import win32com.client as win32
from sqlalchemy.orm import Session

from app.db.models import CurrentSupplierPrice, PriceHistory, Product, ProductStock, Supplier
from app.services.cost_calculation_service import CostCalculationService
from app.services.supplier_service import SupplierService
from app.services.supplier_currency_cost_service import SupplierCurrencyCostService
from app.exports.excel_column_format import apply_standard_worksheet_format, excel_value_by_header
from app.utils.excel_fast_writer import write_excel_table
from app.utils.output_headers import standardize_output_header


class OrderPlanningExporter:
    def __init__(self, session: Session):
        self.session = session
        self.cost_calculation = CostCalculationService(session)
        self.supplier_service = SupplierService(session)
        self.currency_cost_service = SupplierCurrencyCostService(session)
        self._xl_center = -4108
        self._xl_vcenter = -4160

    @staticmethod
    def _rgb(r: int, g: int, b: int) -> int:
        return r + g * 256 + b * 65536

    @staticmethod
    def _excel_column_letter(col_num: int) -> str:
        result = ""
        while col_num > 0:
            col_num, rem = divmod(col_num - 1, 26)
            result = chr(65 + rem) + result
        return result

    @staticmethod
    def _to_decimal(value: object) -> Decimal:
        if value is None:
            return Decimal("0")
        if isinstance(value, Decimal):
            return value
        return Decimal(str(value))

    @staticmethod
    def _excel_value(header: str, value: object) -> Any:
        excel_value = excel_value_by_header(header, value)
        if isinstance(excel_value, (int, float)) and not isinstance(excel_value, bool) and excel_value == 0:
            return ""
        return excel_value

    def _create_excel_app(self):
        pythoncom.CoInitialize()
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        return excel

    def _apply_header_common(self, ws, headers_count: int) -> None:
        ws.Cells.Font.Name = "Aptos Narrow"
        ws.Cells.Font.Size = 11
        last_col = self._excel_column_letter(headers_count)
        rng = ws.Range(f"A1:{last_col}1")
        rng.Font.Name = "Aptos Narrow"
        rng.Font.Size = 11
        rng.Font.Bold = True
        rng.WrapText = True
        rng.HorizontalAlignment = self._xl_center
        rng.VerticalAlignment = self._xl_vcenter
        ws.Rows(1).RowHeight = 60

    @staticmethod
    def _excel_invariant_number_format(fmt: str | None) -> str | None:
        if not fmt:
            return fmt
        return (
            str(fmt)
            .replace("ДД", "dd")
            .replace("ММ", "mm")
            .replace("ГГ", "yy")
            .replace(",0000", ".0000")
            .replace(",00", ".00")
            .replace(",0", ".0")
        )

    def _set_format(self, target, fmt: str) -> None:
        candidates = [
            ("NumberFormatLocal", fmt),
            ("NumberFormat", self._excel_invariant_number_format(fmt) or fmt),
            ("NumberFormat", fmt),
            ("NumberFormat", "General"),
        ]
        for attr, value in candidates:
            try:
                setattr(target, attr, value)
                return
            except Exception:
                pass

    def _header_map(self, headers: Sequence[str]) -> dict[str, int]:
        return {str(header): idx + 1 for idx, header in enumerate(headers)}

    def _col(self, header_map: dict[str, int], header: str) -> str | None:
        idx = header_map.get(header)
        return self._excel_column_letter(idx) if idx else None

    def _color_headers(self, ws, header_map: dict[str, int], first: str, last: str, color: int, font_color: int | None = None) -> None:
        a = header_map.get(first)
        b = header_map.get(last)
        if not a or not b:
            return
        rng = ws.Range(f"{self._excel_column_letter(a)}1:{self._excel_column_letter(b)}1")
        rng.Interior.Color = color
        if font_color is not None:
            rng.Font.Color = font_color

    def _set_width(self, ws, header_map: dict[str, int], header: str, width: float) -> None:
        letter = self._col(header_map, header)
        if letter:
            ws.Columns(f"{letter}:{letter}").ColumnWidth = width

    def _format_cols(self, ws, header_map: dict[str, int], headers: Sequence[str], fmt: str) -> None:
        for header in headers:
            letter = self._col(header_map, header)
            if letter:
                self._set_format(ws.Columns(f"{letter}:{letter}"), fmt)

    def _calc_supplier_option(
        self,
        supplier: Supplier,
        product_id: int,
        supplier_price: object,
        price_date,
        price_currency_code: object | None = None,
    ):
        try:
            calc = self.currency_cost_service.calculate_costs_for_price_record(
                supplier_id=supplier.id,
                product_id=product_id,
                supplier_price=self._to_decimal(supplier_price),
                price_currency_code=price_currency_code or supplier.base_currency,
            )
            return {
                "supplier": supplier.name or "",
                "cost_novo": calc.cost_novo_wvat,
                "full_cost": calc.full_cost_msk,
                "date": price_date,
                "currency": calc.currency_code,
            }
        except Exception:
            return None

    def _get_supplier_options(self, product_id: int) -> list[dict]:
        supplier_ids = set()
        supplier_ids.update(row[0] for row in self.session.query(CurrentSupplierPrice.supplier_id).filter(CurrentSupplierPrice.product_id == product_id).all())
        supplier_ids.update(row[0] for row in self.session.query(PriceHistory.supplier_id).filter(PriceHistory.product_id == product_id).all())

        options: list[dict] = []
        for supplier_id in supplier_ids:
            supplier = self.session.query(Supplier).filter(Supplier.id == supplier_id).first()
            if not supplier or not bool(getattr(supplier, "rating_calc", True)):
                continue
            current = self.session.query(CurrentSupplierPrice).filter(
                CurrentSupplierPrice.supplier_id == supplier_id,
                CurrentSupplierPrice.product_id == product_id,
            ).first()
            if current is not None:
                option = self._calc_supplier_option(
                    supplier,
                    product_id,
                    current.price,
                    current.last_update,
                    current.currency,
                )
            else:
                history = self.session.query(PriceHistory).filter(
                    PriceHistory.supplier_id == supplier_id,
                    PriceHistory.product_id == product_id,
                ).order_by(PriceHistory.price_date.desc()).first()
                option = self._calc_supplier_option(
                    supplier,
                    product_id,
                    history.price,
                    history.price_date,
                    history.currency,
                ) if history else None
            if option and option["full_cost"] is not None:
                options.append(option)
        options.sort(key=lambda x: (self._to_decimal(x["full_cost"]), str(x["supplier"]).lower()))
        return options

    def _base_headers(self) -> list[str]:
        return [
            "Brand",
            "Product Name",
            "Pack",
            "Ср.Продажи мес",
            "Safe Stock (st), mnth",
            "Safe Stock (st+tr), mnth",
            "Safe Stock (+ord), mnth",
            "к Быстрому Заказу, шт",
            "к Быстрому Заказу, л",
            "к Заказу, шт",
            "к Заказу, л",
            "Дистр цена",
            "Промо цена",
            "Stock",
            "Transit",
            "Purchase Order",
            "Order IS",
            "Stock IS",
            "Reserve cust",
            "Reserve E-Comm",
            "Damaged",
        ]

    def build_export_data(self, display_rows: list[dict]) -> tuple[list[str], list[list[object]]]:
        max_suppliers = 0
        prepared = []
        for row in display_rows:
            product_id = row.get("product_id")
            options = self._get_supplier_options(int(product_id)) if product_id else []
            max_suppliers = max(max_suppliers, len(options))
            prepared.append((row, options))

        headers = self._base_headers()
        for idx in range(1, max_suppliers + 1):
            headers.extend([
                f"Cost Novo with VAT_{idx}",
                f"Full Cost Msk_{idx}",
                f"Supplier_{idx}",
                f"last update_{idx}",
                f"Currency_{idx}",
            ])

        rows: list[list[object]] = []
        for row, options in prepared:
            values = [
                row.get("brand", ""),
                row.get("product_name", ""),
                row.get("pack"),
                row.get("avg_sales_month"),
                row.get("safe_stock_st_month"),
                row.get("safe_stock_st_tr_month"),
                row.get("safe_stock_ord_month"),
                row.get("quick_order_pcs"),
                row.get("quick_order_l"),
                row.get("std_order_pcs"),
                row.get("std_order_l"),
                row.get("distr_price"),
                row.get("promo_price"),
                row.get("stock"),
                row.get("transit"),
                row.get("purchase_order"),
                row.get("order_is"),
                row.get("stock_is"),
                row.get("reserve"),
                row.get("reserve_ecomm"),
                row.get("markdown"),
            ]
            for option in options:
                values.extend([option["cost_novo"], option["full_cost"], option["supplier"], option["date"], option["currency"]])
            while len(values) < len(headers):
                values.append("")
            rows.append(values)
        return headers, rows

    def export_report(self, *, display_rows: list[dict], output_path: str | Path) -> Path:
        output_path = Path(output_path)
        if output_path.suffix.lower() != ".xlsx":
            output_path = output_path.with_suffix(".xlsx")
        output_path.parent.mkdir(parents=True, exist_ok=True)
        target_path = output_path.resolve()
        if target_path.exists():
            try:
                target_path.unlink()
            except PermissionError:
                raise PermissionError(f"Не удается перезаписать файл:\n{target_path}\n\nСкорее всего, он открыт в Excel. Закрой файл и попробуй снова.")

        headers, rows = self.build_export_data(display_rows)
        excel = None
        wb = None
        try:
            excel = self._create_excel_app()
            wb = excel.Workbooks.Add()
            ws = wb.Worksheets(1)
            ws.Name = "Sheet1"

            write_excel_table(
                ws,
                headers,
                rows,
                header_getter=standardize_output_header,
                value_getter=lambda row, header, col_index: self._excel_value(
                    str(header),
                    row[col_index] if col_index < len(row) else "",
                ),
            )

            apply_standard_worksheet_format(ws, headers, freeze_cell="L2", zoom=85)

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
