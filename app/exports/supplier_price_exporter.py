from __future__ import annotations

from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Any

import pythoncom
import win32com.client as win32
from sqlalchemy.orm import Session

from app.db.models import PriceHistory, Product, ProductStock, Supplier, SupplierPriceCalculation, TempPriceImport, OrderPlanningCalculation
from app.services.cost_calculation_service import CostCalculationService
from app.services.price_repository import PriceRepository


class SupplierPriceExporter:
    def __init__(self, session: Session):
        self.session = session
        self.cost_calculation = CostCalculationService(session)
        self.price_repository = PriceRepository(session)

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
    def _safe_filename(value: str) -> str:
        s = (value or "").strip()
        for ch in ["\\", "/", ":", "*", "?", '"', "<", ">", "|"]:
            s = s.replace(ch, "_")
        return s or "Supplier"

    @staticmethod
    def _excel_value(value: object) -> Any:
        if value is None:
            return None
        if isinstance(value, Decimal):
            return float(value)
        return value

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
    def _base_order_header(header: object) -> str:
        text = str(header or "")
        if text.startswith("к Быстрому заказу, л"):
            return "к Быстрому заказу, л"
        if text.startswith("к Заказу, л"):
            return "к Заказу, л"
        return text

    def _row_value_by_export_header(self, row: dict, header: object):
        """Return row value for the visible Excel header.

        Export headers include months, for example
        "к Быстрому заказу, л (3 м)", while row dictionaries use
        stable keys without the suffix. Normalizing the header prevents
        empty Excel cells after the suffix is added.
        """
        key = self._base_order_header(header)
        return row.get(key)

    @classmethod
    def _is_order_plan_export_header(cls, header: object) -> bool:
        key = cls._base_order_header(header)
        return key in {"Ср.Продажи мес", "к Быстрому заказу, л", "к Заказу, л"}

    @staticmethod
    def _round_fx_rate(value: object):
        if value is None or value == "":
            return ""
        try:
            return int(SupplierPriceExporter._to_decimal(value).to_integral_value(rounding=ROUND_HALF_UP))
        except Exception:
            return ""

    def _get_supplier_currency_rate(self, supplier_id: int | None):
        if not supplier_id:
            return "", None
        supplier = self.session.query(Supplier).filter(Supplier.id == supplier_id).first()
        if supplier is None:
            return "", None
        try:
            from app.services.supplier_service import SupplierService
            rate = SupplierService(self.session).get_rate_to_rub(supplier.base_currency)
        except Exception:
            rate = None
        return supplier.base_currency or "", rate

    def _copy_header_style(self, ws, from_header: str, to_header: str, headers: list[str]):
        if from_header not in headers or to_header not in headers:
            return
        src = self._excel_column_letter(headers.index(from_header) + 1)
        dst = self._excel_column_letter(headers.index(to_header) + 1)
        try:
            ws.Range(f"{dst}1").Interior.Color = ws.Range(f"{src}1").Interior.Color
            # В CostCalc_ цвет шрифта должен оставаться Automatic, если он не задан явно
            # для конкретного блока. Нельзя копировать белый шрифт со старых диапазонов.
            ws.Range(f"{dst}1").Font.ColorIndex = -4105
        except Exception:
            pass

    def _format_fx_column(self, ws, header: str, headers: list[str]):
        if header not in headers:
            return
        col = self._excel_column_letter(headers.index(header) + 1)
        ws.Columns(f"{col}:{col}").NumberFormatLocal = "# ##0"
        ws.Columns(f"{col}:{col}").ColumnWidth = 7.29

    @staticmethod
    def _calc_pack_price(price_per_l: object, pack: object):
        if price_per_l is None or pack is None:
            return None
        d_price = SupplierPriceExporter._to_decimal(price_per_l)
        d_pack = SupplierPriceExporter._to_decimal(pack)
        if d_pack == 0:
            return None
        return (d_price * d_pack).quantize(Decimal("0.0001"))

    @staticmethod
    def _is_blank_excel_value(value: object) -> bool:
        if value is None:
            return True
        if isinstance(value, str):
            return value.strip() == ""
        return False

    @staticmethod
    def _normalize_header(value: object) -> str:
        return "".join(ch for ch in str(value or "").lower() if ch.isalnum())

    @staticmethod
    def _to_decimal_excel_value(value: object) -> Decimal:
        if isinstance(value, str):
            cleaned = value.strip().replace(" ", "").replace("\u00a0", "").replace(",", ".")
            return Decimal(cleaned)
        return SupplierPriceExporter._to_decimal(value)

    @staticmethod
    def _calc_qty_volume_for_export(qty: object, volume: object, pack: object):
        qty_is_blank = SupplierPriceExporter._is_blank_excel_value(qty)
        volume_is_blank = SupplierPriceExporter._is_blank_excel_value(volume)

        if qty_is_blank and volume_is_blank:
            return None, None

        try:
            d_pack = SupplierPriceExporter._to_decimal(pack)
        except Exception:
            d_pack = Decimal("0")

        if d_pack == 0:
            return qty if not qty_is_blank else None, volume if not volume_is_blank else None

        if not qty_is_blank and volume_is_blank:
            try:
                d_qty = SupplierPriceExporter._to_decimal_excel_value(qty)
                return d_qty, (d_qty * d_pack).quantize(Decimal("0.0001"))
            except Exception:
                return qty, None

        if qty_is_blank and not volume_is_blank:
            try:
                d_volume = SupplierPriceExporter._to_decimal_excel_value(volume)
                return (d_volume / d_pack).quantize(Decimal("0.0001")), d_volume
            except Exception:
                return None, volume

        return qty, volume

    def _read_qty_volume_from_source_excel(self, excel, source_file_path: str | Path | None) -> dict[int, dict]:
        if not source_file_path:
            return {}

        source_path = Path(source_file_path)
        if not source_path.exists():
            return {}

        wb_source = None
        try:
            wb_source = excel.Workbooks.Open(str(source_path.resolve()), ReadOnly=True)
            ws_source = wb_source.Worksheets(1)
            used_cols = int(ws_source.UsedRange.Columns.Count)
            used_rows = int(ws_source.UsedRange.Rows.Count)

            qty_col = None
            volume_col = None

            for col_index in range(1, used_cols + 1):
                header_norm = self._normalize_header(ws_source.Cells(1, col_index).Value)
                if header_norm in {"qtypcs", "qtypc", "qtypieces", "quantitypcs", "quantity"}:
                    qty_col = col_index
                elif header_norm in {"volumel", "volume", "volemul"}:
                    volume_col = col_index

            if qty_col is None and volume_col is None:
                return {}

            result: dict[int, dict] = {}
            for excel_row in range(2, used_rows + 1):
                import_row_no = excel_row - 1
                result[import_row_no] = {
                    "Qty, pcs": ws_source.Cells(excel_row, qty_col).Value if qty_col else None,
                    "Volume, L": ws_source.Cells(excel_row, volume_col).Value if volume_col else None,
                }

            return result
        except Exception:
            return {}
        finally:
            try:
                if wb_source is not None:
                    wb_source.Close(SaveChanges=False)
            except Exception:
                pass

    @staticmethod
    def _excel_column_letter(col_num: int) -> str:
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

    def _set_number_format_safe(self, target, format_en: str, format_local: str | None = None):
        try:
            target.NumberFormat = format_en
        except Exception:
            if format_local:
                target.NumberFormatLocal = format_local
            else:
                raise


    def _header_map(self, headers: list[str]) -> dict[str, int]:
        return {str(header): idx + 1 for idx, header in enumerate(headers)}

    def _column_letter_by_header(self, header_map: dict[str, int], header: str) -> str | None:
        idx = header_map.get(header)
        if not idx:
            return None
        return self._excel_column_letter(idx)

    def _set_format_by_header(self, ws, header_map: dict[str, int], header: str, format_local: str) -> None:
        letter = self._column_letter_by_header(header_map, header)
        if letter:
            ws.Columns(f"{letter}:{letter}").NumberFormatLocal = format_local

    def _set_width_by_header(self, ws, header_map: dict[str, int], header: str, width: float) -> None:
        letter = self._column_letter_by_header(header_map, header)
        if letter:
            ws.Columns(f"{letter}:{letter}").ColumnWidth = width

    def _apply_calculated_export_formats_by_header(self, ws, headers: list[str]) -> None:
        header_map = self._header_map(headers)

        text_headers = ["Supplier Article"]
        price_decimal_headers = ["Price, L", "Price, pack", "Price, L (prev)"]
        rub_headers = [
            "Cost Novo with VAT",
            "Full Cost Msk",
            "Cost Novo with VAT (prev)",
            "Full Cost Msk (prev)",
            "Дистр цена",
            "Промо цена",
            "curr LPC",
            "curr Landed cost",
            "Best full Price, L",
            "Best full Price, L 2",
        ]
        date_headers = ["last update", "last update (prev)", "last update Best1", "last update Best2"]
        integer_headers = [
            "Qty, pcs",
            "Volume, L",
            "Stock",
            "Transit",
            "Purchase Order",
            "Order IS",
            "Stock IS",
            "Reserve cust",
            "Reserve E-Comm",
            "Damaged",
            "Ср.Продажи мес",
        ]
        integer_headers += [h for h in headers if h.startswith("к Быстрому заказу, л") or h.startswith("к Заказу, л")]
        fx_headers = [h for h in headers if h == "FX rate" or h in {"FX rate Best1", "FX rate Best2"} or h.startswith("FX rate_")]

        for header in text_headers:
            self._set_format_by_header(ws, header_map, header, "@")
        for header in price_decimal_headers:
            self._set_format_by_header(ws, header_map, header, "# ##0,00_ ;[Red]-# ##0,00_ ;'-'")
        for header in rub_headers:
            self._set_format_by_header(ws, header_map, header, "# ##0 ₽")
        for header in date_headers:
            self._set_format_by_header(ws, header_map, header, "ДД.ММ.ГГ;@")
        for header in integer_headers:
            self._set_format_by_header(ws, header_map, header, '# ##0;[Red]-# ##0;"-"')
        for header in fx_headers:
            self._set_format_by_header(ws, header_map, header, "# ##0")

        for header in ("Supplier Product Name", "Our Product Name"):
            self._set_width_by_header(ws, header_map, header, 31.14)
        for header in ("Qty, pcs", "Volume, L"):
            self._set_width_by_header(ws, header_map, header, 10.50)
        self._set_width_by_header(ws, header_map, "last update", 11.00)
        self._set_width_by_header(ws, header_map, "last update (prev)", 11.00)
        self._set_width_by_header(ws, header_map, "Best Suppl", 16.14)
        self._set_width_by_header(ws, header_map, "Best Suppl 2", 16.14)
        self._set_width_by_header(ws, header_map, "last update Best1", 9.43)
        self._set_width_by_header(ws, header_map, "last update Best2", 9.86)
        for header in integer_headers:
            self._set_width_by_header(ws, header_map, header, 8.14)
        for header in fx_headers:
            self._set_width_by_header(ws, header_map, header, 7.29)
        for header in ("Currency", "Currency Best1", "Currency Best2"):
            self._set_width_by_header(ws, header_map, header, 8.14)

    def _calc_supplier_full_cost_from_db(self, supplier_id: int, product_id: int, supplier_price: object):
        if supplier_price is None:
            return None

        supplier = self.session.query(Supplier).filter(Supplier.id == supplier_id).first()
        if supplier is None:
            return None

        from app.services.supplier_service import SupplierService

        supplier_service = SupplierService(self.session)
        rate_to_rub = supplier_service.get_rate_to_rub(supplier.base_currency)
        if rate_to_rub is None or float(rate_to_rub) == 0:
            return None

        try:
            result = self.cost_calculation.calculate_supplier_costs(
                supplier_id=supplier_id,
                product_id=product_id,
                supplier_price=self._to_decimal(supplier_price),
                fx_rate=self._to_decimal(rate_to_rub),
                currency_code=supplier.base_currency,
            )
            return result.full_cost_msk
        except Exception:
            return None

    def _consider_best_candidate(self, cand_supplier_name, cand_full_cost, cand_date, best1, best2, cand_fx_rate=None, cand_currency=""):
        if not cand_supplier_name or cand_full_cost is None:
            return best1, best2

        candidate = {
            "supplier": cand_supplier_name,
            "price": cand_full_cost,
            "date": cand_date,
            "fx_rate": cand_fx_rate,
            "currency": cand_currency or "",
        }

        if best1["price"] is None:
            best1 = candidate
            return best1, best2

        if self._to_decimal(cand_full_cost) < self._to_decimal(best1["price"]):
            best2 = best1
            best1 = candidate
            return best1, best2

        if best2["price"] is None:
            best2 = candidate
            return best1, best2

        if self._to_decimal(cand_full_cost) < self._to_decimal(best2["price"]):
            best2 = candidate

        return best1, best2

    def _get_best_two_suppliers_for_export(
        self,
        current_supplier_id: int,
        product_id: int,
        current_supplier_name: str,
        current_imported_full_cost,
        current_imported_date,
    ):
        best1 = {"supplier": "", "price": None, "date": None, "fx_rate": None, "currency": ""}
        best2 = {"supplier": "", "price": None, "date": None, "fx_rate": None, "currency": ""}
        seen = set()

        current_supplier = self.session.query(Supplier).filter(Supplier.id == current_supplier_id).first()
        current_supplier_in_rating = bool(current_supplier.rating_calc) if current_supplier else False

        if current_supplier_in_rating:
            best1, best2 = self._consider_best_candidate(
                current_supplier_name,
                current_imported_full_cost,
                current_imported_date,
                best1,
                best2,
                self._get_supplier_currency_rate(current_supplier_id)[1],
                self._get_supplier_currency_rate(current_supplier_id)[0],
            )

        seen.add(current_supplier_id)

        current_rows = self.price_repository.get_suppliers_with_current_prices_for_product(
            product_id=product_id,
            only_rating_calc=True,
        )
        for current_row in current_rows:
            if current_row.supplier_id == current_supplier_id or current_row.supplier_id in seen:
                continue
            v_full_cost = self._calc_supplier_full_cost_from_db(
                current_row.supplier_id,
                product_id,
                current_row.price,
            )
            best1, best2 = self._consider_best_candidate(
                current_row.supplier_name,
                v_full_cost,
                current_row.price_date,
                best1,
                best2,
                self._get_supplier_currency_rate(current_row.supplier_id)[1],
                self._get_supplier_currency_rate(current_row.supplier_id)[0],
            )
            seen.add(current_row.supplier_id)

        history_rows = (
            self.session.query(PriceHistory, Supplier)
            .join(Supplier, Supplier.id == PriceHistory.supplier_id)
            .filter(
                PriceHistory.product_id == product_id,
                PriceHistory.supplier_id != current_supplier_id,
                PriceHistory.price.isnot(None),
                Supplier.rating_calc.is_(True),
            )
            .order_by(PriceHistory.supplier_id.asc(), PriceHistory.price_date.desc(), PriceHistory.id.desc())
            .all()
        )
        latest_history_by_supplier: dict[int, tuple[PriceHistory, Supplier]] = {}
        for history_row, supplier_row in history_rows:
            if history_row.supplier_id not in latest_history_by_supplier:
                latest_history_by_supplier[history_row.supplier_id] = (history_row, supplier_row)

        for supplier_key, data in latest_history_by_supplier.items():
            if supplier_key in seen:
                continue
            history_row, supplier_row = data
            v_full_cost = self._calc_supplier_full_cost_from_db(supplier_row.id, product_id, history_row.price)
            best1, best2 = self._consider_best_candidate(
                supplier_row.name,
                v_full_cost,
                history_row.price_date,
                best1,
                best2,
                self._get_supplier_currency_rate(supplier_row.id)[1],
                self._get_supplier_currency_rate(supplier_row.id)[0],
            )
            seen.add(supplier_key)

        return best1, best2

    def _get_previous_same_supplier_values(self, supplier_id: int, product_id: int, current_price_date):
        if not supplier_id or not product_id or current_price_date is None:
            return None, None, None, None

        prev_snapshot = self.price_repository.get_previous_supplier_price_snapshot(
            supplier_id=supplier_id,
            product_id=product_id,
            last_price_date=current_price_date,
        )
        if prev_snapshot is None:
            return None, None, None, None

        prev_price = prev_snapshot.price
        prev_full_cost = self._calc_supplier_full_cost_from_db(supplier_id, product_id, prev_price)
        prev_cost_novo = None
        if prev_full_cost is not None:
            try:
                supplier = self.session.query(Supplier).filter(Supplier.id == supplier_id).first()
                if supplier is not None:
                    from app.services.supplier_service import SupplierService

                    supplier_service = SupplierService(self.session)
                    rate_to_rub = supplier_service.get_rate_to_rub(supplier.base_currency)
                    if rate_to_rub is not None and float(rate_to_rub) != 0:
                        calc_result = self.cost_calculation.calculate_supplier_costs(
                            supplier_id=supplier_id,
                            product_id=product_id,
                            supplier_price=self._to_decimal(prev_price),
                            fx_rate=self._to_decimal(rate_to_rub),
                            currency_code=supplier.base_currency,
                        )
                        prev_cost_novo = calc_result.cost_novo_wvat
            except Exception:
                prev_cost_novo = None

        return prev_price, prev_cost_novo, prev_full_cost, prev_snapshot.price_date


    def _get_latest_avg_sales_month(self, product_id: int):
        if not product_id:
            return None

        row = (
            self.session.query(OrderPlanningCalculation)
            .filter(OrderPlanningCalculation.product_id == int(product_id))
            .order_by(
                OrderPlanningCalculation.period_to.desc(),
                OrderPlanningCalculation.period_from.desc(),
                OrderPlanningCalculation.id.desc(),
            )
            .first()
        )
        return row.avg_sales_month if row else None

    def _calc_order_liters(self, months: int | None, avg_sales_month: object, stock_month_base: object, pack: object):
        if months is None:
            return Decimal("0")

        avg = self._to_decimal(avg_sales_month)
        pack_value = self._to_decimal(pack)
        base_qty = self._to_decimal(stock_month_base)

        if avg <= 0 or pack_value <= 0:
            return Decimal("0")

        raw_liters = (Decimal(str(months)) * avg) - base_qty
        if raw_liters <= 0:
            return Decimal("0")

        packs = (raw_liters / pack_value).to_integral_value(rounding="ROUND_CEILING")
        return packs * pack_value

    def _calc_order_planning_export_values(
        self,
        *,
        product_id: int,
        stock,
        pack: object,
        quick_months: int | None,
        order_months: int | None,
    ) -> dict:
        avg_sales_month = self._get_latest_avg_sales_month(product_id)

        if avg_sales_month is None:
            return {
                "Ср.Продажи мес": None,
                "к Быстрому заказу, л": Decimal("0"),
                "к Заказу, л": Decimal("0"),
            }

        stock_qty = self._to_decimal(getattr(stock, "stock_qty", None))
        transit_qty = self._to_decimal(getattr(stock, "transit_qty", None))
        is_confirmed_order_qty = self._to_decimal(getattr(stock, "is_confirmed_order_qty", None))
        order_qty = self._to_decimal(getattr(stock, "order_qty", None))
        is_order_qty = self._to_decimal(getattr(stock, "is_order_qty", None))
        reserve_qty = self._to_decimal(getattr(stock, "reserve_qty", None))
        reserve_ecomm_qty = self._to_decimal(getattr(stock, "reserve_ecomm_qty", None))
        free_base = stock_qty

        free_st_tr = free_base + transit_qty + is_confirmed_order_qty
        free_plus_ord = free_base + transit_qty + order_qty + is_order_qty

        return {
            "Ср.Продажи мес": avg_sales_month,
            "к Быстрому заказу, л": self._calc_order_liters(quick_months, avg_sales_month, free_st_tr, pack),
            "к Заказу, л": self._calc_order_liters(order_months, avg_sales_month, free_plus_ord, pack),
        }


    def build_export_rows(self, batch_id: str, imported_by: str, quick_months: int | None = None, order_months: int | None = None) -> list[dict]:
        rows = (
            self.session.query(TempPriceImport, SupplierPriceCalculation, Product, ProductStock, Supplier)
            .outerjoin(
                SupplierPriceCalculation,
                (TempPriceImport.batch_id == SupplierPriceCalculation.batch_id)
                & (TempPriceImport.imported_by == SupplierPriceCalculation.imported_by)
                & (TempPriceImport.import_row_no == SupplierPriceCalculation.import_row_no),
            )
            .outerjoin(Product, TempPriceImport.selected_product_id == Product.id)
            .outerjoin(ProductStock, TempPriceImport.selected_product_id == ProductStock.product_id)
            .outerjoin(Supplier, TempPriceImport.supplier_id == Supplier.id)
            .filter(
                TempPriceImport.batch_id == batch_id,
                TempPriceImport.imported_by == imported_by,
            )
            .order_by(TempPriceImport.import_row_no.asc())
            .all()
        )

        out_rows: list[dict] = []

        for temp_row, calc_row, product, stock, supplier in rows:
            current_supplier_name = supplier.name if supplier else ""
            product_id_for_row = temp_row.selected_product_id or 0
            current_price_date = temp_row.import_date

            best1 = {"supplier": "", "price": None, "date": None, "fx_rate": None, "currency": ""}
            best2 = {"supplier": "", "price": None, "date": None, "fx_rate": None, "currency": ""}

            if product_id_for_row > 0:
                best1, best2 = self._get_best_two_suppliers_for_export(
                    current_supplier_id=temp_row.supplier_id,
                    product_id=product_id_for_row,
                    current_supplier_name=current_supplier_name,
                    current_imported_full_cost=calc_row.full_cost_msk if calc_row else None,
                    current_imported_date=calc_row.calc_date if calc_row else current_price_date,
                )

            pack_value = product.pack if product is not None else temp_row.new_pack
            price_per_l = temp_row.price
            price_pack_export = (
                temp_row.price_pack if temp_row.price_pack is not None else self._calc_pack_price(price_per_l, pack_value)
            )

            prev_price = None
            prev_cost_novo = None
            prev_full_cost = None
            prev_price_date = None
            if temp_row.supplier_id and product_id_for_row > 0:
                prev_price, prev_cost_novo, prev_full_cost, prev_price_date = self._get_previous_same_supplier_values(
                    supplier_id=temp_row.supplier_id,
                    product_id=product_id_for_row,
                    current_price_date=current_price_date,
                )

            transit_total = None
            if stock is not None:
                transit_total = self._to_decimal(stock.transit_qty) + self._to_decimal(stock.is_confirmed_order_qty)

            order_plan_values = self._calc_order_planning_export_values(
                product_id=product_id_for_row,
                stock=stock,
                pack=pack_value,
                quick_months=quick_months,
                order_months=order_months,
            )

            out_rows.append(
                {
                    "Supplier Article": temp_row.supplier_article or "",
                    "Supplier Product Name": temp_row.product_name or "",
                    "Our Product Name": product.name if product else "",
                    "Pack": self._excel_value(pack_value),
                    "Qty, pcs": self._excel_value(temp_row.qty_pcs),
                    "Volume, L": self._excel_value(temp_row.volume_l),
                    "Price, L": self._excel_value(price_per_l),
                    "Price, pack": self._excel_value(price_pack_export),
                    "Currency": calc_row.currency_code if calc_row else (supplier.base_currency if supplier else ""),
                    "FX rate": self._round_fx_rate(calc_row.fx_rate_used if calc_row else None),
                    "Cost Novo with VAT": self._excel_value(calc_row.cost_novo_wvat if calc_row else None),
                    "Full Cost Msk": self._excel_value(calc_row.full_cost_msk if calc_row else None),
                    "last update (prev)": prev_price_date,
                    "Price, L (prev)": self._excel_value(prev_price),
                    "Cost Novo with VAT (prev)": self._excel_value(prev_cost_novo),
                    "Full Cost Msk (prev)": self._excel_value(prev_full_cost),
                    "Дистр цена": self._excel_value(stock.distr_price if stock else None),
                    "Промо цена": self._excel_value(stock.promo_price if stock else None),
                    "curr LPC": self._excel_value(stock.lpc if stock else None),
                    "curr Landed cost": self._excel_value(stock.landed_cost if stock else None),
                    "Best Suppl": best1["supplier"],
                    "Best full Price, L": self._excel_value(best1["price"]),
                    "last update Best1": best1["date"],
                    "FX rate Best1": self._round_fx_rate(best1.get("fx_rate")),
                    "Currency Best1": best1.get("currency", ""),
                    "Best Suppl 2": best2["supplier"],
                    "Best full Price, L 2": self._excel_value(best2["price"]),
                    "last update Best2": best2["date"],
                    "FX rate Best2": self._round_fx_rate(best2.get("fx_rate")),
                    "Currency Best2": best2.get("currency", ""),
                    "Stock": self._excel_value(stock.stock_qty if stock else None),
                    "Transit": self._excel_value(transit_total),
                    "Purchase Order": self._excel_value(stock.order_qty if stock else None),
                    "Order IS": self._excel_value(stock.is_order_qty if stock else None),
                    "Stock IS": self._excel_value(stock.is_stock_qty if stock else None),
                    "Reserve cust": self._excel_value(stock.reserve_qty if stock else None),
                    "Reserve E-Comm": self._excel_value(getattr(stock, "reserve_ecomm_qty", 0) if stock else None),
                    "Damaged": self._excel_value(stock.markdown_qty if stock else None),
                    "Ср.Продажи мес": self._excel_value(order_plan_values["Ср.Продажи мес"]),
                    "к Быстрому заказу, л": self._excel_value(order_plan_values["к Быстрому заказу, л"]),
                    "к Заказу, л": self._excel_value(order_plan_values["к Заказу, л"]),
                }
            )

        return out_rows

    def export_template(self, file_path: str | Path) -> Path:
        file_path = Path(file_path)
        if file_path.suffix.lower() != ".xlsx":
            file_path = file_path.with_suffix(".xlsx")
        file_path.parent.mkdir(parents=True, exist_ok=True)

        excel = None
        wb = None
        try:
            target_path = file_path.resolve()

            # если файл уже существует, удаляем его заранее,
            # чтобы SaveAs не упирался в блокировку/перезапись
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
            ws.Name = "Sheet1"

            headers = ["Material number", "Material", "Price, L", "Price, pack", "Qty, pcs", "Volume, L"]
            for col_index, header in enumerate(headers, start=1):
                ws.Cells(1, col_index).Value = header

            self._apply_header_common(ws, len(headers))
            ws.Range("A1:F1").Interior.Color = 0xCDCDCD

            ws.Columns("A:A").ColumnWidth = 18
            ws.Columns("B:B").ColumnWidth = 31.14
            ws.Columns("C:F").ColumnWidth = 12

            self._set_number_format_safe(ws.Columns("A:A"), "@", "@")
            self._set_number_format_safe(ws.Columns("C:F"), "0.00", "0,00")

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


    def export_calculated(
        self,
        batch_id: str,
        imported_by: str,
        supplier_id: int,
        output_path: str | Path | None = None,
        source_file_path: str | Path | None = None,
        quick_order_months: int | None = None,
        safe_stock_months: int | None = None,
    ) -> Path:
        supplier = self.session.query(Supplier).filter(Supplier.id == supplier_id).first()
        supplier_name = self._safe_filename(supplier.name if supplier else "Supplier")

        if output_path is None:
            filename = f"CostCalc_{supplier_name}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
            output_path = Path(filename)
        else:
            output_path = Path(output_path)

        if output_path.suffix.lower() != ".xlsx":
            output_path = output_path.with_suffix(".xlsx")
        output_path.parent.mkdir(parents=True, exist_ok=True)

        rows = self.build_export_rows(batch_id=batch_id, imported_by=imported_by, quick_months=quick_order_months, order_months=safe_stock_months)

        excel = None
        wb = None
        try:
            excel = self._create_excel_app()
            wb = excel.Workbooks.Add()
            ws = wb.Worksheets(1)
            ws.Name = "Sheet1"

            headers = [
                "Supplier Article",
                "Supplier Product Name",
                "Our Product Name",
                "Pack",
                "Qty, pcs",
                "Volume, L",
                "Price, L",
                "Price, pack",
                "Currency",
                "FX rate",
                "Cost Novo with VAT",
                "Full Cost Msk",
                "last update (prev)",
                "Price, L (prev)",
                "Cost Novo with VAT (prev)",
                "Full Cost Msk (prev)",
                "Дистр цена",
                "Промо цена",
                "curr LPC",
                "curr Landed cost",
                "Best Suppl",
                "Best full Price, L",
                "last update Best1",
                "FX rate Best1",
                "Currency Best1",
                "Best Suppl 2",
                "Best full Price, L 2",
                "last update Best2",
                "FX rate Best2",
                "Currency Best2",
                "Stock",
                "Transit",
                "Purchase Order",
                "Order IS",
                "Stock IS",
                "Reserve cust",
                "Reserve E-Comm",
                "Damaged",
                "Ср.Продажи мес",
                f"к Быстрому заказу, л ({quick_order_months} м)" if quick_order_months is not None else "к Быстрому заказу, л",
                f"к Заказу, л ({safe_stock_months} м)" if safe_stock_months is not None else "к Заказу, л",
            ]

            for col_index, header in enumerate(headers, start=1):
                ws.Cells(1, col_index).Value = header

            row_num = 2
            for row in rows:
                qty_value, volume_value = self._calc_qty_volume_for_export(
                    row.get("Qty, pcs"),
                    row.get("Volume, L"),
                    row.get("Pack"),
                )
                row["Qty, pcs"] = self._excel_value(qty_value)
                row["Volume, L"] = self._excel_value(volume_value)

                for col_index, header in enumerate(headers, start=1):
                    value = self._row_value_by_export_header(row, header)
                    if self._is_order_plan_export_header(header):
                        # Для колонок заказа 0 — это значение, а не пустая ячейка.
                        # Формат Excel сам покажет ноль как "-".
                        ws.Cells(row_num, col_index).Value = self._excel_value(value)
                    else:
                        ws.Cells(row_num, col_index).Value = self._excel_value_or_blank(value)
                row_num += 1

            self._apply_header_common(ws, len(headers))

            # ===== Header colors: one-to-one with the Access export layout,
            # shifted for the inserted "(prev)" block =====
            ws.Range("A1:K1").Interior.Color = self._rgb(205, 205, 205)

            # prev block
            ws.Range("L1:O1").Interior.Color = self._rgb(166, 166, 166)

            # price / stock price block
            ws.Range("P1:S1").Interior.Color = self._rgb(192, 0, 0)
            ws.Range("P1:S1").Font.Color = self._rgb(255, 255, 255)

            # best supplier 1
            ws.Range("T1:V1").Interior.Color = self._rgb(0, 176, 240)

            # best supplier 2
            ws.Range("W1:Y1").Interior.Color = self._rgb(146, 208, 80)

            # stock / transit / purchase order
            ws.Range("Z1:AB1").Interior.Color = self._rgb(33, 92, 152)
            ws.Range("Z1:AB1").Font.Color = self._rgb(255, 255, 255)

            # order is / stock is
            ws.Range("AC1:AD1").Interior.Color = self._rgb(192, 0, 0)
            ws.Range("AC1:AD1").Font.Color = self._rgb(255, 255, 255)

            # reserve / damaged
            ws.Range("AE1:AG1").Interior.Color = self._rgb(33, 92, 152)
            ws.Range("AE1:AG1").Font.Color = self._rgb(255, 255, 255)

            # order planning
            ws.Range("AH1:AJ1").Interior.Color = self._rgb(160, 43, 147)
            ws.Range("AH1:AJ1").Font.Color = self._rgb(255, 255, 255)

            # ===== Number/date formats and widths =====
            # Важно: после добавления новых колонок нельзя форматировать по буквам A/B/C.
            # Форматы применяем по названию колонки, чтобы даты не превращались в числа при сдвиге.
            self._apply_calculated_export_formats_by_header(ws, headers)

            def _paint_header_block(first_header, last_header, color, font_color=None):
                if first_header in headers and last_header in headers:
                    c1 = self._excel_column_letter(headers.index(first_header) + 1)
                    c2 = self._excel_column_letter(headers.index(last_header) + 1)
                    rng = ws.Range(f"{c1}1:{c2}1")
                    rng.Interior.Color = color
                    if font_color is not None:
                        rng.Font.Color = font_color
                    else:
                        # Automatic font color. Старые фиксированные диапазоны могли
                        # поставить белый шрифт на сдвинутые колонки, поэтому всегда
                        # сбрасываем цвет для блоков, где белый не нужен.
                        rng.Font.ColorIndex = -4105

            _paint_header_block("Supplier Article", "Full Cost Msk", self._rgb(205, 205, 205))
            _paint_header_block("last update (prev)", "Full Cost Msk (prev)", self._rgb(166, 166, 166))
            _paint_header_block("Дистр цена", "curr Landed cost", self._rgb(192, 0, 0), self._rgb(255, 255, 255))
            _paint_header_block("Best Suppl", "Currency Best1", self._rgb(0, 176, 240))
            _paint_header_block("Best Suppl 2", "Currency Best2", self._rgb(146, 208, 80))
            _paint_header_block("Stock", "Purchase Order", self._rgb(33, 92, 152), self._rgb(255, 255, 255))
            _paint_header_block("Order IS", "Stock IS", self._rgb(192, 0, 0), self._rgb(255, 255, 255))
            _paint_header_block("Reserve cust", "Damaged", self._rgb(33, 92, 152), self._rgb(255, 255, 255))
            _paint_header_block("Ср.Продажи мес", headers[-1], self._rgb(160, 43, 147), self._rgb(255, 255, 255))

            for _src, _dst in [
                ("Currency", "FX rate"),
                ("last update Best1", "FX rate Best1"),
                ("last update Best1", "Currency Best1"),
                ("last update Best2", "FX rate Best2"),
                ("last update Best2", "Currency Best2"),
            ]:
                self._copy_header_style(ws, _src, _dst, headers)
            for _fx_header in ("FX rate", "FX rate Best1", "FX rate Best2"):
                self._format_fx_column(ws, _fx_header, headers)
            for _cur_header in ("Currency", "Currency Best1", "Currency Best2"):
                if _cur_header in headers:
                    ws.Columns(f"{self._excel_column_letter(headers.index(_cur_header)+1)}:{self._excel_column_letter(headers.index(_cur_header)+1)}").ColumnWidth = 8.14

            ws.Range(f"A1:{self._excel_column_letter(len(headers))}1").AutoFilter(1)

            try:
                ws.Activate()
                window = excel.ActiveWindow
                window.FreezePanes = False
                window.SplitRow = 1
                window.SplitColumn = 6  # freeze after column F, start visible moving part from G
                window.ScrollRow = 1
                window.ScrollColumn = 1
                window.Zoom = 85
                window.FreezePanes = True
                ws.Range("A1").Select()
                window.ScrollRow = 1
                window.ScrollColumn = 1
            except Exception:
                pass

            wb.SaveAs(str(output_path.resolve()))
            return output_path
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

    @staticmethod
    def _rgb(r: int, g: int, b: int) -> int:
        return r + g * 256 + b * 65536




# Backward-compatible alias.
SupplierPriceExport = SupplierPriceExporter
