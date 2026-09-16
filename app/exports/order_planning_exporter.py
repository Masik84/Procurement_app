from __future__ import annotations

import logging
from decimal import Decimal
from pathlib import Path
from typing import Any

import pythoncom
import win32com.client as win32
from sqlalchemy.orm import Session

from app.db.models import CurrentSupplierPrice, PriceHistory, ProductStock, Supplier
from app.exports.excel_column_format import apply_standard_worksheet_format, excel_value_by_header
from app.services.cost_calculation_service import CostCalculationService
from app.services.price_repository import PriceRepository
from app.services.product_uc3_service import ProductUc3Service
from app.services.supplier_currency_cost_service import SupplierCurrencyCostService
from app.services.supplier_service import SupplierService
from app.utils.excel_fast_writer import write_excel_table
from app.utils.excel_format_rules import save_workbook_xlsx
from app.utils.money import round4, to_decimal
from app.utils.output_headers import standardize_output_header

logger = logging.getLogger(__name__)


class OrderPlanningExporter:
    def __init__(self, session: Session):
        self.session = session
        self.cost_calculation = CostCalculationService(session)
        self.supplier_service = SupplierService(session)
        self.currency_cost_service = SupplierCurrencyCostService(
            session,
            cost_calculation=self.cost_calculation,
        )
        self.price_repository = PriceRepository(session)
        self.uc3_service = ProductUc3Service(session)
        self._stocks_by_product: dict[int, ProductStock] = {}
        self._current_uc3_by_product = {}
        self._vat = Decimal("0")

    _to_decimal = staticmethod(to_decimal)

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

    def _calc_uc3_from_full_cost(self, *, product_id: int, full_cost: object) -> Decimal | None:
        stock = self._stocks_by_product.get(int(product_id))
        sale_price = ProductUc3Service.sales_reference_price(stock)
        if sale_price is None or full_cost is None:
            return None
        try:
            denominator = Decimal("1") + self._to_decimal(self._vat)
            if denominator == 0:
                return None
            return round4((sale_price - self._to_decimal(full_cost)) / denominator)
        except Exception:
            return None

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
                "uc3": self._calc_uc3_from_full_cost(
                    product_id=product_id,
                    full_cost=calc.full_cost_msk,
                ),
                "date": price_date,
                "fx_rate": calc.fx_rate_used,
                "currency": calc.currency_code,
            }
        except Exception:
            logger.exception(
                "Не удалось рассчитать вариант поставщика %s для product_id=%s",
                getattr(supplier, "name", ""),
                product_id,
            )
            return None

    def _get_supplier_options(self, product_id: int, min_price_date=None) -> list[dict]:
        bulk_prices = getattr(self, "_bulk_prices_by_product", None)
        if bulk_prices is not None:
            options: list[dict] = []
            for snapshot in bulk_prices.get(int(product_id), []):
                supplier = self.cost_calculation.get_supplier(snapshot.supplier_id)
                option = self._calc_supplier_option(
                    supplier,
                    product_id,
                    snapshot.price,
                    snapshot.price_date,
                    snapshot.currency_code,
                )
                if option and option["full_cost"] is not None:
                    options.append(option)
            options.sort(key=lambda x: (self._to_decimal(x["full_cost"]), str(x["supplier"]).lower()))
            return options

        current_query = self.session.query(CurrentSupplierPrice.supplier_id).filter(
            CurrentSupplierPrice.product_id == product_id,
            CurrentSupplierPrice.price.isnot(None),
        )
        history_query = self.session.query(PriceHistory.supplier_id).filter(
            PriceHistory.product_id == product_id,
            PriceHistory.price.isnot(None),
        )
        if min_price_date is not None:
            current_query = current_query.filter(CurrentSupplierPrice.last_update >= min_price_date)
            history_query = history_query.filter(PriceHistory.price_date >= min_price_date)

        supplier_ids = {row[0] for row in current_query.all()}
        supplier_ids.update(row[0] for row in history_query.all())

        options: list[dict] = []
        for supplier_id in supplier_ids:
            supplier = self.session.query(Supplier).filter(Supplier.id == supplier_id).first()
            if not supplier or not bool(getattr(supplier, "rating_calc", True)):
                continue

            snapshot = self.price_repository.get_last_supplier_price_snapshot(
                supplier_id=supplier_id,
                product_id=product_id,
                min_price_date=min_price_date,
            )
            if snapshot is None:
                continue

            option = self._calc_supplier_option(
                supplier,
                product_id,
                snapshot.price,
                snapshot.price_date,
                snapshot.currency_code,
            )
            if option and option["full_cost"] is not None:
                options.append(option)

        options.sort(key=lambda x: (self._to_decimal(x["full_cost"]), str(x["supplier"]).lower()))
        return options

    @staticmethod
    def _base_headers() -> list[str]:
        # Quick-order columns remain in GUI/calculation, but are intentionally
        # excluded from the Order Planning Excel export.
        return [
            "Brand",
            "Product Name",
            "Pack",
            "Категория ABC",
            "Ср.Продажи мес",
            "Safe Stock (st), mnth",
            "Safe Stock (st+tr), mnth",
            "Safe Stock (+ord), mnth",
            "к Заказу, шт",
            "к Заказу, л",
            "Дистр цена",
            "Промо цена",
            "curr LPC",
            "curr Landed cost",
            "Target uC3",
            "Walk-Away uC3",
            "Stock",
            "Transit",
            "Purchase Order",
            "Order IS",
            "Stock IS",
            "Reserve cust",
            "Reserve E-Comm",
            "Damaged",
        ]

    def build_export_data(
        self,
        display_rows: list[dict],
        supplier_price_age_months: int = 3,
    ) -> tuple[list[str], list[list[object]]]:
        min_price_date = PriceRepository.supplier_price_cutoff_from_months(supplier_price_age_months)
        product_ids = {
            int(row["product_id"])
            for row in display_rows
            if row.get("product_id")
        }

        self._stocks_by_product = {
            int(stock.product_id): stock
            for stock in (
                self.session.query(ProductStock).filter(ProductStock.product_id.in_(product_ids)).all()
                if product_ids else []
            )
        }
        self._current_uc3_by_product = self.uc3_service.get_current_map(product_ids)
        try:
            fixed_costs = self.cost_calculation.get_fixed_costs()
            self._vat = self._to_decimal(getattr(fixed_costs, "vat", None))
        except Exception:
            self._vat = Decimal("0")

        self._bulk_prices_by_product = self.price_repository.get_supplier_prices_for_products(
            product_ids,
            only_rating_calc=True,
            min_price_date=min_price_date,
            exclude_manual=False,
        )
        supplier_ids = {
            price.supplier_id
            for prices in self._bulk_prices_by_product.values()
            for price in prices
        }
        self.currency_cost_service.preload_reference_data(
            product_ids=product_ids,
            supplier_ids=supplier_ids,
        )

        max_suppliers = 0
        prepared = []
        for row in display_rows:
            product_id = row.get("product_id")
            options = self._get_supplier_options(int(product_id), min_price_date=min_price_date) if product_id else []
            max_suppliers = max(max_suppliers, len(options))
            prepared.append((row, options))

        headers = self._base_headers()
        for idx in range(1, max_suppliers + 1):
            headers.extend([
                f"Supplier_{idx}",
                f"Cost Novo with VAT_{idx}",
                f"Full Cost Msk_{idx}",
                f"uC3_{idx}",
                f"last update_{idx}",
                f"FX rate_{idx}",
                f"Currency_{idx}",
            ])

        rows: list[list[object]] = []
        for row, options in prepared:
            product_id = int(row["product_id"]) if row.get("product_id") else None
            stock = self._stocks_by_product.get(product_id) if product_id else None
            current_uc3 = self._current_uc3_by_product.get(product_id) if product_id else None

            values = [
                row.get("brand", ""),
                row.get("product_name", ""),
                row.get("pack"),
                row.get("abc_category") or "-",
                row.get("avg_sales_month"),
                row.get("safe_stock_st_month"),
                row.get("safe_stock_st_tr_month"),
                row.get("safe_stock_ord_month"),
                row.get("std_order_pcs"),
                row.get("std_order_l"),
                row.get("distr_price"),
                row.get("promo_price"),
                getattr(stock, "lpc", None) if stock else None,
                getattr(stock, "landed_cost", None) if stock else None,
                getattr(current_uc3, "target_uc3", None) if current_uc3 else None,
                getattr(current_uc3, "walk_away_uc3", None) if current_uc3 else None,
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
                values.extend([
                    option["supplier"],
                    option["cost_novo"],
                    option["full_cost"],
                    option["uc3"],
                    option["date"],
                    option["fx_rate"],
                    option["currency"],
                ])
            while len(values) < len(headers):
                values.append("")
            rows.append(values)
        return headers, rows

    def export_report(
        self,
        *,
        display_rows: list[dict],
        output_path: str | Path,
        supplier_price_age_months: int = 3,
    ) -> Path:
        output_path = Path(output_path)
        if output_path.suffix.lower() != ".xlsx":
            output_path = output_path.with_suffix(".xlsx")
        output_path.parent.mkdir(parents=True, exist_ok=True)
        target_path = output_path.resolve()
        if target_path.exists():
            try:
                target_path.unlink()
            except PermissionError:
                raise PermissionError(
                    f"Не удается перезаписать файл:\n{target_path}\n\n"
                    "Скорее всего, он открыт в Excel. Закрой файл и попробуй снова."
                )

        headers, rows = self.build_export_data(
            display_rows,
            supplier_price_age_months=supplier_price_age_months,
        )
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

            # Freeze through the standard order columns; current cost/target
            # columns and supplier blocks remain scrollable to the right.
            apply_standard_worksheet_format(ws, headers, freeze_cell="K2", zoom=85)

            save_workbook_xlsx(wb, target_path)
            return target_path
        finally:
            try:
                if wb is not None:
                    wb.Close(SaveChanges=False)
            except Exception:
                logger.exception("Подавленная ошибка (см. traceback выше)")
            try:
                if excel is not None:
                    excel.Quit()
            except Exception:
                logger.exception("Подавленная ошибка (см. traceback выше)")
            pythoncom.CoUninitialize()
