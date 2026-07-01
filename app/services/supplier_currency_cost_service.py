from __future__ import annotations

from decimal import Decimal, ROUND_HALF_UP

from sqlalchemy.orm import Session

from app.services.cost_calculation_service import CostCalculationResult, CostCalculationService
from app.services.supplier_service import SupplierService
from app.utils.text import clean_multi_spaces


class SupplierCurrencyCostService:
    """Calculate supplier costs when a saved price currency differs from the supplier's current currency.

    Supplier absolute costs (`transport_cost_per_l`, `agent_fee`) are stored on the supplier card
    in the supplier's current base currency. Historical price rows keep their own currency in
    PriceHistory/CurrentSupplierPrice. When these currencies differ, transport and agent fee are
    converted from the supplier base currency to the saved price currency before cost calculation.
    """

    def __init__(self, session: Session) -> None:
        self.session = session
        self.cost_calculation = CostCalculationService(session)
        self.supplier_service = SupplierService(session)

    @staticmethod
    def _to_decimal(value: object) -> Decimal:
        if value is None or value == "":
            return Decimal("0")
        if isinstance(value, Decimal):
            return value
        return Decimal(str(value).replace("\u00a0", "").replace(" ", "").replace(",", "."))

    @staticmethod
    def _round4(value: Decimal) -> Decimal:
        return value.quantize(Decimal("0.0001"), rounding=ROUND_HALF_UP)

    @staticmethod
    def _currency(value: object) -> str:
        return clean_multi_spaces(value).upper()

    def get_rate_to_rub(self, currency_code: object) -> Decimal:
        code = self._currency(currency_code)
        if not code:
            raise ValueError("Не указана валюта для расчета себестоимости.")

        rate = self.supplier_service.get_rate_to_rub(code)
        rate_value = self._to_decimal(rate)
        if rate_value == 0:
            raise ValueError(f"Для валюты '{code}' не найден корректный курс rate_to_rub.")
        return rate_value

    def convert_amount(self, amount: object, from_currency: object, to_currency: object) -> Decimal:
        value = self._to_decimal(amount)
        source = self._currency(from_currency)
        target = self._currency(to_currency)

        if value == 0 or not source or not target or source == target:
            return value

        source_rate = self.get_rate_to_rub(source)
        target_rate = self.get_rate_to_rub(target)
        return self._round4(value * source_rate / target_rate)

    def calculate_costs_for_price_record(
        self,
        *,
        supplier_id: int,
        product_id: int,
        supplier_price: object,
        price_currency_code: object | None,
    ) -> CostCalculationResult:
        supplier = self.cost_calculation.get_supplier(supplier_id)

        supplier_currency = self._currency(supplier.base_currency)
        calc_currency = self._currency(price_currency_code) or supplier_currency
        if not calc_currency:
            raise ValueError(f"Для поставщика '{supplier.name}' не указана валюта.")

        fx_rate = self.get_rate_to_rub(calc_currency)

        # Supplier card absolute costs are entered in the supplier's current base currency.
        # If a historical price is saved in another currency, calculate all per-currency
        # input components in that saved price currency.
        transport = self._to_decimal(getattr(supplier, "transport_cost_per_l", None))
        agent_fee = self._to_decimal(getattr(supplier, "agent_fee", None))
        if supplier_currency and supplier_currency != calc_currency:
            transport = self.convert_amount(transport, supplier_currency, calc_currency)
            agent_fee = self.convert_amount(agent_fee, supplier_currency, calc_currency)

        return self.cost_calculation.calculate_supplier_costs(
            supplier_id=supplier_id,
            product_id=product_id,
            supplier_price=self._to_decimal(supplier_price),
            fx_rate=fx_rate,
            currency_code=calc_currency,
            transport=transport,
            agent_fee=agent_fee,
        )
