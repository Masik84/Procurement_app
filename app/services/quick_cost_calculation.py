from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal, ROUND_HALF_UP

from sqlalchemy.orm import Session

from app.db.models import FixedCosts, MarkingRate, Supplier


@dataclass(slots=True)
class QuickCostCalculationResult:
    supplier_price: Decimal
    cost_novo_wvat: Decimal
    full_cost_msk: Decimal

    fx_rate_used: Decimal
    transport_used: Decimal
    reexport_used: Decimal
    fx_markup_used: Decimal

    has_customs_used: bool
    via_novo_used: bool
    supplier_is_rf_used: bool
    is_excise_used: bool
    marking_used: Decimal


class QuickCostCalculationService:
    def __init__(self, session: Session) -> None:
        self.session = session

    @staticmethod
    def _to_decimal(value: object) -> Decimal:
        if value is None:
            return Decimal("0")
        if isinstance(value, Decimal):
            return value
        return Decimal(str(value))

    @staticmethod
    def _round4(value: Decimal) -> Decimal:
        return value.quantize(Decimal("0.0001"), rounding=ROUND_HALF_UP)

    def get_fixed_costs(self) -> FixedCosts:
        row = (
            self.session.query(FixedCosts)
            .order_by(FixedCosts.id.asc())
            .first()
        )
        if row is None:
            raise ValueError("В таблице fixed_costs нет данных.")
        return row

    def get_supplier(self, supplier_id: int) -> Supplier:
        row = (
            self.session.query(Supplier)
            .filter(Supplier.id == supplier_id)
            .first()
        )
        if row is None:
            raise ValueError(f"Поставщик id={supplier_id} не найден.")
        return row

    def get_marking_cost_by_pack_type(self, pack_type_name: str) -> Decimal:
        if not pack_type_name:
            return Decimal("0")

        row = (
            self.session.query(MarkingRate)
            .filter(MarkingRate.pack_type == pack_type_name)
            .first()
        )
        if row is None:
            return Decimal("0")

        return self._to_decimal(row.cost_per_l)

    def calculate(
        self,
        *,
        supplier_price: Decimal,
        supplier_id: int,
        pack_type_name: str,
        fx_rate: Decimal,
        transport: Decimal,
        reexport: Decimal,
        fx_markup: Decimal,
        has_customs: bool,
        via_novo: bool,
        supplier_is_rf: bool,
        marks_for_us: bool,
        is_excise: bool,
    ) -> QuickCostCalculationResult:
        if self._to_decimal(supplier_price) == Decimal("0"):
            raise ValueError("Поле 'Цена поставщика' не может быть равно 0.")

        supplier = self.get_supplier(supplier_id)
        fixed = self.get_fixed_costs()

        d_price = self._to_decimal(supplier_price)
        d_transport = self._to_decimal(transport)
        d_reexport = self._to_decimal(reexport)
        d_fx_rate = self._to_decimal(fx_rate)
        d_fx_markup = self._to_decimal(fx_markup)

        if d_fx_rate == Decimal("0"):
            raise ValueError("Поле 'Курс' не может быть равно 0.")

        d_customs_clearance = self._to_decimal(fixed.customs_clearance)
        d_additional_customs = self._to_decimal(fixed.additional_customs)
        d_excise = self._to_decimal(fixed.excise)
        d_eco_fee = self._to_decimal(fixed.eco_fee)
        d_vat = self._to_decimal(fixed.vat)
        d_customs_fee = self._to_decimal(fixed.customs_fee)
        d_bank_fee = self._to_decimal(fixed.bank_fee)
        d_money = self._to_decimal(fixed.money)
        d_storage = self._to_decimal(fixed.storage)
        d_move_novo = self._to_decimal(fixed.move_novo_tamozh)
        d_move_msk = self._to_decimal(fixed.move_tamozh_chekhov)

        if marks_for_us:
            d_marking = Decimal("0")
        else:
            d_marking = self.get_marking_cost_by_pack_type(pack_type_name)

        customs_multiplier = Decimal("1") + d_customs_clearance if has_customs else Decimal("1")

        if supplier_is_rf:
            base_before_add = (
                (d_price + d_transport)
                * (Decimal("1") + d_reexport)
                * customs_multiplier
                * d_fx_rate
                * (Decimal("1") + d_fx_markup)
            )
            base = base_before_add + d_marking
        else:
            base_before_add = (
                (d_price + d_transport)
                * (Decimal("1") + d_reexport)
                * customs_multiplier
                * (Decimal("1") + d_bank_fee)
                * d_fx_rate
                * (Decimal("1") + d_fx_markup)
            )
            base = base_before_add + d_additional_customs + d_marking

        if not supplier_is_rf:
            base = (
                base
                + d_customs_fee
                + (d_excise if is_excise else Decimal("0"))
                + d_eco_fee
            )

        cost_novo_wvat = self._round4(base * (Decimal("1") + d_vat))

        logistics = d_storage
        if not supplier_is_rf:
            logistics += d_move_msk
            if via_novo:
                logistics += d_move_novo

        full_cost_msk = self._round4(
            cost_novo_wvat * (Decimal("1") + d_money)
            + logistics * (Decimal("1") + d_vat)
        )

        return QuickCostCalculationResult(
            supplier_price=d_price,
            cost_novo_wvat=cost_novo_wvat,
            full_cost_msk=full_cost_msk,
            fx_rate_used=d_fx_rate,
            transport_used=d_transport,
            reexport_used=d_reexport,
            fx_markup_used=d_fx_markup,
            has_customs_used=has_customs,
            via_novo_used=via_novo,
            supplier_is_rf_used=supplier_is_rf,
            is_excise_used=is_excise,
            marking_used=d_marking,
        )
