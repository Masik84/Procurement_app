from __future__ import annotations

from dataclasses import dataclass
from typing import Optional

from sqlalchemy.orm import Session

from app.db.models import FixedCosts, MarkingRate, PackType, Product, Supplier


@dataclass(slots=True)
class CostCalculationResult:
    supplier_price: float
    cost_novo_wvat: float
    full_cost_msk: float

    currency_code: str
    fx_rate_used: float
    fx_markup_used: float
    transport_used: float
    reexport_used: float

    has_customs_used: bool
    via_novo_used: bool
    bank_fee_used: float
    customs_fee_used: float
    move_novo_used: float
    move_msk_used: float
    is_excise_used: bool
    additional_customs_used: float
    storage_used: float
    marking_used: float


class CostCalculationService:
    def __init__(self, session: Session) -> None:
        self.session = session

    def get_fixed_costs(self) -> FixedCosts:
        row = (
            self.session.query(FixedCosts)
            .order_by(FixedCosts.id.asc())
            .first()
        )
        if row is None:
            raise ValueError("В таблице fixed_costs нет данных.")
        return row

    def get_product(self, product_id: int) -> Product:
        row = (
            self.session.query(Product)
            .filter(Product.id == product_id)
            .first()
        )
        if row is None:
            raise ValueError(f"Product id={product_id} не найден.")
        return row

    def get_supplier(self, supplier_id: int) -> Supplier:
        row = (
            self.session.query(Supplier)
            .filter(Supplier.id == supplier_id)
            .first()
        )
        if row is None:
            raise ValueError(f"Supplier id={supplier_id} не найден.")
        return row

    def get_pack_type_by_volume(self, pack_value: object) -> Optional[PackType]:
        if pack_value is None:
            return None

        pack_num = float(pack_value)

        return (
            self.session.query(PackType)
            .filter(PackType.volume == pack_num)
            .first()
        )

    def get_marking_rate_by_pack_type(self, pack_type_name: str) -> Optional[MarkingRate]:
        return (
            self.session.query(MarkingRate)
            .filter(MarkingRate.pack_type == pack_type_name)
            .first()
        )

    def get_marking_cost(self, product_id: int) -> float:
        product = self.get_product(product_id)
        pack_type = self.get_pack_type_by_volume(product.pack)

        if pack_type is None:
            return 0.0

        marking_rate = self.get_marking_rate_by_pack_type(pack_type.name)
        if marking_rate is None:
            return 0.0

        return float(marking_rate.cost_per_l)

    def calc_cost_novo_wvat(
        self,
        *,
        supplier_price: float,
        product_id: int,
        supplier_id: int,
        transport: float,
        reexport: float,
        fx_rate: float,
        fx_markup: float,
        has_customs: bool,
    ) -> Optional[float]:
        if supplier_price is None or float(supplier_price) == 0:
            return None

        product = self.get_product(product_id)
        supplier = self.get_supplier(supplier_id)
        fixed = self.get_fixed_costs()

        d_price = float(supplier_price)
        d_transport = float(transport)
        d_reexport = float(reexport)
        d_fx_rate = float(fx_rate)
        d_fx_markup = float(fx_markup)

        d_customs_clearance = float(fixed.customs_clearance)
        d_additional_customs = float(fixed.additional_customs)
        d_excise = float(fixed.excise)
        d_eco_fee = float(fixed.eco_fee)
        d_vat = float(fixed.vat)
        d_customs_fee = float(fixed.customs_fee)
        d_bank_fee = float(fixed.bank_fee)

        if supplier.marks_for_us:
            d_marking = 0.0
        else:
            d_marking = self.get_marking_cost(product_id)

        is_excise = bool(product.is_excise)
        supplier_is_rf = bool(supplier.is_rf)

        customs_multiplier = 1.0 + d_customs_clearance if has_customs else 1.0

        if supplier_is_rf:
            base_before_add = (
                (d_price + d_transport)
                * (1.0 + d_reexport)
                * customs_multiplier
                * d_fx_rate
                * (1.0 + d_fx_markup)
            )
            base = base_before_add + d_marking
        else:
            base_before_add = (
                (d_price + d_transport)
                * (1.0 + d_reexport)
                * customs_multiplier
                * (1.0 + d_bank_fee)
                * d_fx_rate
                * (1.0 + d_fx_markup)
            )
            base = base_before_add + d_additional_customs + d_marking

        if not supplier_is_rf:
            base = (
                base
                + d_customs_fee
                + (d_excise if is_excise else 0.0)
                + d_eco_fee
            )

        return round(base * (1.0 + d_vat), 4)

    def calc_full_cost_msk(
        self,
        *,
        supplier_price: float,
        product_id: int,
        supplier_id: int,
        transport: float,
        reexport: float,
        fx_rate: float,
        fx_markup: float,
        has_customs: bool,
        via_novo: bool,
    ) -> Optional[float]:
        cost_novo = self.calc_cost_novo_wvat(
            supplier_price=supplier_price,
            product_id=product_id,
            supplier_id=supplier_id,
            transport=transport,
            reexport=reexport,
            fx_rate=fx_rate,
            fx_markup=fx_markup,
            has_customs=has_customs,
        )

        if cost_novo is None:
            return None

        supplier = self.get_supplier(supplier_id)
        fixed = self.get_fixed_costs()

        d_cost_novo = float(cost_novo)
        d_money = float(fixed.money)
        d_storage = float(fixed.storage)
        d_move_novo = float(fixed.move_novo_tamozh)
        d_move_msk = float(fixed.move_tamozh_chekhov)
        d_vat = float(fixed.vat)

        supplier_is_rf = bool(supplier.is_rf)

        logistics = d_storage

        if not supplier_is_rf:
            logistics += d_move_msk
            if via_novo:
                logistics += d_move_novo

        return round(d_cost_novo * (1.0 + d_money) + logistics * (1.0 + d_vat), 4)

    def calculate_supplier_costs(
        self,
        *,
        supplier_id: int,
        product_id: int,
        supplier_price: float,
        fx_rate: float,
        currency_code: str,
        transport: Optional[float] = None,
        reexport: Optional[float] = None,
        fx_markup: Optional[float] = None,
        has_customs: Optional[bool] = None,
        via_novo: Optional[bool] = None,
    ) -> CostCalculationResult:
        supplier = self.get_supplier(supplier_id)
        product = self.get_product(product_id)
        fixed = self.get_fixed_costs()

        transport_used = float(
            supplier.transport_cost_per_l if transport is None else transport
        )
        reexport_used = float(
            supplier.reexport_percent if reexport is None else reexport
        )
        fx_markup_used = float(
            supplier.fx_rate_markup if fx_markup is None else fx_markup
        )
        has_customs_used = bool(
            supplier.has_import_duty if has_customs is None else has_customs
        )
        via_novo_used = bool(
            supplier.is_via_novo if via_novo is None else via_novo
        )

        cost_novo = self.calc_cost_novo_wvat(
            supplier_price=float(supplier_price),
            product_id=product_id,
            supplier_id=supplier_id,
            transport=transport_used,
            reexport=reexport_used,
            fx_rate=float(fx_rate),
            fx_markup=fx_markup_used,
            has_customs=has_customs_used,
        )
        if cost_novo is None:
            raise ValueError("Не удалось рассчитать cost_novo_wvat.")

        full_cost = self.calc_full_cost_msk(
            supplier_price=float(supplier_price),
            product_id=product_id,
            supplier_id=supplier_id,
            transport=transport_used,
            reexport=reexport_used,
            fx_rate=float(fx_rate),
            fx_markup=fx_markup_used,
            has_customs=has_customs_used,
            via_novo=via_novo_used,
        )
        if full_cost is None:
            raise ValueError("Не удалось рассчитать full_cost_msk.")

        marking_used = 0.0 if supplier.marks_for_us else self.get_marking_cost(product_id)

        customs_fee_used = 0.0 if supplier.is_rf else float(fixed.customs_fee)
        move_novo_used = 0.0 if supplier.is_rf else float(fixed.move_novo_tamozh)
        move_msk_used = 0.0 if supplier.is_rf else float(fixed.move_tamozh_chekhov)
        is_excise_used = False if supplier.is_rf else bool(product.is_excise)

        return CostCalculationResult(
            supplier_price=float(supplier_price),
            cost_novo_wvat=float(cost_novo),
            full_cost_msk=float(full_cost),
            currency_code=currency_code,
            fx_rate_used=float(fx_rate),
            fx_markup_used=fx_markup_used,
            transport_used=transport_used,
            reexport_used=reexport_used,
            has_customs_used=has_customs_used,
            via_novo_used=via_novo_used,
            bank_fee_used=float(fixed.bank_fee),
            customs_fee_used=customs_fee_used,
            move_novo_used=move_novo_used,
            move_msk_used=move_msk_used,
            is_excise_used=is_excise_used,
            additional_customs_used=float(fixed.additional_customs),
            storage_used=float(fixed.storage),
            marking_used=float(marking_used),
        )