from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable, Optional
from decimal import Decimal, ROUND_HALF_UP

from sqlalchemy.orm import Session

from app.db.models import FixedCosts, MarkingRate, PackType, Product, Supplier


@dataclass(slots=True)
class CostCalculationResult:
    supplier_price: Decimal
    cost_novo_wvat: Decimal
    full_cost_msk: Decimal

    currency_code: str
    fx_rate_used: Decimal
    fx_markup_used: Decimal
    fx_markup_abs_used: Decimal
    transport_used: Decimal
    reexport_used: Decimal
    insurance_used: Decimal
    agent_fee_used: Decimal

    has_customs_used: bool
    via_novo_used: bool
    bank_fee_used: Decimal
    customs_fee_used: Decimal
    move_novo_used: Decimal
    move_msk_used: Decimal
    is_excise_used: bool
    additional_customs_used: Decimal
    storage_used: Decimal
    marking_used: Decimal


class CostCalculationService:
    def __init__(self, session: Session) -> None:
        self.session = session
        # One service instance is used for a whole import/save operation.  The
        # calculation below asks for the same reference rows several times, so
        # keep them in memory for the lifetime of that operation instead of
        # issuing repeated round-trips to PostgreSQL for every price row.
        self._fixed_costs_cache: Optional[FixedCosts] = None
        self._products_cache: dict[int, Product] = {}
        self._suppliers_cache: dict[int, Supplier] = {}
        self._pack_types_cache: dict[float, Optional[PackType]] = {}
        self._marking_rates_cache: dict[str, Optional[MarkingRate]] = {}
        self._pack_types_loaded = False
        self._marking_rates_loaded = False

    def preload_reference_data(
        self,
        *,
        product_ids: Iterable[int] = (),
        supplier_ids: Iterable[int] = (),
    ) -> None:
        """Warm calculation reference caches in a fixed number of queries."""
        missing_products = {
            int(product_id) for product_id in product_ids
            if product_id is not None and int(product_id) not in self._products_cache
        }
        if missing_products:
            products = self.session.query(Product).filter(Product.id.in_(missing_products)).all()
            self._products_cache.update({int(product.id): product for product in products})

        missing_suppliers = {
            int(supplier_id) for supplier_id in supplier_ids
            if supplier_id is not None and int(supplier_id) not in self._suppliers_cache
        }
        if missing_suppliers:
            suppliers = self.session.query(Supplier).filter(Supplier.id.in_(missing_suppliers)).all()
            self._suppliers_cache.update({int(supplier.id): supplier for supplier in suppliers})

        if self._fixed_costs_cache is None:
            self.get_fixed_costs()

        if not self._pack_types_loaded:
            for pack_type in self.session.query(PackType).all():
                if pack_type.volume is not None:
                    self._pack_types_cache[float(pack_type.volume)] = pack_type
            self._pack_types_loaded = True

        if not self._marking_rates_loaded:
            for marking_rate in self.session.query(MarkingRate).all():
                self._marking_rates_cache[str(marking_rate.pack_type)] = marking_rate
            self._marking_rates_loaded = True

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
        if self._fixed_costs_cache is not None:
            return self._fixed_costs_cache

        row = (
            self.session.query(FixedCosts)
            .order_by(FixedCosts.id.asc())
            .first()
        )
        if row is None:
            raise ValueError("В таблице fixed_costs нет данных.")
        self._fixed_costs_cache = row
        return row

    def get_product(self, product_id: int) -> Product:
        product_id = int(product_id)
        cached = self._products_cache.get(product_id)
        if cached is not None:
            return cached

        row = (
            self.session.query(Product)
            .filter(Product.id == product_id)
            .first()
        )
        if row is None:
            raise ValueError(f"Product id={product_id} не найден.")
        self._products_cache[product_id] = row
        return row

    def get_supplier(self, supplier_id: int) -> Supplier:
        supplier_id = int(supplier_id)
        cached = self._suppliers_cache.get(supplier_id)
        if cached is not None:
            return cached

        row = (
            self.session.query(Supplier)
            .filter(Supplier.id == supplier_id)
            .first()
        )
        if row is None:
            raise ValueError(f"Supplier id={supplier_id} не найден.")
        self._suppliers_cache[supplier_id] = row
        return row

    def get_pack_type_by_volume(self, pack_value: object) -> Optional[PackType]:
        if pack_value is None:
            return None

        pack_num = float(pack_value)
        if pack_num in self._pack_types_cache:
            return self._pack_types_cache[pack_num]
        if self._pack_types_loaded:
            return None

        row = (
            self.session.query(PackType)
            .filter(PackType.volume == pack_num)
            .first()
        )
        self._pack_types_cache[pack_num] = row
        return row

    def get_marking_rate_by_pack_type(self, pack_type_name: str) -> Optional[MarkingRate]:
        if pack_type_name in self._marking_rates_cache:
            return self._marking_rates_cache[pack_type_name]
        if self._marking_rates_loaded:
            return None

        row = (
            self.session.query(MarkingRate)
            .filter(MarkingRate.pack_type == pack_type_name)
            .first()
        )
        self._marking_rates_cache[pack_type_name] = row
        return row

    def get_marking_cost(self, product_id: int) -> Decimal:
        product = self.get_product(product_id)
        pack_type = self.get_pack_type_by_volume(product.pack)

        if pack_type is None:
            return Decimal("0")

        marking_rate = self.get_marking_rate_by_pack_type(pack_type.name)
        if marking_rate is None:
            return Decimal("0")

        return self._to_decimal(marking_rate.cost_per_l)

    def calc_cost_novo_wvat(
        self,
        *,
        supplier_price: Decimal,
        product_id: int,
        supplier_id: int,
        transport: Decimal,
        reexport: Decimal,
        insurance: Decimal,
        fx_rate: Decimal,
        fx_markup: Decimal,
        fx_markup_abs: Decimal,
        has_customs: bool,
        agent_fee: Optional[Decimal] = None,
    ) -> Optional[float]:
        if supplier_price is None or self._to_decimal(supplier_price) == Decimal("0"):
            return None

        product = self.get_product(product_id)
        supplier = self.get_supplier(supplier_id)
        fixed = self.get_fixed_costs()

        d_price = self._to_decimal(supplier_price)
        d_transport = self._to_decimal(transport)
        d_reexport = self._to_decimal(reexport)
        d_insurance = self._to_decimal(insurance)
        d_fx_rate = self._to_decimal(fx_rate)
        d_fx_markup = self._to_decimal(fx_markup)
        d_fx_markup_abs = self._to_decimal(fx_markup_abs)
        d_effective_fx_rate = d_fx_rate * (Decimal("1") + d_fx_markup) + d_fx_markup_abs

        d_customs_clearance = self._to_decimal(fixed.customs_clearance)
        d_additional_customs = self._to_decimal(fixed.additional_customs)
        d_excise = self._to_decimal(fixed.excise)
        d_eco_fee = self._to_decimal(fixed.eco_fee)
        d_vat = self._to_decimal(fixed.vat)
        d_customs_fee = self._to_decimal(fixed.customs_fee)
        d_bank_fee = self._to_decimal(fixed.bank_fee)
        d_agent_fee = self._to_decimal(
            getattr(supplier, "agent_fee", None) if agent_fee is None else agent_fee
        )

        if supplier.marks_for_us:
            d_marking = Decimal("0")
        else:
            d_marking = self.get_marking_cost(product_id)

        is_excise = bool(product.is_excise)
        supplier_is_rf = bool(supplier.is_rf)

        customs_multiplier = Decimal("1") + d_customs_clearance if has_customs else Decimal("1")
        customs_and_insurance_multiplier = customs_multiplier + d_insurance

        if supplier_is_rf:
            base_before_add = (
                (d_price + d_transport)
                * (Decimal("1") + d_reexport)
                * customs_and_insurance_multiplier
                * d_effective_fx_rate
            )
            base = base_before_add + d_marking + (d_agent_fee * d_fx_rate)
        else:
            base_before_add = (
                (d_price + d_transport)
                * (Decimal("1") + d_reexport)
                * customs_and_insurance_multiplier
                * (Decimal("1") + d_bank_fee)
                * d_effective_fx_rate
            )
            base = base_before_add + d_additional_customs + d_marking + (d_agent_fee * d_fx_rate)

        if not supplier_is_rf:
            base = (
                base
                + d_customs_fee
                + (d_excise if is_excise else Decimal("0"))
                + d_eco_fee
            )

        return self._round4(base * (Decimal("1") + d_vat))

    def calc_full_cost_msk(
        self,
        *,
        supplier_price: Decimal,
        product_id: int,
        supplier_id: int,
        transport: Decimal,
        reexport: Decimal,
        insurance: Decimal,
        fx_rate: Decimal,
        fx_markup: Decimal,
        fx_markup_abs: Decimal,
        has_customs: bool,
        via_novo: bool,
        agent_fee: Optional[Decimal] = None,
    ) -> Optional[float]:
        cost_novo = self.calc_cost_novo_wvat(
            supplier_price=supplier_price,
            product_id=product_id,
            supplier_id=supplier_id,
            transport=transport,
            reexport=reexport,
            insurance=insurance,
            fx_rate=fx_rate,
            fx_markup=fx_markup,
            fx_markup_abs=fx_markup_abs,
            has_customs=has_customs,
            agent_fee=agent_fee,
        )

        if cost_novo is None:
            return None

        supplier = self.get_supplier(supplier_id)
        fixed = self.get_fixed_costs()

        d_cost_novo = self._to_decimal(cost_novo)
        d_money = self._to_decimal(fixed.money)
        d_storage = self._to_decimal(fixed.storage)
        d_move_novo = self._to_decimal(fixed.move_novo_tamozh)
        d_move_msk = self._to_decimal(fixed.move_tamozh_chekhov)
        d_vat = self._to_decimal(fixed.vat)

        supplier_is_rf = bool(supplier.is_rf)

        logistics = d_storage

        if not supplier_is_rf:
            logistics += d_move_msk
            if via_novo:
                logistics += d_move_novo

        return self._round4(
            d_cost_novo * (Decimal("1") + d_money) +
            logistics * (Decimal("1") + d_vat)
        )

    def calculate_supplier_costs(
        self,
        *,
        supplier_id: int,
        product_id: int,
        supplier_price: Decimal,
        fx_rate: Decimal,
        currency_code: str,
        transport: Optional[Decimal] = None,
        reexport: Optional[Decimal] = None,
        insurance: Optional[Decimal] = None,
        fx_markup: Optional[Decimal] = None,
        fx_markup_abs: Optional[Decimal] = None,
        has_customs: Optional[bool] = None,
        via_novo: Optional[bool] = None,
        agent_fee: Optional[Decimal] = None,
    ) -> CostCalculationResult:
        supplier = self.get_supplier(supplier_id)
        product = self.get_product(product_id)
        fixed = self.get_fixed_costs()

        transport_used = self._to_decimal(
            supplier.transport_cost_per_l if transport is None else transport
        )
        reexport_used = self._to_decimal(
            supplier.reexport_percent if reexport is None else reexport
        )
        insurance_used = self._to_decimal(
            supplier.insurance_percent if insurance is None else insurance
        )
        fx_markup_used = self._to_decimal(
            supplier.fx_rate_markup if fx_markup is None else fx_markup
        )
        fx_markup_abs_used = self._to_decimal(
            supplier.fx_rate_markup_abs if fx_markup_abs is None else fx_markup_abs
        )
        has_customs_used = bool(
            supplier.has_import_duty if has_customs is None else has_customs
        )
        via_novo_used = bool(
            supplier.is_via_novo if via_novo is None else via_novo
        )
        agent_fee_used = self._to_decimal(
            getattr(supplier, "agent_fee", None) if agent_fee is None else agent_fee
        )

        cost_novo = self.calc_cost_novo_wvat(
            supplier_price=self._to_decimal(supplier_price),
            product_id=product_id,
            supplier_id=supplier_id,
            transport=transport_used,
            reexport=reexport_used,
            insurance=insurance_used,
            fx_rate=self._to_decimal(fx_rate),
            fx_markup=fx_markup_used,
            fx_markup_abs=fx_markup_abs_used,
            has_customs=has_customs_used,
            agent_fee=agent_fee_used,
        )
        if cost_novo is None:
            raise ValueError("Не удалось рассчитать cost_novo_wvat.")

        full_cost = self.calc_full_cost_msk(
            supplier_price=self._to_decimal(supplier_price),
            product_id=product_id,
            supplier_id=supplier_id,
            transport=transport_used,
            reexport=reexport_used,
            insurance=insurance_used,
            fx_rate=self._to_decimal(fx_rate),
            fx_markup=fx_markup_used,
            fx_markup_abs=fx_markup_abs_used,
            has_customs=has_customs_used,
            via_novo=via_novo_used,
            agent_fee=agent_fee_used,
        )
        if full_cost is None:
            raise ValueError("Не удалось рассчитать full_cost_msk.")

        marking_used = Decimal("0") if supplier.marks_for_us else self.get_marking_cost(product_id)

        customs_fee_used = Decimal("0") if supplier.is_rf else self._to_decimal(fixed.customs_fee)
        move_novo_used = Decimal("0") if supplier.is_rf else self._to_decimal(fixed.move_novo_tamozh)
        move_msk_used = Decimal("0") if supplier.is_rf else self._to_decimal(fixed.move_tamozh_chekhov)
        is_excise_used = False if supplier.is_rf else bool(product.is_excise)

        return CostCalculationResult(
            supplier_price=self._to_decimal(supplier_price),
            cost_novo_wvat=self._to_decimal(cost_novo),
            full_cost_msk=self._to_decimal(full_cost),
            currency_code=currency_code,
            fx_rate_used=self._to_decimal(fx_rate),
            fx_markup_used=fx_markup_used,
            fx_markup_abs_used=fx_markup_abs_used,
            transport_used=transport_used,
            reexport_used=reexport_used,
            insurance_used=insurance_used,
            agent_fee_used=agent_fee_used,
            has_customs_used=has_customs_used,
            via_novo_used=via_novo_used,
            bank_fee_used=self._to_decimal(fixed.bank_fee),
            customs_fee_used=customs_fee_used,
            move_novo_used=move_novo_used,
            move_msk_used=move_msk_used,
            is_excise_used=is_excise_used,
            additional_customs_used=self._to_decimal(fixed.additional_customs),
            storage_used=self._to_decimal(fixed.storage),
            marking_used=self._to_decimal(marking_used),
        )
