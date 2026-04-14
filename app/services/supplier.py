from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal
from typing import Optional

from sqlalchemy.orm import Session

from app.db.models import ExchangeRate, Supplier
from app.utils.text import clean_multi_spaces


@dataclass(slots=True)
class SupplierUpsertData:
    name: str
    base_currency: str
    transport_cost_per_l: Decimal = Decimal("0")
    reexport_percent: Decimal = Decimal("0")
    fx_rate_markup: Decimal = Decimal("0")
    is_via_novo: bool = False
    has_import_duty: bool = False
    rating_calc: bool = True
    marks_for_us: bool = False
    is_rf: bool = False


class SupplierService:
    def __init__(self, session: Session) -> None:
        self.session = session

    def get_supplier_by_id(self, supplier_id: int) -> Optional[Supplier]:
        return (
            self.session.query(Supplier)
            .filter(Supplier.id == supplier_id)
            .first()
        )

    def get_supplier_by_name(self, supplier_name: object) -> Optional[Supplier]:
        name = clean_multi_spaces(supplier_name)
        if not name:
            return None

        return (
            self.session.query(Supplier)
            .filter(Supplier.name == name)
            .first()
        )

    def get_all_suppliers(self) -> list[Supplier]:
        return (
            self.session.query(Supplier)
            .order_by(Supplier.name.asc())
            .all()
        )

    def get_exchange_rate(self, currency_code: object) -> Optional[ExchangeRate]:
        code = clean_multi_spaces(currency_code).upper()
        if not code:
            return None

        return (
            self.session.query(ExchangeRate)
            .filter(ExchangeRate.currency_code == code)
            .first()
        )

    def get_all_exchange_rates(self) -> list[ExchangeRate]:
        return (
            self.session.query(ExchangeRate)
            .order_by(ExchangeRate.currency_code.asc())
            .all()
        )

    def get_rate_to_rub(self, currency_code: object) -> Optional[float]:
        rate = self.get_exchange_rate(currency_code)
        if rate is None:
            return None
        return rate.rate_to_rub

    @staticmethod
    def normalize_supplier_name(value: object) -> str:
        return clean_multi_spaces(value)

    @staticmethod
    def normalize_currency_code(value: object) -> str:
        return clean_multi_spaces(value).upper()

    def validate_supplier_data(self, data: SupplierUpsertData) -> SupplierUpsertData:
        name = self.normalize_supplier_name(data.name)
        currency = self.normalize_currency_code(data.base_currency)

        if not name:
            raise ValueError("Введите название поставщика.")

        if not currency:
            raise ValueError("Выберите валюту поставщика.")

        return SupplierUpsertData(
            name=name,
            base_currency=currency,
            transport_cost_per_l=Decimal(str(data.transport_cost_per_l)),
            reexport_percent=Decimal(str(data.reexport_percent)),
            fx_rate_markup=Decimal(str(data.fx_rate_markup)),
            is_via_novo=bool(data.is_via_novo),
            has_import_duty=bool(data.has_import_duty),
            rating_calc=bool(data.rating_calc),
            marks_for_us=bool(data.marks_for_us),
            is_rf=bool(data.is_rf),
        )

    def save_exchange_rate(self, currency_code: object, rate_to_rub: float) -> ExchangeRate:
        code = self.normalize_currency_code(currency_code)
        if not code:
            raise ValueError("Currency code is required.")

        rate_value = Decimal(str(rate_to_rub))
        row = self.get_exchange_rate(code)

        if row is None:
            row = ExchangeRate(currency_code=code, rate_to_rub=rate_value)
            self.session.add(row)
        else:
            row.rate_to_rub = rate_value

        self.session.flush()
        return row

    def create_supplier(self, data: SupplierUpsertData) -> Supplier:
        validated = self.validate_supplier_data(data)

        existing = self.get_supplier_by_name(validated.name)
        if existing is not None:
            raise ValueError(f"Поставщик '{validated.name}' уже существует.")

        supplier = Supplier(
            name=validated.name,
            base_currency=validated.base_currency,
            transport_cost_per_l=validated.transport_cost_per_l,
            reexport_percent=validated.reexport_percent,
            fx_rate_markup=validated.fx_rate_markup,
            is_via_novo=validated.is_via_novo,
            has_import_duty=validated.has_import_duty,
            rating_calc=validated.rating_calc,
            marks_for_us=validated.marks_for_us,
            is_rf=validated.is_rf,
        )
        self.session.add(supplier)
        self.session.flush()
        return supplier

    def update_supplier(self, supplier: Supplier, data: SupplierUpsertData) -> Supplier:
        validated = self.validate_supplier_data(data)

        if supplier.name != validated.name:
            existing = self.get_supplier_by_name(validated.name)
            if existing is not None and existing.id != supplier.id:
                raise ValueError(f"Поставщик '{validated.name}' уже существует.")

        supplier.name = validated.name
        supplier.base_currency = validated.base_currency
        supplier.transport_cost_per_l = validated.transport_cost_per_l
        supplier.reexport_percent = validated.reexport_percent
        supplier.fx_rate_markup = validated.fx_rate_markup
        supplier.is_via_novo = validated.is_via_novo
        supplier.has_import_duty = validated.has_import_duty
        supplier.rating_calc = validated.rating_calc
        supplier.marks_for_us = validated.marks_for_us
        supplier.is_rf = validated.is_rf

        self.session.flush()
        return supplier

    def create_or_update_supplier(
        self,
        *,
        supplier_id: Optional[int] = None,
        data: SupplierUpsertData,
    ) -> Supplier:
        validated = self.validate_supplier_data(data)

        if supplier_id is not None:
            supplier = self.get_supplier_by_id(supplier_id)
            if supplier is None:
                raise ValueError(f"Поставщик id={supplier_id} не найден.")
            return self.update_supplier(supplier, validated)

        existing = self.get_supplier_by_name(validated.name)
        if existing is not None:
            return self.update_supplier(existing, validated)

        return self.create_supplier(validated)

    def ensure_supplier(
        self,
        *,
        supplier_id: Optional[int] = None,
        data: SupplierUpsertData,
    ) -> Supplier:
        return self.create_or_update_supplier(supplier_id=supplier_id, data=data)

    def load_supplier_snapshot(self, supplier_id: int) -> SupplierUpsertData:
        supplier = self.get_supplier_by_id(supplier_id)
        if supplier is None:
            raise ValueError(f"Поставщик id={supplier_id} не найден.")

        return SupplierUpsertData(
            name=supplier.name,
            base_currency=supplier.base_currency,
            transport_cost_per_l=supplier.transport_cost_per_l,
            reexport_percent=supplier.reexport_percent,
            fx_rate_markup=supplier.fx_rate_markup,
            is_via_novo=bool(supplier.is_via_novo),
            has_import_duty=bool(supplier.has_import_duty),
            rating_calc=bool(supplier.rating_calc),
            marks_for_us=bool(supplier.marks_for_us),
            is_rf=bool(supplier.is_rf),
        )
