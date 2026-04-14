from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal
from typing import Optional

from sqlalchemy.orm import Session

from app.db.models import CurrentSupplierPrice, PriceHistory, Supplier


@dataclass(slots=True)
class SupplierPriceSnapshot:
    supplier_id: int
    product_id: int
    price: Decimal
    currency_code: str
    price_date: datetime


@dataclass(slots=True)
class SupplierPriceWithSupplier:
    supplier_id: int
    supplier_name: str
    product_id: int
    price: Decimal
    currency_code: str
    price_date: datetime


class PriceRepository:
    def __init__(self, session: Session) -> None:
        self.session = session

    @staticmethod
    def _to_decimal(value: object) -> Decimal:
        if value is None:
            return Decimal("0")
        if isinstance(value, Decimal):
            return value
        return Decimal(str(value))

    def get_current_supplier_price(self, supplier_id: int, product_id: int) -> Optional[CurrentSupplierPrice]:
        return self.session.query(CurrentSupplierPrice).filter(
            CurrentSupplierPrice.supplier_id == supplier_id,
            CurrentSupplierPrice.product_id == product_id,
        ).first()

    def get_last_price_history_row(self, supplier_id: int, product_id: int) -> Optional[PriceHistory]:
        return self.session.query(PriceHistory).filter(
            PriceHistory.supplier_id == supplier_id,
            PriceHistory.product_id == product_id,
            PriceHistory.price.isnot(None),
        ).order_by(PriceHistory.price_date.desc(), PriceHistory.id.desc()).first()

    def get_previous_price_history_row(self, supplier_id: int, product_id: int, last_price_date: datetime) -> Optional[PriceHistory]:
        return self.session.query(PriceHistory).filter(
            PriceHistory.supplier_id == supplier_id,
            PriceHistory.product_id == product_id,
            PriceHistory.price.isnot(None),
            PriceHistory.price_date < last_price_date,
        ).order_by(PriceHistory.price_date.desc(), PriceHistory.id.desc()).first()

    def get_last_supplier_price_snapshot(self, supplier_id: int, product_id: int) -> Optional[SupplierPriceSnapshot]:
        current_row = self.get_current_supplier_price(supplier_id=supplier_id, product_id=product_id)
        if current_row is not None:
            return SupplierPriceSnapshot(
                supplier_id=supplier_id,
                product_id=product_id,
                price=self._to_decimal(current_row.price),
                currency_code=current_row.currency,
                price_date=current_row.last_update,
            )

        history_row = self.get_last_price_history_row(supplier_id=supplier_id, product_id=product_id)
        if history_row is None:
            return None

        return SupplierPriceSnapshot(
            supplier_id=supplier_id,
            product_id=product_id,
            price=self._to_decimal(history_row.price),
            currency_code=history_row.currency,
            price_date=history_row.price_date,
        )

    def get_previous_supplier_price_snapshot(self, supplier_id: int, product_id: int, last_price_date: datetime) -> Optional[SupplierPriceSnapshot]:
        history_row = self.get_previous_price_history_row(
            supplier_id=supplier_id,
            product_id=product_id,
            last_price_date=last_price_date,
        )
        if history_row is None:
            return None

        return SupplierPriceSnapshot(
            supplier_id=supplier_id,
            product_id=product_id,
            price=self._to_decimal(history_row.price),
            currency_code=history_row.currency,
            price_date=history_row.price_date,
        )

    def get_suppliers_with_current_prices_for_product(self, product_id: int, only_rating_calc: bool = True) -> list[SupplierPriceWithSupplier]:
        query = self.session.query(CurrentSupplierPrice, Supplier).join(
            Supplier, Supplier.id == CurrentSupplierPrice.supplier_id
        ).filter(
            CurrentSupplierPrice.product_id == product_id,
            CurrentSupplierPrice.price.isnot(None),
        )

        if only_rating_calc:
            query = query.filter(Supplier.rating_calc.is_(True))

        rows = query.order_by(Supplier.name.asc(), CurrentSupplierPrice.last_update.desc()).all()

        result: list[SupplierPriceWithSupplier] = []
        for current_price, supplier in rows:
            result.append(
                SupplierPriceWithSupplier(
                    supplier_id=supplier.id,
                    supplier_name=supplier.name,
                    product_id=current_price.product_id,
                    price=self._to_decimal(current_price.price),
                    currency_code=current_price.currency,
                    price_date=current_price.last_update,
                )
            )
        return result

    def save_supplier_price(self, *, supplier_id: int, product_id: int, price: object, currency_code: str, price_date: datetime) -> None:
        price_value = self._to_decimal(price)

        history_row = PriceHistory(
            supplier_id=supplier_id,
            product_id=product_id,
            price_date=price_date,
            price=price_value,
            currency=currency_code,
        )
        self.session.add(history_row)

        current_row = self.get_current_supplier_price(supplier_id=supplier_id, product_id=product_id)

        if current_row is None:
            current_row = CurrentSupplierPrice(
                supplier_id=supplier_id,
                product_id=product_id,
                price=price_value,
                currency=currency_code,
                last_update=price_date,
            )
            self.session.add(current_row)
        else:
            if current_row.last_update is None or current_row.last_update <= price_date:
                current_row.price = price_value
                current_row.currency = currency_code
                current_row.last_update = price_date

        self.session.flush()