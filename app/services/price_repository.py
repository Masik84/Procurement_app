from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal
from typing import Iterable, Mapping, Optional

from sqlalchemy.orm import Session
from sqlalchemy import func, tuple_

from app.db.models import CurrentSupplierPrice, PriceHistory, Supplier
from app.services.supplier_service import MANUAL_SUPPLIER_NAME


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

    @staticmethod
    def supplier_price_cutoff_from_months(months: int | None) -> datetime | None:
        try:
            months_value = int(months or 0)
        except (TypeError, ValueError):
            return None

        if months_value <= 0:
            return None

        today = datetime.now()
        year = today.year
        month = today.month - months_value
        while month <= 0:
            month += 12
            year -= 1
        return datetime(year, month, 1)

    def get_current_supplier_price(self, supplier_id: int, product_id: int) -> Optional[CurrentSupplierPrice]:
        return self.session.query(CurrentSupplierPrice).filter(
            CurrentSupplierPrice.supplier_id == supplier_id,
            CurrentSupplierPrice.product_id == product_id,
        ).first()

    def get_last_price_history_row(
        self,
        supplier_id: int,
        product_id: int,
        min_price_date: datetime | None = None,
    ) -> Optional[PriceHistory]:
        query = self.session.query(PriceHistory).filter(
            PriceHistory.supplier_id == supplier_id,
            PriceHistory.product_id == product_id,
            PriceHistory.price.isnot(None),
        )
        if min_price_date is not None:
            query = query.filter(PriceHistory.price_date >= min_price_date)
        return query.order_by(PriceHistory.price_date.desc(), PriceHistory.id.desc()).first()

    def get_previous_price_history_row(self, supplier_id: int, product_id: int, last_price_date: datetime) -> Optional[PriceHistory]:
        return self.session.query(PriceHistory).filter(
            PriceHistory.supplier_id == supplier_id,
            PriceHistory.product_id == product_id,
            PriceHistory.price.isnot(None),
            PriceHistory.price_date < last_price_date,
        ).order_by(PriceHistory.price_date.desc(), PriceHistory.id.desc()).first()

    def get_last_supplier_price_snapshot(
        self,
        supplier_id: int,
        product_id: int,
        min_price_date: datetime | None = None,
    ) -> Optional[SupplierPriceSnapshot]:
        current_row = self.get_current_supplier_price(supplier_id=supplier_id, product_id=product_id)
        if current_row is not None and (
            min_price_date is None
            or (current_row.last_update is not None and current_row.last_update >= min_price_date)
        ):
            return SupplierPriceSnapshot(
                supplier_id=supplier_id,
                product_id=product_id,
                price=self._to_decimal(current_row.price),
                currency_code=current_row.currency,
                price_date=current_row.last_update,
            )

        history_row = self.get_last_price_history_row(
            supplier_id=supplier_id,
            product_id=product_id,
            min_price_date=min_price_date,
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

    def get_suppliers_with_current_prices_for_product(
        self,
        product_id: int,
        only_rating_calc: bool = True,
        min_price_date: datetime | None = None,
    ) -> list[SupplierPriceWithSupplier]:
        query = self.session.query(CurrentSupplierPrice, Supplier).join(
            Supplier, Supplier.id == CurrentSupplierPrice.supplier_id
        ).filter(
            CurrentSupplierPrice.product_id == product_id,
            CurrentSupplierPrice.price.isnot(None),
            Supplier.name != MANUAL_SUPPLIER_NAME,
        )

        if min_price_date is not None:
            query = query.filter(CurrentSupplierPrice.last_update >= min_price_date)

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

    def get_latest_history_prices_for_product(
        self,
        product_id: int,
        only_rating_calc: bool = True,
        min_price_date: datetime | None = None,
    ) -> list[SupplierPriceWithSupplier]:
        max_dates_query = (
            self.session.query(
                PriceHistory.supplier_id.label("supplier_id"),
                PriceHistory.product_id.label("product_id"),
                func.max(PriceHistory.price_date).label("max_price_date"),
            )
            .filter(
                PriceHistory.product_id == product_id,
                PriceHistory.price.isnot(None),
            )
        )

        if min_price_date is not None:
            max_dates_query = max_dates_query.filter(PriceHistory.price_date >= min_price_date)

        max_dates_subq = (
            max_dates_query
            .group_by(
                PriceHistory.supplier_id,
                PriceHistory.product_id,
            )
            .subquery()
        )

        rows_query = (
            self.session.query(PriceHistory, Supplier)
            .join(
                max_dates_subq,
                (PriceHistory.supplier_id == max_dates_subq.c.supplier_id)
                & (PriceHistory.product_id == max_dates_subq.c.product_id)
                & (PriceHistory.price_date == max_dates_subq.c.max_price_date),
            )
            .join(Supplier, Supplier.id == PriceHistory.supplier_id)
            .filter(
                PriceHistory.product_id == product_id,
                PriceHistory.price.isnot(None),
            )
        )

        if min_price_date is not None:
            rows_query = rows_query.filter(PriceHistory.price_date >= min_price_date)

        rows = (
            rows_query
            .order_by(Supplier.name.asc(), PriceHistory.price_date.desc(), PriceHistory.id.desc())
            .all()
        )

        result: list[SupplierPriceWithSupplier] = []
        seen_supplier_ids: set[int] = set()

        for history_row, supplier in rows:
            if only_rating_calc and not supplier.rating_calc:
                continue

            if supplier.id in seen_supplier_ids:
                continue

            seen_supplier_ids.add(supplier.id)
            result.append(
                SupplierPriceWithSupplier(
                    supplier_id=supplier.id,
                    supplier_name=supplier.name,
                    product_id=history_row.product_id,
                    price=self._to_decimal(history_row.price),
                    currency_code=history_row.currency,
                    price_date=history_row.price_date,
                )
            )

        return result

    def get_supplier_prices_for_products(
        self,
        product_ids: Iterable[int],
        *,
        only_rating_calc: bool = True,
        min_price_date: datetime | None = None,
        exclude_manual: bool = True,
    ) -> dict[int, list[SupplierPriceWithSupplier]]:
        """Return one effective price per product/supplier using two bulk queries.

        A qualifying current price has the same priority as in the former
        per-product implementation.  History supplies only pairs which have no
        qualifying current row, and ties on the latest date are resolved by the
        greatest history id exactly as before.
        """
        ids = sorted({int(product_id) for product_id in product_ids if product_id is not None})
        result: dict[int, list[SupplierPriceWithSupplier]] = {product_id: [] for product_id in ids}
        if not ids:
            return result

        current_query = (
            self.session.query(CurrentSupplierPrice, Supplier)
            .join(Supplier, Supplier.id == CurrentSupplierPrice.supplier_id)
            .filter(
                CurrentSupplierPrice.product_id.in_(ids),
                CurrentSupplierPrice.price.isnot(None),
            )
        )
        if exclude_manual:
            current_query = current_query.filter(Supplier.name != MANUAL_SUPPLIER_NAME)
        if min_price_date is not None:
            current_query = current_query.filter(CurrentSupplierPrice.last_update >= min_price_date)
        if only_rating_calc:
            current_query = current_query.filter(Supplier.rating_calc.is_(True))

        current_rows = current_query.order_by(
            CurrentSupplierPrice.product_id.asc(),
            Supplier.name.asc(),
            CurrentSupplierPrice.last_update.desc(),
        ).all()
        seen_pairs: set[tuple[int, int]] = set()
        for current_price, supplier in current_rows:
            product_id = int(current_price.product_id)
            supplier_id = int(supplier.id)
            seen_pairs.add((product_id, supplier_id))
            result[product_id].append(SupplierPriceWithSupplier(
                supplier_id=supplier_id,
                supplier_name=supplier.name,
                product_id=product_id,
                price=self._to_decimal(current_price.price),
                currency_code=current_price.currency,
                price_date=current_price.last_update,
            ))

        max_dates_query = self.session.query(
            PriceHistory.supplier_id.label("supplier_id"),
            PriceHistory.product_id.label("product_id"),
            func.max(PriceHistory.price_date).label("max_price_date"),
        ).filter(
            PriceHistory.product_id.in_(ids),
            PriceHistory.price.isnot(None),
        )
        if min_price_date is not None:
            max_dates_query = max_dates_query.filter(PriceHistory.price_date >= min_price_date)
        max_dates_subq = max_dates_query.group_by(
            PriceHistory.supplier_id,
            PriceHistory.product_id,
        ).subquery()

        history_query = (
            self.session.query(PriceHistory, Supplier)
            .join(
                max_dates_subq,
                (PriceHistory.supplier_id == max_dates_subq.c.supplier_id)
                & (PriceHistory.product_id == max_dates_subq.c.product_id)
                & (PriceHistory.price_date == max_dates_subq.c.max_price_date),
            )
            .join(Supplier, Supplier.id == PriceHistory.supplier_id)
            .filter(
                PriceHistory.product_id.in_(ids),
                PriceHistory.price.isnot(None),
            )
        )
        if min_price_date is not None:
            history_query = history_query.filter(PriceHistory.price_date >= min_price_date)
        history_rows = history_query.order_by(
            PriceHistory.product_id.asc(),
            Supplier.name.asc(),
            PriceHistory.price_date.desc(),
            PriceHistory.id.desc(),
        ).all()

        history_seen: set[tuple[int, int]] = set()
        for history_row, supplier in history_rows:
            product_id = int(history_row.product_id)
            supplier_id = int(supplier.id)
            pair = (product_id, supplier_id)
            if pair in seen_pairs or pair in history_seen:
                continue
            if only_rating_calc and not supplier.rating_calc:
                continue
            history_seen.add(pair)
            result[product_id].append(SupplierPriceWithSupplier(
                supplier_id=supplier_id,
                supplier_name=supplier.name,
                product_id=product_id,
                price=self._to_decimal(history_row.price),
                currency_code=history_row.currency,
                price_date=history_row.price_date,
            ))

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

    def save_supplier_prices_batch(self, prices: Iterable[Mapping[str, object]]) -> int:
        """Save history/current supplier prices with one lookup and one flush.

        Rows are applied in the supplied order so duplicate supplier/product
        pairs keep the same last-update behaviour as ``save_supplier_price``.
        """
        records = list(prices)
        if not records:
            return 0

        pairs = {
            (int(record["supplier_id"]), int(record["product_id"]))
            for record in records
        }
        current_rows = (
            self.session.query(CurrentSupplierPrice)
            .filter(
                tuple_(
                    CurrentSupplierPrice.supplier_id,
                    CurrentSupplierPrice.product_id,
                ).in_(pairs)
            )
            .all()
        )
        current_by_pair = {
            (int(row.supplier_id), int(row.product_id)): row
            for row in current_rows
        }

        history_rows: list[PriceHistory] = []
        new_current_rows: list[CurrentSupplierPrice] = []
        for record in records:
            supplier_id = int(record["supplier_id"])
            product_id = int(record["product_id"])
            price_value = self._to_decimal(record["price"])
            currency_code = str(record["currency_code"])
            price_date = record["price_date"]

            history_rows.append(PriceHistory(
                supplier_id=supplier_id,
                product_id=product_id,
                price_date=price_date,
                price=price_value,
                currency=currency_code,
            ))

            key = (supplier_id, product_id)
            current_row = current_by_pair.get(key)
            if current_row is None:
                current_row = CurrentSupplierPrice(
                    supplier_id=supplier_id,
                    product_id=product_id,
                    price=price_value,
                    currency=currency_code,
                    last_update=price_date,
                )
                current_by_pair[key] = current_row
                new_current_rows.append(current_row)
            elif current_row.last_update is None or current_row.last_update <= price_date:
                current_row.price = price_value
                current_row.currency = currency_code
                current_row.last_update = price_date

        self.session.add_all(history_rows)
        if new_current_rows:
            self.session.add_all(new_current_rows)
        self.session.flush()
        return len(records)
