from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal
from typing import Iterable, Mapping, Optional

from sqlalchemy.dialects.postgresql import insert as pg_insert
from sqlalchemy.orm import Session
from sqlalchemy import func, or_, tuple_

from app.db.models import CurrentSupplierPrice, PriceHistory, Supplier
from app.services.supplier_service import MANUAL_SUPPLIER_NAME
from app.utils.money import to_decimal


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

    _to_decimal = staticmethod(to_decimal)

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
        self.session.flush()

        self.upsert_current_prices(
            [{
                "supplier_id": supplier_id,
                "product_id": product_id,
                "price": price_value,
                "currency": currency_code,
                "last_update": price_date,
            }],
            only_if_newer=True,
        )

    def upsert_current_prices(self, values: Iterable[Mapping[str, object]], *, only_if_newer: bool) -> None:
        """Atomically insert/update current_supplier_prices rows.

        A SELECT-then-INSERT let two users who saved the same new
        supplier/product pair at the same time fail with a primary-key
        violation, rolling back the whole save. ON CONFLICT resolves that
        inside PostgreSQL.

        only_if_newer=True keeps the date priority: an existing row is replaced
        only when its last_update is empty or not newer than the new price.
        """
        rows = [dict(value) for value in values]
        if not rows:
            return
        table = CurrentSupplierPrice.__table__
        for offset in range(0, len(rows), 1000):
            stmt = pg_insert(table).values(rows[offset:offset + 1000])
            stmt = stmt.on_conflict_do_update(
                index_elements=[table.c.supplier_id, table.c.product_id],
                set_={
                    "price": stmt.excluded.price,
                    "currency": stmt.excluded.currency,
                    "last_update": stmt.excluded.last_update,
                },
                where=(
                    or_(
                        table.c.last_update.is_(None),
                        table.c.last_update <= stmt.excluded.last_update,
                    )
                    if only_if_newer else None
                ),
            )
            self.session.execute(stmt)

        # The statement bypassed the ORM: drop stale CurrentSupplierPrice
        # objects so later reads in this session see the stored values.
        for obj in list(self.session.identity_map.values()):
            if isinstance(obj, CurrentSupplierPrice):
                self.session.expire(obj)

    def save_supplier_prices_batch(self, prices: Iterable[Mapping[str, object]]) -> int:
        """Write only new supplier-price observations to history.

        History is compared by supplier/product/price type, not merely by date.
        Reimporting unchanged prices never creates more history rows. Current
        prices are still refreshed according to their existing date priority.
        All reads are set based (no SELECT per imported row).
        """
        records = list(prices)
        if not records:
            return 0

        pairs = {
            (int(record["supplier_id"]), int(record["product_id"]))
            for record in records
        }
        pair_filter = tuple_(
            PriceHistory.supplier_id, PriceHistory.product_id,
        ).in_(pairs)
        # Final current-price candidate per pair within this batch. The
        # comparison with the row already stored is done atomically by
        # upsert_current_prices(only_if_newer=True).
        current_by_pair: dict[tuple[int, int], dict] = {}

        # Find the latest saved observation for each price type separately.
        # A Window function avoids loading the entire price history into Python.
        ranked = (
            self.session.query(
                PriceHistory.id.label("history_id"),
                func.row_number().over(
                    partition_by=(
                        PriceHistory.supplier_id,
                        PriceHistory.product_id,
                        PriceHistory.price_type,
                    ),
                    order_by=(PriceHistory.price_date.desc(), PriceHistory.id.desc()),
                ).label("row_rank"),
            )
            .filter(pair_filter)
            .subquery()
        )
        latest_rows = (
            self.session.query(PriceHistory)
            .join(ranked, PriceHistory.id == ranked.c.history_id)
            .filter(ranked.c.row_rank == 1)
            .all()
        )
        last_by_type = {
            (int(row.supplier_id), int(row.product_id), row.price_type or "Price List"):
                (self._to_decimal(row.price), str(row.currency).strip().upper())
            for row in latest_rows
        }

        # Reimporting an old-date file must not duplicate a price observation
        # already present on that exact date, even if newer prices now exist.
        dates = {record["price_date"] for record in records}
        same_date_rows = (
            self.session.query(PriceHistory)
            .filter(pair_filter, PriceHistory.price_date.in_(dates))
            .all()
        )
        existing_observations = {
            (
                int(row.supplier_id), int(row.product_id),
                row.price_date, row.price_type or "Price List",
                self._to_decimal(row.price), str(row.currency).strip().upper(),
            )
            for row in same_date_rows
        }

        history_rows: list[PriceHistory] = []
        for record in records:
            supplier_id = int(record["supplier_id"])
            product_id = int(record["product_id"])
            price_value = self._to_decimal(record["price"])
            currency_code = str(record["currency_code"]).strip().upper()
            price_date = record["price_date"]
            price_type = str(record.get("price_type") or "Price List").strip() or "Price List"

            type_key = (supplier_id, product_id, price_type)
            observation = (
                supplier_id, product_id, price_date, price_type,
                price_value, currency_code,
            )
            value = (price_value, currency_code)
            if value != last_by_type.get(type_key) and observation not in existing_observations:
                history_rows.append(PriceHistory(
                    supplier_id=supplier_id,
                    product_id=product_id,
                    price_date=price_date,
                    price=price_value,
                    currency=currency_code,
                    price_type=price_type,
                ))
                existing_observations.add(observation)
                last_by_type[type_key] = value

            # Preserve the previous current-price date-priority semantics even
            # when the history record was skipped as an unchanged observation:
            # a later record with the same or a newer date wins.
            key = (supplier_id, product_id)
            candidate = current_by_pair.get(key)
            if candidate is None or candidate["last_update"] <= price_date:
                current_by_pair[key] = {
                    "supplier_id": supplier_id,
                    "product_id": product_id,
                    "price": price_value,
                    "currency": currency_code,
                    "last_update": price_date,
                }

        self.session.add_all(history_rows)
        self.session.flush()
        self.upsert_current_prices(current_by_pair.values(), only_if_newer=True)
        return len(history_rows)
