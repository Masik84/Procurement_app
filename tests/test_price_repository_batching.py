from __future__ import annotations

import unittest
from datetime import datetime
from decimal import Decimal

from sqlalchemy import create_engine, event
from sqlalchemy.orm import sessionmaker

from app.db.db import Base
from app.db.models import CurrentSupplierPrice, PriceHistory, Product, Supplier
from app.services.price_repository import PriceRepository


class PriceRepositoryBatchingTests(unittest.TestCase):
    def setUp(self) -> None:
        self.engine = create_engine("sqlite:///:memory:")
        Base.metadata.create_all(
            self.engine,
            tables=[
                Product.__table__,
                Supplier.__table__,
                PriceHistory.__table__,
                CurrentSupplierPrice.__table__,
            ],
        )
        self.session = sessionmaker(bind=self.engine)()

        products = [
            Product(id=1, name="Product 1", brand="Brand", pack=1, is_excise=False),
            Product(id=2, name="Product 2", brand="Brand", pack=4, is_excise=False),
        ]
        suppliers = [
            Supplier(id=1, name="Supplier A", base_currency="USD", rating_calc=True),
            Supplier(id=2, name="Supplier B", base_currency="EUR", rating_calc=False),
            Supplier(id=3, name="Manual", base_currency="USD", rating_calc=True),
        ]
        self.session.add_all(products + suppliers)
        self.session.add_all([
            CurrentSupplierPrice(
                supplier_id=1,
                product_id=1,
                price=Decimal("10"),
                currency="USD",
                last_update=datetime(2026, 8, 10),
            ),
            CurrentSupplierPrice(
                supplier_id=2,
                product_id=1,
                price=Decimal("20"),
                currency="EUR",
                last_update=datetime(2026, 1, 1),
            ),
            CurrentSupplierPrice(
                supplier_id=3,
                product_id=1,
                price=Decimal("30"),
                currency="USD",
                last_update=datetime(2026, 8, 10),
            ),
        ])
        self.session.add_all([
            PriceHistory(supplier_id=1, product_id=1, price=11, currency="USD", price_date=datetime(2026, 8, 12)),
            PriceHistory(supplier_id=2, product_id=1, price=21, currency="EUR", price_date=datetime(2026, 8, 11)),
            PriceHistory(supplier_id=2, product_id=1, price=22, currency="EUR", price_date=datetime(2026, 8, 11)),
            PriceHistory(supplier_id=1, product_id=2, price=40, currency="USD", price_date=datetime(2026, 8, 9)),
        ])
        self.session.commit()

    def tearDown(self) -> None:
        self.session.close()
        self.engine.dispose()

    @staticmethod
    def _signature(prices) -> list[tuple]:
        return sorted(
            (
                int(price.supplier_id),
                int(price.product_id),
                Decimal(price.price),
                price.currency_code,
                price.price_date,
            )
            for price in prices
        )

    def _legacy_prices(self, repository: PriceRepository, product_id: int, cutoff: datetime):
        current = repository.get_suppliers_with_current_prices_for_product(
            product_id,
            only_rating_calc=False,
            min_price_date=cutoff,
        )
        seen = {int(price.supplier_id) for price in current}
        history = [
            price
            for price in repository.get_latest_history_prices_for_product(
                product_id,
                only_rating_calc=False,
                min_price_date=cutoff,
            )
            if int(price.supplier_id) not in seen
        ]
        return current + history

    def test_bulk_prices_equal_legacy_selection_rules(self) -> None:
        repository = PriceRepository(self.session)
        cutoff = datetime(2026, 8, 1)
        bulk = repository.get_supplier_prices_for_products(
            [1, 2],
            only_rating_calc=False,
            min_price_date=cutoff,
        )

        for product_id in (1, 2):
            self.assertEqual(
                self._signature(self._legacy_prices(repository, product_id, cutoff)),
                self._signature(bulk[product_id]),
            )

        # The latest duplicate history date must retain the former highest-id rule.
        supplier_b = next(price for price in bulk[1] if price.supplier_id == 2)
        self.assertEqual(Decimal("22"), supplier_b.price)
        self.assertNotIn(3, {price.supplier_id for price in bulk[1]})

    def test_bulk_prices_use_two_queries_for_multiple_products(self) -> None:
        statements = 0

        def count_statement(*_args, **_kwargs):
            nonlocal statements
            statements += 1

        event.listen(self.engine, "before_cursor_execute", count_statement)
        try:
            PriceRepository(self.session).get_supplier_prices_for_products(
                [1, 2],
                only_rating_calc=False,
                min_price_date=datetime(2026, 8, 1),
            )
        finally:
            event.remove(self.engine, "before_cursor_execute", count_statement)

        self.assertEqual(2, statements)


if __name__ == "__main__":
    unittest.main()
