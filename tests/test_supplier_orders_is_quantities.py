from __future__ import annotations

from datetime import datetime
from decimal import Decimal
import unittest
from unittest.mock import patch

import pandas as pd
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker

from app.db.db import Base
from app.db.models import (
    Product,
    ProductArticle,
    ProductStock,
    TempIsImport,
    TempStockImport,
    TempSupplierOrdersImport,
)
from app.imports.stock_importer import StockImporter, parse_source_is_excise
from app.imports.supplier_orders_importer import SupplierOrdersImporter
from app.services.product_stock_service import ProductStockService


class SupplierOrdersImporterIsQuantityTests(unittest.TestCase):
    def test_splits_coral_quantities_from_regular_supplier_orders(self):
        headers = [
            "Статус",
            "Brand",
            "Supplier 1",
            "Артикул",
            "Продукт + упаковка",
            "Упаковка",
            "Акциз (да/нет)",
            "ABC",
            "Назв на англ",
            "Кол-во, л",
        ]
        frame = pd.DataFrame([
            [None] * len(headers),
            headers,
            ["order", "Brand A", "Supplier A", "A-1", "Product A", "1", "да", "A", "Product A 1L", 10],
            ["order", "Brand A", "  CoRaL ", "A-1", "Product A", "1", "да", "A", "Product A 1L", 3],
            ["confirmed", "Brand A", "CORAL", "A-1", "Product A", "1", "да", "A", "Product A 1L", 4],
            ["confirmed", "Brand A", "Supplier A", "A-1", "Product A", "1", "да", "A", "Product A 1L", 100],
            ["done", "Brand A", "CORAL", "A-1", "Product A", "1", "да", "A", "Product A 1L", 200],
        ])

        with patch("app.imports.supplier_orders_importer.read_excel_raw", return_value=frame):
            rows = SupplierOrdersImporter().read_excel(__file__)

        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["order_qty"], 10.0)
        self.assertEqual(rows[0]["is_order_qty"], 3.0)
        self.assertEqual(rows[0]["is_confirmed_order_qty"], 4.0)
        self.assertEqual(rows[0]["new_brand"], "Brand A")
        self.assertEqual(rows[0]["new_pack"], Decimal("1"))
        self.assertIs(rows[0]["new_is_excise"], True)

    def test_stock_product_type_excise_rule(self):
        self.assertIs(parse_source_is_excise("PVL", product_type=True), True)
        self.assertIs(parse_source_is_excise(" cvl ", product_type=True), True)
        self.assertIs(parse_source_is_excise("Transmission", product_type=True), False)
        self.assertIs(parse_source_is_excise("Industry", product_type=True), False)
        self.assertIs(parse_source_is_excise("", product_type=True), False)

    def test_stock_importer_populates_existing_new_product_fields(self):
        headers = [
            "Бренд",
            "Английское наименование продукта",
            "Code 1C",
            "SKU",
            "Упаковка",
            "Prod.type",
            "Страна происхождения",
            "LPC",
            "Цена/л c НДС",
            "Цена/л c НДС",
            "Landed Cost+VAT/L",
            "Группа бренда",
            "Категория ABC",
            "Свободный сток (Новороссийск)",
            "Общий Транзит, л",
        ]
        top_headers = [""] * len(headers)
        top_headers[8] = "Цена Дистр"
        top_headers[9] = "Цена Промо"
        row = [
            "SINOPEC",
            "SINOPEC 4306 20KG",
            "0609200062418184",
            "SINOPEC 4306 20кг",
            "20,0",
            "CVL",
            "CN",
            1,
            2,
            3,
            4,
            "import",
            "A",
            5,
            6,
        ]
        frame = pd.DataFrame([top_headers, [None] * len(headers), [None] * len(headers), headers, row])

        with patch("app.imports.stock_importer.read_excel_raw", return_value=frame):
            rows = StockImporter().read_excel(__file__)

        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["new_brand"], "SINOPEC")
        self.assertEqual(rows[0]["new_pack"], Decimal("20.0"))
        self.assertIs(rows[0]["new_is_excise"], True)
        self.assertEqual(rows[0]["new_product_name"], "SINOPEC 4306 20KG")
        self.assertEqual(rows[0]["source_product_name"], "SINOPEC 4306 20KG")


class ProductStockServiceIsQuantityTests(unittest.TestCase):
    def setUp(self):
        self.engine = create_engine("sqlite+pysqlite:///:memory:", future=True)
        Base.metadata.create_all(
            self.engine,
            tables=[
                Product.__table__,
                ProductArticle.__table__,
                ProductStock.__table__,
                TempStockImport.__table__,
                TempSupplierOrdersImport.__table__,
                TempIsImport.__table__,
            ],
        )
        self.session = sessionmaker(bind=self.engine, autoflush=False, future=True)()

    def tearDown(self):
        self.session.close()
        self.engine.dispose()

    def _add_product(self, name: str) -> Product:
        product = Product(brand="Brand", name=name, pack=1, is_excise=False)
        self.session.add(product)
        self.session.flush()
        return product

    def test_supplier_orders_update_all_order_fields(self):
        product = self._add_product("Product A")
        untouched_product = self._add_product("Product B")
        self.session.add_all([
            ProductStock(
                product_id=product.id,
                product_name=product.name,
                order_qty=99,
                is_order_qty=98,
                is_confirmed_order_qty=97,
            ),
            ProductStock(
                product_id=untouched_product.id,
                product_name=untouched_product.name,
                order_qty=9,
                is_order_qty=8,
                is_confirmed_order_qty=7,
            ),
            TempSupplierOrdersImport(
                batch_id="batch",
                imported_by="tester",
                import_date=datetime.now(),
                import_row_no=3,
                source_article="A-1",
                source_product_name=product.name,
                order_qty=Decimal("10"),
                is_order_qty=Decimal("3"),
                is_confirmed_order_qty=Decimal("4"),
                selected_product_id=product.id,
            ),
        ])
        self.session.flush()

        saved = ProductStockService(self.session).save_supplier_orders_to_product_stock("batch", "tester")

        self.assertEqual(saved, 1)
        self.session.expire_all()
        stock = self.session.get(ProductStock, product.id)
        self.assertEqual(stock.order_qty, Decimal("10.0000000000"))
        self.assertEqual(stock.is_order_qty, Decimal("3.0000000000"))
        self.assertEqual(stock.is_confirmed_order_qty, Decimal("4.0000000000"))

        untouched_stock = self.session.get(ProductStock, untouched_product.id)
        self.assertEqual(untouched_stock.order_qty, Decimal("0E-10"))
        self.assertEqual(untouched_stock.is_order_qty, Decimal("0E-10"))
        self.assertEqual(untouched_stock.is_confirmed_order_qty, Decimal("0E-10"))

    def test_is_update_changes_only_is_stock_quantity(self):
        product = self._add_product("Product A")
        self.session.add_all([
            ProductStock(
                product_id=product.id,
                product_name=product.name,
                is_order_qty=Decimal("3"),
                is_confirmed_order_qty=Decimal("4"),
                is_stock_qty=Decimal("99"),
            ),
            TempIsImport(
                batch_id="batch",
                imported_by="tester",
                import_date=datetime.now(),
                import_row_no=3,
                source_article="A-1",
                source_product_name=product.name,
                remains_qty=Decimal("100"),
                confirmed_qty=Decimal("200"),
                stock_qty=Decimal("5"),
                selected_product_id=product.id,
            ),
        ])
        self.session.flush()

        saved = ProductStockService(self.session).save_is_to_product_stock("batch", "tester")

        self.assertEqual(saved, 1)
        self.session.expire_all()
        stock = self.session.get(ProductStock, product.id)
        self.assertEqual(stock.is_order_qty, Decimal("3.0000000000"))
        self.assertEqual(stock.is_confirmed_order_qty, Decimal("4.0000000000"))
        self.assertEqual(stock.is_stock_qty, Decimal("5.0000000000"))

    def test_creates_stock_product_and_article_link_from_existing_new_fields(self):
        row = TempStockImport(
            batch_id="batch",
            imported_by="tester",
            import_date=datetime.now(),
            import_row_no=5,
            source_article="0609200062418184",
            source_sku="SINOPEC 4306 20кг",
            source_product_name="SINOPEC 4306 20KG",
            new_product_name="SINOPEC 4306 20KG",
            new_brand="SINOPEC",
            new_pack=Decimal("20"),
            new_is_excise=False,
        )
        self.session.add(row)
        self.session.flush()

        created = ProductStockService(self.session).create_stock_products_from_temp("batch", "tester")

        self.assertEqual(created, 1)
        self.assertIsNotNone(row.selected_product_id)
        product = self.session.get(Product, row.selected_product_id)
        self.assertEqual(product.name, "SINOPEC 4306 20KG")
        self.assertEqual(product.brand, "SINOPEC")
        self.assertEqual(product.pack, Decimal("20.0000000000"))
        self.assertIs(product.is_excise, False)
        link = self.session.query(ProductArticle).filter_by(product_id=product.id).one()
        self.assertEqual(link.article, "0609200062418184")
        self.assertEqual(link.name, "SINOPEC 4306 20KG")

    def test_stock_automatch_flushes_updates_in_bounded_batches(self):
        products = [self._add_product(f"BATCH PRODUCT {index} 1L") for index in range(3)]
        self.session.add_all([
            TempStockImport(
                batch_id="batch",
                imported_by="tester",
                import_date=datetime.now(),
                import_row_no=index + 5,
                source_product_name=product.name,
                new_product_name=product.name,
                new_brand=product.brand,
                new_pack=product.pack,
                new_is_excise=product.is_excise,
            )
            for index, product in enumerate(products)
        ])
        self.session.flush()

        service = ProductStockService(self.session)
        service.AUTOMATCH_FLUSH_BATCH_SIZE = 2
        with patch.object(self.session, "flush", wraps=self.session.flush) as flush_mock:
            matched = service.automatch_stock_rows("batch", "tester")

        self.assertEqual(matched, 3)
        self.assertEqual(flush_mock.call_count, 2)
        rows = self.session.query(TempStockImport).filter_by(batch_id="batch").all()
        self.assertTrue(all(row.selected_product_id is not None for row in rows))
        self.assertTrue(all(row.new_product_name is None for row in rows))
        self.assertTrue(all(row.new_brand is None for row in rows))
        self.assertTrue(all(row.new_pack is None for row in rows))
        self.assertTrue(all(row.new_is_excise is None for row in rows))

    def test_auto_creates_supplier_order_product_with_explicit_excise_value(self):
        row = TempSupplierOrdersImport(
            batch_id="batch",
            imported_by="tester",
            import_date=datetime.now(),
            import_row_no=3,
            source_article="550033463CN",
            source_product_name="SHELL SPIRAX S4 CX 30 209L",
            new_product_name="SHELL SPIRAX S4 CX 30 209L",
            new_brand="SHELL",
            new_pack=Decimal("209"),
            new_is_excise=True,
            order_qty=Decimal("10"),
        )
        self.session.add(row)
        self.session.flush()

        stats = ProductStockService(self.session).save_supplier_orders("batch", "tester")

        self.assertEqual(stats["created_products_count"], 1)
        product = self.session.query(Product).filter_by(name="SHELL SPIRAX S4 CX 30 209L").one()
        self.assertEqual(product.name, "SHELL SPIRAX S4 CX 30 209L")
        self.assertIs(product.is_excise, True)
        link = self.session.query(ProductArticle).filter_by(product_id=product.id).one()
        self.assertEqual(link.article, "550033463CN")
        self.assertEqual(link.name, "SHELL SPIRAX S4 CX 30 209L")
        stock = self.session.get(ProductStock, product.id)
        self.assertEqual(stock.order_qty, Decimal("10.0000000000"))
        self.assertEqual(
            self.session.query(TempSupplierOrdersImport).filter_by(batch_id="batch").count(),
            0,
        )


if __name__ == "__main__":
    unittest.main()
