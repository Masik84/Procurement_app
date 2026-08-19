import unittest
from types import SimpleNamespace
from unittest.mock import MagicMock, patch

import pandas as pd

from app.services.order_planning_service import OrderPlanningService
from app.services.product_matching_service import ProductMatchingService


def _product(product_id: int, name: str, brand: str):
    return SimpleNamespace(id=product_id, name=name, brand=brand)


class ProductMatchingBrandTests(unittest.TestCase):
    def test_sales_products_exclude_kit_pack_types(self):
        service = OrderPlanningService(MagicMock())
        engine = MagicMock()
        service._sales_engine = MagicMock(return_value=engine)

        with patch(
            "app.services.order_planning_service.pd.read_sql",
            return_value=pd.DataFrame(),
        ) as read_sql:
            service._read_sales_products()

        query = str(read_sql.call_args.args[0])
        params = read_sql.call_args.kwargs["params"]
        self.assertIn('"Вид_упаковки"', query)
        self.assertEqual(
            params["excluded_pack_types"],
            ["комплект", "комплект 4+1"],
        )
        engine.connect.assert_called_once_with()

    def test_brand_aliases_have_the_same_automatic_matching_key(self):
        aliases = {
            "VAG": ("Volkswagen", "VW", "AUDI", "SKODA", "ŠKODA"),
            "GM": ("General Motors",),
        }

        for canonical, variants in aliases.items():
            canonical_key = ProductMatchingService._brand_key(canonical)
            for variant in variants:
                with self.subTest(canonical=canonical, variant=variant):
                    self.assertEqual(
                        ProductMatchingService._brand_key(variant),
                        canonical_key,
                    )

    def test_article_lookup_accepts_vag_and_gm_aliases(self):
        matcher = ProductMatchingService(MagicMock())
        vag = SimpleNamespace(product=_product(1, "VOLKSWAGEN PRODUCT", "VAG"))
        gm = SimpleNamespace(product=_product(2, "GENERAL MOTORS PRODUCT", "GM"))
        matcher._article_links_cache = {
            "vag-article": vag,
            "gm-article": gm,
        }
        matcher._article_links_by_brand_cache = {
            ("vag-article", ProductMatchingService._brand_key("VAG")): vag,
            ("gm-article", ProductMatchingService._brand_key("GM")): gm,
        }

        for source_brand in ("Volkswagen", "VW", "AUDI", "SKODA", "ŠKODA"):
            with self.subTest(source_brand=source_brand):
                self.assertIs(
                    matcher._get_article_link_by_exact_article(
                        "VAG-ARTICLE",
                        brand=source_brand,
                    ),
                    vag,
                )
        self.assertIs(
            matcher._get_article_link_by_exact_article(
                "GM-ARTICLE",
                brand="General Motors",
            ),
            gm,
        )

    def test_article_lookup_uses_article_and_brand_as_one_key(self):
        matcher = ProductMatchingService(MagicMock())
        castrol = SimpleNamespace(product=_product(1, "CASTROL PRODUCT", "CASTROL"))
        bmw = SimpleNamespace(product=_product(2, "BMW PRODUCT", "BMW"))
        matcher._article_links_cache = {"shared": bmw}
        matcher._article_links_by_brand_cache = {
            ("shared", "CASTROL"): castrol,
            ("shared", "BMW"): bmw,
        }

        result = matcher._get_article_link_by_exact_article("SHARED", brand="CASTROL")

        self.assertIs(result, castrol)

    def test_reused_sales_code_is_rematched_from_current_article_and_brand(self):
        service = OrderPlanningService(MagicMock())
        old_bmw = _product(2, "BMW GROUP HYPOID REAR AXLE OIL SAF 75W-85 0,5L", "BMW")
        correct_castrol = _product(1, "CASTROL MAGNATEC 10W-40 A3/B4 1L", "CASTROL")
        old_link = SimpleNamespace(
            product=old_bmw,
            sales_article="83120445832",
            sales_product_name="BMW GROUP HYPOID REAR AXLE OIL SAF 75W-85 0,5L",
            sales_pack=0.5,
            sales_brand="BMW",
            sales_is_excise=True,
        )
        service._read_sales_products = MagicMock(return_value=pd.DataFrame([{
            "Код": "new007",
            "Артикул": "1601A7",
            "Продукт_упаковка": "CASTROL MAGNATEC 10W-40 A3/B4 1L",
            "Упаковка": 1,
            "Brand": "CASTROL",
            "Акциз_да_нет": "да",
        }]))
        service._link_map = MagicMock(return_value={"new007": old_link})
        service.matcher.find_customer_product = MagicMock(return_value=correct_castrol)

        result = service.check_products()

        self.assertEqual(len(result.rows), 1)
        self.assertEqual(result.rows[0]["product_id"], correct_castrol.id)
        self.assertEqual(result.rows[0]["product_name"], correct_castrol.name)
        service.matcher.find_customer_product.assert_called_once_with(
            "1601A7",
            "CASTROL MAGNATEC 10W-40 A3/B4 1L",
            1,
            brand="CASTROL",
        )


if __name__ == "__main__":
    unittest.main()
