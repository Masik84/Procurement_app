import unittest
from types import SimpleNamespace
from unittest.mock import MagicMock

import pandas as pd

from app.services.order_planning_service import OrderPlanningService
from app.services.product_matching_service import ProductMatchingService


def _product(product_id: int, name: str, brand: str):
    return SimpleNamespace(id=product_id, name=name, brand=brand)


class ProductMatchingBrandTests(unittest.TestCase):
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
