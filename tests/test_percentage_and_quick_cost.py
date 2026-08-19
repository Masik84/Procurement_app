from decimal import Decimal
from types import SimpleNamespace
import unittest

from app.page_functions.fixed_costs_page import FixedCostsPage
from app.services.cost_calculation_service import CostCalculationService
from app.services.quick_cost_calculation_service import QuickCostCalculationService
from app.services.target_price_service import TargetPriceService
from app.utils.parsers import parse_user_percent


class PercentageInputTests(unittest.TestCase):
    def test_user_percentage_input_always_means_percentage_points(self):
        self.assertEqual(parse_user_percent("1"), Decimal("0.01"))
        self.assertEqual(parse_user_percent("1%"), Decimal("0.01"))
        self.assertEqual(parse_user_percent("0,5"), Decimal("0.005"))
        self.assertEqual(parse_user_percent("3,5"), Decimal("0.035"))
        self.assertEqual(parse_user_percent("3,5%"), Decimal("0.035"))
        self.assertEqual(parse_user_percent("100"), Decimal("1"))

    def test_fixed_cost_percentage_columns_use_the_same_rule(self):
        page = SimpleNamespace(PERCENT_COLUMNS=FixedCostsPage.PERCENT_COLUMNS)

        self.assertEqual(FixedCostsPage._to_decimal(page, "22", "vat"), Decimal("0.22"))
        self.assertEqual(FixedCostsPage._to_decimal(page, "0,25", "bank_fee"), Decimal("0.0025"))
        self.assertEqual(FixedCostsPage._to_decimal(page, "2,4", "money"), Decimal("0.024"))
        self.assertEqual(FixedCostsPage._to_decimal(page, "5", "customs_clearance"), Decimal("0.05"))

    def test_fixed_cost_absolute_columns_are_not_divided_by_100(self):
        page = SimpleNamespace(PERCENT_COLUMNS=FixedCostsPage.PERCENT_COLUMNS)

        self.assertEqual(FixedCostsPage._to_decimal(page, "3,92", "additional_customs"), Decimal("3.92"))
        self.assertEqual(FixedCostsPage._to_decimal(page, "17,45", "storage"), Decimal("17.45"))


class QuickCostCalculationTests(unittest.TestCase):
    def test_geeta_barrel_matches_supplier_price_calculation(self):
        service = QuickCostCalculationService(session=None)
        service.get_supplier = lambda _supplier_id: SimpleNamespace(agent_fee=Decimal("0"))
        service.get_fixed_costs = lambda: SimpleNamespace(
            customs_clearance=Decimal("0.05"),
            additional_customs=Decimal("3.92"),
            excise=Decimal("8.503"),
            eco_fee=Decimal("1.7044"),
            vat=Decimal("0.22"),
            customs_fee=Decimal("2.4"),
            bank_fee=Decimal("0.0025"),
            money=Decimal("0.024"),
            storage=Decimal("17.45"),
            move_novo_tamozh=Decimal("5.06"),
            move_tamozh_chekhov=Decimal("1.04"),
        )
        service.get_marking_cost_by_pack_type = lambda _pack_type: Decimal("0")

        result = service.calculate(
            supplier_price=Decimal("3.36"),
            supplier_id=15,
            pack_type_name="бочка",
            fx_rate=Decimal("85"),
            transport=Decimal("0.28"),
            reexport=parse_user_percent("1"),
            insurance=Decimal("0"),
            fx_markup=parse_user_percent("0"),
            fx_markup_abs=Decimal("0"),
            has_customs=True,
            via_novo=True,
            supplier_is_rf=False,
            marks_for_us=True,
            is_excise=False,
            agent_fee=Decimal("0"),
        )

        self.assertEqual(result.cost_novo_wvat, Decimal("411.0953"))
        self.assertEqual(result.full_cost_msk, Decimal("449.6926"))

    def test_insurance_is_added_from_price_transport_reexport_base(self):
        service = QuickCostCalculationService(session=None)
        service.get_supplier = lambda _supplier_id: SimpleNamespace(agent_fee=Decimal("0"))
        service.get_fixed_costs = lambda: SimpleNamespace(
            customs_clearance=Decimal("0.05"),
            additional_customs=Decimal("0"),
            excise=Decimal("0"),
            eco_fee=Decimal("0"),
            vat=Decimal("0"),
            customs_fee=Decimal("0"),
            bank_fee=Decimal("0"),
            money=Decimal("0"),
            storage=Decimal("0"),
            move_novo_tamozh=Decimal("0"),
            move_tamozh_chekhov=Decimal("0"),
        )
        service.get_marking_cost_by_pack_type = lambda _pack_type: Decimal("0")

        result = service.calculate(
            supplier_price=Decimal("3.36"),
            supplier_id=1,
            pack_type_name="бочка",
            fx_rate=Decimal("1"),
            transport=Decimal("0.52"),
            reexport=parse_user_percent("1.5"),
            insurance=parse_user_percent("1"),
            fx_markup=Decimal("0"),
            fx_markup_abs=Decimal("0"),
            has_customs=True,
            via_novo=False,
            supplier_is_rf=True,
            marks_for_us=True,
            is_excise=False,
            agent_fee=Decimal("0"),
        )

        self.assertEqual(result.cost_novo_wvat, Decimal("4.1745"))
        self.assertEqual(result.full_cost_msk, Decimal("4.1745"))

        standard = CostCalculationService(session=None)
        standard.get_product = lambda _product_id: SimpleNamespace(is_excise=False)
        standard.get_supplier = lambda _supplier_id: SimpleNamespace(
            agent_fee=Decimal("0"),
            marks_for_us=True,
            is_rf=True,
        )
        standard.get_fixed_costs = service.get_fixed_costs
        standard.get_marking_cost = lambda _product_id: Decimal("0")
        standard_cost = standard.calc_cost_novo_wvat(
            supplier_price=Decimal("3.36"),
            product_id=1,
            supplier_id=1,
            transport=Decimal("0.52"),
            reexport=parse_user_percent("1.5"),
            insurance=parse_user_percent("1"),
            fx_rate=Decimal("1"),
            fx_markup=Decimal("0"),
            fx_markup_abs=Decimal("0"),
            has_customs=True,
            agent_fee=Decimal("0"),
        )
        self.assertEqual(standard_cost, result.cost_novo_wvat)

        supplier = SimpleNamespace(is_rf=True, marks_for_us=True, agent_fee=Decimal("0"))
        product = SimpleNamespace(is_excise=False)
        fixed = service.get_fixed_costs()

        class FakeQuery:
            def __init__(self, value):
                self.value = value

            def filter(self, *_args):
                return self

            def order_by(self, *_args):
                return self

            def first(self):
                return self.value

        class FakeSession:
            def query(self, model):
                from app.db.models import FixedCosts, Product, Supplier

                return FakeQuery({Supplier: supplier, Product: product, FixedCosts: fixed}[model])

        target = TargetPriceService(FakeSession())
        target.cost_calculation.get_marking_cost = lambda _product_id: Decimal("0")
        reverse_cost, target_price = target.reverse_calculate_target_price(
            target_supplier_id=1,
            product_id=1,
            full_cost_msk=result.full_cost_msk,
            currency_code="USD",
            fx_rate=Decimal("1"),
            transport=Decimal("0.52"),
            reexport=parse_user_percent("1.5"),
            insurance=parse_user_percent("1"),
            fx_markup=Decimal("0"),
            fx_markup_abs=Decimal("0"),
            has_customs=True,
            via_novo=False,
            agent_fee=Decimal("0"),
        )
        self.assertEqual(reverse_cost, Decimal("4.1745"))
        self.assertEqual(target_price, Decimal("3.3600"))
if __name__ == "__main__":
    unittest.main()
