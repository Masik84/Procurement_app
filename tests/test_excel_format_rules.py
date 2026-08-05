from __future__ import annotations

import unittest

from app.utils.excel_format_rules import (
    ExcelNumberFormatError,
    FORMATS,
    cost_calc_headers,
    set_number_format_safe,
    to_invariant_number_format,
    to_local_number_format,
)


class FakeExcelApplication:
    def __init__(self, decimal_separator: str, thousands_separator: str):
        self.values = {3: decimal_separator, 4: thousands_separator}

    def International(self, key):
        return self.values[key]


class FakeTupleExcelApplication:
    International = ("country", "date-order", ",", " ", ";")


class FakeExcelTarget:
    def __init__(
        self,
        *,
        reject_invariant: bool = False,
        silent_general: bool = False,
        separators: tuple[str, str] | None = None,
    ):
        self._number_format = FORMATS.GENERAL
        self.reject_invariant = reject_invariant
        self.silent_general = silent_general
        self.assignments: list[tuple[str, str]] = []
        if separators is not None:
            self.Application = FakeExcelApplication(*separators)

    @property
    def NumberFormat(self):
        return self._number_format

    @NumberFormat.setter
    def NumberFormat(self, value):
        self.assignments.append(("NumberFormat", value))
        if self.reject_invariant:
            raise RuntimeError("invariant format rejected")
        self._number_format = FORMATS.GENERAL if self.silent_general else value

    @property
    def NumberFormatLocal(self):
        return self._number_format

    @NumberFormatLocal.setter
    def NumberFormatLocal(self, value):
        self.assignments.append(("NumberFormatLocal", value))
        self._number_format = FORMATS.GENERAL if self.silent_general else value


class ExcelFormatRulesTests(unittest.TestCase):
    def test_converts_project_local_formats_to_invariant(self):
        cases = {
            '# ##0,00;[Red]-# ##0,00;"-"': FORMATS.DECIMAL_2,
            '# ##0;[Red]-# ##0;"-"': FORMATS.INTEGER,
            '# ##0 ₽;[Red]-# ##0 ₽;"-"': FORMATS.MONEY_RUB,
            "ДД.ММ.ГГ;@": FORMATS.DATE,
            "# ##0,0###": FORMATS.FX_FLEX,
        }
        for local, invariant in cases.items():
            with self.subTest(local=local):
                self.assertEqual(to_invariant_number_format(local), invariant)

    def test_uses_invariant_number_format_first(self):
        target = FakeExcelTarget()
        applied = set_number_format_safe(target, '# ##0,00;[Red]-# ##0,00;"-"')
        self.assertEqual(target.assignments[0], ("NumberFormat", FORMATS.DECIMAL_2))
        self.assertEqual(applied, FORMATS.DECIMAL_2)

    def test_builds_russian_local_fallback(self):
        self.assertEqual(
            to_local_number_format(FORMATS.DECIMAL_2),
            '# ##0,00;[Red]\\-# ##0,00;"-"',
        )

    def test_real_excel_locale_uses_number_format_local_first(self):
        target = FakeExcelTarget(separators=(",", " "))
        set_number_format_safe(target, FORMATS.DECIMAL_2)
        self.assertEqual(target.assignments[0][0], "NumberFormatLocal")
        self.assertEqual(
            target.assignments[0][1],
            '# ##0,00;[Red]\\-# ##0,00;"-"',
        )

    def test_pywin32_tuple_international_property_uses_zero_based_indexes(self):
        target = FakeExcelTarget()
        target.Application = FakeTupleExcelApplication()
        set_number_format_safe(target, FORMATS.INTEGER)
        self.assertEqual(
            target.assignments[0],
            ("NumberFormatLocal", '# ##0;[Red]\\-# ##0;"-"'),
        )

    def test_uses_local_fallback_when_invariant_is_rejected(self):
        target = FakeExcelTarget(reject_invariant=True)
        applied = set_number_format_safe(
            target,
            FORMATS.DECIMAL_2,
            '# ##0,00;[Red]-# ##0,00;"-"',
        )
        self.assertEqual(target.assignments[-1][0], "NumberFormatLocal")
        self.assertFalse(applied.casefold() == FORMATS.GENERAL.casefold())

    def test_rejects_silent_general_fallback(self):
        target = FakeExcelTarget(silent_general=True)
        with self.assertRaises(ExcelNumberFormatError):
            set_number_format_safe(target, FORMATS.DECIMAL_2)

    def test_general_format_is_allowed_when_requested(self):
        target = FakeExcelTarget()
        self.assertEqual(set_number_format_safe(target, FORMATS.GENERAL), FORMATS.GENERAL)

    def test_cost_calc_order_columns_are_before_volume_to_take(self):
        headers, standard_order_header = cost_calc_headers(
            quick_order_months=3,
            safe_stock_months=5,
        )
        volume_index = headers.index("Volume to take")
        self.assertEqual(
            headers[volume_index - 2:volume_index],
            ["Ср.Продажи мес", standard_order_header],
        )
        self.assertEqual(standard_order_header, "к Заказу, л (5 м)")

    def test_cost_calc_quick_order_column_is_hidden(self):
        headers, _ = cost_calc_headers(quick_order_months=3, safe_stock_months=5)
        self.assertFalse(any(header.startswith("к Быстрому заказу") for header in headers))


if __name__ == "__main__":
    unittest.main()
