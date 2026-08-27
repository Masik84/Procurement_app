from __future__ import annotations

import unittest

from app.utils.table_sort import numeric_id_value


class NumericIdValueTests(unittest.TestCase):
    def test_text_id_becomes_integer(self) -> None:
        self.assertEqual(numeric_id_value("9883"), 9883)

    def test_temporary_negative_id_becomes_integer(self) -> None:
        self.assertEqual(numeric_id_value(" -12 "), -12)

    def test_invalid_or_empty_id_is_left_unconverted(self) -> None:
        self.assertIsNone(numeric_id_value(""))
        self.assertIsNone(numeric_id_value("99A"))
        self.assertIsNone(numeric_id_value(True))

    def test_ids_from_the_ui_example_have_numeric_order(self) -> None:
        values = ["1", "10", "100", "1000", "1001", "2", "3", "11"]

        result = sorted(values, key=numeric_id_value)

        self.assertEqual(result, ["1", "2", "3", "10", "11", "100", "1000", "1001"])


if __name__ == "__main__":
    unittest.main()
