from __future__ import annotations

from pathlib import Path
from typing import Any

from openpyxl import load_workbook

from app.utils.parsers import parse_loose_number
from app.utils.text import clean_multi_spaces


class SupplierPriceImporter:
    """Reads supplier price Excel templates into rows for SupplierPriceService."""

    HEADER_ALIASES = {
        "material number": "supplier_article",
        "article": "supplier_article",
        "supplier article": "supplier_article",
        "material": "product_name",
        "supplier product name": "product_name",
        "product name": "product_name",
        "price, l": "price",
        "price, lt": "price",
        "price l": "price",
        "price, pack": "price_pack",
        "price (pack)": "price_pack",
        "price pack": "price_pack",
        "qty, pcs": "qty_pcs",
        "qty pcs": "qty_pcs",
        "volume, l": "volume_l",
        "volume l": "volume_l",
    }

    @staticmethod
    def _norm_header(value: Any) -> str:
        return clean_multi_spaces(str(value or "")).strip().lower()

    @staticmethod
    def _clean_text(value: Any) -> str | None:
        if value is None:
            return None
        text = clean_multi_spaces(str(value))
        if not text or text.lower() == "nan":
            return None
        return text

    @staticmethod
    def _num(value: Any):
        if value is None or value == "":
            return None
        return parse_loose_number(value)

    def read_excel(self, file_path: str | Path) -> list[dict]:
        wb = load_workbook(file_path, data_only=True, read_only=True)
        ws = wb.active

        header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
        if not header_row:
            return []

        columns: dict[int, str] = {}
        for index, header in enumerate(header_row):
            key = self.HEADER_ALIASES.get(self._norm_header(header))
            if key:
                columns[index] = key

        rows: list[dict] = []
        for excel_row_no, values in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            row = {
                "import_row_no": excel_row_no,
                "supplier_article": None,
                "product_name": None,
                "price": None,
                "price_pack": None,
                "qty_pcs": None,
                "volume_l": None,
            }
            for index, key in columns.items():
                value = values[index] if index < len(values) else None
                if key in {"supplier_article", "product_name"}:
                    row[key] = self._clean_text(value)
                else:
                    row[key] = self._num(value)

            if any(row.get(k) not in (None, "") for k in ("supplier_article", "product_name", "price", "price_pack", "qty_pcs", "volume_l")):
                rows.append(row)

        wb.close()
        return rows
