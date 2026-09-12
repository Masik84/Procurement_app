from __future__ import annotations

from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path

from app.utils.excel_import import read_excel_raw
from app.utils.text import clean_multi_spaces


class ProductUc3Importer:
    """Read Product Target/Walk-Away uC3 values from the code-generated template."""

    HEADERS = ["Product Name", "Target uC3", "Walk-Away uC3"]

    @staticmethod
    def _round_integer(value):
        if value is None or value == "":
            return None
        try:
            if str(value).strip().lower() in {"nan", "none"}:
                return None
            number = Decimal(str(value).replace(" ", "").replace(",", "."))
            if not number.is_finite():
                return None
            return int(number.to_integral_value(rounding=ROUND_HALF_UP))
        except (InvalidOperation, ValueError, TypeError):
            raise ValueError(f"Некорректное числовое значение uC3: {value}")

    def read_excel(self, file_path: str | Path) -> list[dict]:
        file_path = Path(file_path)
        if not file_path.exists():
            raise FileNotFoundError(f"Файл не найден: {file_path}")

        df = read_excel_raw(file_path, header=0)
        if df.shape[1] < 3:
            raise ValueError(
                "Файл Target uC3 должен содержать 3 колонки: "
                "Product Name, Target uC3, Walk-Away uC3."
            )

        # Read by the first three columns so the generated template and a user-filled
        # copy remain compatible even if Excel changes header formatting.
        df = df.iloc[:, :3].copy()
        df.columns = ["product_name", "target_uc3", "walk_away_uc3"]

        rows: list[dict] = []
        for index, row in df.iterrows():
            product_name = clean_multi_spaces(row.get("product_name"))
            if not product_name:
                continue
            rows.append(
                {
                    "import_row_no": int(index) + 2,
                    "product_name": product_name,
                    "target_uc3": self._round_integer(row.get("target_uc3")),
                    "walk_away_uc3": self._round_integer(row.get("walk_away_uc3")),
                }
            )
        return rows
