from __future__ import annotations

from pathlib import Path

import pandas as pd

from app.utils.parsers import parse_loose_number
from app.utils.text import clean_multi_spaces


class PriceHistoryImporter:
    def read_excel(self, file_path: str | Path) -> list[dict]:
        file_path = Path(file_path)

        if not file_path.exists():
            raise FileNotFoundError(f"Файл не найден: {file_path}")

        df = pd.read_excel(file_path, header=0)

        if df.shape[1] < 7:
            raise ValueError(
                "Файл импорта истории цен должен содержать минимум 7 колонок: "
                "supplier_name, article, supplier_product_name, our_product_name, price_date, price, currency."
            )

        df = df.iloc[:, :7].copy()
        df.columns = [
            "supplier_name",
            "supplier_article",
            "supplier_product_name",
            "our_product_name",
            "price_date",
            "price",
            "currency",
        ]

        df["import_row_no"] = df.index + 2

        for col in ["supplier_name", "supplier_article", "supplier_product_name", "our_product_name", "currency"]:
            df[col] = df[col].apply(clean_multi_spaces)

        df["price"] = df["price"].apply(parse_loose_number)
        df["price_date"] = pd.to_datetime(df["price_date"], errors="coerce")

        df = df[
            (df["supplier_article"] != "")
            | (df["supplier_product_name"] != "")
            | (df["our_product_name"] != "")
        ].copy()

        df = df.where(pd.notna(df), None)
        return df.to_dict(orient="records")
