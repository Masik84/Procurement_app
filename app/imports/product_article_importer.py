from __future__ import annotations

from pathlib import Path

import pandas as pd

from app.utils.excel_import import excel_text, read_excel_raw
from app.utils.text import clean_multi_spaces


class ProductArticleImporter:
    REQUIRED_COLUMNS = ["Product name", "Article", "Product name (variant)"]

    def read_excel(self, file_path: str | Path) -> list[dict]:
        file_path = Path(file_path)

        if not file_path.exists():
            raise FileNotFoundError(f"Файл не найден: {file_path}")

        df = read_excel_raw(file_path, header=0)

        for column in self.REQUIRED_COLUMNS:
            if column not in df.columns:
                raise ValueError(f"Отсутствует обязательная колонка: {column}")

        df = df[self.REQUIRED_COLUMNS].copy()
        df = df.where(pd.notna(df), None)

        rows: list[dict] = []
        for _, row in df.iterrows():
            product_name = clean_multi_spaces(row["Product name"]).upper()
            article = excel_text(row["Article"])
            variant_name = clean_multi_spaces(row["Product name (variant)"]).upper()

            if not product_name and not article and not variant_name:
                continue

            rows.append(
                {
                    "product_name": product_name,
                    "article": article,
                    "variant_name": variant_name,
                }
            )

        return rows
