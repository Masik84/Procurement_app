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

        out = pd.DataFrame({
            "product_name": df["Product name"].map(clean_multi_spaces).str.upper(),
            "article": df["Article"].map(excel_text),
            "variant_name": df["Product name (variant)"].map(clean_multi_spaces).str.upper(),
        })
        non_empty_mask = (
            (out["product_name"] != "") | (out["article"] != "") | (out["variant_name"] != "")
        )
        return out.loc[non_empty_mask].to_dict(orient="records")
