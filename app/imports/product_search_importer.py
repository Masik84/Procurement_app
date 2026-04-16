from pathlib import Path

import pandas as pd

from app.utils.text import clean_multi_spaces


class ProductSearchImporter:
    def read_excel(self, file_path: str | Path) -> list[dict]:
        file_path = Path(file_path)

        if not file_path.exists():
            raise FileNotFoundError(f"Файл не найден: {file_path}")

        df = pd.read_excel(file_path, header=0)

        if df.shape[1] < 2:
            raise ValueError(
                "Файл поиска продуктов должен содержать минимум 2 колонки: article, product_name."
            )

        df = df.iloc[:, :2].copy()
        df.columns = ["source_article", "source_product_name"]

        df["source_article"] = df["source_article"].apply(clean_multi_spaces)
        df["source_product_name"] = df["source_product_name"].apply(clean_multi_spaces)

        df = df[(df["source_article"] != "") | (df["source_product_name"] != "")].copy()
        df = df.reset_index(drop=True)
        df = df.where(pd.notna(df), None)
        df["import_row_no"] = df.index + 2

        return df.to_dict(orient="records")
