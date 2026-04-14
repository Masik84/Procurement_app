from __future__ import annotations

from pathlib import Path

import pandas as pd

from app.utils.parsers import parse_loose_number
from app.utils.text import clean_multi_spaces


class SupplierPriceImporter:
    """
    Expected file layout:
        column A -> supplier_article
        column B -> product_name
        column C -> price
        column D -> price_pack

    Data starts from row 2 in Excel, so header=0.
    Import stops logically on rows where product_name is empty:
    such rows are simply dropped.
    """

    def read_excel(self, file_path: str | Path) -> list[dict]:
        file_path = Path(file_path)

        if not file_path.exists():
            raise FileNotFoundError(f"Файл не найден: {file_path}")

        df = pd.read_excel(file_path, header=0)

        if df.shape[1] < 4:
            raise ValueError(
                "Файл импорта прайса должен содержать минимум 4 колонки: "
                "article, product_name, price, price_pack."
            )

        df = df.iloc[:, :4].copy()
        df.columns = ["supplier_article", "product_name", "price", "price_pack"]

        df["supplier_article"] = df["supplier_article"].apply(clean_multi_spaces)
        df["product_name"] = df["product_name"].apply(clean_multi_spaces)
        df["price"] = df["price"].apply(parse_loose_number)
        df["price_pack"] = df["price_pack"].apply(parse_loose_number)

        df = df[df["product_name"] != ""].copy()
        df = df.reset_index(drop=True)

        # Excel row numbers: header is row 1, data starts from row 2
        df["import_row_no"] = df.index + 2

        rows = df.to_dict(orient="records")
        return rows