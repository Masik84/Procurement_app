from pathlib import Path

import pandas as pd

from app.utils.parsers import parse_loose_number
from app.utils.text import clean_multi_spaces


class SupplierPriceImporter:
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

        # Берем первые 4 колонки шаблона:
        # 1) supplier_article
        # 2) product_name
        # 3) price
        # 4) price_pack
        df = df.iloc[:, :4].copy()
        df.columns = ["supplier_article", "product_name", "price", "price_pack"]

        # Сохраняем исходный номер строки Excel:
        # header = строка 1, данные начинаются со строки 2
        df["import_row_no"] = df.index + 2

        df["supplier_article"] = df["supplier_article"].apply(clean_multi_spaces)
        df["product_name"] = df["product_name"].apply(clean_multi_spaces)
        df["price"] = df["price"].apply(parse_loose_number)
        df["price_pack"] = df["price_pack"].apply(parse_loose_number)

        # Удаляем только полностью пустые строки:
        # если есть либо article, либо product_name — строку сохраняем
        df = df[
            (df["supplier_article"] != "") |
            (df["product_name"] != "")
        ].copy()

        df = df.where(pd.notna(df), None)

        rows = df.to_dict(orient="records")
        return rows