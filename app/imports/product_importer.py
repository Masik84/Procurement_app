from __future__ import annotations

from pathlib import Path

import pandas as pd

from app.services.product_matching_service import ProductMatchingService
from app.utils.parsers import parse_loose_number
from app.utils.text import clean_multi_spaces


class ProductImporter:
    TRUE_VALUES = {"ДА", "TRUE", "1"}
    FALSE_VALUES = {"НЕТ", "FALSE", "0"}

    def _parse_id_value(self, value):
        if value is None:
            return None

        text = clean_multi_spaces(value)
        if not text:
            return None

        try:
            number = parse_loose_number(text)
            if number is None:
                return None
            return int(float(number))
        except Exception:
            return None

    def read_excel(self, file_path: str | Path) -> list[dict]:
        file_path = Path(file_path)
        if not file_path.exists():
            raise FileNotFoundError(f"Файл не найден: {file_path}")

        df = pd.read_excel(file_path, header=0)
        if df.shape[1] < 5:
            raise ValueError(
                "Файл должен содержать минимум 5 колонок: ID, Product name, Brand, Pack, is_excise."
            )

        df = df.iloc[:, :5].copy()
        df.columns = ["id", "name", "brand", "pack", "is_excise"]
        df = df.where(pd.notna(df), None)

        rows: list[dict] = []
        for idx, item in enumerate(df.to_dict(orient="records"), start=2):
            product_id = self._parse_id_value(item.get("id"))
            name = clean_multi_spaces(item.get("name")).upper()
            brand = clean_multi_spaces(item.get("brand")).upper()
            pack_raw = item.get("pack")
            excise_raw = clean_multi_spaces(item.get("is_excise")).upper()

            if not any([product_id is not None, name, brand, pack_raw not in (None, ""), excise_raw]):
                continue

            if not name:
                raise ValueError(f"Строка {idx}: не заполнено поле Product name.")
            if not brand:
                raise ValueError(f"Строка {idx}: не заполнено поле Brand.")

            pack = parse_loose_number(pack_raw)
            if pack is None:
                raise ValueError(f"Строка {idx}: поле Pack должно быть числом.")

            ProductMatchingService.validate_product_name_pack_format(
                product_name=name,
                pack_value=pack,
            )
            family = ProductMatchingService.build_product_family_from_name(name, pack)

            if excise_raw in self.TRUE_VALUES:
                is_excise = True
            elif excise_raw in self.FALSE_VALUES:
                is_excise = False
            else:
                raise ValueError(
                    f"Строка {idx}: поле is_excise должно быть 'да/нет' или 'true/false'."
                )

            rows.append(
                {
                    "id": product_id,
                    "name": name,
                    "brand": brand,
                    "pack": pack,
                    "is_excise": is_excise,
                    "family": family,
                }
            )

        if not rows:
            raise ValueError("Файл не содержит данных для импорта.")

        return rows