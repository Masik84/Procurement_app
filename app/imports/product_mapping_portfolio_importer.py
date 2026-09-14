from __future__ import annotations

from pathlib import Path
from typing import Any
import warnings

import pandas as pd


class ProductMappingPortfolioImporter:
    """Read the Product Mapping source from the 1C Portfolio workbook."""

    SHEETS = (
        "Масла",
        "Масла (TEBOIL)",
        "НЕПРОВЕРЕННАЯ номенклатура",
    )

    @staticmethod
    def _text(value: Any) -> str:
        if value is None:
            return ""
        try:
            if pd.isna(value):
                return ""
        except Exception:
            pass
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return str(value)

    @classmethod
    def _upper_text(cls, value: Any) -> str:
        return cls._text(value).upper()

    @staticmethod
    def _float_series(series: pd.Series) -> pd.Series:
        # The portfolio uses both comma and dot decimal separators in text
        # cells. Do not touch any text fields; numeric source columns are the
        # only columns converted here.
        normalized = series.map(
            lambda value: value.replace(",", ".") if isinstance(value, str) else value
        )
        return pd.to_numeric(normalized, errors="coerce").astype(float)

    def read_excel(self, file_path: str | Path) -> pd.DataFrame:
        file_path = Path(file_path)
        if not file_path.exists():
            raise FileNotFoundError(f"Файл не найден: {file_path}")

        # Portfolio files contain Excel data-validation extensions that
        # openpyxl can read only partially.  We only READ the workbook here and
        # never save it back, so nothing is removed from the user's file.
        # Suppress only this exact openpyxl warning; all other warnings remain.
        with warnings.catch_warnings():
            warnings.filterwarnings(
                "ignore",
                message=r"Data Validation extension is not supported and will be removed",
                category=UserWarning,
                module=r"openpyxl\.worksheet\._reader",
            )

            book = pd.ExcelFile(file_path)
            if any(sheet not in book.sheet_names for sheet in self.SHEETS):
                raise ValueError("выберите файл Портфель")

            frames: list[pd.DataFrame] = []
            for sheet in self.SHEETS:
                frame = pd.read_excel(
                    book,
                    sheet_name=sheet,
                    dtype={"Артикул 1С": str},
                    keep_default_na=False,
                )
                frames.append(frame)

        source = pd.concat(frames, ignore_index=True, sort=False)

        # User-defined source filters.  Text itself is not stripped/collapsed.
        category_upper = source["Категория"].map(self._upper_text)
        pack_type_upper = source["Вид упаковки"].map(self._upper_text)
        type_upper = source["Type"].map(self._upper_text)
        keep_mask = (
            category_upper.ne("АВТОХИМИЯ")
            & ~pack_type_upper.str.contains("КОМПЛЕКТ", regex=False, na=False)
            & ~type_upper.isin({"SPARES", "FILTER", "PROMOTION"})
        )
        source = source.loc[keep_mask].copy()

        # Return the source schema already expected by ProductMappingService.
        # Product/brand names are upper-cased, but spaces and all other text are
        # preserved exactly as they are in the Portfolio.
        result = pd.DataFrame(
            {
                "Код": source["Код 1С"].map(self._text),
                "Артикул": source["Артикул 1С"].map(self._text),
                "Продукт_упаковка": source["Английское наименование продукта"].map(self._upper_text),
                "Упаковка": self._float_series(source["Упаковка"]),
                "Кол_во_в_упак": self._float_series(source["Кол-во шт в упаковке"]),
                "Brand": source["Brand"].map(self._upper_text),
                "Акциз_да_нет": source["Акциз"].map(self._text),
                "Type": source["Type"].map(self._text),
                "Вид упаковки": source["Вид упаковки"].map(self._text),
                "УЕ": source["УЕ"].map(self._text),
                "Плотность": self._float_series(source["Плотность"]),
                "Вес Нетто кг": self._float_series(source["Вес Нетто кг"]),
            }
        )
        return result
