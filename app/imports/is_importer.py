from __future__ import annotations

from pathlib import Path
import pandas as pd

from app.utils.parsers import parse_loose_number
from app.utils.text import clean_multi_spaces, normalize_product_name


def _norm(v: object) -> str:
    return clean_multi_spaces(v)


def _norm_header(v: object) -> str:
    return _norm(v).upper()


def _find_header_index(columns: list[str], exact: str, *contains: str) -> int:
    exact_norm = _norm_header(exact)
    for i, col in enumerate(columns):
        if _norm_header(col) == exact_norm:
            return i
    if contains:
        for i, col in enumerate(columns):
            col_norm = _norm_header(col)
            ok = True
            for token in contains:
                ok = ok and (_norm_header(token) in col_norm)
            if ok:
                return i
    return -1


class ISImporter:
    orders_sheet = "IS orders tracking"
    stock_sheet = "Stock IS"

    def read_excel(self, file_path: str | Path) -> list[dict]:
        file_path = Path(file_path)
        if not file_path.exists():
            raise FileNotFoundError(f"Файл не найден: {file_path}")

        book = pd.ExcelFile(file_path)
        agg = {}

        if self.orders_sheet in book.sheet_names:
            df = pd.read_excel(file_path, sheet_name=self.orders_sheet, header=None)
            if len(df) >= 4:
                headers = [str(x).strip() if pd.notna(x) else "" for x in df.iloc[2].tolist()]
                data = df.iloc[3:].copy().reset_index(drop=True)
                data.columns = headers

                col_code = _find_header_index(headers, "Phoenix code")
                col_products = _find_header_index(headers, "PRODUCTS")
                col_remains = _find_header_index(headers, "Remains w confirmed", "REMAINS", "CONFIRMED")
                col_confirmed = _find_header_index(headers, "Confirmed")

                required = [col_code, col_products, col_remains, col_confirmed]
                if any(x < 0 for x in required):
                    raise ValueError("Не найдены обязательные колонки на листе IS orders tracking.")

                for idx, rec in data.iterrows():
                    src_article = _norm(rec.iloc[col_code])
                    src_prod_name = _norm(rec.iloc[col_products])
                    remains_qty = max(float(parse_loose_number(rec.iloc[col_remains]) or 0), 0)
                    confirmed_qty = max(float(parse_loose_number(rec.iloc[col_confirmed]) or 0), 0)

                    if src_article == "" and src_prod_name == "":
                        continue
                    if remains_qty + confirmed_qty == 0:
                        continue

                    key = f"A|{src_article.upper()}" if src_article else f"N|{normalize_product_name(src_prod_name)}"
                    if key not in agg:
                        agg[key] = {
                            "import_row_no": idx + 4,
                            "source_article": src_article or None,
                            "source_product_name": src_prod_name or None,
                            "remains_qty": 0.0,
                            "confirmed_qty": 0.0,
                            "stock_qty": 0.0,
                        }

                    agg[key]["remains_qty"] += remains_qty
                    agg[key]["confirmed_qty"] += confirmed_qty

        if self.stock_sheet in book.sheet_names:
            df = pd.read_excel(file_path, sheet_name=self.stock_sheet, header=None)
            if len(df) >= 3:
                headers = [str(x).strip() if pd.notna(x) else "" for x in df.iloc[1].tolist()]
                data = df.iloc[2:].copy().reset_index(drop=True)
                data.columns = headers

                col_material = _find_header_index(headers, "Material", "MATERIAL")
                col_description = _find_header_index(headers, "Description", "DESCRIPTION")
                col_volume = _find_header_index(headers, "Volume", "VOLUME")

                required = [col_material, col_description, col_volume]
                if any(x < 0 for x in required):
                    raise ValueError("Не найдены обязательные колонки на листе Stock IS.")

                for idx, rec in data.iterrows():
                    src_article = _norm(rec.iloc[col_material])
                    src_prod_name = _norm(rec.iloc[col_description])
                    stock_qty = max(float(parse_loose_number(rec.iloc[col_volume]) or 0), 0)

                    if src_article == "" and src_prod_name == "":
                        continue
                    if stock_qty == 0:
                        continue

                    key = f"A|{src_article.upper()}" if src_article else f"N|{normalize_product_name(src_prod_name)}"
                    if key not in agg:
                        agg[key] = {
                            "import_row_no": idx + 3,
                            "source_article": src_article or None,
                            "source_product_name": src_prod_name or None,
                            "remains_qty": 0.0,
                            "confirmed_qty": 0.0,
                            "stock_qty": 0.0,
                        }

                    if not agg[key]["source_article"] and src_article:
                        agg[key]["source_article"] = src_article
                    if not agg[key]["source_product_name"] and src_prod_name:
                        agg[key]["source_product_name"] = src_prod_name

                    agg[key]["stock_qty"] += stock_qty

        return [row for row in agg.values() if row["remains_qty"] + row["confirmed_qty"] + row["stock_qty"] > 0]