from __future__ import annotations

from pathlib import Path
import pandas as pd

from app.imports.stock_importer import is_excluded_brand
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


class SupplierOrdersImporter:
    sheet_name = "Закупки в пути"

    def read_excel(self, file_path: str | Path) -> list[dict]:
        file_path = Path(file_path)
        if not file_path.exists():
            raise FileNotFoundError(f"Файл не найден: {file_path}")

        df = pd.read_excel(file_path, sheet_name=self.sheet_name, header=None)
        if len(df) < 3:
            return []

        headers = [str(x).strip() if pd.notna(x) else "" for x in df.iloc[1].tolist()]
        data = df.iloc[2:].copy().reset_index(drop=True)
        data.columns = headers

        col_status = _find_header_index(headers, "Статус", "СТАТУС")
        col_brand = _find_header_index(headers, "Brand", "BRAND")
        col_supplier1 = _find_header_index(headers, "Supplier 1", "SUPPLIER")
        col_article = _find_header_index(headers, "Артикул", "АРТИКУЛ")
        col_prod = _find_header_index(headers, "Назв на англ", "НАЗВ", "АНГЛ")
        col_qty = _find_header_index(headers, "Кол-во, л", "КОЛ", "Л")

        required = [col_status, col_brand, col_supplier1, col_article, col_prod, col_qty]
        if any(x < 0 for x in required):
            raise ValueError("Не найдены обязательные колонки в листе Закупки в пути.")

        agg = {}

        for idx, rec in data.iterrows():
            src_status = _norm(rec.iloc[col_status]).lower()
            if src_status == "":
                break
            if src_status not in {"in transit", "order"}:
                continue

            src_brand = _norm(rec.iloc[col_brand])
            if is_excluded_brand(src_brand):
                continue

            src_supplier = _norm(rec.iloc[col_supplier1])
            if src_status == "order" and normalize_product_name(src_supplier) == "coral":
                continue

            src_article = _norm(rec.iloc[col_article])
            src_prod_name = _norm(rec.iloc[col_prod])
            qty_num = float(parse_loose_number(rec.iloc[col_qty]) or 0)

            key = f"A|{src_article.upper()}" if src_article else f"R|{idx+3}"
            if key not in agg:
                agg[key] = {
                    "import_row_no": idx + 3,
                    "source_article": src_article or None,
                    "source_product_name": src_prod_name or None,
                    "transit_qty": 0.0,
                    "order_qty": 0.0,
                }

            if src_status == "in transit":
                agg[key]["transit_qty"] += qty_num
            else:
                agg[key]["order_qty"] += qty_num

        return list(agg.values())