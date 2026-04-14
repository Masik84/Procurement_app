from __future__ import annotations

from pathlib import Path
import pandas as pd

from app.utils.parsers import parse_loose_number
from app.utils.text import clean_multi_spaces, normalize_product_name


EXCLUDED_BRANDS = {
    "-", "phoenixoil", "gazpromneft", "нефтемастер", "teboil", "glc", "cnrg",
    "coolstream", "kansler", "mannol", "synthetium", "siberia", "foxy",
    "oilright", "лавр", "lavr", "eltrans", "astrohim",
}


def _norm(v: object) -> str:
    return clean_multi_spaces(v)


def _norm_header(v: object) -> str:
    return _norm(v).upper()


def _header_contains(header: str, token: str) -> bool:
    return _norm_header(token) in _norm_header(header)


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


def is_excluded_brand(brand: object) -> bool:
    raw = _norm(brand)
    if raw == "-":
        return True
    return normalize_product_name(raw) in EXCLUDED_BRANDS


class StockImporter:
    sheet_name = "Stock&Price"

    def read_excel(self, file_path: str | Path) -> list[dict]:
        file_path = Path(file_path)
        if not file_path.exists():
            raise FileNotFoundError(f"Файл не найден: {file_path}")

        df = pd.read_excel(file_path, sheet_name=self.sheet_name, header=None)
        if len(df) < 5:
            return []

        headers = [str(x).strip() if pd.notna(x) else "" for x in df.iloc[0].tolist()]
        data = df.iloc[4:].copy().reset_index(drop=True)
        data.columns = headers

        col_brand = _find_header_index(headers, "Бренд", "БРЕНД")
        col_prod = _find_header_index(headers, "Английское наименование продукта", "АНГЛИЙСКОЕ", "НАИМЕНОВАНИЕ", "ПРОДУКТА")
        col_article = _find_header_index(headers, "Code 1C", "CODE", "1C")
        col_sku = _find_header_index(headers, "SKU", "SKU")
        col_origin = _find_header_index(headers, "Страна происхождения", "СТРАНА", "ПРОИСХ")
        col_lpc = _find_header_index(headers, "LPC", "LPC")
        col_landed = _find_header_index(headers, "Landed Cost+VAT/L", "LANDED", "VAT")
        col_distr = _find_header_index(headers, "Цена Дистр ", "ЦЕНА", "ДИСТР")
        col_promo = _find_header_index(headers, "Цена Промо", "ЦЕНА", "ПРОМО")
        col_brand_group = _find_header_index(headers, "Группа бренда", "ГРУППА", "БРЕНДА")
        col_stock_start = _find_header_index(headers, "Свободный сток (Новороссийск)", "СВОБОДНЫЙ", "СТОК", "НОВОРОССИЙСК")
        col_stock_end = _find_header_index(headers, "Заказы Клиентов Оплачено/Частично оплачено", "ЗАКАЗЫ", "КЛИЕНТОВ", "ОПЛАЧЕНО")

        required = [col_brand, col_prod, col_article, col_sku, col_origin, col_lpc, col_landed, col_distr, col_promo, col_brand_group, col_stock_start, col_stock_end]
        if any(x < 0 for x in required):
            raise ValueError("Не найдены обязательные колонки в листе Stock&Price.")
        if col_stock_end <= col_stock_start:
            raise ValueError("Некорректный диапазон колонок остатков.")

        rows = []
        stock_headers = headers[col_stock_start:col_stock_end]

        for idx, rec in data.iterrows():
            src_sku = _norm(rec.iloc[col_sku])
            if src_sku == "":
                break

            src_brand = _norm(rec.iloc[col_brand])
            if is_excluded_brand(src_brand):
                continue

            src_article = _norm(rec.iloc[col_article])
            src_prod_name = _norm(rec.iloc[col_prod])
            src_origin = _norm(rec.iloc[col_origin])
            src_brand_group = _norm(rec.iloc[col_brand_group])

            lpc_val = parse_loose_number(rec.iloc[col_lpc])
            landed_val = parse_loose_number(rec.iloc[col_landed])
            distr_val = parse_loose_number(rec.iloc[col_distr])
            promo_val = parse_loose_number(rec.iloc[col_promo])

            stock_qty = 0.0
            markdown_qty = 0.0
            reserve_qty = 0.0

            for offset, h in enumerate(stock_headers):
                q = parse_loose_number(rec.iloc[col_stock_start + offset])
                if q is None:
                    continue
                h_norm = _norm_header(h)
                if _header_contains(h_norm, "УЦЕНКА") or _header_contains(h_norm, "БРАК"):
                    markdown_qty += float(q)
                elif _header_contains(h_norm, "РЕЗЕРВ"):
                    reserve_qty += float(q)
                else:
                    stock_qty += float(q)

            total_qty = stock_qty + markdown_qty + reserve_qty

            rows.append({
                "import_row_no": idx + 5,
                "source_article": src_article or None,
                "source_sku": src_sku or None,
                "source_product_name": src_prod_name or None,
                "source_origin": src_origin or None,
                "source_brand_group": src_brand_group or None,
                "lpc": lpc_val,
                "landed_cost": landed_val,
                "distr_price": distr_val,
                "promo_price": promo_val,
                "stock_qty": stock_qty,
                "markdown_qty": markdown_qty,
                "reserve_qty": reserve_qty,
                "has_lpc_warning": bool(total_qty > 0 and (lpc_val is None or float(lpc_val) == 0)),
            })

        return rows