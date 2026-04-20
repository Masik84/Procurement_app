from __future__ import annotations

from pathlib import Path
import pandas as pd

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

    @staticmethod
    def _series_to_number(series: pd.Series) -> pd.Series:
        return pd.to_numeric(series, errors="coerce").fillna(0.0)

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
        col_prod = _find_header_index(
            headers,
            "Английское наименование продукта",
            "АНГЛИЙСКОЕ",
            "НАИМЕНОВАНИЕ",
            "ПРОДУКТА",
        )
        col_article = _find_header_index(headers, "Code 1C", "CODE", "1C")
        col_sku = _find_header_index(headers, "SKU", "SKU")
        col_origin = _find_header_index(headers, "Страна происхождения", "СТРАНА", "ПРОИСХ")
        col_lpc = _find_header_index(headers, "LPC", "LPC")
        col_landed = _find_header_index(headers, "Landed Cost+VAT/L", "LANDED", "VAT")
        col_distr = _find_header_index(headers, "Цена Дистр ", "ЦЕНА", "ДИСТР")
        col_promo = _find_header_index(headers, "Цена Промо", "ЦЕНА", "ПРОМО")
        col_brand_group = _find_header_index(headers, "Группа бренда", "ГРУППА", "БРЕНДА")
        col_transit = _find_header_index(headers, "Общий Транзит, л", "ОБЩИЙ", "ТРАНЗИТ", "Л")
        col_stock_start = _find_header_index(
            headers,
            "Свободный сток (Новороссийск)",
            "СВОБОДНЫЙ",
            "СТОК",
            "НОВОРОССИЙСК",
        )
        col_stock_end = _find_header_index(
            headers,
            "Заказы Клиентов Оплачено/Частично оплачено",
            "ЗАКАЗЫ",
            "КЛИЕНТОВ",
            "ОПЛАЧЕНО",
        )

        required = [
            col_brand, col_prod, col_article, col_sku, col_origin,
            col_lpc, col_landed, col_distr, col_promo,
            col_brand_group, col_transit, col_stock_start, col_stock_end,
        ]
        if any(x < 0 for x in required):
            raise ValueError("Не найдены обязательные колонки в листе Stock&Price.")
        if col_stock_end <= col_stock_start:
            raise ValueError("Некорректный диапазон колонок остатков.")

        stock_headers = headers[col_stock_start:col_stock_end]
        stock_columns = headers[col_stock_start:col_stock_end]

        markdown_cols: list[str] = []
        reserve_cols: list[str] = []
        plain_stock_cols: list[str] = []
        for h in stock_headers:
            h_norm = _norm_header(h)
            if _header_contains(h_norm, "УЦЕНКА") or _header_contains(h_norm, "БРАК"):
                markdown_cols.append(h)
            elif _header_contains(h_norm, "РЕЗЕРВ"):
                reserve_cols.append(h)
            else:
                plain_stock_cols.append(h)

        compact = pd.DataFrame({
            "source_brand": data.iloc[:, col_brand].map(_norm),
            "source_product_name": data.iloc[:, col_prod].map(_norm),
            "source_article": data.iloc[:, col_article].map(_norm),
            "source_sku": data.iloc[:, col_sku].map(_norm),
            "source_origin": data.iloc[:, col_origin].map(_norm),
            "source_brand_group": data.iloc[:, col_brand_group].map(_norm),
            "lpc": self._series_to_number(data.iloc[:, col_lpc]),
            "landed_cost": self._series_to_number(data.iloc[:, col_landed]),
            "distr_price": self._series_to_number(data.iloc[:, col_distr]),
            "promo_price": self._series_to_number(data.iloc[:, col_promo]),
            "transit_qty": self._series_to_number(data.iloc[:, col_transit]),
        })

        if stock_columns:
            stock_matrix = data.loc[:, stock_columns].apply(pd.to_numeric, errors="coerce").fillna(0.0)
        else:
            stock_matrix = pd.DataFrame(index=data.index)

        compact["stock_qty"] = stock_matrix[plain_stock_cols].sum(axis=1) if plain_stock_cols else 0.0
        compact["markdown_qty"] = stock_matrix[markdown_cols].sum(axis=1) if markdown_cols else 0.0
        compact["reserve_qty"] = stock_matrix[reserve_cols].sum(axis=1) if reserve_cols else 0.0

        compact["total_qty"] = (
            compact["stock_qty"] + compact["markdown_qty"] + compact["reserve_qty"] + compact["transit_qty"]
        )
        compact["has_lpc_warning"] = (compact["total_qty"] > 0) & (compact["lpc"] == 0)

        non_empty_mask = ~((compact["source_sku"] == "") & (compact["source_article"] == "") & (compact["source_product_name"] == ""))
        filtered = compact.loc[non_empty_mask].copy()
        filtered = filtered.loc[~filtered["source_brand"].map(is_excluded_brand)].copy()

        rows: list[dict] = []
        for idx, rec in filtered.iterrows():
            rows.append({
                "import_row_no": int(idx) + 5,
                "source_article": rec["source_article"] or None,
                "source_sku": rec["source_sku"] or None,
                "source_product_name": rec["source_product_name"] or None,
                "source_origin": rec["source_origin"] or None,
                "source_brand_group": rec["source_brand_group"] or None,
                "lpc": float(rec["lpc"] or 0),
                "landed_cost": float(rec["landed_cost"] or 0),
                "distr_price": float(rec["distr_price"] or 0),
                "promo_price": float(rec["promo_price"] or 0),
                "stock_qty": float(rec["stock_qty"] or 0),
                "transit_qty": float(rec["transit_qty"] or 0),
                "markdown_qty": float(rec["markdown_qty"] or 0),
                "reserve_qty": float(rec["reserve_qty"] or 0),
                "has_lpc_warning": bool(rec["has_lpc_warning"]),
            })

        return rows
