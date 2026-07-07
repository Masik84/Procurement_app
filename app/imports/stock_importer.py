from __future__ import annotations

from pathlib import Path
import pandas as pd

from app.utils.excel_import import excel_text, read_excel_raw
from app.utils.text import clean_multi_spaces, normalize_product_name


EXCLUDED_BRANDS = {
    "-", "phoenixoil", "gazpromneft", "нефтемастер", "glc", "cnrg",
    "coolstream", "kansler", "mannol", "synthetium", "siberia", "foxy",
    "oilright", "лавр", "lavr", "eltrans", "astrohim", "лукойл",
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

        df = read_excel_raw(file_path, sheet_name=self.sheet_name, header=None)
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
        required = [
            col_brand, col_prod, col_article, col_sku, col_origin,
            col_lpc, col_landed, col_distr, col_promo,
            col_brand_group, col_transit, col_stock_start,
        ]
        if any(x < 0 for x in required):
            raise ValueError("Не найдены обязательные колонки в листе Stock&Price.")
        if col_transit <= col_stock_start:
            raise ValueError("Некорректный диапазон колонок остатков: 'Общий Транзит, л' должен быть правее 'Свободный сток (Новороссийск)'.")

        # Остатки ищем по названиям колонок, но только в безопасном диапазоне:
        # от "Свободный сток (Новороссийск)" до колонки "Общий Транзит, л".
        # Так колонки могут сдвигаться, но мы не захватываем заказы клиентов и другие блоки справа.
        # Важно: работаем с индексами колонок, а не с data.loc[:, names], потому что в Excel
        # бывают повторяющиеся/пустые заголовки, и pandas тогда может подтянуть лишние колонки.
        stock_col_indexes = list(range(col_stock_start, col_transit))

        # В оригинальном файле нет отдельной строки с типами колонок.
        # Поэтому раскладываем складской блок по РЕАЛЬНЫМ названиям колонок.
        # Free Stock / Stock = все складские колонки, которые по бизнес-логике являются "сток":
        #   Свободный сток, Переборка, Чужая маркировка, своб.остаток, Проблема с КМ, Перемещения.
        # Отдельно считаем: Уценка, Резерв, E-com.
        # Заказы клиентов внутри диапазона до "Общий Транзит, л" специально игнорируем.
        markdown_indexes: list[int] = []
        reserve_indexes: list[int] = []
        reserve_ecomm_indexes: list[int] = []
        plain_stock_indexes: list[int] = []

        for i in stock_col_indexes:
            h_norm = _norm_header(headers[i])
            if not h_norm:
                continue

            if _header_contains(h_norm, "УЦЕНКА") or _header_contains(h_norm, "БРАК"):
                markdown_indexes.append(i)
                continue

            if _header_contains(h_norm, "E-COM") or _header_contains(h_norm, "ECOM"):
                reserve_ecomm_indexes.append(i)
                continue

            if _header_contains(h_norm, "РЕЗЕРВ"):
                reserve_indexes.append(i)
                continue

            is_free_stock = (
                (_header_contains(h_norm, "СВОБОДНЫЙ") and _header_contains(h_norm, "СТОК"))
                or (_header_contains(h_norm, "СВОБ") and _header_contains(h_norm, "ОСТАТОК"))
                or _header_contains(h_norm, "ПЕРЕБОРКА")
                or (_header_contains(h_norm, "ЧУЖАЯ") and _header_contains(h_norm, "МАРКИРОВ"))
                or (_header_contains(h_norm, "ПРОБЛЕМА") and _header_contains(h_norm, "КМ"))
                or _header_contains(h_norm, "ПЕРЕМЕЩЕНИ")
            )
            if is_free_stock:
                plain_stock_indexes.append(i)

        compact = pd.DataFrame({
            "source_brand": data.iloc[:, col_brand].map(_norm),
            "source_product_name": data.iloc[:, col_prod].map(_norm),
            "source_article": data.iloc[:, col_article].map(excel_text),
            "source_sku": data.iloc[:, col_sku].map(excel_text),
            "source_origin": data.iloc[:, col_origin].map(_norm),
            "source_brand_group": data.iloc[:, col_brand_group].map(_norm),
            "lpc": self._series_to_number(data.iloc[:, col_lpc]),
            "landed_cost": self._series_to_number(data.iloc[:, col_landed]),
            "distr_price": self._series_to_number(data.iloc[:, col_distr]),
            "promo_price": self._series_to_number(data.iloc[:, col_promo]),
            "transit_qty": self._series_to_number(data.iloc[:, col_transit]),
        })

        def sum_columns_by_index(indexes: list[int]) -> pd.Series:
            if not indexes:
                return pd.Series(0.0, index=data.index)
            return data.iloc[:, indexes].apply(pd.to_numeric, errors="coerce").fillna(0.0).sum(axis=1)

        compact["stock_qty"] = sum_columns_by_index(plain_stock_indexes)
        compact["markdown_qty"] = sum_columns_by_index(markdown_indexes)
        compact["reserve_qty"] = sum_columns_by_index(reserve_indexes)
        compact["reserve_ecomm_qty"] = sum_columns_by_index(reserve_ecomm_indexes)

        compact["total_qty"] = (
            compact["stock_qty"] + compact["markdown_qty"] + compact["reserve_qty"] + compact["reserve_ecomm_qty"] + compact["transit_qty"]
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
                "reserve_ecomm_qty": float(rec["reserve_ecomm_qty"] or 0),
                "has_lpc_warning": bool(rec["has_lpc_warning"]),
            })

        return rows
