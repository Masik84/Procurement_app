from __future__ import annotations

from pathlib import Path
import pandas as pd

from app.utils.excel_import import excel_text, read_excel_raw
from app.utils.parsers import parse_loose_number
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


def parse_source_is_excise(value: object, *, product_type: bool = False) -> bool | None:
    normalized = normalize_product_name(value)
    if product_type:
        return normalized in {"pvl", "cvl"}
    if not normalized:
        return None
    if normalized in {"да", "yes", "true", "1"}:
        return True
    if normalized in {"нет", "no", "false", "0"}:
        return False
    return None


# Точное распределение колонок складского блока листа Stock&Price.
# Сопоставление выполняется без учёта регистра и с нормализацией обычных/неразрывных пробелов,
# но по полному названию колонки: никакие другие колонки в остатки не включаются.
STOCK_COLUMN_TO_FIELD = {
    "Свободный сток (Новороссийск)": "stock_qty",
    "Свободный сток Чехов": "stock_qty",
    "Уценка (Ново)": "markdown_qty",
    "Уценка Чехов кат. А": "markdown_qty",
    "Уценка Чехов кат. А1": "markdown_qty",
    "Уценка Чехов кат. Б": "markdown_qty",
    "Уценка Чехов кат. С": "markdown_qty",
    "E-com Чехов": "reserve_ecomm_qty",
    "Переборка (Чехов)": "stock_qty",
    "Резерв (Новороссийск)": "reserve_qty",
    "Резерв (Чехов)": "reserve_qty",
    "Тендер (Чехов)": "stock_qty",
    "Чужая маркировка": "stock_qty",
    "Балашиха свободный сток": "stock_qty",
    "Уценка Балашиха кат. А": "markdown_qty",
    "Уценка Балашиха кат. Б": "markdown_qty",
    "Уценка Балашиха кат. С": "markdown_qty",
    "Балашиха резерв": "reserve_qty",
    "Калининград свободный сток": "stock_qty",
    "Уценка Калининград кат. А": "markdown_qty",
    "Уценка Калининград кат. Б": "markdown_qty",
    "Уценка Калининград кат. С": "markdown_qty",
    "Калининград резерв": "reserve_qty",
    "Проблема с КМ Чехов": "stock_qty",
    "Проблема с КМ Балашиха": "stock_qty",
    "Чужая маркировка Ново": "stock_qty",
    "Перемещения Ново-Чехов": "stock_qty",
    "Перемещения Мск-Клд": "stock_qty",
}


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

        # Рабочие названия колонок находятся в 4-й строке Excel.
        # Именно её используем для основных данных, ABC и складского блока.
        # Первая строка нужна только для двух ценовых колонок, потому что в
        # 4-й строке несколько колонок имеют одинаковое имя "Цена/л c НДС".
        top_headers = [str(x).strip() if pd.notna(x) else "" for x in df.iloc[0].tolist()]
        headers = [str(x).strip() if pd.notna(x) else "" for x in df.iloc[3].tolist()]
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
        col_pack = _find_header_index(headers, "Упаковка", "УПАКОВКА")
        col_excise = _find_header_index(headers, "Акциз (да/нет)", "АКЦИЗ")
        col_product_type = _find_header_index(headers, "Prod.type", "PROD.TYPE")
        col_origin = _find_header_index(headers, "Страна происхождения", "СТРАНА", "ПРОИСХ")
        col_lpc = _find_header_index(headers, "LPC", "LPC")
        col_landed = _find_header_index(headers, "Landed Cost+VAT/L", "LANDED", "VAT")
        col_distr = _find_header_index(top_headers, "Цена Дистр", "ЦЕНА", "ДИСТР")
        col_promo = _find_header_index(top_headers, "Цена Промо", "ЦЕНА", "ПРОМО")
        col_brand_group = _find_header_index(headers, "Группа бренда", "ГРУППА", "БРЕНДА")
        col_abc_category = _find_header_index(headers, "Категория ABC", "КАТЕГОРИЯ", "ABC")
        col_transit = _find_header_index(headers, "Общий Транзит, л", "ОБЩИЙ", "ТРАНЗИТ", "Л")
        col_stock_start = _find_header_index(
            headers,
            "Свободный сток (Новороссийск)",
            "СВОБОДНЫЙ",
            "СТОК",
            "НОВОРОССИЙСК",
        )
        required = {
            "Бренд": col_brand,
            "Английское наименование продукта": col_prod,
            "Code 1C": col_article,
            "SKU": col_sku,
            "Упаковка": col_pack,
            "Страна происхождения": col_origin,
            "LPC": col_lpc,
            "Landed Cost+VAT/L": col_landed,
            "Цена Дистр": col_distr,
            "Цена Промо": col_promo,
            "Группа бренда": col_brand_group,
            "Категория ABC": col_abc_category,
            "Общий Транзит, л": col_transit,
            "Свободный сток (Новороссийск)": col_stock_start,
        }
        missing = [name for name, index in required.items() if index < 0]
        if col_excise < 0 and col_product_type < 0:
            missing.append("Акциз (да/нет) или Prod.type")
        if missing:
            raise ValueError(
                "Не найдены обязательные колонки в листе Stock&Price: "
                + ", ".join(missing)
            )
        if col_transit <= col_stock_start:
            raise ValueError("Некорректный диапазон колонок остатков: 'Общий Транзит, л' должен быть правее 'Свободный сток (Новороссийск)'.")

        # Анализируем только складской блок от "Свободный сток (Новороссийск)"
        # до "Общий Транзит, л" и берём только явно перечисленные выше колонки.
        # Работаем с индексами, потому что в Excel могут встречаться повторяющиеся заголовки.
        stock_indexes_by_field = {
            "stock_qty": [],
            "markdown_qty": [],
            "reserve_qty": [],
            "reserve_ecomm_qty": [],
        }
        normalized_column_map = {_norm_header(column_name): field_name for column_name, field_name in STOCK_COLUMN_TO_FIELD.items()}
        for i in range(col_stock_start, col_transit):
            field_name = normalized_column_map.get(_norm_header(headers[i]))
            if field_name:
                stock_indexes_by_field[field_name].append(i)

        compact = pd.DataFrame({
            "source_brand": data.iloc[:, col_brand].map(_norm),
            "source_pack": data.iloc[:, col_pack].map(parse_loose_number),
            "source_is_excise": (
                data.iloc[:, col_excise].map(parse_source_is_excise)
                if col_excise >= 0
                else data.iloc[:, col_product_type].map(
                    lambda value: parse_source_is_excise(value, product_type=True)
                )
            ),
            "source_product_name": data.iloc[:, col_prod].map(_norm),
            "source_article": data.iloc[:, col_article].map(excel_text),
            "source_sku": data.iloc[:, col_sku].map(excel_text),
            "source_origin": data.iloc[:, col_origin].map(_norm),
            "source_brand_group": data.iloc[:, col_brand_group].map(_norm),
            "abc_category": data.iloc[:, col_abc_category].map(_norm).replace("", "-"),
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

        compact["stock_qty"] = sum_columns_by_index(stock_indexes_by_field["stock_qty"])
        compact["markdown_qty"] = sum_columns_by_index(stock_indexes_by_field["markdown_qty"])
        compact["reserve_qty"] = sum_columns_by_index(stock_indexes_by_field["reserve_qty"])
        compact["reserve_ecomm_qty"] = sum_columns_by_index(stock_indexes_by_field["reserve_ecomm_qty"])

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
                "new_product_name": rec["source_product_name"] or None,
                "new_brand": rec["source_brand"] or None,
                "new_pack": rec["source_pack"],
                "new_is_excise": rec["source_is_excise"],
                "source_origin": rec["source_origin"] or None,
                "source_brand_group": rec["source_brand_group"] or None,
                "abc_category": rec["abc_category"] or "-",
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
