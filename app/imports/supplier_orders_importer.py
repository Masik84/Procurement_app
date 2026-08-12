from __future__ import annotations

from pathlib import Path
import pandas as pd

from app.imports.stock_importer import is_excluded_brand, parse_source_is_excise
from app.utils.excel_import import excel_text, read_excel_raw
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

        df = read_excel_raw(file_path, sheet_name=self.sheet_name, header=None)
        if len(df) < 3:
            return []

        headers = [str(x).strip() if pd.notna(x) else "" for x in df.iloc[1].tolist()]
        data = df.iloc[2:].copy().reset_index(drop=True)
        data.columns = headers

        col_status = _find_header_index(headers, "Статус", "СТАТУС")
        col_brand = _find_header_index(headers, "Brand", "BRAND")
        col_supplier1 = _find_header_index(headers, "Supplier 1", "SUPPLIER")
        col_article = _find_header_index(headers, "Артикул", "АРТИКУЛ")
        col_our_product = _find_header_index(headers, "Продукт + упаковка", "ПРОДУКТ", "УПАКОВКА")
        col_pack = _find_header_index(headers, "Упаковка", "УПАКОВКА")
        col_excise = _find_header_index(headers, "Акциз (да/нет)", "АКЦИЗ")
        col_abc_category = _find_header_index(headers, "ABC")
        col_prod = _find_header_index(headers, "Назв на англ", "НАЗВ", "АНГЛ")
        if col_prod < 0:
            col_prod = _find_header_index(headers, "Product name Назв на англ", "PRODUCT", "NAME")
        col_qty = _find_header_index(headers, "Кол-во, л", "КОЛ", "Л")

        required = [
            col_status, col_brand, col_supplier1, col_article, col_our_product, col_pack, col_excise,
            col_abc_category, col_prod, col_qty,
        ]
        if any(x < 0 for x in required):
            raise ValueError("Не найдены обязательные колонки в листе Закупки в пути.")

        df2 = pd.DataFrame({
            'status': data.iloc[:, col_status].fillna('').map(_norm).str.lower(),
            'brand': data.iloc[:, col_brand].fillna('').map(_norm),
            'pack': data.iloc[:, col_pack].map(parse_loose_number),
            'is_excise': data.iloc[:, col_excise].map(parse_source_is_excise),
            'supplier1': data.iloc[:, col_supplier1].fillna('').map(_norm),
            'article': data.iloc[:, col_article].map(excel_text),
            'our_product_name': data.iloc[:, col_our_product].fillna('').map(_norm),
            'abc_category': data.iloc[:, col_abc_category].fillna('').map(_norm).replace('', '-'),
            'product_name': data.iloc[:, col_prod].fillna('').map(_norm),
            'order_qty': pd.to_numeric(data.iloc[:, col_qty], errors='coerce').fillna(0).astype(float),
        })
        df2['import_row_no'] = df2.index + 3

        non_empty_mask = (
            (df2['status'] != '')
            | (df2['article'] != '')
            | (df2['our_product_name'] != '')
            | (df2['product_name'] != '')
        )
        df2 = df2.loc[non_empty_mask]
        if df2.empty:
            return []

        supplier_is_coral = df2['supplier1'].map(normalize_product_name) == 'coral'
        df2 = df2.loc[
            (~supplier_is_coral & (df2['status'] == 'order'))
            | (supplier_is_coral & df2['status'].isin(('order', 'confirmed')))
        ].copy()
        if df2.empty:
            return []

        df2 = df2.loc[~df2['brand'].map(is_excluded_brand)]
        if df2.empty:
            return []

        supplier_is_coral = df2['supplier1'].map(normalize_product_name) == 'coral'
        df2['is_order_qty'] = df2['order_qty'].where(
            supplier_is_coral & (df2['status'] == 'order'),
            0,
        )
        df2['is_confirmed_order_qty'] = df2['order_qty'].where(
            supplier_is_coral & (df2['status'] == 'confirmed'),
            0,
        )
        # CORAL quantities belong only to the dedicated IS fields.
        df2['order_qty'] = df2['order_qty'].where(~supplier_is_coral, 0)

        df2['key'] = df2.apply(
            lambda r: f"A|{r['article'].upper()}" if r['article'] else f"R|{int(r['import_row_no'])}",
            axis=1,
        )

        grouped = (
            df2.groupby('key', sort=False)
            .agg(
                import_row_no=('import_row_no', 'min'),
                source_article=('article', lambda s: next((x for x in s if x), '')),
                source_brand=('brand', lambda s: next((x for x in s if x), '')),
                source_pack=('pack', lambda s: next((x for x in s if pd.notna(x)), None)),
                source_is_excise=('is_excise', lambda s: next((x for x in s if pd.notna(x)), None)),
                source_our_product_name=('our_product_name', lambda s: next((x for x in s if x), '')),
                abc_category=('abc_category', lambda s: next((x for x in s if x and x != '-'), '-')),
                source_product_name=('product_name', lambda s: next((x for x in s if x), '')),
                order_qty=('order_qty', 'sum'),
                is_order_qty=('is_order_qty', 'sum'),
                is_confirmed_order_qty=('is_confirmed_order_qty', 'sum'),
            )
            .reset_index(drop=True)
        )

        rows = []
        for rec in grouped.to_dict('records'):
            rows.append({
                'import_row_no': int(rec['import_row_no']),
                'source_article': rec['source_article'] or None,
                'source_our_product_name': rec['source_our_product_name'] or None,
                'abc_category': rec['abc_category'] or '-',
                'source_product_name': rec['source_product_name'] or None,
                'new_product_name': rec['source_product_name'] or None,
                'new_brand': rec['source_brand'] or None,
                'new_pack': rec['source_pack'],
                'new_is_excise': rec['source_is_excise'],
                'order_qty': float(rec['order_qty'] or 0),
                'is_order_qty': float(rec['is_order_qty'] or 0),
                'is_confirmed_order_qty': float(rec['is_confirmed_order_qty'] or 0),
            })
        return rows
