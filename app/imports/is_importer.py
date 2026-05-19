from __future__ import annotations

from pathlib import Path
import pandas as pd

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
        parts: list[pd.DataFrame] = []

        if self.orders_sheet in book.sheet_names:
            df = pd.read_excel(book, sheet_name=self.orders_sheet, header=None)
            if len(df) >= 4:
                headers = [str(x).strip() if pd.notna(x) else "" for x in df.iloc[3].tolist()]
                data = df.iloc[4:].copy().reset_index(drop=True)
                data.columns = headers

                col_code = _find_header_index(headers, "Phoenix code")
                col_products = _find_header_index(headers, "PRODUCTS")
                col_remains = _find_header_index(headers, "Remains w confirmed")
                col_confirmed = _find_header_index(headers, "Confirmed")

                required = [col_code, col_products, col_remains, col_confirmed]
                if any(x < 0 for x in required):
                    raise ValueError("Не найдены обязательные колонки на листе IS orders tracking.")

                orders = pd.DataFrame({
                    'source_article': data.iloc[:, col_code].fillna('').map(_norm),
                    'source_product_name': data.iloc[:, col_products].fillna('').map(_norm),
                    'remains_qty': pd.to_numeric(data.iloc[:, col_remains], errors='coerce').fillna(0).clip(lower=0).astype(float),
                    'confirmed_qty': pd.to_numeric(data.iloc[:, col_confirmed], errors='coerce').fillna(0).clip(lower=0).astype(float),
                })
                orders['stock_qty'] = 0.0
                orders['import_row_no'] = orders.index + 5
                non_empty_mask = (orders['source_article'] != '') | (orders['source_product_name'] != '')
                qty_mask = (orders['remains_qty'] + orders['confirmed_qty']) > 0
                orders = orders.loc[non_empty_mask & qty_mask]
                if not orders.empty:
                    parts.append(orders)

        if self.stock_sheet in book.sheet_names:
            df = pd.read_excel(book, sheet_name=self.stock_sheet, header=None)
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

                stock = pd.DataFrame({
                    'source_article': data.iloc[:, col_material].fillna('').map(_norm),
                    'source_product_name': data.iloc[:, col_description].fillna('').map(_norm),
                    'stock_qty': pd.to_numeric(data.iloc[:, col_volume], errors='coerce').fillna(0).clip(lower=0).astype(float),
                })
                stock['remains_qty'] = 0.0
                stock['confirmed_qty'] = 0.0
                stock['import_row_no'] = stock.index + 3
                non_empty_mask = (stock['source_article'] != '') | (stock['source_product_name'] != '')
                qty_mask = stock['stock_qty'] > 0
                stock = stock.loc[non_empty_mask & qty_mask]
                if not stock.empty:
                    parts.append(stock)

        if not parts:
            return []

        df_all = pd.concat(parts, ignore_index=True, sort=False)
        df_all['source_article_key'] = df_all['source_article'].str.upper()
        df_all['source_name_key'] = df_all['source_product_name'].map(normalize_product_name)
        df_all['key'] = df_all.apply(
            lambda r: f"A|{r['source_article_key']}" if r['source_article_key'] else f"N|{r['source_name_key']}",
            axis=1,
        )

        grouped = (
            df_all.groupby('key', sort=False)
            .agg(
                import_row_no=('import_row_no', 'min'),
                source_article=('source_article', lambda s: next((x for x in s if x), '')),
                source_product_name=('source_product_name', lambda s: next((x for x in s if x), '')),
                remains_qty=('remains_qty', 'sum'),
                confirmed_qty=('confirmed_qty', 'sum'),
                stock_qty=('stock_qty', 'sum'),
            )
            .reset_index(drop=True)
        )

        rows = []
        for rec in grouped.to_dict('records'):
            total = float(rec['remains_qty'] or 0) + float(rec['confirmed_qty'] or 0) + float(rec['stock_qty'] or 0)
            if total <= 0:
                continue
            rows.append({
                'import_row_no': int(rec['import_row_no']),
                'source_article': rec['source_article'] or None,
                'source_product_name': rec['source_product_name'] or None,
                'remains_qty': float(rec['remains_qty'] or 0),
                'confirmed_qty': float(rec['confirmed_qty'] or 0),
                'stock_qty': float(rec['stock_qty'] or 0),
            })
        return rows
