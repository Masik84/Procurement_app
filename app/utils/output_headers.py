from __future__ import annotations

import re
from typing import Iterable

from app.utils.excel_format_rules import FORMATS, set_number_format_safe

# Only user-visible GUI/Excel header text is standardized here.
# Do not use this for DB column names or business/internal identifiers.
OUTPUT_HEADER_RENAMES: dict[str, str] = {
    "Pack Price, L": "Price, pack",
    "Price (Pack)": "Price, pack",
    "Price, Pack": "Price, pack",
    "Supplier Product name": "Supplier Product Name",
    "Target Price (Pack)": "Target Price, pack",
    "Cost Novo withVAT": "Cost Novo with VAT",
    "Cost Novo withVAT (prev)": "Cost Novo with VAT (prev)",
}

# Common Excel specs approved in Book1.xlsx. Keys are final visible headers.
HEADER_SPECS = {
    "Article": ((205, 205, 205), None, '@'),
    "Best full Price, L": ((0, 176, 240), None, FORMATS.MONEY_RUB),
    "Best full Price, L 2": ((146, 207, 80), None, FORMATS.MONEY_RUB),
    "Best Suppl": ((0, 176, 240), None, None),
    "Best Suppl 2": ((146, 207, 80), None, None),
    "Brand": ((205, 205, 205), None, None),
    "Cost Novo with VAT": ((205, 205, 205), None, FORMATS.MONEY_RUB),
    "Cost Novo with VAT (prev)": ((166, 166, 166), None, FORMATS.MONEY_RUB),
    "curr Landed cost": ((192, 0, 0), "white", FORMATS.MONEY_RUB),
    "curr LPC": ((192, 0, 0), "white", FORMATS.MONEY_RUB),
    "Currency": ((205, 205, 205), None, None),
    "Currency Best1": ((0, 176, 240), None, None),
    "Currency Best2": ((146, 207, 80), None, None),
    "Damaged": ((33, 92, 152), None, FORMATS.INTEGER),
    "Excise duty": ((205, 205, 205), None, None),
    "fin.Supplier for calc": ((205, 205, 205), None, None),
    "Full Cost Msk": ((205, 205, 205), None, FORMATS.MONEY_RUB),
    "uC3": ((0, 176, 240), None, FORMATS.MONEY_RUB),
    "Full Cost Msk (prev)": ((166, 166, 166), None, FORMATS.MONEY_RUB),
    "FX rate": ((205, 205, 205), None, FORMATS.DECIMAL_1),
    "FX rate Best1": ((0, 176, 240), None, FORMATS.DECIMAL_1),
    "FX rate Best2": ((146, 207, 80), None, FORMATS.DECIMAL_1),
    "ID": ((205, 205, 205), None, None),
    "is_excise": ((205, 205, 205), None, None),
    "last update": ((205, 205, 205), None, FORMATS.DATE),
    "last update (prev)": ((166, 166, 166), None, FORMATS.DATE),
    "last update Best1": ((0, 176, 240), None, FORMATS.DATE),
    "last update Best2": ((146, 207, 80), None, FORMATS.DATE),
    "Material": ((205, 205, 205), None, None),
    "Material number": ((205, 205, 205), None, '@'),
    "Order IS": ((192, 0, 0), "white", FORMATS.INTEGER),
    "Our Product Name": ((205, 205, 205), None, None),
    "Pack": ((205, 205, 205), None, None),
    "Категория ABC": ((205, 205, 205), None, "@"),
    "Price, pack": ((205, 205, 205), None, FORMATS.DECIMAL_2),
    "Price": ((205, 205, 205), None, FORMATS.DECIMAL_2),
    "Price, L": ((205, 205, 205), None, FORMATS.DECIMAL_2),
    "Price, L (prev)": ((166, 166, 166), None, FORMATS.DECIMAL_2),
    "Price date": ((205, 205, 205), None, FORMATS.DATE),
    "Product name": ((205, 205, 205), None, None),
    "Product name (variant)": ((205, 205, 205), None, None),
    "Purchase Order": ((33, 92, 152), None, FORMATS.INTEGER),
    "Qty, pcs": ((205, 205, 205), None, FORMATS.INTEGER),
    "Reserve cust": ((33, 92, 152), None, FORMATS.INTEGER),
    "Reserve E-Comm": ((33, 92, 152), None, FORMATS.INTEGER),
    "Stock": ((33, 92, 152), None, FORMATS.INTEGER),
    "Stock IS": ((192, 0, 0), "white", FORMATS.INTEGER),
    "Supplier Article": ((205, 205, 205), None, '@'),
    "Supplier name": ((205, 205, 205), None, None),
    "Supplier Product Name": ((205, 205, 205), None, None),
    "Target Price, L": ((205, 205, 205), None, FORMATS.DECIMAL_2),
    "Target Price, pack": ((205, 205, 205), None, FORMATS.DECIMAL_2),
    "Transit": ((33, 92, 152), None, FORMATS.INTEGER),
    "Volume, L": ((205, 205, 205), None, FORMATS.INTEGER),
    "Volume to take": ((205, 205, 205), None, FORMATS.INTEGER),
    "Валюта": ((146, 207, 80), None, None),
    "Вид закупки": ((205, 205, 205), None, None),
    "Дата": ((205, 205, 205), None, FORMATS.DATE),
    "Дистр цена": ((192, 0, 0), "white", FORMATS.MONEY_RUB),
    "Клиент": ((205, 205, 205), None, None),
    "Код продукта": ((205, 205, 205), None, '@'),
    "Количество": ((205, 205, 205), None, FORMATS.INTEGER),
    "Комментарии": ((205, 205, 205), None, None),
    "Кост руб л с НДС": ((0, 176, 240), None, FORMATS.MONEY_RUB),
    "Курс": ((146, 207, 80), None, FORMATS.INTEGER),
    "Менеджер": ((205, 205, 205), None, None),
    "Название продукта": ((205, 205, 205), None, None),
    "Объем л": ((205, 205, 205), None, FORMATS.INTEGER),
    "Поставщик": ((146, 207, 80), None, None),
    "Промо цена": ((192, 0, 0), "white", FORMATS.MONEY_RUB),
    "Ср.Продажи мес": ((160, 43, 147), "white", FORMATS.INTEGER),
    "Условия оплаты": ((205, 205, 205), None, None),
    "Фасовка": ((205, 205, 205), None, FORMATS.INTEGER),
    "к Быстрому заказу, л": ((160, 43, 147), "white", FORMATS.INTEGER),
    "к Заказу, л": ((160, 43, 147), "white", FORMATS.INTEGER),
}

_INTEGER_HEADERS = {
    "Qty, pcs", "Volume, L", "StockQty", "TransitQty", "MarkdownQty",
    "ReserveQty", "ReserveECommQty", "OrderQty", "ConfirmedQty", "RemainsQty", "LPC",
}
INTEGER_FORMAT = FORMATS.INTEGER


def standardize_output_header(header: object) -> str:
    text = str(header or "")
    suffix = ""
    # Dynamic columns may be like Cost Novo with VAT_1 or Supplier_2.
    m = re.match(r"^(.*?)(_[0-9]+)$", text)
    if m:
        text, suffix = m.group(1), m.group(2)
    text = OUTPUT_HEADER_RENAMES.get(text, text)
    return text + suffix


def display_headers(headers: Iterable[object]) -> list[str]:
    return [standardize_output_header(h) for h in headers]


def base_header(header: object) -> str:
    text = standardize_output_header(header)
    text = re.sub(r"_[0-9]+$", "", text)
    text = re.sub(r" \([0-9]+ м\)$", "", text)
    return text


def excel_spec(header: object):
    b = base_header(header)
    if b in HEADER_SPECS:
        return HEADER_SPECS[b]
    if b in _INTEGER_HEADERS:
        return ((205, 205, 205), None, INTEGER_FORMAT)
    if b.startswith("FX rate"):
        return ((205, 205, 205), None, FORMATS.DECIMAL_1)
    return None


def rgb_to_excel(rgb: tuple[int, int, int]) -> int:
    r, g, b = rgb
    return int(r) + int(g) * 256 + int(b) * 65536


def apply_header_style_and_formats(ws, headers: list[str], column_letter_func) -> None:
    """Apply Book1-based header fill/font/number formats by visible header names."""
    for idx, raw_header in enumerate(headers, start=1):
        visible = standardize_output_header(raw_header)
        spec = excel_spec(visible)
        if not spec:
            continue
        fill, font, number_format = spec
        letter = column_letter_func(idx)
        try:
            cell = ws.Cells(1, idx)
            cell.Interior.Color = rgb_to_excel(fill)
            if font == "white":
                cell.Font.Color = rgb_to_excel((255, 255, 255))
            else:
                # Automatic/black. Avoid stale white font from fixed color blocks.
                cell.Font.ColorIndex = -4105
        except Exception:
            pass
        if number_format:
            set_number_format_safe(ws.Columns(f"{letter}:{letter}"), number_format)
