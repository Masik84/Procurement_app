from __future__ import annotations

import logging
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any

logger = logging.getLogger(__name__)


"""Central Excel display rules for Procurement App.

This module is the single source of truth for reusable Excel presentation:
visible header captions, number formats, widths, alignment and value-based
highlighting. Exporters may still define report structure/order, but standard
column presentation must be resolved here.
"""


@dataclass(frozen=True)
class ExcelFormats:
    GENERAL: str = "General"
    TEXT: str = "@"
    DATE: str = "ДД.ММ.ГГ;@"
    INTEGER: str = '# ##0;[Red]-# ##0;"-"'
    DECIMAL_1: str = '# ##0,0;[Red]-# ##0,0;"-"'
    DECIMAL_2: str = '# ##0,00;[Red]-# ##0,00;"-"'
    DECIMAL_2_SIMPLE: str = "0,00"
    DECIMAL_2_PLAIN: str = "# ##0,00"
    PRICE_DECIMAL: str = '# ##0,00_ ;[Red]-# ##0,00_ ;"-"'
    DECIMAL_4: str = '# ##0,0000;[Red]-# ##0,0000;"-"'
    # Optional decimal digits: 10 -> 10; 0.75 -> 0,75; 203.7 -> 203,7.
    DECIMAL_OPTIONAL_4: str = '# ##0,####;[Red]-# ##0,####;"-"'
    DECIMAL_FLEX: str = '# ##0,00##;[Red]-# ##0,00##;0'
    MONEY_RUB: str = '# ##0 ₽;[Red]-# ##0 ₽;"-"'
    MONEY_RUB_SIMPLE: str = "# ##0 ₽"
    PERCENT_1: str = "0,0%"
    PERCENT_2: str = "0,00%"
    PERCENT_FLEX: str = "0,0#%"
    FX_INTEGER: str = "# ##0"
    FX_FLEX: str = "# ##0,0###"


FORMATS = ExcelFormats()
LOCAL_FORMATS = FORMATS


class ExcelNumberFormatError(RuntimeError):
    """Kept for compatibility with existing callers."""


# ---------------------------------------------------------------------------
# Header normalization / visible captions
# ---------------------------------------------------------------------------

DISPLAY_HEADER_RENAMES: dict[str, str] = {
    "Pack Price, L": "Price, pack",
    "Price (Pack)": "Price, pack",
    "Price, Pack": "Price, pack",
    "Supplier Product name": "Supplier Product Name",
    "Target Price (Pack)": "Target Price, pack",
    "Cost Novo withVAT": "Cost Novo with VAT",
    "Cost Novo withVAT (prev)": "Cost Novo with VAT (prev)",
    "Категория ABC": "кат. ABC",
}

# Final visible aliases must resolve back to one logical key for formatting.
_RULE_HEADER_ALIASES = {
    "кат. ABC": "Категория ABC",
}

_DYNAMIC_SUFFIX_RE = re.compile(r"^(?P<base>.+?)(?P<suffix>_\d+)$")
_MONTH_SUFFIX_RE = re.compile(r"^(?P<base>.+?) \(\d+ м\)$")


def split_dynamic_suffix(header: object) -> tuple[str, str]:
    text = str(header or "").strip()
    match = _DYNAMIC_SUFFIX_RE.match(text)
    if match:
        return match.group("base"), match.group("suffix")
    return text, ""


def standardize_output_header(header: object) -> str:
    text, suffix = split_dynamic_suffix(header)
    text = DISPLAY_HEADER_RENAMES.get(text, text)
    return text + suffix


def normalize_rule_header(header: object) -> str:
    """Return logical base header used by the central formatting rules."""
    text, _suffix = split_dynamic_suffix(header)
    text = DISPLAY_HEADER_RENAMES.get(text, text)
    text = _RULE_HEADER_ALIASES.get(text, text)
    month_match = _MONTH_SUFFIX_RE.match(text)
    if month_match:
        text = month_match.group("base")
    return " ".join(text.replace("\n", " ").split()).strip()


def has_dynamic_suffix(header: object) -> bool:
    _base, suffix = split_dynamic_suffix(header)
    return bool(suffix)


def _compact_header(header: object) -> str:
    text = normalize_rule_header(header).casefold().replace("с", "c")
    return re.sub(r"[^0-9a-zа-яё]+", "", text)


def is_uc3_header(header: object) -> bool:
    return "uc3" in _compact_header(header)


def is_article_header_name(header: object) -> bool:
    compact = _compact_header(header)
    return "article" in compact or "артикул" in compact


# ---------------------------------------------------------------------------
# Central semantic groups
# ---------------------------------------------------------------------------

DATE_HEADERS = {
    "Дата",
    "Calc date",
    "Price date",
    "last update",
    "last update (prev)",
    "last update Best1",
    "last update Best2",
    "Change date",
}

BOOL_HEADERS = {
    "Has customs",
    "Via Novo",
    "Excise duty",
    "is_excise",
}

TEXT_HEADERS = {
    "ID",
    "Менеджер",
    "Клиент",
    "Customer Product Name",
    "Our Product Name",
    "Supplier Product Name",
    "Product Name",
    "Product name",
    "Product name (variant)",
    "Supplier",
    "Supplier name",
    "Supplier (donor)",
    "Best Suppl",
    "Best Suppl 2",
    "Currency",
    "Currency Best1",
    "Currency Best2",
    "Currency (donor)",
    "Comments",
    "Комментарии",
    "Brand",
    "Family",
    "Категория ABC",
    "Material",
    "Material number",
    "Код продукта",
    "Название продукта",
    "Вид закупки",
    "Условия оплаты",
    "Поставщик",
    "Валюта",
    "fin.Supplier for calc",
}

INTEGER_HEADERS = {
    "Qty, pcs",
    "Volume, L",
    "Volume to take",
    "StockQty",
    "TransitQty",
    "MarkdownQty",
    "ReserveQty",
    "ReserveECommQty",
    "OrderQty",
    "ConfirmedQty",
    "RemainsQty",
    "LPC",
    "Qty",
    "Quantity",
    "ImportRowNo",
    "Stock",
    "Transit",
    "Purchase Order",
    "Order IS",
    "Stock IS",
    "Reserve cust",
    "Reserve E-Comm",
    "Damaged",
    "Количество",
    "Объем л",
    "Ср.Продажи мес",
    "Safe Stock (st), mnth",
    "Safe Stock (st+tr), mnth",
    "Safe Stock (+ord), mnth",
    "к Быстрому Заказу, шт",
    "к Быстрому Заказу, л",
    "к Быстрому заказу, л",
    "к Заказу, шт",
    "к Заказу, л",
    "Volume PY",
    "Volume 3 mnth",
    "Target uC3",
    "Walk-Away uC3",
}

MONEY_HEADERS = {
    "Cost Novo with VAT",
    "Cost Novo with VAT (prev)",
    "Full Cost Msk",
    "Full Cost Msk (prev)",
    "Дистр цена",
    "Промо цена",
    "curr LPC",
    "curr Landed cost",
    "Best full Price, L",
    "Best full Price, L 2",
    "Кост руб л с НДС",
}

PRICE_DECIMAL_HEADERS = {
    "Price",
    "Supplier Price, L",
    "Price, L",
    "Price, pack",
    "Price, L (prev)",
    "Target Price, L",
    "Target price, L",
    "Target Price, pack",
    "Target price (for suppl)",
    "Supplier Price, L (donor)",
    "Cost per L",
    "Price per L",
}

DECIMAL_HEADERS = {
    "FX markup %",
    "Insurance %",
    "FX markup abs",
    "Transport",
    "Re-export",
    "Agent fee",
    "Bank fee",
    "Customs fee",
    "Additional customs",
    "Storage",
    "Move",
    "Marking",
    "VAT",
    "Money",
    "abs Change",
    "Markup % (from suppl price)",
}


def is_date_header(header: object) -> bool:
    base = normalize_rule_header(header)
    return base in DATE_HEADERS or base.lower().endswith(" date")


def is_bool_header(header: object) -> bool:
    return normalize_rule_header(header) in BOOL_HEADERS


def is_text_header(header: object) -> bool:
    base = normalize_rule_header(header)
    return base in TEXT_HEADERS or is_article_header_name(base)


def is_integer_header(header: object) -> bool:
    base = normalize_rule_header(header)
    return base in INTEGER_HEADERS or is_uc3_header(base)


def is_money_header(header: object) -> bool:
    return normalize_rule_header(header) in MONEY_HEADERS


def is_numeric_header(header: object) -> bool:
    base = normalize_rule_header(header)
    if is_integer_header(base) or is_money_header(base):
        return True
    if base == "Pack":
        return True
    if base.startswith("FX rate"):
        return True
    return base in PRICE_DECIMAL_HEADERS or base in DECIMAL_HEADERS


# ---------------------------------------------------------------------------
# Width / number format / header style rules
# ---------------------------------------------------------------------------

COLUMN_WIDTHS: dict[str, float] = {
    "ID": 10.0,
    "Дата": 11.0,
    "Calc date": 16.0,
    "Price date": 11.0,
    "Change date": 14.0,
    "Менеджер": 18.0,
    "Клиент": 22.0,
    "Customer Product Name": 31.14,
    "Our Product Name": 31.14,
    "Supplier Product Name": 31.14,
    "Product Name": 31.14,
    "Product name": 31.14,
    "Product name (variant)": 31.14,
    "Название продукта": 35.0,
    "Article": 18.0,
    "Supplier Article": 14.0,
    "Material number": 18.0,
    "Brand": 14.0,
    "Family": 16.0,
    "Pack": 8.43,
    "Фасовка": 7.29,
    "Qty in Box": 12.0,
    "is_excise": 14.0,
    "Категория ABC": 4.86,
    "Qty, pcs": 10.0,
    "Volume, L": 10.0,
    "Volume to take": 10.0,
    "Ср.Продажи мес": 8.14,
    "Safe Stock (st), mnth": 8.14,
    "Safe Stock (st+tr), mnth": 8.14,
    "Safe Stock (+ord), mnth": 8.14,
    "к Быстрому Заказу, шт": 8.14,
    "к Быстрому Заказу, л": 8.14,
    "к Быстрому заказу, л": 8.14,
    "к Заказу, шт": 7.57,
    "к Заказу, л": 7.57,
    "Дистр цена": 8.43,
    "Промо цена": 8.43,
    "curr LPC": 7.57,
    "curr Landed cost": 7.57,
    "Target uC3": 7.57,
    "Walk-Away uC3": 7.57,
    "uC3": 7.57,
    "uC3 PY": 7.57,
    "uC3 3 mnth": 7.57,
    "min uC3 stock": 10.50,
    "Best uC3": 5.71,
    "Best 2 uC3": 5.71,
    "Supplier": 16.14,
    "Supplier name": 24.0,
    "Supplier (donor)": 16.14,
    "Best Suppl": 16.14,
    "Best Suppl 2": 16.14,
    "Supplier Price, L": 10.0,
    "Supplier Price, L (donor)": 12.0,
    "Currency": 8.14,
    "Currency Best1": 8.14,
    "Currency Best2": 8.14,
    "Currency (donor)": 10.0,
    "FX rate": 7.29,
    "FX rate (donor)": 10.0,
    "FX rate Best1": 7.29,
    "FX rate Best2": 7.29,
    "Cost Novo with VAT": 8.57,
    "Cost Novo with VAT (prev)": 8.57,
    "Full Cost Msk": 8.57,
    "Full Cost Msk (prev)": 8.57,
    "Target Price, L": 12.0,
    "Target price, L": 7.57,
    "Target Price, pack": 12.0,
    "Target price (for suppl)": 8.14,
    "Price": 13.0,
    "Price, L": 10.0,
    "Price, pack": 10.0,
    "Price, L (prev)": 10.0,
    "last update": 9.43,
    "last update (prev)": 11.0,
    "last update Best1": 9.43,
    "last update Best2": 9.86,
    "Stock": 8.43,
    "Transit": 8.43,
    "Purchase Order": 8.43,
    "Order IS": 8.43,
    "Stock IS": 8.43,
    "Reserve cust": 8.43,
    "Reserve E-Comm": 8.43,
    "Damaged": 8.43,
    "Volume PY": 8.14,
    "Volume 3 mnth": 8.14,
    "Комментарии": 10.0,
    "Comments": 30.0,
    "Код продукта": 12.57,
    "Количество": 7.29,
    "Объем л": 7.29,
    "Вид закупки": 12.0,
    "Условия оплаты": 10.0,
    "Кост руб л с НДС": 10.86,
    "Поставщик": 16.71,
    "Валюта": 7.86,
    "Курс": 8.71,
}

NUMBER_FORMATS_LOCAL: dict[str, str] = {
    "FX rate": FORMATS.FX_INTEGER,
    "FX rate (donor)": FORMATS.FX_INTEGER,
    "FX rate Best1": FORMATS.FX_INTEGER,
    "FX rate Best2": FORMATS.FX_INTEGER,
    "FX markup %": FORMATS.PERCENT_FLEX,
    "Insurance %": FORMATS.PERCENT_FLEX,
    "FX markup abs": FORMATS.DECIMAL_FLEX,
    "abs Change": FORMATS.DECIMAL_2_SIMPLE,
    "Markup % (from suppl price)": "0%",
}

# Only explicit final overrides are stored here. Unknown/report-specific colors
# are left untouched, preserving exporter-specific visual blocks.
DYNAMIC_SUPPLIER_HEADER_FILLS = {
    "Supplier": (146, 208, 80),       # #92D050
    "Cost Novo with VAT": (0, 176, 240),
    "Full Cost Msk": (0, 176, 240),
    "uC3": (0, 176, 240),
    "last update": (205, 205, 205),
    "FX rate": (205, 205, 205),
    "Currency": (205, 205, 205),
}

ABC_VALUE_FILLS = {
    "A": (204, 255, 153),  # #CCFF99
    "B": (255, 255, 153),  # #FFFF99
    "C": (255, 204, 255),  # #FFCCFF
    "D": (150, 150, 150),  # #969696
}

# Shared default header palette used by generic exporters. Report-specific
# exporters may paint larger semantic blocks before the final central pass.
STANDARD_HEADER_FILLS = {
    "Supplier": (146, 208, 80),
    "Supplier Price, L": (146, 208, 80),
    "Currency": (146, 208, 80),
    "FX rate": (146, 208, 80),
    "Supplier (donor)": (146, 208, 80),
    "Supplier Price, L (donor)": (146, 208, 80),
    "Currency (donor)": (146, 208, 80),
    "FX rate (donor)": (146, 208, 80),
    "Target Price, L": (0, 176, 240),
    "Target Price, pack": (0, 176, 240),
    "Cost Novo with VAT": (0, 176, 240),
    "Full Cost Msk": (0, 176, 240),
    "uC3": (0, 176, 240),
    "Дистр цена": (192, 0, 0),
    "Промо цена": (192, 0, 0),
    "curr LPC": (192, 0, 0),
    "curr Landed cost": (192, 0, 0),
    "Stock": (33, 92, 152),
    "Transit": (33, 92, 152),
    "Purchase Order": (33, 92, 152),
    "Order IS": (192, 0, 0),
    "Stock IS": (192, 0, 0),
    "Reserve cust": (33, 92, 152),
    "Reserve E-Comm": (33, 92, 152),
    "Damaged": (33, 92, 152),
}

STANDARD_HEADER_WHITE_FONT = {
    "Дистр цена", "Промо цена", "curr LPC", "curr Landed cost",
    "Stock", "Transit", "Purchase Order", "Order IS", "Stock IS",
    "Reserve cust", "Reserve E-Comm", "Damaged",
}


def standard_header_fill_for_header(header: object) -> tuple[int, int, int] | None:
    return STANDARD_HEADER_FILLS.get(normalize_rule_header(header))


def standard_header_white_font(header: object) -> bool:
    return normalize_rule_header(header) in STANDARD_HEADER_WHITE_FONT


def number_format_for_header(header: object) -> str | None:
    base = normalize_rule_header(header)
    if is_date_header(base):
        return FORMATS.DATE
    if is_article_header_name(base):
        return FORMATS.TEXT
    if base in TEXT_HEADERS:
        return FORMATS.TEXT
    if is_uc3_header(base):
        return FORMATS.INTEGER
    if base in NUMBER_FORMATS_LOCAL:
        return NUMBER_FORMATS_LOCAL[base]
    if base in INTEGER_HEADERS:
        return FORMATS.INTEGER
    if base in MONEY_HEADERS:
        # Keep the long-standing CostCalc format.  The simple RUB mask is
        # accepted reliably by Excel COM and is the format users already see
        # in the stable CostCalc export.
        return FORMATS.MONEY_RUB_SIMPLE
    if base in PRICE_DECIMAL_HEADERS:
        return FORMATS.PRICE_DECIMAL
    if base in DECIMAL_HEADERS:
        return FORMATS.DECIMAL_FLEX
    return None


def width_for_header(header: object, default: float | None = None) -> float | None:
    base = normalize_rule_header(header)
    if has_dynamic_suffix(header):
        dynamic_widths = {
            "Supplier": 16.14,
            "Cost Novo with VAT": 8.57,
            "Full Cost Msk": 8.57,
            "uC3": 7.57,
            "last update": 9.43,
            "FX rate": 7.29,
            "Currency": 8.14,
        }
        if base in dynamic_widths:
            return dynamic_widths[base]
    return COLUMN_WIDTHS.get(base, default)


def header_fill_for_header(header: object) -> tuple[int, int, int] | None:
    base = normalize_rule_header(header)
    if has_dynamic_suffix(header):
        return DYNAMIC_SUPPLIER_HEADER_FILLS.get(base)
    return None


def data_alignment_for_header(header: object) -> str | None:
    if normalize_rule_header(header) == "Категория ABC":
        return "center"
    return None


def rgb_to_excel(color: tuple[int, int, int]) -> int:
    r, g, b = color
    return int(r) + int(g) * 256 + int(b) * 65536


# ---------------------------------------------------------------------------
# Central COM finalizer
# ---------------------------------------------------------------------------


def _excel_column_letter(col_num: int) -> str:
    result = ""
    while col_num > 0:
        col_num, remainder = divmod(col_num - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _sheet_last_row(ws: Any) -> int:
    try:
        return max(int(ws.UsedRange.Rows.Count), 1)
    except Exception:
        return 1


def _apply_abc_highlighting_com(ws: Any, column_letter: str, last_row: int) -> None:
    if last_row < 2:
        return
    data_range = ws.Range(f"{column_letter}2:{column_letter}{last_row}")
    try:
        data_range.HorizontalAlignment = -4108  # xlCenter
    except Exception:
        logger.exception("Не удалось центрировать колонку кат. ABC")

    # Conditional formatting is much faster than cell-by-cell COM formatting.
    try:
        data_range.FormatConditions.Delete()
        for value, color in ABC_VALUE_FILLS.items():
            condition = data_range.FormatConditions.Add(
                Type=2,  # xlExpression
                Formula1=f'=${column_letter}2="{value}"',
            )
            condition.Interior.Color = rgb_to_excel(color)
        return
    except Exception:
        logger.exception("Не удалось применить conditional formatting для кат. ABC; используется fallback")

    # Fallback for Excel versions/locales where FormatConditions named args fail.
    try:
        values = data_range.Value
        if last_row == 2:
            values = ((values,),) if not isinstance(values, tuple) else values
        for offset, row_values in enumerate(values or (), start=2):
            raw = row_values[0] if isinstance(row_values, tuple) else row_values
            color = ABC_VALUE_FILLS.get(str(raw or "").strip().upper())
            if color is not None:
                ws.Cells(offset, int(ws.Range(f"{column_letter}1").Column)).Interior.Color = rgb_to_excel(color)
    except Exception:
        logger.exception("Не удалось применить fallback-подсветку кат. ABC")


def apply_central_worksheet_rules(ws: Any) -> None:
    """Apply authoritative shared column rules to one COM worksheet.

    Exporter-specific layout can be applied before this call. Shared formats and
    widths listed in this module win last, which prevents report-to-report drift.
    """
    try:
        column_count = int(ws.UsedRange.Columns.Count)
    except Exception:
        return
    if column_count <= 0:
        return

    last_row = _sheet_last_row(ws)
    for col_index in range(1, column_count + 1):
        try:
            cell = ws.Cells(1, col_index)
            raw_header = cell.Value
        except Exception:
            continue
        if raw_header is None or str(raw_header).strip() == "":
            continue

        original_header = str(raw_header)
        visible_header = standardize_output_header(original_header)
        if visible_header != original_header:
            try:
                cell.Value = visible_header
            except Exception:
                logger.exception("Не удалось переименовать Excel-заголовок %s", original_header)

        letter = _excel_column_letter(col_index)
        fmt = number_format_for_header(original_header)
        if fmt:
            set_number_format_safe(ws.Columns(f"{letter}:{letter}"), fmt)

        width = width_for_header(original_header)
        if width is not None:
            try:
                ws.Columns(f"{letter}:{letter}").ColumnWidth = float(width)
            except Exception:
                logger.exception("Не удалось применить ширину для %s", original_header)

        fill = header_fill_for_header(original_header)
        if fill is not None:
            try:
                cell.Interior.Color = rgb_to_excel(fill)
                cell.Font.ColorIndex = -4105  # Automatic / black
            except Exception:
                logger.exception("Не удалось применить цвет заголовка для %s", original_header)

        if normalize_rule_header(original_header) == "Категория ABC":
            _apply_abc_highlighting_com(ws, letter, last_row)


def apply_central_workbook_rules(workbook: Any) -> None:
    try:
        count = int(workbook.Worksheets.Count)
    except Exception:
        return
    for index in range(1, count + 1):
        try:
            apply_central_worksheet_rules(workbook.Worksheets(index))
        except Exception:
            logger.exception("Не удалось применить централизованные Excel-правила к листу %s", index)


# ---------------------------------------------------------------------------
# CostCalc header order
# ---------------------------------------------------------------------------


def cost_calc_headers(
    *,
    quick_order_months: int | None,
    safe_stock_months: int | None,
) -> tuple[list[str], str]:
    standard_order_header = (
        f"к Заказу, л ({safe_stock_months} м)"
        if safe_stock_months is not None
        else "к Заказу, л"
    )
    headers = [
        "Supplier Article",
        "Supplier Product Name",
        "Our Product Name",
        "Pack",
        "Категория ABC",
        "Qty, pcs",
        "Volume, L",
        "Ср.Продажи мес",
        standard_order_header,
        "Volume to take",
        "Target price (for suppl)",
        "Price, L",
        "Price, pack",
        "Currency",
        "FX rate",
        "Cost Novo with VAT",
        "Full Cost Msk",
        "uC3",
        "Target price, L",
        "uC3 PY",
        "uC3 3 mnth",
        "Target uC3",
        "Walk-Away uC3",
        "Markup % (from suppl price)",
        "last update (prev)",
        "Price, L (prev)",
        "abs Change",
        "Cost Novo with VAT (prev)",
        "Full Cost Msk (prev)",
        "Дистр цена",
        "Промо цена",
        "curr LPC",
        "curr Landed cost",
        "min uC3 stock",
        "Best Suppl",
        "Best full Price, L",
        "Best uC3",
        "last update Best1",
        "FX rate Best1",
        "Currency Best1",
        "Best Suppl 2",
        "Best full Price, L 2",
        "Best 2 uC3",
        "last update Best2",
        "FX rate Best2",
        "Currency Best2",
        "Volume PY",
        "Volume 3 mnth",
        "Stock",
        "Transit",
        "Purchase Order",
        "Order IS",
        "Stock IS",
        "Reserve cust",
        "Reserve E-Comm",
        "Damaged",
    ]
    return headers, standard_order_header


# ---------------------------------------------------------------------------
# Excel number format safety / save helpers
# ---------------------------------------------------------------------------


def to_invariant_number_format(format_code: str | None) -> str | None:
    if not format_code:
        return format_code
    return (
        str(format_code)
        .replace("ДД", "dd")
        .replace("ММ", "mm")
        .replace("ГГГГ", "yyyy")
        .replace("ГГ", "yy")
        .replace(",0000", ".0000")
        .replace(",####", ".####")
        .replace(",00##", ".00##")
        .replace(",00", ".00")
        .replace(",0###", ".0###")
        .replace(",0#", ".0#")
        .replace(",0", ".0")
        .replace("# ##", "#,##")
    )


def to_local_number_format(
    format_code: str | None,
    *,
    decimal_separator: str = ",",
    thousands_separator: str = " ",
) -> str | None:
    return format_code


def _bounded_number_format_target(target: Any) -> Any | None:
    """Return the used data area for a whole-column COM range.

    Some Excel builds reject NumberFormat/NumberFormatLocal for an entire
    worksheet column (for example when the header intersects a merged or
    table-managed area). Falling back to rows 2..UsedRange preserves the
    exported data formatting without touching the header.
    """
    try:
        try:
            ws = target.Worksheet
        except Exception:
            ws = target.Parent
        first_col = int(target.Column)
        column_count = max(int(target.Columns.Count), 1)
        used_range = ws.UsedRange
        last_row = max(int(used_range.Row) + int(used_range.Rows.Count) - 1, 2)
        return ws.Range(
            ws.Cells(2, first_col),
            ws.Cells(last_row, first_col + column_count - 1),
        )
    except Exception:
        return None


def set_number_format_safe(
    target: Any,
    format_en: str = FORMATS.GENERAL,
    format_local: str | None = None,
    *,
    verify: bool = True,
) -> str:
    """Apply a real Excel number format with quiet, ordered fallbacks.

    General means that no special presentation is required. Newly created
    export workbooks already use General, and some localized Excel builds reject
    explicitly assigning it to an entire column. In that case there is nothing
    to apply, so General is a deliberate no-op.

    For real formats we try the localized mask first and then its invariant
    NumberFormat equivalent. If Excel rejects formatting the whole column, the
    same candidates are retried on the used data area only (rows 2..UsedRange).
    """
    local_code = format_local or format_en or FORMATS.GENERAL

    if (
        format_local is not None
        and format_en
        and str(format_en).strip().casefold() != FORMATS.GENERAL.casefold()
        and format_en != local_code
    ):
        invariant_code = format_en
    else:
        invariant_code = to_invariant_number_format(local_code) or FORMATS.GENERAL

    general_key = FORMATS.GENERAL.casefold()
    if (
        str(local_code).strip().casefold() == general_key
        and str(invariant_code).strip().casefold() == general_key
    ):
        return FORMATS.GENERAL

    candidates = [
        ("NumberFormatLocal", local_code),
        ("NumberFormat", invariant_code),
    ]

    failures: list[tuple[str, str, Exception]] = []

    def apply_candidates(current_target: Any) -> str | None:
        seen: set[tuple[str, str]] = set()
        for attr, fmt in candidates:
            key = (attr, fmt)
            if not fmt or key in seen:
                continue
            seen.add(key)
            try:
                setattr(current_target, attr, fmt)
                if verify:
                    try:
                        actual = getattr(current_target, attr)
                        if str(actual or "").strip().casefold() == general_key:
                            continue
                    except Exception:
                        # Some COM proxies do not allow reliable read-back even
                        # when the assignment itself succeeded.
                        pass
                return fmt
            except Exception as exc:
                failures.append((attr, fmt, exc))
        return None

    applied = apply_candidates(target)
    if applied is not None:
        return applied

    bounded_target = _bounded_number_format_target(target)
    if bounded_target is not None:
        applied = apply_candidates(bounded_target)
        if applied is not None:
            return applied

    if failures:
        attr, fmt, exc = failures[-1]
        logger.error(
            "Не удалось применить Excel NumberFormat: requested_local=%r, "
            "requested_invariant=%r; last_attempt=%s=%r: %s",
            local_code,
            invariant_code,
            attr,
            fmt,
            exc,
        )
    return FORMATS.GENERAL


def ensure_xlsx_path(file_path: str | Path) -> Path:
    path = Path(file_path)
    if path.suffix.lower() != ".xlsx":
        path = path.with_suffix(".xlsx")
    return path


def save_workbook_xlsx(workbook: Any, file_path: str | Path) -> Path:
    target_path = ensure_xlsx_path(file_path)
    # Final authoritative pass. This intentionally happens after exporter-local
    # layout code so common column rules cannot drift between reports.
    apply_central_workbook_rules(workbook)
    workbook.SaveAs(str(target_path))
    return target_path
