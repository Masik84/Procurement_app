from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from typing import Any

from app.utils.excel_format_rules import (
    ABC_VALUE_FILLS,
    BOOL_HEADERS,
    COLUMN_WIDTHS,
    DATE_HEADERS,
    DECIMAL_HEADERS,
    FORMATS,
    INTEGER_HEADERS,
    MONEY_HEADERS,
    NUMBER_FORMATS_LOCAL,
    PRICE_DECIMAL_HEADERS,
    STANDARD_HEADER_FILLS,
    STANDARD_HEADER_WHITE_FONT,
    TEXT_HEADERS,
    apply_central_worksheet_rules,
    header_fill_for_header,
    is_bool_header,
    is_date_header,
    is_integer_header,
    is_numeric_header,
    is_text_header,
    normalize_rule_header,
    number_format_for_header as central_number_format_for_header,
    rgb_to_excel,
    set_number_format_safe,
    standard_header_fill_for_header,
    standard_header_white_font,
    width_for_header as central_width_for_header,
)
from app.utils.excel_freeze import apply_freeze_panes
from app.utils.excel_headers import article_text, is_article_header


DEFAULT_FONT_NAME = "Aptos Narrow"
DEFAULT_FONT_SIZE = 11
DEFAULT_HEADER_FILL = rgb_to_excel((205, 205, 205))
DEFAULT_HEADER_FONT = rgb_to_excel((0, 0, 0))

# Compatibility aliases retained for older imports.  The actual values live in
# app.utils.excel_format_rules; these names only prevent older callers from
# breaking while the project migrates to the central rule API.
NUMERIC_HEADERS = (
    set(INTEGER_HEADERS)
    | set(MONEY_HEADERS)
    | set(PRICE_DECIMAL_HEADERS)
    | set(DECIMAL_HEADERS)
    | {"Pack", "FX rate", "FX rate (donor)", "FX rate Best1", "FX rate Best2"}
)
TEXT_LEFT_HEADERS = set(TEXT_HEADERS)
HEADER_WIDTHS = COLUMN_WIDTHS
INTEGER_NUMERIC_HEADERS = {"FX rate", "FX rate (donor)", "FX rate Best1", "FX rate Best2"}
HEADER_FILL_COLORS = {key: rgb_to_excel(value) for key, value in STANDARD_HEADER_FILLS.items()}
HEADER_FONT_COLORS = {key: rgb_to_excel((255, 255, 255)) for key in STANDARD_HEADER_WHITE_FONT}
DEFAULT_NUMERIC_FORMAT_LOCAL = FORMATS.DECIMAL_FLEX


def rgb(r: int, g: int, b: int) -> int:
    return rgb_to_excel((r, g, b))


def normalize_header(header: str) -> str:
    return normalize_rule_header(header)


def parse_excel_number(value: object) -> float | int | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, Decimal):
        return float(value)
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        return value

    text = str(value).strip()
    if not text:
        return None
    text = (
        text.replace("\u00a0", " ")
        .replace("₽", "")
        .replace("EUR", "")
        .replace("USD", "")
        .replace("RUB", "")
        .strip()
    )
    if text in {"-", "—"}:
        return None
    negative = False
    if text.startswith("(") and text.endswith(")"):
        negative = True
        text = text[1:-1]
    text = text.replace(" ", "")

    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        text = text.replace(",", ".")

    try:
        number = Decimal(text)
    except (InvalidOperation, ValueError):
        return None
    if negative:
        number = -number
    if number == number.to_integral_value():
        return int(number)
    return float(number)


def parse_excel_date(value: object) -> object:
    if value is None or value == "":
        return ""
    if isinstance(value, (datetime, date)):
        return value
    text = str(value).strip()
    if not text:
        return ""
    for fmt in ("%d.%m.%Y", "%d.%m.%y", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            pass
    return value


def excel_value_by_header(header: str, value: object) -> Any:
    base = normalize_rule_header(header)
    if value is None:
        return ""
    if is_article_header(base):
        return article_text(value)
    if is_date_header(base):
        return parse_excel_date(value)
    if is_bool_header(base):
        if isinstance(value, bool):
            return "Да" if value else "Нет"
        text = str(value).strip()
        if text in {"1", "True", "true", "Да", "да"}:
            return "Да"
        if text in {"0", "False", "false", "Нет", "нет"}:
            return "Нет"
        return text
    if is_numeric_header(base):
        number = parse_excel_number(value)
        return "" if number is None else number
    if isinstance(value, Decimal):
        return float(value)
    if isinstance(value, (datetime, date)):
        return value
    if isinstance(value, bool):
        return "Да" if value else "Нет"
    return value


def number_format_for_header(header: str) -> str | None:
    return central_number_format_for_header(header)


def width_for_header(header: str, default: float = 10.0) -> float:
    width = central_width_for_header(header, default)
    return float(default if width is None else width)


def excel_column_letter(col_num: int) -> str:
    result = ""
    while col_num > 0:
        col_num, remainder = divmod(col_num - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _apply_abc_data_style(ws: Any, headers: list[str]) -> None:
    # The final central pass applies conditional formatting. This helper only
    # exists so callers of apply_standard_worksheet_format get the same result
    # even when they save by a custom path.
    apply_central_worksheet_rules(ws)


def apply_standard_worksheet_format(
    ws: Any,
    headers: list[str],
    *,
    freeze_cell: str = "D2",
    zoom: int = 85,
) -> None:
    xl_center = -4108
    xl_left = -4131
    xl_vcenter = -4160

    headers_count = len(headers)
    if headers_count <= 0:
        return
    last_col = excel_column_letter(headers_count)

    ws.Cells.Font.Name = DEFAULT_FONT_NAME
    ws.Cells.Font.Size = DEFAULT_FONT_SIZE

    header_range = ws.Range(f"A1:{last_col}1")
    header_range.Font.Name = DEFAULT_FONT_NAME
    header_range.Font.Size = DEFAULT_FONT_SIZE
    header_range.Font.Bold = True
    header_range.WrapText = True
    header_range.HorizontalAlignment = xl_center
    header_range.VerticalAlignment = xl_vcenter

    for col_index, header in enumerate(headers, start=1):
        col_letter = excel_column_letter(col_index)
        base = normalize_rule_header(header)
        cell = ws.Cells(1, col_index)

        fill = standard_header_fill_for_header(header)
        if fill is not None:
            cell.Interior.Color = rgb_to_excel(fill)
        else:
            cell.Interior.Color = DEFAULT_HEADER_FILL
        if standard_header_white_font(header):
            cell.Font.Color = rgb_to_excel((255, 255, 255))
        else:
            cell.Font.Color = DEFAULT_HEADER_FONT

        fmt = central_number_format_for_header(header)
        if fmt is None:
            fmt = FORMATS.TEXT if is_text_header(base) or is_article_header(base) else FORMATS.GENERAL
        set_number_format_safe(ws.Columns(f"{col_letter}:{col_letter}"), fmt)

        width = central_width_for_header(header, 12.0)
        if width is not None:
            ws.Columns(f"{col_letter}:{col_letter}").ColumnWidth = float(width)

        if base == "Категория ABC":
            ws.Columns(f"{col_letter}:{col_letter}").HorizontalAlignment = xl_center
        elif is_text_header(base):
            ws.Columns(f"{col_letter}:{col_letter}").HorizontalAlignment = xl_left

    ws.Rows(1).EntireRow.AutoFit()
    ws.Range(f"A1:{last_col}1").AutoFilter(1)
    apply_freeze_panes(ws, freeze_cell=freeze_cell, zoom=zoom)

    # Authoritative shared override (renames visible ABC header, optional Pack
    # decimals, central widths, dynamic supplier block and ABC value colors).
    apply_central_worksheet_rules(ws)


def apply_target_price_calculated_worksheet_format(
    ws: Any,
    headers: list[str],
    *,
    freeze_cell: str = "D2",
    zoom: int = 85,
) -> None:
    # Dynamic supplier columns are recognized centrally by their _N suffix.
    apply_standard_worksheet_format(ws, headers, freeze_cell=freeze_cell, zoom=zoom)
