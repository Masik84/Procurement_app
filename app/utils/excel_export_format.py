from __future__ import annotations

import logging
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Any, Mapping, Sequence

from app.utils.excel_fast_writer import write_excel_table
from app.utils.excel_format_rules import (
    ABC_VALUE_FILLS,
    FORMATS,
    header_fill_for_header,
    is_article_header_name,
    is_bool_header,
    is_date_header,
    is_integer_header,
    is_money_header,
    is_numeric_header,
    is_text_header,
    normalize_rule_header,
    number_format_for_header,
    rgb_to_excel,
    standard_header_fill_for_header,
    standard_header_white_font,
    standardize_output_header,
    width_for_header,
)
from app.utils.excel_headers import article_text, is_article_header

logger = logging.getLogger(__name__)

XL_CENTER = -4108
XL_LEFT = -4131
XL_RIGHT = -4152
XL_VCENTER = -4160

FONT_NAME = "Aptos Narrow"
FONT_SIZE = 11
HEADER_ROW_HEIGHT = 45

TEXT_FORMAT = FORMATS.TEXT
DATE_FORMAT_LOCAL = FORMATS.DATE
INTEGER_FORMAT_LOCAL = FORMATS.INTEGER
DECIMAL_FORMAT_LOCAL = FORMATS.DECIMAL_2
DECIMAL4_FORMAT_LOCAL = FORMATS.DECIMAL_4
MONEY_FORMAT_LOCAL = FORMATS.MONEY_RUB
GENERAL_FORMAT = FORMATS.GENERAL

DEFAULT_HEADER_COLOR = (205, 205, 205)
SUPPLIER_HEADER_COLOR = (146, 208, 80)
COST_HEADER_COLOR = (0, 176, 240)
WHITE = (255, 255, 255)
BLACK = (0, 0, 0)


def rgb(r: int, g: int, b: int) -> int:
    return rgb_to_excel((r, g, b))


def excel_column_letter(col_num: int) -> str:
    result = ""
    while col_num > 0:
        col_num, remainder = divmod(col_num - 1, 26)
        result = chr(65 + remainder) + result
    return result


def header_map(headers: Sequence[str]) -> dict[str, int]:
    return {str(header): idx + 1 for idx, header in enumerate(headers)}


def col_letter(header_to_index: Mapping[str, int], header: str) -> str | None:
    idx = header_to_index.get(header)
    return excel_column_letter(idx) if idx else None


def apply_base_table_style(ws, headers_count: int) -> None:
    ws.Cells.Font.Name = FONT_NAME
    ws.Cells.Font.Size = FONT_SIZE
    last_col = excel_column_letter(headers_count)
    header_range = ws.Range(f"A1:{last_col}1")
    header_range.Font.Name = FONT_NAME
    header_range.Font.Size = FONT_SIZE
    header_range.Font.Bold = True
    header_range.WrapText = True
    header_range.HorizontalAlignment = XL_CENTER
    header_range.VerticalAlignment = XL_VCENTER
    ws.Rows(1).RowHeight = HEADER_ROW_HEIGHT


def normalize_header(header: str) -> str:
    return normalize_rule_header(header).casefold()


def is_decimal4_header(header: str) -> bool:
    base = normalize_rule_header(header)
    return base in {
        "Supplier Price, L", "Supplier Price, L (donor)", "Price, L",
        "Price, pack", "Cost per L", "Price per L",
    }


def is_decimal_header(header: str) -> bool:
    return is_numeric_header(header) and not is_integer_header(header)


def parse_decimal(value: Any) -> Decimal | None:
    if value is None or value == "":
        return None
    if isinstance(value, Decimal):
        return value
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return Decimal(str(value))
    text = str(value).strip()
    if not text or text in {"-", "—"}:
        return None
    text = text.replace("\u00a0", " ").replace("₽", "").replace("руб.", "").replace("руб", "")
    text = text.replace(" ", "")
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        text = text.replace(",", ".")
    try:
        return Decimal(text)
    except InvalidOperation:
        return None


def parse_date(value: Any) -> Any:
    if isinstance(value, (datetime, date)):
        return value
    text = str(value or "").strip()
    if not text:
        return ""
    for fmt in ("%d.%m.%Y", "%d.%m.%y", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            pass
    return value


def excel_cell_value(header: str, value: Any) -> Any:
    if value is None or value == "":
        return ""
    if is_article_header(header):
        return article_text(value)
    if is_date_header(header):
        return parse_date(value)
    if is_bool_header(header):
        if isinstance(value, bool):
            return "Да" if value else "Нет"
        return value
    if is_numeric_header(header):
        number = parse_decimal(value)
        if number is None or number == 0:
            return ""
        # Display formatting can be integer while keeping the underlying
        # decimal value (Safe Stock / uC3 etc.). Only inherently integer
        # quantity values should be rounded here.
        base = normalize_rule_header(header)
        hard_integer_values = {
            "Qty, pcs", "Volume, L", "Volume to take", "Stock", "Transit",
            "Purchase Order", "Order IS", "Stock IS", "Reserve cust",
            "Reserve E-Comm", "Damaged", "Количество", "Объем л",
            "к Быстрому Заказу, шт", "к Быстрому Заказу, л",
            "к Быстрому заказу, л", "к Заказу, шт", "к Заказу, л",
            "Volume PY", "Volume 3 mnth",
        }
        if base in hard_integer_values:
            return int(number.quantize(Decimal("1"), rounding=ROUND_HALF_UP))
        return float(number)
    if isinstance(value, Decimal):
        return float(value) if value != 0 else ""
    if isinstance(value, bool):
        return "Да" if value else "Нет"
    return value


def column_format_local(header: str) -> str:
    return number_format_for_header(header) or GENERAL_FORMAT


def apply_column_formats(ws, headers: Sequence[str]) -> None:
    from app.utils.excel_format_rules import set_number_format_safe

    for idx, header in enumerate(headers, start=1):
        letter = excel_column_letter(idx)
        fmt = column_format_local(str(header))
        set_number_format_safe(ws.Columns(f"{letter}:{letter}"), fmt)


def write_table(ws, headers: Sequence[str], rows: Sequence[Sequence[Any]]) -> None:
    write_excel_table(
        ws,
        headers,
        rows,
        header_getter=standardize_output_header,
        value_getter=lambda row, header, col_index: excel_cell_value(
            str(header),
            row[col_index] if col_index < len(row) else "",
        ),
    )


def set_widths_by_headers(
    ws,
    headers: Sequence[str],
    widths: Mapping[str, float] | None = None,
    default_width: float = 10.0,
) -> None:
    hm = header_map(headers)
    overrides = widths or {}
    for header in headers:
        letter = col_letter(hm, header)
        if not letter:
            continue
        central = width_for_header(header)
        # Standard columns always take the central width. Explicit caller widths
        # are only a fallback for report-specific/unregistered columns.
        width = central if central is not None else overrides.get(header, default_width)
        ws.Columns(f"{letter}:{letter}").ColumnWidth = width


def color_headers(
    ws,
    headers: Sequence[str],
    color_map: Mapping[str, tuple[int, int, int]] | None = None,
    default_color: tuple[int, int, int] = DEFAULT_HEADER_COLOR,
) -> None:
    hm = header_map(headers)
    last_col = excel_column_letter(len(headers))
    ws.Range(f"A1:{last_col}1").Interior.Color = rgb(*default_color)
    ws.Range(f"A1:{last_col}1").Font.Color = rgb(*BLACK)
    overrides = color_map or {}
    for header in headers:
        letter = col_letter(hm, header)
        if not letter:
            continue
        central = header_fill_for_header(header) or standard_header_fill_for_header(header)
        color = central or overrides.get(header)
        if color is not None:
            ws.Range(f"{letter}1").Interior.Color = rgb_to_excel(color)
        if standard_header_white_font(header):
            ws.Range(f"{letter}1").Font.Color = rgb(*WHITE)


def write_dict_table(ws, headers: Sequence[str], rows: Sequence[Mapping[str, Any]]) -> None:
    write_excel_table(
        ws,
        headers,
        rows,
        header_getter=standardize_output_header,
        value_getter=lambda row, header, _col_index: excel_cell_value(str(header), row.get(header, "")),
    )


def apply_standard_table_format(
    ws,
    headers: Sequence[str],
    *,
    widths: Mapping[str, float] | None = None,
    color_map: Mapping[str, tuple[int, int, int]] | None = None,
    apply_filter: bool = True,
    default_width: float = 12.0,
) -> None:
    if not headers:
        return
    apply_base_table_style(ws, len(headers))
    color_headers(ws, headers, color_map or {})
    apply_column_formats(ws, headers)
    set_widths_by_headers(ws, headers, widths, default_width=default_width)
    if apply_filter:
        try:
            last_col = excel_column_letter(len(headers))
            ws.Range(f"A1:{last_col}1").AutoFilter(1)
        except Exception:
            logger.exception("Подавленная ошибка (см. traceback выше)")


def write_and_format_table(
    ws,
    headers: Sequence[str],
    rows: Sequence[Sequence[Any]] | Sequence[Mapping[str, Any]],
    *,
    widths: Mapping[str, float] | None = None,
    color_map: Mapping[str, tuple[int, int, int]] | None = None,
    apply_filter: bool = True,
) -> None:
    if rows and isinstance(rows[0], Mapping):
        write_dict_table(ws, headers, rows)  # type: ignore[arg-type]
    else:
        write_table(ws, headers, rows)  # type: ignore[arg-type]
    apply_standard_table_format(ws, headers, widths=widths, color_map=color_map, apply_filter=apply_filter)


def openpyxl_cell_value(header: str, value: Any) -> Any:
    return excel_cell_value(header, value)


def openpyxl_number_format(header: str) -> str:
    return number_format_for_header(header) or FORMATS.GENERAL


def write_openpyxl_dict_sheet(
    ws,
    rows: Sequence[Mapping[str, Any]],
    *,
    widths: Mapping[str, float] | None = None,
) -> None:
    if not rows:
        return
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    headers = list(rows[0].keys())
    ws.append([standardize_output_header(h) for h in headers])
    for row in rows:
        ws.append([openpyxl_cell_value(h, row.get(h, "")) for h in headers])

    for col_index, header in enumerate(headers, start=1):
        col_letter = get_column_letter(col_index)
        header_cell = ws.cell(row=1, column=col_index)
        fill = header_fill_for_header(header) or standard_header_fill_for_header(header) or DEFAULT_HEADER_COLOR
        header_cell.fill = PatternFill("solid", fgColor="%02X%02X%02X" % fill)
        header_cell.font = Font(
            name=FONT_NAME,
            size=FONT_SIZE,
            bold=True,
            color="FFFFFF" if standard_header_white_font(header) else "000000",
        )
        header_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        fmt = openpyxl_number_format(header)
        for cell in ws[col_letter][1:]:
            cell.font = Font(name=FONT_NAME, size=FONT_SIZE)
            cell.number_format = fmt
            if is_article_header(header):
                cell.value = article_text(cell.value)
                cell.data_type = "s"
            if normalize_rule_header(header) == "Категория ABC":
                cell.alignment = Alignment(horizontal="center", vertical="center")
                fill_color = ABC_VALUE_FILLS.get(str(cell.value or "").strip().upper())
                if fill_color:
                    cell.fill = PatternFill("solid", fgColor="%02X%02X%02X" % fill_color)

        central_width = width_for_header(header)
        if central_width is not None:
            ws.column_dimensions[col_letter].width = central_width
        elif widths and header in widths:
            ws.column_dimensions[col_letter].width = widths[header]
        else:
            max_len = max(len(str(c.value or "")) for c in ws[col_letter])
            ws.column_dimensions[col_letter].width = min(max(max_len + 2, 12), 40)

    ws.row_dimensions[1].height = HEADER_ROW_HEIGHT
    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "A2"
