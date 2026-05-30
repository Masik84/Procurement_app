from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from typing import Any, Mapping, Sequence

XL_CENTER = -4108
XL_LEFT = -4131
XL_RIGHT = -4152
XL_VCENTER = -4160

FONT_NAME = "Aptos Narrow"
FONT_SIZE = 11
HEADER_ROW_HEIGHT = 45

TEXT_FORMAT = "@"
DATE_FORMAT_LOCAL = "ДД.ММ.ГГ;@"
INTEGER_FORMAT_LOCAL = '# ##0;[Red]-# ##0;"-"'
DECIMAL_FORMAT_LOCAL = '# ##0,00;[Red]-# ##0,00;"-"'
DECIMAL4_FORMAT_LOCAL = '# ##0,0000;[Red]-# ##0,0000;"-"'
MONEY_FORMAT_LOCAL = '# ##0 ₽;[Red]-# ##0 ₽;"-"'
GENERAL_FORMAT = "General"

DEFAULT_HEADER_COLOR = (205, 205, 205)
SUPPLIER_HEADER_COLOR = (146, 208, 80)
COST_HEADER_COLOR = (0, 176, 240)
WHITE = (255, 255, 255)
BLACK = (0, 0, 0)


def rgb(r: int, g: int, b: int) -> int:
    return r + g * 256 + b * 65536


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


def set_number_format_safe(target, format_en: str = GENERAL_FORMAT, format_local: str | None = None) -> None:
    try:
        target.NumberFormat = format_en
    except Exception:
        if format_local:
            target.NumberFormatLocal = format_local
        else:
            raise


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
    text = " ".join(str(header or "").strip().lower().replace("\n", " ").split())
    return text


def normalize_header_base(header: str) -> str:
    h = normalize_header(header)
    # Для динамических колонок вида Supplier_1 / Full Cost Msk_2 / ...
    # формат должен определяться по базовому названию колонки.
    if "_" in h and h.rsplit("_", 1)[1].isdigit():
        h = h.rsplit("_", 1)[0].strip()
    # Для колонок сравнения с предыдущими значениями формат такой же, как у базовой колонки.
    for suffix in (" (prev)", " prev"):
        if h.endswith(suffix):
            h = h[: -len(suffix)].strip()
    return h


def is_date_header(header: str) -> bool:
    h = normalize_header_base(header)
    return h in {"дата", "price date", "date"} or h.endswith(" date")


def is_text_header(header: str) -> bool:
    h = normalize_header_base(header)
    return h in {
        "id", "менеджер", "клиент", "customer product name", "our product name",
        "supplier", "supplier article", "currency", "comments", "has customs", "via novo",
        "excise duty", "final supplier", "product name", "manager name", "customer name",
        "article", "brand", "family",
    }


def is_integer_header(header: str) -> bool:
    h = normalize_header_base(header)
    return h in {"qty, pcs", "qty pcs", "qty", "quantity", "pack", "volume, l", "volume l"}


def is_money_header(header: str) -> bool:
    h = normalize_header_base(header)
    return h in {
        "final price",
        "price rub",
        "cost novo withvat",
        "cost novo with vat",
        "full cost msk",
    }


def is_decimal4_header(header: str) -> bool:
    h = normalize_header_base(header)
    return h in {
        "supplier price, l", "supplier price", "price, l", "cost per l", "price per l",
        "cost novo",
    }


def is_decimal_header(header: str) -> bool:
    h = normalize_header_base(header)
    return (
        is_decimal4_header(header)
        or is_money_header(header)
        or h in {
            "fx markup", "transport", "re-export", "agent fee", "bank fee", "customs fee",
            "additional customs", "storage", "move novo", "move msk", "marking", "volume, l",
        }
    )


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
        # 1,234.56 -> 1234.56, 1.234,56 -> 1234.56
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
    if is_date_header(header):
        return parse_date(value)
    if is_decimal_header(header) or is_integer_header(header):
        number = parse_decimal(value)
        if number is None or number == 0:
            return ""
        if is_integer_header(header):
            return int(number.quantize(Decimal("1")))
        return float(number)
    if isinstance(value, Decimal):
        return float(value) if value != 0 else ""
    if isinstance(value, bool):
        return "Да" if value else "Нет"
    return value


def column_format_local(header: str) -> str:
    if is_date_header(header):
        return DATE_FORMAT_LOCAL
    if is_text_header(header):
        return TEXT_FORMAT
    if is_integer_header(header):
        return INTEGER_FORMAT_LOCAL
    if is_money_header(header):
        return MONEY_FORMAT_LOCAL
    if is_decimal4_header(header):
        return DECIMAL4_FORMAT_LOCAL
    if is_decimal_header(header):
        return DECIMAL_FORMAT_LOCAL
    return GENERAL_FORMAT


def apply_column_formats(ws, headers: Sequence[str]) -> None:
    for idx, header in enumerate(headers, start=1):
        letter = excel_column_letter(idx)
        fmt = column_format_local(header)
        column = ws.Columns(f"{letter}:{letter}")
        if fmt == TEXT_FORMAT:
            set_number_format_safe(column, TEXT_FORMAT, TEXT_FORMAT)
        elif fmt == GENERAL_FORMAT:
            set_number_format_safe(column, GENERAL_FORMAT, GENERAL_FORMAT)
        else:
            # Все форматы ниже заданы в русской локали Excel: пробел как разделитель тысяч,
            # запятая как десятичный разделитель, знак ₽. Нельзя сначала ставить General:
            # Excel примет его успешно и локальный формат так и не применится.
            try:
                column.NumberFormatLocal = fmt
            except Exception:
                column.NumberFormat = fmt


def write_table(ws, headers: Sequence[str], rows: Sequence[Sequence[Any]]) -> None:
    for col_index, header in enumerate(headers, start=1):
        ws.Cells(1, col_index).Value = header
    for row_index, row in enumerate(rows, start=2):
        for col_index, header in enumerate(headers, start=1):
            value = row[col_index - 1] if col_index - 1 < len(row) else ""
            ws.Cells(row_index, col_index).Value = excel_cell_value(header, value)


def set_widths_by_headers(ws, headers: Sequence[str], widths: Mapping[str, float], default_width: float = 10.0) -> None:
    hm = header_map(headers)
    for header in headers:
        letter = col_letter(hm, header)
        if letter:
            ws.Columns(f"{letter}:{letter}").ColumnWidth = widths.get(header, default_width)


def color_headers(ws, headers: Sequence[str], color_map: Mapping[str, tuple[int, int, int]], default_color: tuple[int, int, int] = DEFAULT_HEADER_COLOR) -> None:
    hm = header_map(headers)
    last_col = excel_column_letter(len(headers))
    ws.Range(f"A1:{last_col}1").Interior.Color = rgb(*default_color)
    ws.Range(f"A1:{last_col}1").Font.Color = rgb(*BLACK)
    for header, color in color_map.items():
        letter = col_letter(hm, header)
        if letter:
            ws.Range(f"{letter}1").Interior.Color = rgb(*color)
