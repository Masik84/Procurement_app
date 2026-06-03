from __future__ import annotations

from datetime import date, datetime, time
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Any


DEFAULT_FONT_NAME = "Aptos Narrow"
DEFAULT_FONT_SIZE = 11
DEFAULT_ZOOM = 85

# Excel constants for Win32 COM.
XL_CENTER = -4108
XL_TOP = -4160


def rgb(r: int, g: int, b: int) -> int:
    return r + g * 256 + b * 65536


def normalize_header(header: str) -> str:
    """Return base column name for shared Excel formatting.

    Some reports add suffixes like _1/_2. Formatting should still be selected by
    the actual column name.
    """
    text = str(header or "").strip()
    if "_" in text:
        base, suffix = text.rsplit("_", 1)
        if suffix.isdigit():
            return base
    return text


DATE_HEADERS = {
    "Дата",
    "Calc date",
    "Price date",
    "last update",
    "last update (prev)",
    "last update Best1",
    "last update Best2",
}

BOOL_HEADERS = {
    "Has customs",
    "Via Novo",
    "Excise duty",
}

FX_RATE_HEADERS = {
    "FX rate",
    "FX rate (donor)",
}

NUMERIC_HEADERS = {
    "Pack",
    "Qty, pcs",
    "Volume, L",
    "Supplier Price, L",
    "Supplier Price, L (donor)",
    "Price, L",
    "Price, pack",
    "Price, L (prev)",
    "FX rate",
    "FX rate (donor)",
    "Target Price, L",
    "Target Price, pack",
    "Cost Novo with VAT",
    "Full Cost Msk",
    "FX markup",
    "Transport",
    "Re-export",
    "Agent fee",
    "Bank fee",
    "Customs fee",
    "Additional customs",
    "Storage",
    "Move Novo",
    "Move Msk",
    "Marking",
    "VAT",
    "Money",
    "Дистр цена",
    "Промо цена",
    "curr LPC",
    "curr Landed cost",
}

TEXT_HEADERS = {
    "Менеджер",
    "Клиент",
    "Customer Product Name",
    "Our Product Name",
    "Supplier",
    "Supplier (donor)",
    "fin.Supplier for calc",
    "Supplier Article",
    "Supplier Product Name",
    "Material number",
    "Material",
    "Currency",
    "Currency (donor)",
    "Comments",
}

HEADER_WIDTHS = {
    "Дата": 11.0,
    "Calc date": 16.0,
    "Price date": 11.0,
    "last update": 11.0,
    "Менеджер": 18.0,
    "Клиент": 22.0,
    "Customer Product Name": 31.14,
    "Our Product Name": 31.14,
    "Supplier Product Name": 31.14,
    "Material number": 18.0,
    "Material": 31.14,
    "Brand": 14.0,
    "Pack": 8.43,
    "Qty, pcs": 10.0,
    "Volume, L": 10.0,
    "Supplier": 16.14,
    "Supplier (donor)": 16.14,
    "fin.Supplier for calc": 18.0,
    "Supplier Article": 18.0,
    "Supplier Price, L": 10.0,
    "Supplier Price, L (donor)": 12.0,
    "Currency": 8.14,
    "Currency (donor)": 10.0,
    "FX rate": 7.29,
    "FX rate (donor)": 10.0,
    "Target Price, L": 12.0,
    "Target Price, pack": 12.0,
    "Cost Novo with VAT": 12.0,
    "Full Cost Msk": 12.0,
    "FX markup": 10.0,
    "Transport": 10.0,
    "Re-export": 10.0,
    "Agent fee": 10.0,
    "Has customs": 10.0,
    "Via Novo": 10.0,
    "Bank fee": 10.0,
    "Customs fee": 10.0,
    "Additional customs": 12.0,
    "Storage": 10.0,
    "Move Novo": 10.0,
    "Move Msk": 10.0,
    "Marking": 10.0,
    "Excise duty": 10.0,
    "VAT": 10.0,
    "Money": 10.0,
    "Дистр цена": 11.0,
    "Промо цена": 11.0,
    "curr LPC": 11.0,
    "curr Landed cost": 11.0,
    "Comments": 30.0,
}

# Local Excel format strings for Russian locale.
DATE_FORMAT_LOCAL = "ДД.ММ.ГГ;@"
PRICE_FORMAT_LOCAL = "# ##0,00;[Red]-# ##0,00;0"
PRICE_RUB_FORMAT_LOCAL = "# ##0 ₽"
FX_RATE_FORMAT_LOCAL = "# ##0"
QTY_FORMAT_LOCAL = "# ##0,00##;[Red]-# ##0,00##;0"

NUMBER_FORMATS_LOCAL = {
    "Дата": DATE_FORMAT_LOCAL,
    "Calc date": DATE_FORMAT_LOCAL,
    "Price date": DATE_FORMAT_LOCAL,
    "last update": DATE_FORMAT_LOCAL,
    "last update (prev)": DATE_FORMAT_LOCAL,
    "last update Best1": DATE_FORMAT_LOCAL,
    "last update Best2": DATE_FORMAT_LOCAL,
    "FX rate": FX_RATE_FORMAT_LOCAL,
    "FX rate (donor)": FX_RATE_FORMAT_LOCAL,
    "Target Price, L": PRICE_FORMAT_LOCAL,
    "Target Price, pack": PRICE_FORMAT_LOCAL,
    "Cost Novo with VAT": PRICE_FORMAT_LOCAL,
    "Full Cost Msk": PRICE_FORMAT_LOCAL,
    "Supplier Price, L": PRICE_FORMAT_LOCAL,
    "Supplier Price, L (donor)": PRICE_FORMAT_LOCAL,
    "Price, L": PRICE_FORMAT_LOCAL,
    "Price, pack": PRICE_FORMAT_LOCAL,
    "Price, L (prev)": PRICE_FORMAT_LOCAL,
    "Qty, pcs": QTY_FORMAT_LOCAL,
    "Volume, L": QTY_FORMAT_LOCAL,
    "Pack": QTY_FORMAT_LOCAL,
    "Дистр цена": PRICE_RUB_FORMAT_LOCAL,
    "Промо цена": PRICE_RUB_FORMAT_LOCAL,
    "curr LPC": PRICE_RUB_FORMAT_LOCAL,
    "curr Landed cost": PRICE_RUB_FORMAT_LOCAL,
}

DEFAULT_NUMERIC_FORMAT_LOCAL = PRICE_FORMAT_LOCAL

HEADER_FILL_COLORS = {
    "Supplier": rgb(146, 208, 80),
    "Supplier (donor)": rgb(146, 208, 80),
    "Supplier Price, L": rgb(146, 208, 80),
    "Supplier Price, L (donor)": rgb(146, 208, 80),
    "Currency": rgb(146, 208, 80),
    "Currency (donor)": rgb(146, 208, 80),
    "FX rate": rgb(146, 208, 80),
    "FX rate (donor)": rgb(146, 208, 80),
    "last update": rgb(146, 208, 80),
    "Target Price, L": rgb(0, 176, 240),
    "Target Price, pack": rgb(0, 176, 240),
    "Cost Novo with VAT": rgb(0, 176, 240),
    "Full Cost Msk": rgb(0, 176, 240),
    "Дистр цена": rgb(192, 0, 0),
    "Промо цена": rgb(192, 0, 0),
    "curr LPC": rgb(192, 0, 0),
    "curr Landed cost": rgb(192, 0, 0),
}

HEADER_FONT_COLORS = {
    "Дистр цена": rgb(255, 255, 255),
    "Промо цена": rgb(255, 255, 255),
    "curr LPC": rgb(255, 255, 255),
    "curr Landed cost": rgb(255, 255, 255),
}

DEFAULT_HEADER_FILL = rgb(205, 205, 205)
DEFAULT_HEADER_FONT = rgb(0, 0, 0)


def excel_column_letter(col_num: int) -> str:
    result = ""
    while col_num > 0:
        col_num, remainder = divmod(col_num - 1, 26)
        result = chr(65 + remainder) + result
    return result


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


def parse_fx_rate(value: object) -> int | str:
    number = parse_excel_number(value)
    if number is None:
        return ""
    return int(Decimal(str(number)).to_integral_value(rounding=ROUND_HALF_UP))


def _to_datetime(value: object) -> datetime | object:
    if value is None or value == "":
        return ""
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime.combine(value, time.min)
    text = str(value).strip().replace(",", ".")
    if not text:
        return ""
    for fmt in ("%d.%m.%Y", "%d.%m.%y", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M:%S.%f"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            pass
    return value


def parse_excel_date(value: object) -> object:
    value = _to_datetime(value)
    if isinstance(value, datetime):
        try:
            import pywintypes  # type: ignore
            return pywintypes.Time(value)
        except Exception:
            return value
    return value


def excel_value_by_header(header: str, value: object) -> Any:
    base = normalize_header(header)
    if value is None:
        return ""
    if base in DATE_HEADERS:
        return parse_excel_date(value)
    if base in BOOL_HEADERS:
        if isinstance(value, bool):
            return "Да" if value else "Нет"
        text = str(value).strip()
        if text in {"1", "True", "true", "Да", "да"}:
            return "Да"
        if text in {"0", "False", "false", "Нет", "нет"}:
            return "Нет"
        return text
    if base in FX_RATE_HEADERS:
        return parse_fx_rate(value)
    if base in NUMERIC_HEADERS:
        number = parse_excel_number(value)
        return "" if number is None else number
    if isinstance(value, Decimal):
        return float(value)
    if isinstance(value, (datetime, date)):
        return parse_excel_date(value)
    if isinstance(value, bool):
        return "Да" if value else "Нет"
    return value


def number_format_for_header(header: str) -> str | None:
    base = normalize_header(header)
    if base in NUMBER_FORMATS_LOCAL:
        return NUMBER_FORMATS_LOCAL[base]
    if base in NUMERIC_HEADERS:
        return DEFAULT_NUMERIC_FORMAT_LOCAL
    return None


def width_for_header(header: str, default: float = 10.0) -> float:
    return HEADER_WIDTHS.get(normalize_header(header), default)


def header_fill_for_header(header: str) -> int:
    return HEADER_FILL_COLORS.get(normalize_header(header), DEFAULT_HEADER_FILL)


def header_font_color_for_header(header: str) -> int:
    return HEADER_FONT_COLORS.get(normalize_header(header), DEFAULT_HEADER_FONT)


def set_number_format_safe(target: Any, format_code: str) -> None:
    try:
        target.NumberFormatLocal = format_code
    except Exception:
        try:
            target.NumberFormat = format_code
        except Exception:
            pass


def apply_standard_worksheet_format(ws: Any, headers: list[str], *, freeze_cell: str = "D2", zoom: int = DEFAULT_ZOOM) -> None:
    """Apply shared report formatting.

    Alignment is intentionally applied only to the header row. Data/body cells
    keep Excel's default alignment unless a report explicitly changes it.
    """
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
    header_range.HorizontalAlignment = XL_CENTER
    header_range.VerticalAlignment = XL_TOP

    for col_index, header in enumerate(headers, start=1):
        col_letter = excel_column_letter(col_index)
        base = normalize_header(header)
        cell = ws.Cells(1, col_index)
        cell.Interior.Color = header_fill_for_header(header)
        cell.Font.Color = header_font_color_for_header(header)

        fmt = number_format_for_header(header)
        if fmt is None:
            fmt = "@" if base in TEXT_HEADERS else "General"
        set_number_format_safe(ws.Columns(f"{col_letter}:{col_letter}"), fmt)
        ws.Columns(f"{col_letter}:{col_letter}").ColumnWidth = width_for_header(header, 12.0)

    ws.Rows(1).EntireRow.AutoFit()
    header_range.AutoFilter(1)

    try:
        app = ws.Application
        ws.Activate()
        app.ActiveWindow.Zoom = zoom
        ws.Range(freeze_cell).Select()
        app.ActiveWindow.FreezePanes = True
    except Exception:
        pass


def apply_target_price_calculated_worksheet_format(
    ws: Any,
    headers: list[str],
    *,
    freeze_cell: str = "D2",
    zoom: int = DEFAULT_ZOOM,
) -> None:
    """Apply formatting for TargetPriceCalc_ reports.

    This report intentionally keeps its historical money format for donor-cost
    option columns: whole rubles with the ₽ sign. Generic Target Price reports
    still use the standard two-decimal price format.
    """
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
    header_range.HorizontalAlignment = XL_CENTER
    header_range.VerticalAlignment = XL_TOP

    for col_index, header in enumerate(headers, start=1):
        col_letter = excel_column_letter(col_index)
        base = normalize_header(header)
        cell = ws.Cells(1, col_index)

        # Header colors are report-specific, but still centralized here.
        if col_index <= 3:
            cell.Interior.Color = DEFAULT_HEADER_FILL
            cell.Font.Color = DEFAULT_HEADER_FONT
        elif 4 <= col_index <= 7:
            cell.Interior.Color = rgb(192, 0, 0)
            cell.Font.Color = rgb(255, 255, 255)
        elif base in {"Cost Novo with VAT", "Full Cost Msk"}:
            cell.Interior.Color = rgb(0, 176, 240)
            cell.Font.Color = DEFAULT_HEADER_FONT
        elif base in {"Supplier", "last update", "Currency"}:
            cell.Interior.Color = rgb(146, 208, 80)
            cell.Font.Color = DEFAULT_HEADER_FONT
        else:
            cell.Interior.Color = header_fill_for_header(header)
            cell.Font.Color = header_font_color_for_header(header)

        # Number formats for TargetPriceCalc_ must match the original report.
        if col_index == 1:
            fmt = "@"
        elif 4 <= col_index <= 7:
            fmt = PRICE_RUB_FORMAT_LOCAL
        elif base in {"Cost Novo with VAT", "Full Cost Msk"}:
            fmt = PRICE_RUB_FORMAT_LOCAL
        elif base in {"Supplier", "Currency"}:
            fmt = "@"
        elif base == "last update":
            fmt = DATE_FORMAT_LOCAL
        else:
            fmt = number_format_for_header(header)
            if fmt is None:
                fmt = "@" if base in TEXT_HEADERS else "General"
        set_number_format_safe(ws.Columns(f"{col_letter}:{col_letter}"), fmt)

        # Widths for TargetPriceCalc_ must match the original report.
        if col_index == 1:
            width = 18.0
        elif col_index in {2, 3}:
            width = 31.14
        elif 4 <= col_index <= 7:
            width = 11.0
        elif base in {"Cost Novo with VAT", "Full Cost Msk"}:
            width = 10.0
        elif base == "Supplier":
            width = 16.0
        elif base == "last update":
            width = 11.0
        elif base == "Currency":
            width = 9.0
        else:
            width = width_for_header(header, 12.0)
        ws.Columns(f"{col_letter}:{col_letter}").ColumnWidth = width

    ws.Rows(1).EntireRow.AutoFit()
    header_range.AutoFilter(1)

    try:
        app = ws.Application
        ws.Activate()
        app.ActiveWindow.Zoom = zoom
        ws.Range(freeze_cell).Select()
        app.ActiveWindow.FreezePanes = True
    except Exception:
        pass

