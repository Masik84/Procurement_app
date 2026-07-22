from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from typing import Any


DEFAULT_FONT_NAME = "Aptos Narrow"
DEFAULT_FONT_SIZE = 11


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

INTEGER_NUMERIC_HEADERS = {
    "FX rate",
    "FX rate (donor)",
    "FX rate Best1",
    "FX rate Best2",
}

NUMERIC_HEADERS = {
    "Qty, pcs",
    "Volume, L",
    "Volume to take",
    "Supplier Price, L",
    "Price, L",
    "Price, pack",
    "Price, L (prev)",
    "FX rate",
    "FX rate (donor)",
    "Pack",
    "Target Price, L",
    "Target Price, pack",
    "Supplier Price, L (donor)",
    "Cost Novo with VAT",
    "Full Cost Msk",
    "uC3",
    "FX markup %",
    "FX markup abs",
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
    "Best full Price, L 2",
    "Best full Price, L",
    "curr LPC",
    "curr Landed cost",
    "Ср.Продажи мес",
    "Safe Stock (st), mnth",
    "Safe Stock (st+tr), mnth",
    "Safe Stock (+ord), mnth",
    "к Быстрому Заказу, шт",
    "к Быстрому Заказу, л",
    "к Заказу, шт",
    "к Заказу, л",
    "Stock",
    "Transit",
    "Purchase Order",
    "Order IS",
    "Stock IS",
    "Reserve cust",
    "Reserve E-Comm",
    "Damaged",
}

TEXT_LEFT_HEADERS = {
    "Article",
    "Product name",
    "Product name (variant)",
    "Менеджер",
    "Клиент",
    "Customer Product Name",
    "Our Product Name",
    "Supplier Product Name",
    "Supplier",
    "Best Suppl 2",
    "Best Suppl",
    "Supplier (donor)",
    "Supplier Article",
    "Currency",
    "Currency Best2",
    "Currency Best1",
    "Currency (donor)",
    "Comments",
    "Brand",
    "Product Name",
}

HEADER_WIDTHS = {
    "Дата": 11.0,
    "Менеджер": 18.0,
    "Клиент": 22.0,
    "Customer Product Name": 31.14,
    "Our Product Name": 31.14,
    "Pack": 8.43,
    "Qty, pcs": 10.0,
    "Volume, L": 10.0,
    "Volume to take": 10.0,
    "Supplier": 16.14,
    "Supplier Article": 14.0,
    "Supplier Price, L": 10.0,
    "Currency": 8.14,
    "FX rate": 7.29,
    "Cost Novo with VAT": 12.0,
    "Full Cost Msk": 12.0,
    "uC3": 12.0,
    "FX markup %": 10.0,
    "FX markup abs": 10.0,
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
    "Calc date": 16.0,
    "Price date": 11.0,
    "Brand": 14.0,
    "Supplier (donor)": 16.14,
    "Supplier Price, L (donor)": 12.0,
    "Currency (donor)": 10.0,
    "FX rate (donor)": 10.0,
    "Target Price, L": 12.0,
    "Target Price, pack": 12.0,
    "VAT": 10.0,
    "Money": 10.0,
    "Comments": 30.0,
    "Product Name": 31.14,
    "Product name": 31.14,
    "Product name (variant)": 31.14,
    "Article": 18.0,
    "Ср.Продажи мес": 10.50,
    "Safe Stock (st), mnth": 10.50,
    "Safe Stock (st+tr), mnth": 10.50,
    "Safe Stock (+ord), mnth": 10.50,
    "к Быстрому Заказу, шт": 8.43,
    "к Быстрому Заказу, л": 8.43,
    "к Заказу, шт": 8.43,
    "к Заказу, л": 8.43,
    "Дистр цена": 8.43,
    "Best Suppl": 16.14,
    "Best Suppl 2": 16.14,
    "Best full Price, L": 12.0,
    "Best full Price, L 2": 12.0,
    "FX rate Best1": 7.29,
    "FX rate Best2": 7.29,
    "Currency Best1": 8.14,
    "Currency Best2": 8.14,
    "Промо цена": 8.43,
    "Stock": 8.43,
    "Transit": 8.43,
    "Purchase Order": 8.43,
    "Order IS": 8.43,
    "Stock IS": 8.43,
    "Reserve cust": 8.43,
    "Reserve E-Comm": 8.43,
    "Damaged": 8.43,
    "last update": 9.43,
}

# Number formats are local Excel format strings for Russian locale. They are
# intentionally mapped by header name so all exports can reuse the same rules.
NUMBER_FORMATS_LOCAL = {
    "Дата": "ДД.ММ.ГГ;@",
    "Price date": "ДД.ММ.ГГ;@",
    "last update": "ДД.ММ.ГГ;@",
    "last update (prev)": "ДД.ММ.ГГ;@",
    "last update Best1": "ДД.ММ.ГГ;@",
    "last update Best2": "ДД.ММ.ГГ;@",
    "FX rate": "# ##0,0###",
    "FX rate (donor)": "# ##0,0###",
    "FX markup %": "0,0#%",
    "FX markup abs": "# ##0,00##;[Red]-# ##0,00##;0",
    "Qty, pcs": "# ##0,00##;[Red]-# ##0,00##;0",
    "Volume, L": "# ##0,00##;[Red]-# ##0,00##;0",
    "Volume to take": "# ##0,00##;[Red]-# ##0,00##;0",
    "Ср.Продажи мес": "# ##0,00",
    "Safe Stock (st), mnth": "# ##0,00",
    "Safe Stock (st+tr), mnth": "# ##0,00",
    "Safe Stock (+ord), mnth": "# ##0,00",
    "к Быстрому Заказу, шт": '# ##0;[Red]-# ##0;"-"',
    "к Быстрому Заказу, л": '# ##0;[Red]-# ##0;"-"',
    "к Заказу, шт": '# ##0;[Red]-# ##0;"-"',
    "к Заказу, л": '# ##0;[Red]-# ##0;"-"',
    "Дистр цена": "# ##0 ₽",
    "Промо цена": "# ##0 ₽",
    "Cost Novo with VAT": '# ##0 ₽;[Red]-# ##0 ₽;"-"',
    "Full Cost Msk": '# ##0 ₽;[Red]-# ##0 ₽;"-"',
    "uC3": '# ##0 ₽;[Red]-# ##0 ₽;"-"',
    "curr LPC": '# ##0 ₽;[Red]-# ##0 ₽;"-"',
    "curr Landed cost": '# ##0 ₽;[Red]-# ##0 ₽;"-"',
    "Best full Price, L": '# ##0 ₽;[Red]-# ##0 ₽;"-"',
    "Best full Price, L 2": '# ##0 ₽;[Red]-# ##0 ₽;"-"',
    "FX rate Best1": "# ##0",
    "FX rate Best2": "# ##0",
    "Stock": '# ##0;[Red]-# ##0;"-"',
    "Transit": '# ##0;[Red]-# ##0;"-"',
    "Purchase Order": '# ##0;[Red]-# ##0;"-"',
    "Order IS": '# ##0;[Red]-# ##0;"-"',
    "Stock IS": '# ##0;[Red]-# ##0;"-"',
    "Reserve cust": '# ##0;[Red]-# ##0;"-"',
    "Reserve E-Comm": '# ##0;[Red]-# ##0;"-"',
    "Damaged": '# ##0;[Red]-# ##0;"-"',
}

DEFAULT_NUMERIC_FORMAT_LOCAL = "# ##0,00##;[Red]-# ##0,00##;0"

HEADER_FILL_COLORS = {
    "Supplier": rgb(146, 208, 80),
    "Supplier Price, L": rgb(146, 208, 80),
    "Currency": rgb(146, 208, 80),
    "FX rate": rgb(146, 208, 80),
    "Currency (donor)": rgb(146, 208, 80),
    "FX rate (donor)": rgb(146, 208, 80),
    "Supplier (donor)": rgb(146, 208, 80),
    "Supplier Price, L (donor)": rgb(146, 208, 80),
    "Target Price, L": rgb(0, 176, 240),
    "Target Price, pack": rgb(0, 176, 240),
    "Cost Novo with VAT": rgb(0, 176, 240),
    "Full Cost Msk": rgb(0, 176, 240),
    "uC3": rgb(0, 176, 240),
    "Volume to take": rgb(205, 205, 205),
    "к Быстрому Заказу, шт": rgb(33, 92, 152),
    "к Быстрому Заказу, л": rgb(33, 92, 152),
    "к Заказу, шт": rgb(33, 92, 152),
    "к Заказу, л": rgb(33, 92, 152),
    "Дистр цена": rgb(192, 0, 0),
    "Промо цена": rgb(192, 0, 0),
    "curr LPC": rgb(192, 0, 0),
    "curr Landed cost": rgb(192, 0, 0),
    "Best Suppl": rgb(0, 176, 240),
    "Best full Price, L": rgb(0, 176, 240),
    "last update Best1": rgb(0, 176, 240),
    "FX rate Best1": rgb(0, 176, 240),
    "Currency Best1": rgb(0, 176, 240),
    "Best Suppl 2": rgb(146, 208, 80),
    "Best full Price, L 2": rgb(146, 208, 80),
    "last update Best2": rgb(146, 208, 80),
    "FX rate Best2": rgb(146, 208, 80),
    "Currency Best2": rgb(146, 208, 80),
    "Stock": rgb(33, 92, 152),
    "Transit": rgb(33, 92, 152),
    "Purchase Order": rgb(33, 92, 152),
    "Order IS": rgb(192, 0, 0),
    "Stock IS": rgb(192, 0, 0),
    "Reserve cust": rgb(33, 92, 152),
    "Reserve E-Comm": rgb(33, 92, 152),
    "Damaged": rgb(33, 92, 152),
}

DEFAULT_HEADER_FILL = rgb(205, 205, 205)
DEFAULT_HEADER_FONT = rgb(0, 0, 0)

HEADER_FONT_COLORS = {
    "к Быстрому Заказу, шт": rgb(255, 255, 255),
    "к Быстрому Заказу, л": rgb(255, 255, 255),
    "к Заказу, шт": rgb(255, 255, 255),
    "к Заказу, л": rgb(255, 255, 255),
    "Дистр цена": rgb(255, 255, 255),
    "Промо цена": rgb(255, 255, 255),
    "curr LPC": rgb(255, 255, 255),
    "curr Landed cost": rgb(255, 255, 255),
    "Stock": rgb(255, 255, 255),
    "Transit": rgb(255, 255, 255),
    "Purchase Order": rgb(255, 255, 255),
    "Order IS": rgb(255, 255, 255),
    "Stock IS": rgb(255, 255, 255),
    "Reserve cust": rgb(255, 255, 255),
    "Reserve E-Comm": rgb(255, 255, 255),
    "Damaged": rgb(255, 255, 255),
}


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
        # Last separator is decimal separator.
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
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
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
    if base in NUMERIC_HEADERS:
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
    base = normalize_header(header)
    if base in NUMBER_FORMATS_LOCAL:
        return NUMBER_FORMATS_LOCAL[base]
    if base in NUMERIC_HEADERS:
        return DEFAULT_NUMERIC_FORMAT_LOCAL
    return None


def width_for_header(header: str, default: float = 10.0) -> float:
    return HEADER_WIDTHS.get(normalize_header(header), default)

# Shared Win32 Excel formatting helpers. Exporters should call these helpers
# instead of keeping report-specific color/width/format rules locally.
def excel_column_letter(col_num: int) -> str:
    result = ""
    while col_num > 0:
        col_num, remainder = divmod(col_num - 1, 26)
        result = chr(65 + remainder) + result
    return result


def set_number_format_safe(target: Any, format_code: str) -> None:
    try:
        target.NumberFormat = format_code
    except Exception:
        try:
            target.NumberFormatLocal = format_code
        except Exception:
            pass


def apply_standard_worksheet_format(ws: Any, headers: list[str], *, freeze_cell: str = "D2", zoom: int = 85) -> None:
    xl_center = -4108
    xl_left = -4131
    xl_vcenter = -4160

    headers_count = len(headers)
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
        base = normalize_header(header)
        cell = ws.Cells(1, col_index)
        cell.Interior.Color = HEADER_FILL_COLORS.get(base, DEFAULT_HEADER_FILL)
        cell.Font.Color = HEADER_FONT_COLORS.get(base, DEFAULT_HEADER_FONT)

        fmt = number_format_for_header(header)
        if fmt is None:
            fmt = "@" if base in TEXT_LEFT_HEADERS else "General"
        set_number_format_safe(ws.Columns(f"{col_letter}:{col_letter}"), fmt)
        ws.Columns(f"{col_letter}:{col_letter}").ColumnWidth = width_for_header(header, 12.0)

    ws.Rows(1).EntireRow.AutoFit()
    ws.Range(f"A1:{last_col}1").AutoFilter(1)

    try:
        app = ws.Application
        app.ActiveWindow.Zoom = zoom
        ws.Activate()
        ws.Range(freeze_cell).Select()
        app.ActiveWindow.FreezePanes = True
    except Exception:
        pass



def _paint_header_range(
    ws: Any,
    headers: list[str],
    first_header: str,
    last_header: str,
    fill_color: int,
    font_color: int | None = None,
) -> None:
    header_map = {str(header): idx + 1 for idx, header in enumerate(headers)}
    first_col = header_map.get(first_header)
    last_col = header_map.get(last_header)
    if not first_col or not last_col:
        return

    c1 = excel_column_letter(first_col)
    c2 = excel_column_letter(last_col)
    rng = ws.Range(f"{c1}1:{c2}1")
    rng.Interior.Color = fill_color
    if font_color is None:
        try:
            rng.Font.ColorIndex = -4105  # Automatic
        except Exception:
            rng.Font.Color = DEFAULT_HEADER_FONT
    else:
        rng.Font.Color = font_color


def _format_header_column(ws: Any, headers: list[str], header: str, *, number_format: str | None = None, width: float | None = None) -> None:
    if header not in headers:
        return
    letter = excel_column_letter(headers.index(header) + 1)
    if number_format:
        set_number_format_safe(ws.Columns(f"{letter}:{letter}"), number_format)
    if width is not None:
        ws.Columns(f"{letter}:{letter}").ColumnWidth = width


def apply_target_price_calculated_worksheet_format(ws: Any, headers: list[str], *, freeze_cell: str = "D2", zoom: int = 85) -> None:
    """Apply shared visual rules for Target Prices / Calculated export.

    Layout:
    - base product columns are gray;
    - current stock price columns are red with white font;
    - each calculated supplier option has a blue cost block and a green supplier/date/currency block.
    """
    apply_standard_worksheet_format(ws, headers, freeze_cell=freeze_cell, zoom=zoom)

    red = rgb(192, 0, 0)
    white = rgb(255, 255, 255)
    blue = rgb(0, 176, 240)
    green = rgb(146, 208, 80)

    _paint_header_range(ws, headers, "Supplier Article", "Our Product Name", DEFAULT_HEADER_FILL)
    _paint_header_range(ws, headers, "Дистр цена", "curr Landed cost", red, white)

    option_idx = 1
    while f"Cost Novo with VAT_{option_idx}" in headers:
        _paint_header_range(
            ws,
            headers,
            f"Cost Novo with VAT_{option_idx}",
            f"Full Cost Msk_{option_idx}",
            blue,
        )
        _paint_header_range(
            ws,
            headers,
            f"Supplier_{option_idx}",
            f"Currency_{option_idx}",
            green,
        )

        _format_header_column(ws, headers, f"Cost Novo with VAT_{option_idx}", number_format="# ##0 ₽;[Red]-# ##0 ₽;\"-\"", width=9.0)
        _format_header_column(ws, headers, f"Full Cost Msk_{option_idx}", number_format="# ##0 ₽;[Red]-# ##0 ₽;\"-\"", width=9.0)
        _format_header_column(ws, headers, f"Supplier_{option_idx}", number_format="@", width=16.14)
        _format_header_column(ws, headers, f"last update_{option_idx}", number_format="ДД.ММ.ГГ;@", width=9.43)
        _format_header_column(ws, headers, f"Currency_{option_idx}", number_format="@", width=8.14)
        option_idx += 1
