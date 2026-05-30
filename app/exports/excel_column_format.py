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

NUMERIC_HEADERS = {
    "Pack",
    "Qty, pcs",
    "Volume, L",
    "Supplier Price, L",
    "Price, L",
    "Price (Pack)",
    "Price, L (prev)",
    "FX rate",
    "Cost Novo withVAT",
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
    "Дистр цена",
    "Промо цена",
    "curr LPC",
    "curr Landed cost",
}

TEXT_LEFT_HEADERS = {
    "Менеджер",
    "Клиент",
    "Customer Product Name",
    "Our Product Name",
    "Supplier",
    "Supplier Article",
    "Currency",
    "Comments",
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
    "Supplier": 16.14,
    "Supplier Article": 14.0,
    "Supplier Price, L": 10.0,
    "Currency": 8.14,
    "FX rate": 7.29,
    "Cost Novo withVAT": 12.0,
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
    "Price date": 11.0,
    "Comments": 30.0,
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
    "Pack": "# ##0,00##;[Red]-# ##0,00##;0",
    "Qty, pcs": "# ##0,00##;[Red]-# ##0,00##;0",
    "Volume, L": "# ##0,00##;[Red]-# ##0,00##;0",
}

DEFAULT_NUMERIC_FORMAT_LOCAL = "# ##0,00##;[Red]-# ##0,00##;0"

HEADER_FILL_COLORS = {
    "Supplier": rgb(146, 208, 80),
    "Supplier Price, L": rgb(146, 208, 80),
    "Currency": rgb(146, 208, 80),
    "FX rate": rgb(146, 208, 80),
    "Cost Novo withVAT": rgb(0, 176, 240),
    "Cost Novo with VAT": rgb(0, 176, 240),
    "Full Cost Msk": rgb(0, 176, 240),
}

DEFAULT_HEADER_FILL = rgb(205, 205, 205)
DEFAULT_HEADER_FONT = rgb(0, 0, 0)


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
