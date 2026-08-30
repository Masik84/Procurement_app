from __future__ import annotations

from decimal import Decimal
from numbers import Integral, Real
import re
from typing import Any

# Только для отображения в GUI/Excel. Не использовать для колонок БД или ключей словарей.
_HEADER_RENAMES = {
    "Pack Price, L": "Price, pack",
    "Price (Pack)": "Price, pack",
    "Price, Pack": "Price, pack",
    "Supplier Product name": "Supplier Product Name",
    "Target Price (Pack)": "Target Price, pack",
    "Cost Novo withVAT": "Cost Novo with VAT",
    "Cost Novo withVAT (prev)": "Cost Novo with VAT (prev)",
}

_SUFFIX_RE = re.compile(r"^(?P<base>.+?)(?P<suffix>_\d+)$")
_INTEGER_FLOAT_TEXT_RE = re.compile(r"^[+]?([0-9]+)[.,]0+$")


def display_header(header: object) -> str:
    text = str(header or "")
    match = _SUFFIX_RE.match(text)
    if match:
        base = match.group("base")
        suffix = match.group("suffix")
        return _HEADER_RENAMES.get(base, base) + suffix
    return _HEADER_RENAMES.get(text, text)


def display_headers(headers) -> list[str]:
    return [display_header(h) for h in headers]


def is_article_header(header: object) -> bool:
    """Return True for every Excel column whose value is an article identifier.

    Exporters use several spellings (Article, Supplier Article,
    SourceArticle and Russian Артикул).  Matching the compact header keeps the
    rule working for spaces, underscores, camel case and numbered report
    columns such as Supplier Article_2.
    """
    compact = re.sub(r"[^0-9a-zа-яё]+", "", display_header(header).casefold())
    return "article" in compact or "артикул" in compact


def article_text(value: Any) -> str:
    """Convert an article value to stable Excel text without losing zeroes."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return str(int(value))
    if isinstance(value, Integral):
        return str(int(value))
    if isinstance(value, Decimal):
        if value.is_nan():
            return ""
        if value == value.to_integral_value():
            return str(int(value))
        text = format(value, "f")
    elif isinstance(value, Real):
        numeric = float(value)
        if numeric != numeric:  # NaN
            return ""
        text = str(int(numeric)) if numeric.is_integer() else str(value)
    else:
        text = str(value).strip()

    if text.casefold() in {"", "nan", "nat", "none", "<na>"}:
        return ""
    match = _INTEGER_FLOAT_TEXT_RE.fullmatch(text)
    return match.group(1) if match else text
