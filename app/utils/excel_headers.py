from __future__ import annotations

from decimal import Decimal
from numbers import Integral, Real
import re
from typing import Any

from app.utils.excel_format_rules import standardize_output_header

_SUFFIX_RE = re.compile(r"^(?P<base>.+?)(?P<suffix>_\d+)$")
_INTEGER_FLOAT_TEXT_RE = re.compile(r"^[+]?([0-9]+)[.,]0+$")


def display_header(header: object) -> str:
    return standardize_output_header(header)


def display_headers(headers) -> list[str]:
    return [display_header(h) for h in headers]


def is_article_header(header: object) -> bool:
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
        if numeric != numeric:
            return ""
        text = str(int(numeric)) if numeric.is_integer() else str(value)
    else:
        text = str(value).strip()

    if text.casefold() in {"", "nan", "nat", "none", "<na>"}:
        return ""
    match = _INTEGER_FLOAT_TEXT_RE.fullmatch(text)
    return match.group(1) if match else text
