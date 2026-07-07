from __future__ import annotations

from decimal import Decimal
from typing import Any
import re

import pandas as pd

from app.utils.text import clean_multi_spaces


_INTEGER_FLOAT_TEXT_RE = re.compile(r"^[+]?([0-9]+)[.,]0+$")


def excel_text(value: Any, *, none_if_empty: bool = False) -> str | None:
    """
    Converts Excel cells that are used as text/IDs/articles to a stable string.

    Excel often stores article/code cells as numbers. pandas/openpyxl may then
    return 149610, 149610.0, Decimal('149610') or the string '149610.0'.
    For matching and saving these values must be treated as the same text key:
    '149610'.

    Notes:
    - If a code was stored in Excel as a real numeric value with leading zeroes,
      the leading zeroes are already lost by Excel before import. They can only
      be preserved when the Excel cell is stored as text.
    - This helper is for import-time text conversion, not fuzzy product matching.
    """
    if value is None:
        return None if none_if_empty else ""

    try:
        if pd.isna(value):
            return None if none_if_empty else ""
    except Exception:
        pass

    if isinstance(value, bool):
        text = str(int(value))
    elif isinstance(value, int):
        text = str(value)
    elif isinstance(value, float):
        if value.is_integer():
            text = str(int(value))
        else:
            text = clean_multi_spaces(value)
    elif isinstance(value, Decimal):
        if value == value.to_integral_value():
            text = str(int(value))
        else:
            text = clean_multi_spaces(value)
    else:
        text = clean_multi_spaces(value)

    if not text or text.lower() == "nan":
        return None if none_if_empty else ""

    match = _INTEGER_FLOAT_TEXT_RE.fullmatch(text)
    if match:
        text = match.group(1)

    text = clean_multi_spaces(text)
    if not text:
        return None if none_if_empty else ""
    return text


def read_excel_raw(*args: Any, **kwargs: Any) -> pd.DataFrame:
    """
    Wrapper around pandas.read_excel for importers.

    dtype=object avoids pandas column-level type inference where possible, so
    text-like cells are not additionally coerced because another row is empty or
    numeric. Numeric/date columns are still parsed explicitly in each importer.
    """
    if not kwargs.get("converters"):
        kwargs.setdefault("dtype", object)
    kwargs.setdefault("keep_default_na", False)
    return pd.read_excel(*args, **kwargs)
