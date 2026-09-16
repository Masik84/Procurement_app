from __future__ import annotations

import logging
import re
from typing import Iterable

from app.utils.excel_format_rules import (
    COLUMN_WIDTHS,
    DISPLAY_HEADER_RENAMES,
    FORMATS,
    header_fill_for_header,
    is_article_header_name,
    normalize_rule_header,
    number_format_for_header,
    rgb_to_excel,
    set_number_format_safe,
    standard_header_fill_for_header,
    standard_header_white_font,
    standardize_output_header,
)

logger = logging.getLogger(__name__)

# Compatibility export: callers that imported the old rename map keep working.
OUTPUT_HEADER_RENAMES = DISPLAY_HEADER_RENAMES

# Compatibility: formatting is no longer authored here. The dictionary remains
# available for code that only checks membership; excel_spec() is authoritative.
HEADER_SPECS: dict[str, tuple] = {}
_INTEGER_HEADERS = set()
INTEGER_FORMAT = FORMATS.INTEGER


def display_headers(headers: Iterable[object]) -> list[str]:
    return [standardize_output_header(h) for h in headers]


def base_header(header: object) -> str:
    return normalize_rule_header(header)


def excel_spec(header: object):
    """Return (fill, font, number_format) from the central Excel rules."""
    base = normalize_rule_header(header)
    if is_article_header_name(base):
        return ((205, 205, 205), None, FORMATS.TEXT)

    fill = header_fill_for_header(header) or standard_header_fill_for_header(header) or (205, 205, 205)
    font = "white" if standard_header_white_font(header) else None
    number_format = number_format_for_header(header)
    if number_format is None and base == "Категория ABC":
        number_format = FORMATS.TEXT
    return (fill, font, number_format)


def apply_header_style_and_formats(ws, headers: list[str], column_letter_func) -> None:
    """Apply header style/number formats from excel_format_rules.py only."""
    for idx, raw_header in enumerate(headers, start=1):
        spec = excel_spec(raw_header)
        if not spec:
            continue
        fill, font, number_format = spec
        letter = column_letter_func(idx)
        try:
            cell = ws.Cells(1, idx)
            cell.Value = standardize_output_header(raw_header)
            cell.Interior.Color = rgb_to_excel(fill)
            if font == "white":
                cell.Font.Color = rgb_to_excel((255, 255, 255))
            else:
                cell.Font.ColorIndex = -4105
        except Exception:
            logger.exception("Подавленная ошибка (см. traceback выше)")
        if number_format:
            set_number_format_safe(ws.Columns(f"{letter}:{letter}"), number_format)


# Populate the legacy dictionary from the same central source so older code
# that reads HEADER_SPECS directly remains compatible.
HEADER_SPECS.update({header: excel_spec(header) for header in COLUMN_WIDTHS})
