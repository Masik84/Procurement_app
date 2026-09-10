from __future__ import annotations

import logging
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Optional

logger = logging.getLogger(__name__)


def to_decimal(value: object, default: Decimal = Decimal("0")) -> Decimal:
    """Convert a DB / UI / Excel value to Decimal.

    This is the single implementation for what used to be a `_to_decimal`
    helper copy-pasted (with small, silent variations) into ~20 files.

    Rules:
    - None and "" return `default` (default is Decimal("0") unless the
      caller asks for something else, e.g. `to_decimal(value, None)`).
    - Decimal values are returned unchanged.
    - Strings are normalized before parsing: non-breaking/regular spaces are
      removed and "," is treated as a decimal separator ("1 234,56" -> 1234.56).
    - Anything that still fails to parse returns `default` instead of raising,
      and is logged as a warning (module + traceback) so a bad value can be
      traced back to its source instead of crashing or failing silently.
    """
    if value is None or value == "":
        return default
    if isinstance(value, Decimal):
        return value
    try:
        text = str(value)
        if isinstance(value, str):
            text = text.strip().replace("\xa0", "").replace(" ", "").replace(",", ".")
        return Decimal(text)
    except (InvalidOperation, ValueError, TypeError):
        logger.warning(
            "to_decimal: не удалось преобразовать %r в Decimal, возвращён default=%r",
            value,
            default,
            exc_info=True,
        )
        return default


def round4(value: Decimal) -> Decimal:
    """Round a Decimal to 4 places, ROUND_HALF_UP.

    Single source of truth for money rounding. Previously most services
    quantized with ROUND_HALF_UP, but one file (price_reports_page.py) quantized
    without specifying a rounding mode, which silently falls back to Python's
    default ROUND_HALF_EVEN ("banker's rounding") - the same computation could
    round differently depending on which module happened to do it. Everything
    now goes through this one function.
    """
    return to_decimal(value).quantize(Decimal("0.0001"), rounding=ROUND_HALF_UP)


class FieldValueError(ValueError):
    """Raised by `parse_decimal_field` when a UI/table cell isn't a valid number."""


def parse_decimal_field(
    value: object,
    field_name: str,
    *,
    empty: str = "raise",
) -> Optional[Decimal]:
    """Parse a Decimal typed by the user into a form field / table cell.

    This replaces the 7 near-identical `_to_decimal(self, value, field_name)`
    copies in page_functions/*.py, which had silently diverged: some treated
    an empty field as None, others as Decimal("0"), one raised. That
    divergence is now an explicit, visible parameter instead of a hidden
    difference between files:

    empty:
        "raise" (default) -> an empty value ("") raises FieldValueError.
            Use this for fields written to a NOT NULL DB column: silently
            saving 0 (or None, which would fail at commit anyway with a
            far less clear DB error) can quietly zero out a value the user
            only meant to edit, not clear.
        "none" -> empty value ("") returns None (field left blank / unset).
            Use this only for a nullable DB column where "unset" is a
            genuinely valid state.
        "zero" -> empty value ("") returns Decimal("0").

    Raises:
        FieldValueError: if the text isn't a valid number (or is empty and
        `empty="raise"`), with `field_name` in the message so the user sees
        exactly which field is wrong.
    """
    if isinstance(value, Decimal):
        return value

    text = str(value).strip().replace("\xa0", "").replace(",", ".")
    if text == "":
        if empty == "raise":
            raise FieldValueError(f"Поле '{field_name}' обязательно и должно быть числом")
        return None if empty == "none" else Decimal("0")

    try:
        return Decimal(text)
    except (InvalidOperation, ValueError):
        raise FieldValueError(f"Поле '{field_name}' должно быть числом")
