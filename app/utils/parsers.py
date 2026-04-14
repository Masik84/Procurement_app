from __future__ import annotations

import re
from datetime import date, datetime
from typing import Optional
from decimal import Decimal, InvalidOperation


_NON_NUMERIC_RE = re.compile(r"[^0-9,.\-]")


def parse_loose_number(value: object) -> Optional[Decimal]:
    """
    Parses numbers written in a loose human/Excel format.

    Returns:
        Decimal | None
    """
    if value is None:
        return None

    if isinstance(value, bool):
        return Decimal(int(value))

    if isinstance(value, Decimal):
        return value

    if isinstance(value, int):
        return Decimal(value)

    if isinstance(value, float):
        return Decimal(str(value))

    s = str(value)
    s = s.replace("\xa0", " ")
    s = s.strip()

    if not s:
        return None

    s = s.replace(" ", "").replace("\t", "")
    s = _NON_NUMERIC_RE.sub("", s)

    if not s or s == "-":
        return None

    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "")
            s = s.replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")

    try:
        return Decimal(s)
    except (InvalidOperation, ValueError):
        return None

def parse_flexible_date(value: object, default_to_today: bool = True) -> Optional[date]:
    """
    Parses flexible date input similar to the Access VBA behavior.

    Supported examples:
    - datetime/date objects
    - "31.12.2025"
    - "31/12/2025"
    - "31-12-2025"
    - "31.12.25"
    - "31.12" -> current year is assumed

    Returns:
        date | None
    """
    if value is None:
        return date.today() if default_to_today else None

    if isinstance(value, datetime):
        return value.date()

    if isinstance(value, date):
        return value

    s = str(value).strip()
    if not s:
        return date.today() if default_to_today else None

    # First try Python's ISO/date parsing for common machine-friendly formats.
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%d.%m.%Y", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass

    normalized = (
        s.replace("\\", ".")
        .replace("/", ".")
        .replace("-", ".")
        .replace(",", ".")
        .strip()
    )

    parts = [p.strip() for p in normalized.split(".") if p.strip()]

    try:
        if len(parts) == 2:
            day = int(parts[0])
            month = int(parts[1])
            year = date.today().year
            return date(year, month, day)

        if len(parts) == 3:
            day = int(parts[0])
            month = int(parts[1])
            year = int(parts[2])

            if year < 100:
                year = 2000 + year

            return date(year, month, day)
    except ValueError:
        return date.today() if default_to_today else None

    return date.today() if default_to_today else None