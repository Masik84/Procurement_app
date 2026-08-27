from __future__ import annotations

from typing import Any


def numeric_id_value(value: Any) -> int | None:
    """Return an integer ID suitable for Qt's numeric DisplayRole."""
    if isinstance(value, int) and not isinstance(value, bool):
        return value

    text = str(value or "").strip()
    if not text:
        return None

    try:
        return int(text)
    except (TypeError, ValueError):
        return None
