"""Validate integer primary keys before staging-table DELETE operations.

The Qt UserRole of a cell can contain its *column name*, not the primary key.
Do not pass those values to a PostgreSQL integer ID predicate.
"""

from __future__ import annotations

from collections.abc import Iterable


def validated_staging_ids(values: Iterable[object]) -> set[int]:
    """Return valid staging PKs or reject the entire deletion, without partial save."""
    ids: set[int] = set()
    invalid: list[str] = []
    for value in values:
        if isinstance(value, bool):
            invalid.append(repr(value))
            continue
        if isinstance(value, int):
            row_id = value
        elif isinstance(value, str) and value.isascii() and value.isdecimal():
            row_id = int(value)
        else:
            invalid.append(repr(value))
            continue
        if row_id <= 0:
            invalid.append(repr(value))
        else:
            ids.add(row_id)
    if invalid:
        examples = ", ".join(invalid[:5])
        raise ValueError(
            "Не удалось определить ID удалённых строк импорта: " + examples + ". "
            "Сохранение остановлено без удаления строк. Обновите текущую партию "
            "и повторите удаление через таблицу."
        )
    return ids
