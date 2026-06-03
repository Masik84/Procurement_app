from __future__ import annotations

import re

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
