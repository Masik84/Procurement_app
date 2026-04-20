from __future__ import annotations

import re


_MULTI_SPACES_RE = re.compile(r"\s+")
_TRAILING_L_RE = re.compile(r"([0-9]+(?:[.,][0-9]+)?)\s*(л|l)\s*$", re.IGNORECASE)
_TRAILING_KG_RE = re.compile(r"([0-9]+(?:[.,][0-9]+)?)\s*(кг|kg)\s*$", re.IGNORECASE)


def clean_multi_spaces(value: object) -> str:
    """
    Normalizes whitespace:
    - converts non-breaking spaces to normal spaces
    - trims
    - collapses repeated spaces
    - treats None/NaN as empty string
    """
    if value is None:
        return ""

    try:
        # pandas NaN != NaN
        if value != value:
            return ""
    except Exception:
        pass

    s = str(value)
    if s.strip().lower() == "nan":
        return ""

    s = s.replace("\xa0", " ")
    s = s.strip()
    s = _MULTI_SPACES_RE.sub(" ", s)
    return s


def normalize_product_name(value: object) -> str:
    """
    Aggressive normalization for product name matching.

    Mirrors the Access idea:
    - lowercase
    - trim / collapse spaces
    - remove spaces and punctuation
    - keep letters and digits in a compact comparable form

    Important:
    This function is intended for comparison, not for display.
    """
    s = clean_multi_spaces(value).lower()

    replacements = {
        "č": "c",
        "ć": "c",
        "š": "s",
        "ž": "z",
        "đ": "d",
    }
    for src, dst in replacements.items():
        s = s.replace(src, dst)

    for ch in (" ", "-", "/", "\\", ".", ",", ";", ":", "(", ")", "'", '"', "+"):
        s = s.replace(ch, "")

    return s


def normalize_customer_product_name(value: object) -> str:
    """
    Softer normalization for customer-facing names.

    Behavior mirrors your VBA logic:
    - trim and normalize spaces
    - if name ends with "<number> л/l" -> convert to "<number>L"
    - if name ends with "<number> кг/kg" -> convert to "<number>KG"
    - otherwise return cleaned original
    """
    s = clean_multi_spaces(value)
    if not s:
        return ""

    if _TRAILING_L_RE.search(s):
        return _TRAILING_L_RE.sub(r"\1L", s).strip()

    if _TRAILING_KG_RE.search(s):
        return _TRAILING_KG_RE.sub(r"\1KG", s).strip()

    return s