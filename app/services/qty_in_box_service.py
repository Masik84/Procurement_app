from __future__ import annotations

from decimal import Decimal

from sqlalchemy.orm import Session

from app.db.models import PackType
from app.utils.parsers import parse_loose_number
from app.utils.text import clean_multi_spaces


SINGLE_ITEM_PACK_MARKERS = ("БОЧ", "ВЕДР", "DRUM", "BUCKET", "PAIL")


def normalize_qty_in_box(value: object, *, field_name: str = "Qty in Box") -> int | None:
    """Return a positive whole-number box quantity or None for an empty value."""
    number = parse_loose_number(value)
    if number is None:
        return None
    if number <= 0:
        raise ValueError(f"Поле '{field_name}' должно быть положительным целым числом.")
    integral = number.to_integral_value()
    if number != integral:
        raise ValueError(
            f"Поле '{field_name}' должно быть целым числом; значение {number} нельзя округлять автоматически."
        )
    return int(integral)


def default_qty_in_box_for_pack(session: Session, pack: object) -> int | None:
    """Return the agreed default 1 for drum/bucket pack types, otherwise None."""
    pack_number = parse_loose_number(pack)
    if pack_number is None:
        return None
    pack_type = (
        session.query(PackType)
        .filter(PackType.volume == pack_number)
        .order_by(PackType.id.asc())
        .first()
    )
    if pack_type is None:
        return None
    pack_name = clean_multi_spaces(pack_type.name).upper()
    if any(marker in pack_name for marker in SINGLE_ITEM_PACK_MARKERS):
        return 1
    return None


def calculate_qty_in_box_candidates(
    *,
    qty_pcs: object = None,
    qty_box: object = None,
    volume_l: object = None,
    pack: object = None,
) -> dict[str, Decimal]:
    """Calculate all usable Qty in Box candidates without rounding."""
    qty_pcs_num = parse_loose_number(qty_pcs)
    qty_box_num = parse_loose_number(qty_box)
    volume_num = parse_loose_number(volume_l)
    pack_num = parse_loose_number(pack)
    candidates: dict[str, Decimal] = {}

    if qty_box_num is None or qty_box_num <= 0:
        return candidates

    if qty_pcs_num is not None and qty_pcs_num > 0:
        candidates["Qty, pcs / Qty, box"] = qty_pcs_num / qty_box_num

    if (
        volume_num is not None
        and volume_num > 0
        and pack_num is not None
        and pack_num > 0
    ):
        candidates["Volume, L / Pack / Qty, box"] = volume_num / pack_num / qty_box_num

    return candidates


def whole_qty_in_box_candidate(value: Decimal) -> int | None:
    if value <= 0 or value != value.to_integral_value():
        return None
    return int(value)
