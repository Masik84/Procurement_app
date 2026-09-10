from __future__ import annotations

from decimal import Decimal
from typing import Mapping

from PySide6.QtWidgets import QInputDialog, QMessageBox, QWidget

from app.db.db import SessionLocal
from app.db.models import PackType
from app.utils.parsers import parse_loose_number
from app.utils.text import clean_multi_spaces


def _format_decimal(value: object) -> str:
    number = parse_loose_number(value)
    if number is None:
        return str(value or "").strip()
    text = format(Decimal(number), "f").rstrip("0").rstrip(".")
    if not text:
        text = "0"
    return text.replace(".", ",")


def ask_pack_type(parent: QWidget, error) -> Decimal | None:
    """Create a missing PackType volume using a package name selected from the DB."""
    options = list(getattr(error, "options", []) or [])
    if not options:
        QMessageBox.warning(
            parent,
            "Вид упаковки не найден",
            f"Нет вида упаковки для {_format_decimal(getattr(error, 'requested_pack', ''))}.\n"
            "Справочник видов упаковки пуст. Сначала добавьте упаковку в справочник Pack types.",
        )
        return None

    requested_pack = parse_loose_number(getattr(error, "requested_pack", None))
    if requested_pack is None:
        QMessageBox.warning(
            parent,
            "Вид упаковки не найден",
            "Не удалось определить объем упаковки.",
        )
        return None
    requested_pack = Decimal(requested_pack)

    # The user chooses only the package NAME.  Volumes from existing rows are
    # irrelevant here: the missing volume itself must be added to pack_types.
    names_by_key: dict[str, str] = {}
    for _volume, name in options:
        clean_name = clean_multi_spaces(name)
        if not clean_name:
            continue
        names_by_key.setdefault(clean_name.casefold(), clean_name)

    names = sorted(names_by_key.values(), key=str.casefold)
    if not names:
        QMessageBox.warning(
            parent,
            "Вид упаковки не найден",
            "В справочнике Pack types нет заполненных названий упаковок.",
        )
        return None

    selected_name, ok = QInputDialog.getItem(
        parent,
        "Вид упаковки не найден",
        (
            f"Нет вида упаковки для {_format_decimal(requested_pack)}л.\n"
            "Выберите название упаковки:"
        ),
        names,
        0,
        False,
    )
    if not ok:
        return None

    selected_name = clean_multi_spaces(selected_name)
    if not selected_name:
        return None

    with SessionLocal() as session:
        # Re-check first: another action/user may already have added this
        # volume while the dialog was open. PackType.volume is unique in DB.
        existing = (
            session.query(PackType)
            .filter(PackType.volume == requested_pack)
            .order_by(PackType.id.asc())
            .first()
        )
        if existing is None:
            session.add(PackType(name=selected_name, volume=requested_pack))
            session.commit()

    # Keep the originally requested volume in the product/temp row.  The
    # selected name is used only to create its new PackType pair.
    return requested_pack


def replace_temp_pack(
    *,
    model,
    batch_id: str,
    imported_by: str,
    old_pack: object,
    new_pack: object,
    pending_changes: Mapping[int, dict] | None = None,
) -> int:
    """Replace an invalid new_pack in the current temp batch and matching in-memory edits."""
    old_num = parse_loose_number(old_pack)
    new_num = parse_loose_number(new_pack)
    if old_num is None or new_num is None:
        return 0

    changed = 0
    with SessionLocal() as session:
        query = session.query(model)
        if hasattr(model, "batch_id"):
            query = query.filter(model.batch_id == batch_id)
        if hasattr(model, "imported_by"):
            query = query.filter(model.imported_by == imported_by)
        if hasattr(model, "selected_product_id"):
            query = query.filter(model.selected_product_id.is_(None))
        for row in query.all():
            row_pack = parse_loose_number(getattr(row, "new_pack", None))
            if row_pack == old_num:
                row.new_pack = new_num
                changed += 1
        session.commit()

    if pending_changes:
        for changes in pending_changes.values():
            if "new_pack" not in changes:
                continue
            if parse_loose_number(changes.get("new_pack")) == old_num:
                changes["new_pack"] = new_num

    return changed


def resolve_missing_pack_for_temp_rows(
    parent: QWidget,
    error,
    *,
    model,
    batch_id: str,
    imported_by: str,
    pending_changes: Mapping[int, dict] | None = None,
) -> bool:
    selected_pack = ask_pack_type(parent, error)
    if selected_pack is None:
        return False
    replace_temp_pack(
        model=model,
        batch_id=batch_id,
        imported_by=imported_by,
        old_pack=getattr(error, "requested_pack", None),
        new_pack=selected_pack,
        pending_changes=pending_changes,
    )
    return True
