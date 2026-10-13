"""SQL-only counters for an import batch (no ORM row materialization).

Keep the same quantity fields as ProductStockPage.refresh_counters.  A batch
with no rows returns zeros.  This helper must run only outside the GUI paint
loop; load_table can still pass its already-fetched rows to refresh_counters.
"""
from __future__ import annotations

from sqlalchemy import case, func


def aggregate_import_counters(session, model, batch_id, imported_by, quantity_fields):
    quantity = sum(
        (func.coalesce(getattr(model, field), 0) for field in quantity_fields),
        0,
    )
    result = session.query(
        func.count(model.id),
        func.coalesce(func.sum(case((model.selected_product_id.is_(None), 1), else_=0)), 0),
        func.coalesce(func.sum(quantity), 0),
    ).filter(
        model.batch_id == batch_id,
        model.imported_by == imported_by,
    ).one()
    return int(result[0]), int(result[1]), float(result[2])
