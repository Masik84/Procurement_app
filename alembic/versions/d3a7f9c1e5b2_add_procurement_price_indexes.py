"""add Procurement App price lookup indexes

Revision ID: d3a7f9c1e5b2
Revises: c4f1a2b3d4e5
Create Date: 2026-08-15

These indexes belong only to the Procurement App database.  The external
sales database is intentionally not changed by this migration.
"""

from typing import Sequence, Union

from alembic import op


revision: str = "d3a7f9c1e5b2"
down_revision: Union[str, Sequence[str], None] = "c4f1a2b3d4e5"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    op.create_index(
        "ix_current_supplier_prices_product_supplier",
        "current_supplier_prices",
        ["product_id", "supplier_id"],
        unique=False,
    )
    op.create_index(
        "ix_price_history_product_supplier_date_id",
        "price_history",
        ["product_id", "supplier_id", "price_date", "id"],
        unique=False,
    )


def downgrade() -> None:
    op.drop_index(
        "ix_price_history_product_supplier_date_id",
        table_name="price_history",
    )
    op.drop_index(
        "ix_current_supplier_prices_product_supplier",
        table_name="current_supplier_prices",
    )
