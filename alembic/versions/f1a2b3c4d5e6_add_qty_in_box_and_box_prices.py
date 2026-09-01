"""add qty in box and supplier box price fields

Revision ID: f1a2b3c4d5e6
Revises: d3a7f9c1e5b2
Create Date: 2026-09-01
"""

from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa


revision: str = "f1a2b3c4d5e6"
down_revision: Union[str, Sequence[str], None] = "d3a7f9c1e5b2"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


_TEMP_TABLES_WITH_NEW_QTY = (
    "temp_price_import",
    "temp_customer_cost_import",
    "temp_stock_import",
    "temp_supplier_orders_import",
    "temp_is_import",
    "temp_product_search_import",
    "temp_target_price_import",
)


def upgrade() -> None:
    op.add_column("products", sa.Column("qty_in_box", sa.Integer(), nullable=True))

    op.add_column("temp_price_import", sa.Column("price_box", sa.Numeric(), nullable=True))
    op.add_column("temp_price_import", sa.Column("qty_box", sa.Numeric(), nullable=True))

    for table_name in _TEMP_TABLES_WITH_NEW_QTY:
        op.add_column(table_name, sa.Column("new_qty_in_box", sa.Integer(), nullable=True))


def downgrade() -> None:
    for table_name in reversed(_TEMP_TABLES_WITH_NEW_QTY):
        op.drop_column(table_name, "new_qty_in_box")

    op.drop_column("temp_price_import", "qty_box")
    op.drop_column("temp_price_import", "price_box")
    op.drop_column("products", "qty_in_box")
