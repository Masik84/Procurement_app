"""add product ABC category to stock and supplier orders imports

Revision ID: a6c8e4f1b2d7
Revises: 7b1f2d9c4a6e
Create Date: 2026-07-29
"""

from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa


revision: str = "a6c8e4f1b2d7"
down_revision: Union[str, Sequence[str], None] = "7b1f2d9c4a6e"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    op.add_column(
        "products",
        sa.Column("abc_category", sa.String(length=50), nullable=False, server_default=sa.text("'-'")),
    )
    op.add_column(
        "temp_stock_import",
        sa.Column("abc_category", sa.String(length=50), nullable=False, server_default=sa.text("'-'")),
    )
    op.add_column(
        "temp_supplier_orders_import",
        sa.Column("source_our_product_name", sa.String(length=500), nullable=True),
    )
    op.add_column(
        "temp_supplier_orders_import",
        sa.Column("abc_category", sa.String(length=50), nullable=False, server_default=sa.text("'-'")),
    )


def downgrade() -> None:
    op.drop_column("temp_supplier_orders_import", "abc_category")
    op.drop_column("temp_supplier_orders_import", "source_our_product_name")
    op.drop_column("temp_stock_import", "abc_category")
    op.drop_column("products", "abc_category")
