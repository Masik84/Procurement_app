"""add IS quantities to supplier orders import

Revision ID: b8e4c2a91d73
Revises: a6c8e4f1b2d7
Create Date: 2026-08-08
"""

from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa


revision: str = "b8e4c2a91d73"
down_revision: Union[str, Sequence[str], None] = "a6c8e4f1b2d7"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    op.add_column(
        "temp_supplier_orders_import",
        sa.Column("is_order_qty", sa.Numeric(), nullable=False, server_default=sa.text("0")),
    )
    op.add_column(
        "temp_supplier_orders_import",
        sa.Column("is_confirmed_order_qty", sa.Numeric(), nullable=False, server_default=sa.text("0")),
    )


def downgrade() -> None:
    op.drop_column("temp_supplier_orders_import", "is_confirmed_order_qty")
    op.drop_column("temp_supplier_orders_import", "is_order_qty")
