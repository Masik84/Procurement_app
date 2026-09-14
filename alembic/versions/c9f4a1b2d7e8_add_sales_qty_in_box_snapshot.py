"""add saved sales qty in box snapshot

Revision ID: c9f4a1b2d7e8
Revises: b7d4e9f2c6a1
Create Date: 2026-09-13
"""

from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa


revision: str = "c9f4a1b2d7e8"
down_revision: Union[str, Sequence[str], None] = "b7d4e9f2c6a1"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    op.add_column(
        "sales_product_links",
        sa.Column("sales_qty_in_box", sa.Integer(), nullable=True),
    )


def downgrade() -> None:
    op.drop_column("sales_product_links", "sales_qty_in_box")
