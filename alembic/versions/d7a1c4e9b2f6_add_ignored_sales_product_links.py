"""add ignored flag for sales product links

Revision ID: d7a1c4e9b2f6
Revises: c9f4a1b2d7e8
Create Date: 2026-09-13
"""

from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa


revision: str = "d7a1c4e9b2f6"
down_revision: Union[str, Sequence[str], None] = "c9f4a1b2d7e8"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    op.add_column(
        "sales_product_links",
        sa.Column(
            "is_ignored",
            sa.Boolean(),
            nullable=False,
            server_default=sa.false(),
        ),
    )
    op.alter_column("sales_product_links", "is_ignored", server_default=None)


def downgrade() -> None:
    op.drop_column("sales_product_links", "is_ignored")
