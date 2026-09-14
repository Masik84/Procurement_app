"""keep default false for sales_product_links.is_ignored

Revision ID: e3b8c1d4f6a2
Revises: d7a1c4e9b2f6
Create Date: 2026-09-14
"""

from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa


revision: str = "e3b8c1d4f6a2"
down_revision: Union[str, Sequence[str], None] = "d7a1c4e9b2f6"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    # Keep a DB-level default permanently. SalesProductLink is also used by
    # code paths that do not explicitly pass is_ignored on INSERT.
    op.alter_column(
        "sales_product_links",
        "is_ignored",
        existing_type=sa.Boolean(),
        nullable=False,
        server_default=sa.false(),
    )


def downgrade() -> None:
    op.alter_column(
        "sales_product_links",
        "is_ignored",
        existing_type=sa.Boolean(),
        nullable=False,
        server_default=None,
    )
