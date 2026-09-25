"""add product prod group

Revision ID: f2c4a6b8d0e1
Revises: e3b8c1d4f6a2
Create Date: 2026-09-25
"""
from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa


revision: str = "f2c4a6b8d0e1"
down_revision: Union[str, Sequence[str], None] = "e3b8c1d4f6a2"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    op.add_column("products", sa.Column("prod_group", sa.String(length=500), nullable=True))
    op.create_index("ix_products_prod_group", "products", ["prod_group"], unique=False)
    op.execute("UPDATE products SET prod_group = family WHERE prod_group IS NULL")


def downgrade() -> None:
    op.drop_index("ix_products_prod_group", table_name="products")
    op.drop_column("products", "prod_group")
