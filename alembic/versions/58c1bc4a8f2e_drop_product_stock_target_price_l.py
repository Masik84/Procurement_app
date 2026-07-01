"""drop product_stock target_price_l

Revision ID: 58c1bc4a8f2e
Revises: 2f7c9d8e4a61
Create Date: 2026-06-30 14:40:00.000000

"""
from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision: str = "58c1bc4a8f2e"
down_revision: Union[str, Sequence[str], None] = "2f7c9d8e4a61"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def _has_column(table_name: str, column_name: str) -> bool:
    bind = op.get_bind()
    inspector = sa.inspect(bind)
    return any(col.get("name") == column_name for col in inspector.get_columns(table_name))


def upgrade() -> None:
    """Upgrade schema."""
    # Target price depends on the supplier price list export and must not be stored in product_stock.
    if _has_column("product_stock", "target_price_l"):
        op.drop_column("product_stock", "target_price_l")


def downgrade() -> None:
    """Downgrade schema."""
    if not _has_column("product_stock", "target_price_l"):
        op.add_column("product_stock", sa.Column("target_price_l", sa.Numeric(), nullable=False, server_default="0"))
