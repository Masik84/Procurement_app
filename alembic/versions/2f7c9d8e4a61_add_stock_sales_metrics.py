"""add stock sales metrics

Revision ID: 2f7c9d8e4a61
Revises: e0e16be8e6ca
Create Date: 2026-06-30 12:30:00.000000

"""
from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision: str = "2f7c9d8e4a61"
down_revision: Union[str, Sequence[str], None] = "e0e16be8e6ca"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    """Upgrade schema."""
    op.add_column("product_stock", sa.Column("volume_py", sa.Numeric(), nullable=False, server_default="0"))
    op.add_column("product_stock", sa.Column("volume_3m", sa.Numeric(), nullable=False, server_default="0"))
    op.add_column("product_stock", sa.Column("uc3_py", sa.Numeric(), nullable=False, server_default="0"))
    op.add_column("product_stock", sa.Column("uc3_3m", sa.Numeric(), nullable=False, server_default="0"))


def downgrade() -> None:
    """Downgrade schema."""
    op.drop_column("product_stock", "uc3_3m")
    op.drop_column("product_stock", "uc3_py")
    op.drop_column("product_stock", "volume_3m")
    op.drop_column("product_stock", "volume_py")
