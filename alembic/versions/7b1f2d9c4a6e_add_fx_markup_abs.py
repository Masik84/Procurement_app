"""add absolute FX markup columns

Revision ID: 7b1f2d9c4a6e
Revises: 58c1bc4a8f2e
Create Date: 2026-07-22
"""

from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa


revision: str = "7b1f2d9c4a6e"
down_revision: Union[str, Sequence[str], None] = "58c1bc4a8f2e"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


_CALC_TABLES = (
    "supplier_price_calculations",
    "temp_customer_cost_options",
    "temp_target_price_options",
    "target_price_calculations",
    "customer_price_calculations",
)


def upgrade() -> None:
    op.add_column(
        "suppliers",
        sa.Column("fx_rate_markup_abs", sa.Numeric(), nullable=False, server_default=sa.text("0")),
    )
    for table_name in _CALC_TABLES:
        op.add_column(
            table_name,
            sa.Column("fx_markup_abs_used", sa.Numeric(), nullable=False, server_default=sa.text("0")),
        )


def downgrade() -> None:
    for table_name in reversed(_CALC_TABLES):
        op.drop_column(table_name, "fx_markup_abs_used")
    op.drop_column("suppliers", "fx_rate_markup_abs")
