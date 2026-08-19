"""add supplier insurance

Revision ID: c4f1a2b3d4e5
Revises: b8e4c2a91d73
Create Date: 2026-08-15
"""

from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa


revision: str = "c4f1a2b3d4e5"
down_revision: Union[str, Sequence[str], None] = "b8e4c2a91d73"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


CALCULATION_TABLES = (
    "supplier_price_calculations",
    "temp_customer_cost_options",
    "temp_target_price_options",
    "target_price_calculations",
    "customer_price_calculations",
)


def upgrade() -> None:
    op.add_column(
        "suppliers",
        sa.Column(
            "insurance_percent",
            sa.Numeric(),
            nullable=False,
            server_default=sa.text("0"),
        ),
    )
    for table_name in CALCULATION_TABLES:
        op.add_column(
            table_name,
            sa.Column(
                "insurance_used",
                sa.Numeric(),
                nullable=False,
                server_default=sa.text("0"),
            ),
        )


def downgrade() -> None:
    for table_name in reversed(CALCULATION_TABLES):
        op.drop_column(table_name, "insurance_used")
    op.drop_column("suppliers", "insurance_percent")
