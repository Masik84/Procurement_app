"""added clm AgentFee

Revision ID: 480d67e594cc
Revises: 643b0dccea1c
Create Date: 2026-04-17

"""
from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision: str = "480d67e594cc"
down_revision: Union[str, Sequence[str], None] = "643b0dccea1c"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    # 1. suppliers.agent_fee
    op.add_column(
        "suppliers",
        sa.Column(
            "agent_fee",
            sa.Numeric(),
            nullable=False,
            server_default="0",
        ),
    )
    op.execute("UPDATE suppliers SET agent_fee = 0 WHERE agent_fee IS NULL")
    op.alter_column("suppliers", "agent_fee", server_default=None)

    # 2. supplier_price_calculations.agent_fee_used
    op.add_column(
        "supplier_price_calculations",
        sa.Column(
            "agent_fee_used",
            sa.Numeric(),
            nullable=False,
            server_default="0",
        ),
    )
    op.execute(
        """
        UPDATE supplier_price_calculations
        SET agent_fee_used = 0
        WHERE agent_fee_used IS NULL
        """
    )
    op.alter_column("supplier_price_calculations", "agent_fee_used", server_default=None)

    # 3. temp_customer_cost_options.agent_fee_used
    op.add_column(
        "temp_customer_cost_options",
        sa.Column(
            "agent_fee_used",
            sa.Numeric(),
            nullable=False,
            server_default="0",
        ),
    )
    op.execute(
        """
        UPDATE temp_customer_cost_options
        SET agent_fee_used = 0
        WHERE agent_fee_used IS NULL
        """
    )
    op.alter_column("temp_customer_cost_options", "agent_fee_used", server_default=None)

    # 4. customer_price_calculations.agent_fee_used
    op.add_column(
        "customer_price_calculations",
        sa.Column(
            "agent_fee_used",
            sa.Numeric(),
            nullable=False,
            server_default="0",
        ),
    )
    op.execute(
        """
        UPDATE customer_price_calculations
        SET agent_fee_used = 0
        WHERE agent_fee_used IS NULL
        """
    )
    op.alter_column("customer_price_calculations", "agent_fee_used", server_default=None)


def downgrade() -> None:
    op.drop_column("customer_price_calculations", "agent_fee_used")
    op.drop_column("temp_customer_cost_options", "agent_fee_used")
    op.drop_column("supplier_price_calculations", "agent_fee_used")
    op.drop_column("suppliers", "agent_fee")