"""merge Move columns

Revision ID: a4b7c9d2e6f1
Revises: f1a2b3c4d5e6
Create Date: 2026-09-08
"""

from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa


revision: str = "a4b7c9d2e6f1"
down_revision: Union[str, Sequence[str], None] = "f1a2b3c4d5e6"
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
        "fixed_costs",
        sa.Column("move", sa.Numeric(), nullable=False, server_default=sa.text("0")),
    )
    op.execute(
        """
        UPDATE fixed_costs
        SET move = COALESCE(move_novo_tamozh, 0)
                 + COALESCE(move_tamozh_chekhov, 0)
        """
    )
    op.drop_column("fixed_costs", "move_novo_tamozh")
    op.drop_column("fixed_costs", "move_tamozh_chekhov")
    op.alter_column("fixed_costs", "move", server_default=None)

    for table_name in CALCULATION_TABLES:
        op.add_column(
            table_name,
            sa.Column(
                "move_used",
                sa.Numeric(),
                nullable=False,
                server_default=sa.text("0"),
            ),
        )
        op.execute(
            sa.text(
                f"""
                UPDATE {table_name}
                SET move_used = COALESCE(move_novo_used, 0)
                              + COALESCE(move_msk_used, 0)
                """
            )
        )
        op.drop_column(table_name, "move_novo_used")
        op.drop_column(table_name, "move_msk_used")
        op.alter_column(table_name, "move_used", server_default=None)


def downgrade() -> None:
    for table_name in reversed(CALCULATION_TABLES):
        op.add_column(
            table_name,
            sa.Column(
                "move_novo_used",
                sa.Numeric(),
                nullable=False,
                server_default=sa.text("0"),
            ),
        )
        op.add_column(
            table_name,
            sa.Column(
                "move_msk_used",
                sa.Numeric(),
                nullable=False,
                server_default=sa.text("0"),
            ),
        )
        op.execute(
            sa.text(
                f"UPDATE {table_name} SET move_novo_used = COALESCE(move_used, 0)"
            )
        )
        op.drop_column(table_name, "move_used")
        op.alter_column(table_name, "move_novo_used", server_default=None)
        op.alter_column(table_name, "move_msk_used", server_default=None)

    op.add_column(
        "fixed_costs",
        sa.Column(
            "move_novo_tamozh",
            sa.Numeric(),
            nullable=False,
            server_default=sa.text("0"),
        ),
    )
    op.add_column(
        "fixed_costs",
        sa.Column(
            "move_tamozh_chekhov",
            sa.Numeric(),
            nullable=False,
            server_default=sa.text("0"),
        ),
    )
    op.execute(
        "UPDATE fixed_costs SET move_novo_tamozh = COALESCE(move, 0)"
    )
    op.drop_column("fixed_costs", "move")
    op.alter_column("fixed_costs", "move_novo_tamozh", server_default=None)
    op.alter_column("fixed_costs", "move_tamozh_chekhov", server_default=None)
