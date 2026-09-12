"""add product uC3 target history and target price source types

Revision ID: b7d4e9f2c6a1
Revises: a4b7c9d2e6f1
Create Date: 2026-09-11
"""

from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa


revision: str = "b7d4e9f2c6a1"
down_revision: Union[str, Sequence[str], None] = "a4b7c9d2e6f1"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    op.create_table(
        "product_uc3_history",
        sa.Column("id", sa.Integer(), autoincrement=True, nullable=False),
        sa.Column("product_id", sa.Integer(), nullable=False),
        sa.Column("target_uc3", sa.Integer(), nullable=True),
        sa.Column("walk_away_uc3", sa.Integer(), nullable=True),
        sa.Column("change_date", sa.DateTime(), nullable=False),
        sa.ForeignKeyConstraint(["product_id"], ["products.id"], ondelete="CASCADE"),
        sa.PrimaryKeyConstraint("id"),
    )
    op.create_index(
        "ix_product_uc3_history_product_id",
        "product_uc3_history",
        ["product_id"],
        unique=False,
    )
    op.create_index(
        "ix_product_uc3_history_change_date",
        "product_uc3_history",
        ["change_date"],
        unique=False,
    )
    op.create_index(
        "ix_product_uc3_history_product_date_id",
        "product_uc3_history",
        ["product_id", "change_date", "id"],
        unique=False,
    )

    op.add_column(
        "temp_target_price_options",
        sa.Column("source_type", sa.String(length=32), nullable=False, server_default="SUPPLIER"),
    )
    op.alter_column(
        "temp_target_price_options",
        "supplier_id",
        existing_type=sa.Integer(),
        nullable=True,
    )
    op.alter_column("temp_target_price_options", "source_type", server_default=None)

    op.add_column(
        "target_price_calculations",
        sa.Column("source_type", sa.String(length=32), nullable=False, server_default="SUPPLIER"),
    )
    op.alter_column(
        "target_price_calculations",
        "donor_supplier_id",
        existing_type=sa.Integer(),
        nullable=True,
    )
    op.alter_column("target_price_calculations", "source_type", server_default=None)


def downgrade() -> None:
    # Rows based on Target/Walk-Away uC3 have no donor supplier and cannot fit
    # the old NOT NULL schema, so remove those calculation rows before rollback.
    op.execute(
        "DELETE FROM target_price_calculations WHERE donor_supplier_id IS NULL"
    )
    op.alter_column(
        "target_price_calculations",
        "donor_supplier_id",
        existing_type=sa.Integer(),
        nullable=False,
    )
    op.drop_column("target_price_calculations", "source_type")

    op.execute(
        "DELETE FROM temp_target_price_options WHERE supplier_id IS NULL"
    )
    op.alter_column(
        "temp_target_price_options",
        "supplier_id",
        existing_type=sa.Integer(),
        nullable=False,
    )
    op.drop_column("temp_target_price_options", "source_type")

    op.drop_index("ix_product_uc3_history_product_date_id", table_name="product_uc3_history")
    op.drop_index("ix_product_uc3_history_change_date", table_name="product_uc3_history")
    op.drop_index("ix_product_uc3_history_product_id", table_name="product_uc3_history")
    op.drop_table("product_uc3_history")
