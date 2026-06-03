"""Target price chaged2

Revision ID: e0e16be8e6ca
Revises: 4cd8dc16d116
Create Date: 2026-06-01 18:35:38.599585

"""
from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision: str = 'e0e16be8e6ca'
down_revision: Union[str, Sequence[str], None] = '4cd8dc16d116'
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade():

    op.execute(
        """
        ALTER TABLE target_price_calculations
        RENAME COLUMN cost_novo_wvat_recalculated
        TO cost_novo_wvat
        """
    )


def downgrade():

    op.execute(
        """
        ALTER TABLE target_price_calculations
        RENAME COLUMN cost_novo_wvat
        TO cost_novo_wvat_recalculated
        """
    )
