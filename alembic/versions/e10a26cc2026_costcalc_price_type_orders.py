"""CostCalc price types and purchase order weighted costs.

Revision ID: e10a26cc2026
Revises: f2c4a6b8d0e1
"""
from alembic import op
import sqlalchemy as sa
revision = "e10a26cc2026"
down_revision = "f2c4a6b8d0e1"
branch_labels = None
depends_on = None

def upgrade():
    op.add_column("price_history", sa.Column("price_type", sa.String(25), nullable=False, server_default="Price List"))
    op.add_column("supplier_price_calculations", sa.Column("price_type", sa.String(25), nullable=False, server_default="Price List"))
    op.add_column("product_stock", sa.Column("cost_novo_purch_order", sa.Numeric(), nullable=True))
    op.add_column("product_stock", sa.Column("full_cost_msk_purch_order", sa.Numeric(), nullable=True))
    op.add_column("temp_supplier_orders_import", sa.Column("order_cost_total", sa.Numeric(), nullable=True))
    op.add_column("temp_supplier_orders_import", sa.Column("order_cost_volume", sa.Numeric(), nullable=True))

def downgrade():
    op.drop_column("temp_supplier_orders_import", "order_cost_volume")
    op.drop_column("temp_supplier_orders_import", "order_cost_total")
    op.drop_column("product_stock", "full_cost_msk_purch_order")
    op.drop_column("product_stock", "cost_novo_purch_order")
    op.drop_column("supplier_price_calculations", "price_type")
    op.drop_column("price_history", "price_type")
