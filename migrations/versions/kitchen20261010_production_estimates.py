"""Keep per-dish estimates and an independent purchase adjustment.

Revision ID: kitchen20261010_estimates
Revises: kitchen20261007_school_classes
"""
from alembic import op
import sqlalchemy as sa

revision = "kitchen20261010_estimates"
down_revision = "kitchen20261007_school_classes"
branch_labels = None
depends_on = None


def upgrade():
    inspector = sa.inspect(op.get_bind())
    if not inspector.has_table("kitchen_purchase_order_item"):
        return
    columns = {column["name"] for column in inspector.get_columns("kitchen_purchase_order_item")}
    if "dish_estimates_json" not in columns:
        op.add_column("kitchen_purchase_order_item", sa.Column(
            "dish_estimates_json", sa.Text(), nullable=False, server_default="{}"
        ))
    if "dish_adjustment_qty" not in columns:
        op.add_column("kitchen_purchase_order_item", sa.Column(
            "dish_adjustment_qty", sa.Numeric(16, 4), nullable=True,
        ))


def downgrade():
    # Existing purchase estimates must never be erased by a downgrade.
    pass
