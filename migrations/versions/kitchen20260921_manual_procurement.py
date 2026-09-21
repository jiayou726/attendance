"""Add manual procurement item source and school calculation snapshots.

Revision ID: kitchen20260921_manual
Revises: kitchen20260913_ai_menu
"""
from alembic import op
import sqlalchemy as sa

revision = "kitchen20260921_manual"
down_revision = "kitchen20260913_ai_menu"
branch_labels = None
depends_on = None


def upgrade():
    inspector = sa.inspect(op.get_bind())
    if not inspector.has_table("kitchen_purchase_order_item"):
        return
    columns = {column["name"] for column in inspector.get_columns("kitchen_purchase_order_item")}
    if "source_type" not in columns:
        op.add_column("kitchen_purchase_order_item", sa.Column(
            "source_type", sa.String(20), nullable=False, server_default="menu"
        ))
    if "per_person_amount" not in columns:
        op.add_column("kitchen_purchase_order_item", sa.Column("per_person_amount", sa.Numeric(16, 4)))
    if "school_headcounts" not in columns:
        op.add_column("kitchen_purchase_order_item", sa.Column(
            "school_headcounts", sa.Text(), nullable=False, server_default="{}"
        ))


def downgrade():
    # Keep historical manual procurement snapshots.
    pass
