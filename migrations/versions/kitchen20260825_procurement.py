"""Add per-item delivery fields for the simplified procurement workflow.

Revision ID: kitchen20260825
Revises: kitchen20260824
"""

from alembic import op
import sqlalchemy as sa


revision = "kitchen20260825"
down_revision = "kitchen20260824"
branch_labels = None
depends_on = None


def _columns(table):
    inspector = sa.inspect(op.get_bind())
    if not inspector.has_table(table):
        return set()
    return {column["name"] for column in inspector.get_columns(table)}


def upgrade():
    columns = _columns("kitchen_purchase_order_item")
    if "delivery_date" not in columns:
        op.add_column("kitchen_purchase_order_item", sa.Column("delivery_date", sa.Date(), nullable=True))
    if "delivery_slot" not in columns:
        op.add_column("kitchen_purchase_order_item", sa.Column("delivery_slot", sa.String(10), nullable=True))


def downgrade():
    # Non-destructive: delivery choices should survive rollback.
    pass
