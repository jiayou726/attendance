"""Add package conversion and ordered tracking to purchase items.

Revision ID: kitchen20260827
Revises: kitchen20260826
"""

from alembic import op
import sqlalchemy as sa


revision = "kitchen20260827"
down_revision = "kitchen20260826"
branch_labels = None
depends_on = None


def _columns(table):
    inspector = sa.inspect(op.get_bind())
    if not inspector.has_table(table):
        return set()
    return {column["name"] for column in inspector.get_columns(table)}


def upgrade():
    columns = _columns("kitchen_purchase_order_item")
    if "package_qty" not in columns:
        op.add_column("kitchen_purchase_order_item", sa.Column("package_qty", sa.Numeric(16, 4), nullable=True))
    if "package_unit" not in columns:
        op.add_column("kitchen_purchase_order_item", sa.Column("package_unit", sa.String(20), nullable=True))
    if "ordered" not in columns:
        op.add_column(
            "kitchen_purchase_order_item",
            sa.Column("ordered", sa.Boolean(), nullable=False, server_default=sa.false()),
        )


def downgrade():
    # Non-destructive: operational tracking must survive rollback.
    pass
