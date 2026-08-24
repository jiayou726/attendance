"""Link procurement package conversions to supplier catalog items.

Revision ID: kitchen20260828
Revises: kitchen20260827
"""

from alembic import op
import sqlalchemy as sa


revision = "kitchen20260828"
down_revision = "kitchen20260827"
branch_labels = None
depends_on = None


def _columns(table):
    inspector = sa.inspect(op.get_bind())
    if not inspector.has_table(table):
        return set()
    return {column["name"] for column in inspector.get_columns(table)}


def upgrade():
    columns = _columns("kitchen_purchase_order_item")
    if "supplier_item_id" not in columns:
        # Kept as a plain integer for SQLite ALTER TABLE compatibility; the
        # model declares the FK for newly-created databases.
        op.add_column("kitchen_purchase_order_item", sa.Column("supplier_item_id", sa.Integer(), nullable=True))
    if "package_conversion_snapshot" not in columns:
        op.add_column(
            "kitchen_purchase_order_item",
            sa.Column("package_conversion_snapshot", sa.String(120), nullable=True),
        )


def downgrade():
    # Non-destructive: procurement history must retain its conversion source.
    pass
