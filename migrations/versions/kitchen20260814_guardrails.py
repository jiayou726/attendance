"""Add kitchen guardrail columns without touching attendance tables.

Revision ID: kitchen20260814
Revises: kitchen20260813
"""

from alembic import op
import sqlalchemy as sa

revision = "kitchen20260814"
down_revision = "kitchen20260813"
branch_labels = None
depends_on = None


def _columns(table):
    inspector = sa.inspect(op.get_bind())
    if not inspector.has_table(table):
        return set()
    return {column["name"] for column in inspector.get_columns(table)}


def _add_if_missing(table, column):
    if column.name not in _columns(table):
        op.add_column(table, column)


def upgrade():
    _add_if_missing("kitchen_ingredient", sa.Column("base_unit", sa.String(10), nullable=False, server_default="g"))
    _add_if_missing("kitchen_purchase_order", sa.Column("supplier_overridden", sa.Boolean(), nullable=False, server_default=sa.false()))
    _add_if_missing("kitchen_purchase_order_item", sa.Column("base_unit_snapshot", sa.String(10), nullable=False, server_default="g"))
    _add_if_missing("kitchen_purchase_order_item", sa.Column("manual_override", sa.Boolean(), nullable=False, server_default=sa.false()))


def downgrade():
    # Intentionally non-destructive. Kitchen production history should not be
    # dropped automatically; rollback requires a reviewed manual migration.
    pass
