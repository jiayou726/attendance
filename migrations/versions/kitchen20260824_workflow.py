"""Kitchen homepage workflow and imported-data metadata only.

Revision ID: kitchen20260824
Revises: kitchen20260814

This migration intentionally touches only kitchen_* tables.
"""

from alembic import op
import sqlalchemy as sa


revision = "kitchen20260824"
down_revision = "kitchen20260814"
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
    _add_if_missing("kitchen_school", sa.Column("default_headcount", sa.Integer(), nullable=False, server_default="0"))
    _add_if_missing("kitchen_supplier", sa.Column("mobile", sa.String(50)))
    _add_if_missing("kitchen_supplier", sa.Column("fax", sa.String(50)))
    _add_if_missing("kitchen_supplier", sa.Column("contact", sa.String(100)))
    _add_if_missing("kitchen_supplier", sa.Column("address", sa.String(255)))
    _add_if_missing("kitchen_supplier", sa.Column("source_file", sa.String(255)))
    _add_if_missing(
        "kitchen_recipe_ingredient",
        sa.Column("quantity_status", sa.String(20), nullable=False, server_default="manual"),
    )
    _add_if_missing("kitchen_recipe_ingredient", sa.Column("source_note", sa.String(255)))
    if not sa.inspect(op.get_bind()).has_table("kitchen_supplier_item"):
        op.create_table(
            "kitchen_supplier_item",
            sa.Column("id", sa.Integer(), primary_key=True),
            sa.Column("supplier_id", sa.Integer(), sa.ForeignKey("kitchen_supplier.id", ondelete="CASCADE"), nullable=False),
            sa.Column("ingredient_id", sa.Integer(), sa.ForeignKey("kitchen_ingredient.id"), nullable=True),
            sa.Column("source_key", sa.String(160)),
            sa.Column("name", sa.String(120), nullable=False),
            sa.Column("unit", sa.String(20), nullable=False),
            sa.Column("package_conversion", sa.String(120)),
            sa.Column("last_quantity", sa.Numeric(16, 3)),
            sa.Column("last_unit_price", sa.Numeric(16, 4)),
            sa.Column("last_purchase_date", sa.Date()),
            sa.Column("order_count", sa.Integer(), nullable=False, server_default="0"),
            sa.Column("source_file", sa.String(255)),
            sa.Column("manual_override", sa.Boolean(), nullable=False, server_default=sa.false()),
            sa.Column("active", sa.Boolean(), nullable=False, server_default=sa.true()),
            sa.Column("created_at", sa.DateTime(), nullable=False, server_default=sa.func.now()),
            sa.Column("updated_at", sa.DateTime(), nullable=False, server_default=sa.func.now()),
            sa.UniqueConstraint("supplier_id", "name", name="uq_kitchen_supplier_item"),
        )
    else:
        _add_if_missing("kitchen_supplier_item", sa.Column("source_key", sa.String(160)))
        _add_if_missing("kitchen_supplier_item", sa.Column("package_conversion", sa.String(120)))
        _add_if_missing("kitchen_supplier_item", sa.Column("manual_override", sa.Boolean(), nullable=False, server_default=sa.false()))
        _add_if_missing("kitchen_supplier_item", sa.Column("active", sa.Boolean(), nullable=False, server_default=sa.true()))


def downgrade():
    # Non-destructive by design: imported kitchen data must survive rollback.
    pass
