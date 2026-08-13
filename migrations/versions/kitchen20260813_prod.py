"""Kitchen production schema only.

Revision ID: kitchen20260813
Revises: e76227564f17
Create Date: 2026-08-13

IMPORTANT: this migration intentionally touches only kitchen_* tables.
It must not rename/drop/alter employee or checkin tables.
"""

from alembic import op
import sqlalchemy as sa


revision = "kitchen20260813"
down_revision = "e76227564f17"
branch_labels = None
depends_on = None


def _has_table(name):
    return sa.inspect(op.get_bind()).has_table(name)


def _columns(name):
    if not _has_table(name):
        return set()
    return {c["name"] for c in sa.inspect(op.get_bind()).get_columns(name)}


def _add_if_missing(table, column):
    if column.name not in _columns(table):
        op.add_column(table, column)


def upgrade():
    # School
    if not _has_table("kitchen_school"):
        op.create_table(
            "kitchen_school",
            sa.Column("id", sa.Integer(), primary_key=True),
            sa.Column("name", sa.String(100), nullable=False, unique=True),
            sa.Column("code", sa.String(50), unique=True),
            sa.Column("active", sa.Boolean(), nullable=False, server_default=sa.true()),
            sa.Column("created_at", sa.DateTime(), nullable=False, server_default=sa.func.now()),
            sa.Column("updated_at", sa.DateTime(), nullable=False, server_default=sa.func.now()),
        )
    else:
        _add_if_missing("kitchen_school", sa.Column("active", sa.Boolean(), nullable=False, server_default=sa.true()))
        _add_if_missing("kitchen_school", sa.Column("created_at", sa.DateTime(), nullable=False, server_default=sa.func.now()))
        _add_if_missing("kitchen_school", sa.Column("updated_at", sa.DateTime(), nullable=False, server_default=sa.func.now()))

    # Supplier
    if not _has_table("kitchen_supplier"):
        op.create_table(
            "kitchen_supplier",
            sa.Column("id", sa.Integer(), primary_key=True),
            sa.Column("name", sa.String(100), nullable=False, unique=True),
            sa.Column("phone", sa.String(50)),
            sa.Column("note", sa.String(255)),
            sa.Column("active", sa.Boolean(), nullable=False, server_default=sa.true()),
            sa.Column("created_at", sa.DateTime(), nullable=False, server_default=sa.func.now()),
            sa.Column("updated_at", sa.DateTime(), nullable=False, server_default=sa.func.now()),
        )
    else:
        _add_if_missing("kitchen_supplier", sa.Column("active", sa.Boolean(), nullable=False, server_default=sa.true()))
        _add_if_missing("kitchen_supplier", sa.Column("created_at", sa.DateTime(), nullable=False, server_default=sa.func.now()))
        _add_if_missing("kitchen_supplier", sa.Column("updated_at", sa.DateTime(), nullable=False, server_default=sa.func.now()))

    # Ingredient. Existing MVP unit_price meant price/kg; kg defaults preserve that meaning.
    if not _has_table("kitchen_ingredient"):
        op.create_table(
            "kitchen_ingredient",
            sa.Column("id", sa.Integer(), primary_key=True),
            sa.Column("name", sa.String(100), nullable=False, unique=True),
            sa.Column("supplier_id", sa.Integer(), sa.ForeignKey("kitchen_supplier.id")),
            sa.Column("purchase_unit", sa.String(20), nullable=False, server_default="kg"),
            sa.Column("grams_per_purchase_unit", sa.Numeric(14, 3), nullable=False, server_default="1000"),
            sa.Column("unit_price", sa.Numeric(14, 4), nullable=False, server_default="0"),
            sa.Column("order_increment", sa.Numeric(14, 4), nullable=False, server_default="0.001"),
            sa.Column("note", sa.String(255)),
            sa.Column("active", sa.Boolean(), nullable=False, server_default=sa.true()),
            sa.Column("created_at", sa.DateTime(), nullable=False, server_default=sa.func.now()),
            sa.Column("updated_at", sa.DateTime(), nullable=False, server_default=sa.func.now()),
        )
    else:
        _add_if_missing("kitchen_ingredient", sa.Column("purchase_unit", sa.String(20), nullable=False, server_default="kg"))
        _add_if_missing("kitchen_ingredient", sa.Column("grams_per_purchase_unit", sa.Numeric(14, 3), nullable=False, server_default="1000"))
        _add_if_missing("kitchen_ingredient", sa.Column("order_increment", sa.Numeric(14, 4), nullable=False, server_default="0.001"))
        _add_if_missing("kitchen_ingredient", sa.Column("active", sa.Boolean(), nullable=False, server_default=sa.true()))
        _add_if_missing("kitchen_ingredient", sa.Column("created_at", sa.DateTime(), nullable=False, server_default=sa.func.now()))
        _add_if_missing("kitchen_ingredient", sa.Column("updated_at", sa.DateTime(), nullable=False, server_default=sa.func.now()))

    # Recipe
    if not _has_table("kitchen_recipe"):
        op.create_table(
            "kitchen_recipe",
            sa.Column("id", sa.Integer(), primary_key=True),
            sa.Column("name", sa.String(120), nullable=False, unique=True),
            sa.Column("category", sa.String(50)),
            sa.Column("serving_output_g", sa.Numeric(10, 2)),
            sa.Column("note", sa.String(255)),
            sa.Column("active", sa.Boolean(), nullable=False, server_default=sa.true()),
            sa.Column("created_at", sa.DateTime(), nullable=False, server_default=sa.func.now()),
            sa.Column("updated_at", sa.DateTime(), nullable=False, server_default=sa.func.now()),
        )
    else:
        _add_if_missing("kitchen_recipe", sa.Column("active", sa.Boolean(), nullable=False, server_default=sa.true()))
        _add_if_missing("kitchen_recipe", sa.Column("created_at", sa.DateTime(), nullable=False, server_default=sa.func.now()))
        _add_if_missing("kitchen_recipe", sa.Column("updated_at", sa.DateTime(), nullable=False, server_default=sa.func.now()))

    if not _has_table("kitchen_recipe_ingredient"):
        op.create_table(
            "kitchen_recipe_ingredient",
            sa.Column("id", sa.Integer(), primary_key=True),
            sa.Column("recipe_id", sa.Integer(), sa.ForeignKey("kitchen_recipe.id", ondelete="CASCADE"), nullable=False),
            sa.Column("ingredient_id", sa.Integer(), sa.ForeignKey("kitchen_ingredient.id"), nullable=False),
            sa.Column("grams_per_person", sa.Numeric(12, 3), nullable=False),
            sa.UniqueConstraint("recipe_id", "ingredient_id", name="uq_kitchen_recipe_ingredient"),
        )

    # Menu
    if not _has_table("kitchen_menu_plan"):
        op.create_table(
            "kitchen_menu_plan",
            sa.Column("id", sa.Integer(), primary_key=True),
            sa.Column("service_date", sa.Date(), nullable=False, index=True),
            sa.Column("meal_type", sa.String(30), nullable=False, server_default="午餐"),
            sa.Column("name", sa.String(120), nullable=False, server_default="中央菜單"),
            sa.Column("note", sa.String(255)),
            sa.Column("status", sa.String(20), nullable=False, server_default="draft"),
            sa.Column("created_at", sa.DateTime(), nullable=False, server_default=sa.func.now()),
            sa.Column("updated_at", sa.DateTime(), nullable=False, server_default=sa.func.now()),
            sa.UniqueConstraint("service_date", "meal_type", "name", name="uq_kitchen_menu_plan_identity"),
        )
    else:
        _add_if_missing("kitchen_menu_plan", sa.Column("status", sa.String(20), nullable=False, server_default="draft"))
        _add_if_missing("kitchen_menu_plan", sa.Column("created_at", sa.DateTime(), nullable=False, server_default=sa.func.now()))
        _add_if_missing("kitchen_menu_plan", sa.Column("updated_at", sa.DateTime(), nullable=False, server_default=sa.func.now()))

    if not _has_table("kitchen_menu_plan_item"):
        op.create_table(
            "kitchen_menu_plan_item",
            sa.Column("id", sa.Integer(), primary_key=True),
            sa.Column("plan_id", sa.Integer(), sa.ForeignKey("kitchen_menu_plan.id", ondelete="CASCADE"), nullable=False),
            sa.Column("recipe_id", sa.Integer(), sa.ForeignKey("kitchen_recipe.id"), nullable=False),
            sa.Column("sort_order", sa.Integer(), nullable=False, server_default="0"),
            sa.UniqueConstraint("plan_id", "recipe_id", name="uq_kitchen_menu_plan_recipe"),
        )

    if not _has_table("kitchen_menu_assignment"):
        op.create_table(
            "kitchen_menu_assignment",
            sa.Column("id", sa.Integer(), primary_key=True),
            sa.Column("plan_id", sa.Integer(), sa.ForeignKey("kitchen_menu_plan.id", ondelete="CASCADE"), nullable=False),
            sa.Column("school_id", sa.Integer(), sa.ForeignKey("kitchen_school.id"), nullable=False),
            sa.Column("headcount", sa.Integer(), nullable=False, server_default="0"),
            sa.UniqueConstraint("plan_id", "school_id", name="uq_kitchen_menu_assignment"),
        )

    # Persisted purchase orders / snapshots
    if not _has_table("kitchen_purchase_order"):
        op.create_table(
            "kitchen_purchase_order",
            sa.Column("id", sa.Integer(), primary_key=True),
            sa.Column("service_date", sa.Date(), nullable=False, index=True),
            sa.Column("supplier_id", sa.Integer(), sa.ForeignKey("kitchen_supplier.id")),
            sa.Column("supplier_key", sa.String(100), nullable=False),
            sa.Column("supplier_name_snapshot", sa.String(120), nullable=False),
            sa.Column("status", sa.String(20), nullable=False, server_default="draft"),
            sa.Column("note", sa.String(255)),
            sa.Column("created_at", sa.DateTime(), nullable=False, server_default=sa.func.now()),
            sa.Column("updated_at", sa.DateTime(), nullable=False, server_default=sa.func.now()),
            sa.UniqueConstraint("service_date", "supplier_key", name="uq_kitchen_purchase_order_supplier_day"),
        )

    if not _has_table("kitchen_purchase_order_item"):
        op.create_table(
            "kitchen_purchase_order_item",
            sa.Column("id", sa.Integer(), primary_key=True),
            sa.Column("order_id", sa.Integer(), sa.ForeignKey("kitchen_purchase_order.id", ondelete="CASCADE"), nullable=False),
            sa.Column("ingredient_id", sa.Integer(), sa.ForeignKey("kitchen_ingredient.id")),
            sa.Column("ingredient_name_snapshot", sa.String(120), nullable=False),
            sa.Column("required_grams", sa.Numeric(16, 3), nullable=False, server_default="0"),
            sa.Column("required_qty", sa.Numeric(16, 4), nullable=False, server_default="0"),
            sa.Column("purchase_unit_snapshot", sa.String(20), nullable=False),
            sa.Column("grams_per_purchase_unit_snapshot", sa.Numeric(16, 3), nullable=False),
            sa.Column("recommended_order_qty", sa.Numeric(16, 4), nullable=False, server_default="0"),
            sa.Column("actual_order_qty", sa.Numeric(16, 4), nullable=False, server_default="0"),
            sa.Column("unit_price_snapshot", sa.Numeric(16, 4), nullable=False, server_default="0"),
            sa.Column("amount", sa.Numeric(18, 4), nullable=False, server_default="0"),
            sa.Column("note", sa.String(255)),
            sa.UniqueConstraint("order_id", "ingredient_id", name="uq_kitchen_purchase_order_item"),
        )


def downgrade():
    # Downgrade remains kitchen-only. It intentionally never touches attendance tables.
    for table in ("kitchen_purchase_order_item", "kitchen_purchase_order"):
        if _has_table(table):
            op.drop_table(table)

    additions = {
        "kitchen_menu_plan": ("updated_at", "created_at", "status"),
        "kitchen_recipe": ("updated_at", "created_at", "active"),
        "kitchen_ingredient": ("updated_at", "created_at", "active", "order_increment", "grams_per_purchase_unit", "purchase_unit"),
        "kitchen_supplier": ("updated_at", "created_at", "active"),
        "kitchen_school": ("updated_at", "created_at"),
    }
    for table, columns in additions.items():
        if _has_table(table):
            for column in columns:
                if column in _columns(table):
                    op.drop_column(table, column)
