"""add editable class count to the daily kitchen sheet

Revision ID: kitchen20260903_class_count
Revises: kitchen20260903_count_override
"""

from alembic import op
import sqlalchemy as sa


revision = "kitchen20260903_class_count"
down_revision = "kitchen20260903_count_override"
branch_labels = None
depends_on = None


def upgrade():
    inspector = sa.inspect(op.get_bind())
    if not inspector.has_table("kitchen_daily_dish_note"):
        return
    columns = {
        column["name"]
        for column in inspector.get_columns("kitchen_daily_dish_note")
    }
    if "class_count" not in columns:
        op.add_column(
            "kitchen_daily_dish_note",
            sa.Column("class_count", sa.Integer(), nullable=True),
        )


def downgrade():
    # Non-destructive: keep historical manually entered class counts.
    pass
