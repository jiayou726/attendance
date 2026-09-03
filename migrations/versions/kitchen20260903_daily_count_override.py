"""add editable counts to the daily kitchen sheet

Revision ID: kitchen20260903_count_override
Revises: kitchen20260903_daily_note
"""

from alembic import op
import sqlalchemy as sa


revision = "kitchen20260903_count_override"
down_revision = "kitchen20260903_daily_note"
branch_labels = None
depends_on = None


def _columns():
    inspector = sa.inspect(op.get_bind())
    if not inspector.has_table("kitchen_daily_dish_note"):
        return set()
    return {
        column["name"]
        for column in inspector.get_columns("kitchen_daily_dish_note")
    }


def upgrade():
    columns = _columns()
    for column_name in ("combo_count", "bento_count", "small_bento_count"):
        if column_name not in columns:
            op.add_column(
                "kitchen_daily_dish_note",
                sa.Column(column_name, sa.Integer(), nullable=True),
            )


def downgrade():
    # Non-destructive: keep historical manually adjusted production counts.
    pass
