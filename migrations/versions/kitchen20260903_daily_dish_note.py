"""add editable ingredient notes for the daily kitchen sheet

Revision ID: kitchen20260903_daily_note
Revises: kitchen20260830_vegetarian
"""

from alembic import op
import sqlalchemy as sa


revision = "kitchen20260903_daily_note"
down_revision = "kitchen20260830_vegetarian"
branch_labels = None
depends_on = None


def upgrade():
    inspector = sa.inspect(op.get_bind())
    if inspector.has_table("kitchen_daily_dish_note"):
        return
    op.create_table(
        "kitchen_daily_dish_note",
        sa.Column("id", sa.Integer(), nullable=False),
        sa.Column("service_date", sa.Date(), nullable=False),
        sa.Column("variant", sa.String(length=20), nullable=False, server_default="regular"),
        sa.Column("recipe_id", sa.Integer(), nullable=False),
        sa.Column("ingredients_text", sa.Text(), nullable=False, server_default=""),
        sa.Column("updated_at", sa.DateTime(), nullable=False),
        sa.ForeignKeyConstraint(["recipe_id"], ["kitchen_recipe.id"], ondelete="CASCADE"),
        sa.PrimaryKeyConstraint("id"),
        sa.UniqueConstraint(
            "service_date", "variant", "recipe_id",
            name="uq_kitchen_daily_dish_note",
        ),
    )
    op.create_index(
        "ix_kitchen_daily_dish_note_service_date",
        "kitchen_daily_dish_note",
        ["service_date"],
        unique=False,
    )


def downgrade():
    op.drop_index(
        "ix_kitchen_daily_dish_note_service_date",
        table_name="kitchen_daily_dish_note",
    )
    op.drop_table("kitchen_daily_dish_note")
