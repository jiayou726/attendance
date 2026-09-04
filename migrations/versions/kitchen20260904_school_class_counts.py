"""store daily class counts per school

Revision ID: kitchen20260904_school_classes
Revises: kitchen20260903_class_count
"""

from alembic import op
import sqlalchemy as sa


revision = "kitchen20260904_school_classes"
down_revision = "kitchen20260903_class_count"
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
    if "school_class_counts" not in columns:
        op.add_column(
            "kitchen_daily_dish_note",
            sa.Column(
                "school_class_counts",
                sa.Text(),
                nullable=False,
                server_default="{}",
            ),
        )


def downgrade():
    # Non-destructive: keep the per-school counts if an older release is restored.
    pass
