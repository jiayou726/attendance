"""Allow editing a dish category for one menu occurrence without changing the master.

Revision ID: kitchen20261104_category
Revises: kitchen20261010_estimates
"""
from alembic import op
import sqlalchemy as sa

revision = "kitchen20261104_category"
down_revision = "kitchen20261010_estimates"
branch_labels = None
depends_on = None


def upgrade():
    inspector = sa.inspect(op.get_bind())
    if not inspector.has_table("kitchen_menu_plan_item"):
        return
    columns = {column["name"] for column in inspector.get_columns("kitchen_menu_plan_item")}
    if "category_override" not in columns:
        op.add_column("kitchen_menu_plan_item", sa.Column(
            "category_override", sa.String(length=20), nullable=True
        ))


def downgrade():
    # Keep existing per-menu edits rather than deleting their categories.
    pass
