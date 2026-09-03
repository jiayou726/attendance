"""add vegetarian headcount to schools

Revision ID: kitchen20260830_vegetarian
Revises: kitchen20260829_school_no_service
"""

from alembic import op
import sqlalchemy as sa


revision = "kitchen20260830_vegetarian"
down_revision = "kitchen20260829"
branch_labels = None
depends_on = None


def upgrade():
    op.add_column(
        "kitchen_school",
        sa.Column("default_vegetarian_headcount", sa.Integer(), nullable=False, server_default="0"),
    )


def downgrade():
    op.drop_column("kitchen_school", "default_vegetarian_headcount")
