"""Add explicit per-school no-service status for procurement guardrails.

Revision ID: kitchen20260829
Revises: kitchen20260828
"""

from alembic import op
import sqlalchemy as sa


revision = "kitchen20260829"
down_revision = "kitchen20260828"
branch_labels = None
depends_on = None


def _columns(table):
    inspector = sa.inspect(op.get_bind())
    if not inspector.has_table(table):
        return set()
    return {column["name"] for column in inspector.get_columns(table)}


def upgrade():
    if "service_status" not in _columns("kitchen_menu_assignment"):
        op.add_column(
            "kitchen_menu_assignment",
            sa.Column("service_status", sa.String(20), nullable=False, server_default="serving"),
        )


def downgrade():
    # Non-destructive: keep the explicit historical service decision.
    pass
