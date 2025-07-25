"""Extend p_type length

Revision ID: 89e0b1a7a381
Revises: 90d40742f620
Create Date: 2025-07-17 01:20:25.000000
"""
from alembic import op
import sqlalchemy as sa

# revision identifiers, used by Alembic.
revision = '89e0b1a7a381'
down_revision = None
branch_labels = None
depends_on = None


def upgrade():
    op.alter_column('checkin', 'p_type',
                    existing_type=sa.String(length=3),
                    type_=sa.String(length=8),
                    existing_nullable=False)


def downgrade():
    op.alter_column('checkin', 'p_type',
                    existing_type=sa.String(length=8),
                    type_=sa.String(length=3),
                    existing_nullable=False)