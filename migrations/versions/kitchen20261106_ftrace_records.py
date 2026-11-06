"""Add persistent ftrace documents, immutable revisions, and receipt imports.

Revision ID: kitchen20261106_ftrace
Revises: kitchen20261104_category
"""
from alembic import op
from ftrace_models import TABLES
from extensions import db
from sqlalchemy import text
revision='kitchen20261106_ftrace'
down_revision='kitchen20261104_category'
branch_labels=None
depends_on=None

def upgrade():
    connection=op.get_bind()
    db.metadata.create_all(connection,tables=list(TABLES))
    if connection.dialect.name=='postgresql':
        for table in TABLES:
            connection.execute(text(f'ALTER TABLE public.{table.name} ENABLE ROW LEVEL SECURITY'))
            connection.execute(text(f'REVOKE ALL ON TABLE public.{table.name} FROM PUBLIC, anon, authenticated'))

def downgrade():
    # Do not destroy receipt provenance or submitted revision history.
    pass
