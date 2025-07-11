"""initial

Revision ID: 8b92ec97ac5c
Revises: 
Create Date: 2025-07-07 12:07:33.729315

"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = '8b92ec97ac5c'
down_revision = None
branch_labels = None
depends_on = None


def upgrade():
    # ### commands auto generated by Alembic - please adjust! ###
    op.create_table('employees',
    sa.Column('id', sa.String(length=16), nullable=False),
    sa.Column('name', sa.String(length=32), nullable=False),
    sa.Column('area', sa.String(length=16), nullable=False),
    sa.PrimaryKeyConstraint('id')
    )
    op.create_table('checkins',
    sa.Column('id', sa.Integer(), nullable=False),
    sa.Column('employee_id', sa.String(length=16), nullable=True),
    sa.Column('work_date', sa.String(length=10), nullable=True),
    sa.Column('p_type', sa.String(length=8), nullable=True),
    sa.Column('ts', sa.String(length=20), nullable=True),
    sa.ForeignKeyConstraint(['employee_id'], ['employees.id'], ),
    sa.PrimaryKeyConstraint('id'),
    sa.UniqueConstraint('employee_id', 'work_date', 'p_type')
    )
    # ### end Alembic commands ###


def downgrade():
    # ### commands auto generated by Alembic - please adjust! ###
    op.drop_table('checkins')
    op.drop_table('employees')
    # ### end Alembic commands ###
