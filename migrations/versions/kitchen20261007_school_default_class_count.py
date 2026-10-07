"""add default class count to kitchen schools

Revision ID: kitchen20261007_school_classes
Revises: kitchen20260921_manual
"""

from alembic import op
import sqlalchemy as sa


revision = "kitchen20261007_school_classes"
down_revision = "kitchen20260921_manual"
branch_labels = None
depends_on = None


DEFAULT_CLASS_COUNTS = {
    "桃園市中壢區信義國小": 22,
    "桃園市平鎮區新勢國小": 26,
    "桃園市平鎮區平鎮高中": 16,
    "桃園市中壢區新明國中": 13,
    "桃園市中壢區中平國小": 23,
    "明原": 1,
    "桃園市中壢區國立中央大學附設私立幼兒園": 2,
    "信義幼兒園": 2,
    "平鎮高中 (小便當)": 1,
    "廣豐食品股份有限公司": 1,
}


def upgrade():
    inspector = sa.inspect(op.get_bind())
    if not inspector.has_table("kitchen_school"):
        return
    columns = {column["name"] for column in inspector.get_columns("kitchen_school")}
    if "default_class_count" not in columns:
        op.add_column(
            "kitchen_school",
            sa.Column(
                "default_class_count",
                sa.Integer(),
                nullable=False,
                server_default="0",
            ),
        )

    school = sa.table(
        "kitchen_school",
        sa.column("name", sa.String()),
        sa.column("default_class_count", sa.Integer()),
    )
    for name, class_count in DEFAULT_CLASS_COUNTS.items():
        op.execute(
            school.update()
            .where(school.c.name == name)
            .values(default_class_count=class_count)
        )


def downgrade():
    # Non-destructive: keep school defaults if an older release is restored.
    pass
