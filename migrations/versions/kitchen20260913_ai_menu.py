"""add AI menu nutrition, tags, meal profiles and drafts

Revision ID: kitchen20260913_ai_menu
Revises: kitchen20260904_school_classes
"""
from alembic import op
import sqlalchemy as sa

revision = "kitchen20260913_ai_menu"
down_revision = "kitchen20260904_school_classes"
branch_labels = None
depends_on = None

INGREDIENT_COLUMNS = (
    ("kcal_per_100g", sa.Numeric(10, 2), True, None),
    ("edible_grams_per_unit", sa.Numeric(12, 3), True, None),
    ("nutrition_source", sa.String(120), True, None),
    ("nutrition_food_name", sa.String(160), True, None),
    ("nutrition_code", sa.String(40), True, None),
    ("nutrition_confidence", sa.String(20), True, None),
    ("nutrition_status", sa.String(20), False, "pending"),
    ("nutrition_verified", sa.Boolean(), False, sa.false()),
    ("nutrition_note", sa.String(255), True, None),
)


def upgrade():
    bind = op.get_bind()
    inspector = sa.inspect(bind)

    if inspector.has_table("kitchen_ingredient"):
        existing = {c["name"] for c in inspector.get_columns("kitchen_ingredient")}
        for name, type_, nullable, default in INGREDIENT_COLUMNS:
            if name in existing:
                continue
            op.add_column(
                "kitchen_ingredient",
                sa.Column(name, type_, nullable=nullable, server_default=default),
            )

    inspector = sa.inspect(bind)
    if not inspector.has_table("kitchen_recipe_tag"):
        op.create_table(
            "kitchen_recipe_tag",
            sa.Column("id", sa.Integer(), primary_key=True),
            sa.Column("recipe_id", sa.Integer(), sa.ForeignKey("kitchen_recipe.id", ondelete="CASCADE"), nullable=False),
            sa.Column("tag", sa.String(30), nullable=False),
            sa.Column("source", sa.String(20), nullable=False, server_default="manual"),
            sa.Column("created_at", sa.DateTime(), nullable=False),
            sa.UniqueConstraint("recipe_id", "tag", name="uq_kitchen_recipe_tag"),
        )

    inspector = sa.inspect(bind)
    if not inspector.has_table("kitchen_meal_profile"):
        op.create_table(
            "kitchen_meal_profile",
            sa.Column("id", sa.Integer(), primary_key=True),
            sa.Column("code", sa.String(40), nullable=False, unique=True),
            sa.Column("name", sa.String(80), nullable=False, unique=True),
            sa.Column("meal_type", sa.String(30), nullable=False, server_default="午餐"),
            sa.Column("kcal_min", sa.Numeric(8, 2), nullable=True),
            sa.Column("kcal_max", sa.Numeric(8, 2), nullable=True),
            sa.Column("source_document", sa.String(255), nullable=True),
            sa.Column("source_url", sa.String(500), nullable=True),
            sa.Column("source_version", sa.String(60), nullable=True),
            sa.Column("note", sa.String(500), nullable=True),
            sa.Column("sort_order", sa.Integer(), nullable=False, server_default="0"),
            sa.Column("active", sa.Boolean(), nullable=False, server_default=sa.true()),
            sa.Column("created_at", sa.DateTime(), nullable=False),
            sa.Column("updated_at", sa.DateTime(), nullable=False),
        )

    inspector = sa.inspect(bind)
    if not inspector.has_table("kitchen_menu_draft"):
        op.create_table(
            "kitchen_menu_draft",
            sa.Column("id", sa.Integer(), primary_key=True),
            sa.Column("name", sa.String(120), nullable=False, server_default="自動排菜"),
            sa.Column("profile_id", sa.Integer(), sa.ForeignKey("kitchen_meal_profile.id"), nullable=True),
            sa.Column("start_date", sa.Date(), nullable=False),
            sa.Column("end_date", sa.Date(), nullable=False),
            sa.Column("include_weekends", sa.Boolean(), nullable=False, server_default=sa.false()),
            sa.Column("meal_type", sa.String(30), nullable=False, server_default="午餐"),
            sa.Column("structure_json", sa.Text(), nullable=False, server_default="{}"),
            sa.Column("rules_json", sa.Text(), nullable=False, server_default="{}"),
            sa.Column("status", sa.String(20), nullable=False, server_default="draft"),
            sa.Column("applied_at", sa.DateTime(), nullable=True),
            sa.Column("created_at", sa.DateTime(), nullable=False),
            sa.Column("updated_at", sa.DateTime(), nullable=False),
        )

    inspector = sa.inspect(bind)
    if not inspector.has_table("kitchen_menu_draft_item"):
        op.create_table(
            "kitchen_menu_draft_item",
            sa.Column("id", sa.Integer(), primary_key=True),
            sa.Column("draft_id", sa.Integer(), sa.ForeignKey("kitchen_menu_draft.id", ondelete="CASCADE"), nullable=False),
            sa.Column("service_date", sa.Date(), nullable=False, index=True),
            sa.Column("category", sa.String(50), nullable=False),
            sa.Column("slot_index", sa.Integer(), nullable=False, server_default="1"),
            sa.Column("sort_order", sa.Integer(), nullable=False, server_default="0"),
            sa.Column("recipe_id", sa.Integer(), sa.ForeignKey("kitchen_recipe.id"), nullable=True),
            sa.Column("locked", sa.Boolean(), nullable=False, server_default=sa.false()),
            sa.Column("created_at", sa.DateTime(), nullable=False),
            sa.Column("updated_at", sa.DateTime(), nullable=False),
            sa.UniqueConstraint("draft_id", "service_date", "category", "slot_index", name="uq_kitchen_menu_draft_slot"),
        )


def downgrade():
    # 保守回滾：不自動刪除營養資料或使用者已建立的 AI 草稿。
    pass
