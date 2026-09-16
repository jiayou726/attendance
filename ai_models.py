"""AI 菜單的資料模型擴充。只做新增欄位與新增草稿表。"""
from datetime import datetime
from sqlalchemy import inspect, text
from extensions import db
import models as core


def _attach_column(model, name, column):
    if not hasattr(model, name):
        setattr(model, name, column)

_attach_column(core.KitchenIngredient, "kcal_per_100g", db.Column(db.Numeric(10, 2), nullable=True))
_attach_column(core.KitchenIngredient, "edible_grams_per_unit", db.Column(db.Numeric(12, 3), nullable=True))
_attach_column(core.KitchenIngredient, "nutrition_source", db.Column(db.String(120), nullable=True))
_attach_column(core.KitchenIngredient, "nutrition_food_name", db.Column(db.String(160), nullable=True))
_attach_column(core.KitchenIngredient, "nutrition_code", db.Column(db.String(40), nullable=True))
_attach_column(core.KitchenIngredient, "nutrition_confidence", db.Column(db.String(20), nullable=True))
_attach_column(core.KitchenIngredient, "nutrition_status", db.Column(db.String(20), nullable=False, default="pending"))
_attach_column(core.KitchenIngredient, "nutrition_verified", db.Column(db.Boolean, nullable=False, default=False))
_attach_column(core.KitchenIngredient, "nutrition_note", db.Column(db.String(255), nullable=True))

# 葷素主分類是菜色主檔的一部分，不再靠菜名「(素)」或關鍵字推測。
# meat：葷食專用；vegetarian：素食專用；shared：葷素共用。
_attach_column(core.KitchenRecipe, "diet_type", db.Column(db.String(20), nullable=True))


def _has_kcal(self):
    if self.kcal_per_100g is None:
        return False
    if (self.base_unit or "g") == "g":
        return True
    return self.edible_grams_per_unit is not None and self.edible_grams_per_unit > 0

if not hasattr(core.KitchenIngredient, "has_kcal"):
    core.KitchenIngredient.has_kcal = property(_has_kcal)

class KitchenRecipeTag(db.Model):
    __tablename__ = "kitchen_recipe_tag"
    id = db.Column(db.Integer, primary_key=True)
    recipe_id = db.Column(db.Integer, db.ForeignKey("kitchen_recipe.id", ondelete="CASCADE"), nullable=False)
    tag = db.Column(db.String(30), nullable=False)
    source = db.Column(db.String(20), nullable=False, default="manual")
    created_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)
    recipe = db.relationship(core.KitchenRecipe, back_populates="tags")
    __table_args__ = (db.UniqueConstraint("recipe_id", "tag", name="uq_kitchen_recipe_tag"),)

if not hasattr(core.KitchenRecipe, "tags"):
    core.KitchenRecipe.tags = db.relationship(KitchenRecipeTag, back_populates="recipe", cascade="all, delete-orphan", order_by=KitchenRecipeTag.tag)

class KitchenMealProfile(db.Model):
    __tablename__ = "kitchen_meal_profile"
    id = db.Column(db.Integer, primary_key=True)
    code = db.Column(db.String(40), nullable=False, unique=True)
    name = db.Column(db.String(80), nullable=False, unique=True)
    meal_type = db.Column(db.String(30), nullable=False, default="午餐")
    kcal_min = db.Column(db.Numeric(8, 2), nullable=True)
    kcal_max = db.Column(db.Numeric(8, 2), nullable=True)
    source_document = db.Column(db.String(255), nullable=True)
    source_url = db.Column(db.String(500), nullable=True)
    source_version = db.Column(db.String(60), nullable=True)
    note = db.Column(db.String(500), nullable=True)
    sort_order = db.Column(db.Integer, nullable=False, default=0)
    active = db.Column(db.Boolean, nullable=False, default=True)
    created_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow, onupdate=datetime.utcnow)

    @property
    def kcal_label(self):
        if self.kcal_min is None or self.kcal_max is None:
            return "未設定熱量基準"
        return f"{int(self.kcal_min)}～{int(self.kcal_max)} kcal"

class KitchenMenuDraft(db.Model):
    __tablename__ = "kitchen_menu_draft"
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(120), nullable=False, default="自動排菜")
    profile_id = db.Column(db.Integer, db.ForeignKey("kitchen_meal_profile.id"), nullable=True)
    start_date = db.Column(db.Date, nullable=False)
    end_date = db.Column(db.Date, nullable=False)
    include_weekends = db.Column(db.Boolean, nullable=False, default=False)
    meal_type = db.Column(db.String(30), nullable=False, default="午餐")
    structure_json = db.Column(db.Text, nullable=False, default="{}")
    rules_json = db.Column(db.Text, nullable=False, default="{}")
    status = db.Column(db.String(20), nullable=False, default="draft")
    applied_at = db.Column(db.DateTime, nullable=True)
    created_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow, onupdate=datetime.utcnow)
    profile = db.relationship(KitchenMealProfile)
    items = db.relationship("KitchenMenuDraftItem", back_populates="draft", cascade="all, delete-orphan")

class KitchenMenuDraftItem(db.Model):
    __tablename__ = "kitchen_menu_draft_item"
    id = db.Column(db.Integer, primary_key=True)
    draft_id = db.Column(db.Integer, db.ForeignKey("kitchen_menu_draft.id", ondelete="CASCADE"), nullable=False)
    service_date = db.Column(db.Date, nullable=False, index=True)
    category = db.Column(db.String(50), nullable=False)
    slot_index = db.Column(db.Integer, nullable=False, default=1)
    sort_order = db.Column(db.Integer, nullable=False, default=0)
    recipe_id = db.Column(db.Integer, db.ForeignKey("kitchen_recipe.id"), nullable=True)
    locked = db.Column(db.Boolean, nullable=False, default=False)
    created_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow, onupdate=datetime.utcnow)
    draft = db.relationship(KitchenMenuDraft, back_populates="items")
    recipe = db.relationship(core.KitchenRecipe)
    __table_args__ = (db.UniqueConstraint("draft_id", "service_date", "category", "slot_index", name="uq_kitchen_menu_draft_slot"),)

core.KitchenRecipeTag = KitchenRecipeTag
core.KitchenMealProfile = KitchenMealProfile
core.KitchenMenuDraft = KitchenMenuDraft
core.KitchenMenuDraftItem = KitchenMenuDraftItem

NUTRITION_COLUMNS = {
    "kcal_per_100g": "NUMERIC(10, 2)",
    "edible_grams_per_unit": "NUMERIC(12, 3)",
    "nutrition_source": "VARCHAR(120)",
    "nutrition_food_name": "VARCHAR(160)",
    "nutrition_code": "VARCHAR(40)",
    "nutrition_confidence": "VARCHAR(20)",
    "nutrition_status": "VARCHAR(20) NOT NULL DEFAULT 'pending'",
    "nutrition_verified": "BOOLEAN NOT NULL DEFAULT FALSE",
    "nutrition_note": "VARCHAR(255)",
}

RECIPE_COLUMNS = {
    "diet_type": "VARCHAR(20)",
}

def ensure_schema(engine):
    inspector = inspect(engine)
    if inspector.has_table("kitchen_ingredient"):
        existing = {c["name"] for c in inspector.get_columns("kitchen_ingredient")}
        with engine.begin() as connection:
            for name, definition in NUTRITION_COLUMNS.items():
                if name not in existing:
                    connection.execute(text(f"ALTER TABLE kitchen_ingredient ADD COLUMN {name} {definition}"))
    if inspector.has_table("kitchen_recipe"):
        existing = {c["name"] for c in inspector.get_columns("kitchen_recipe")}
        with engine.begin() as connection:
            for name, definition in RECIPE_COLUMNS.items():
                if name not in existing:
                    connection.execute(text(f"ALTER TABLE kitchen_recipe ADD COLUMN {name} {definition}"))
    for model in (KitchenRecipeTag, KitchenMealProfile, KitchenMenuDraft, KitchenMenuDraftItem):
        model.__table__.create(bind=engine, checkfirst=True)
