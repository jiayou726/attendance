# models.py
from extensions import db


class Employee(db.Model):
    __tablename__ = "employee"
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(50), nullable=False)
    area = db.Column(db.String(50))
    default_break = db.Column(db.Float, nullable=False, default=0.0)


class Checkin(db.Model):
    __tablename__ = "checkin"
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey("employee.id"), nullable=False)
    work_date = db.Column(db.String(10), nullable=False)
    p_type = db.Column(db.String(10), nullable=False)
    ts = db.Column(db.String(19), nullable=False)
    note = db.Column(db.String(50))

    __table_args__ = (
        db.UniqueConstraint("employee_id", "work_date", "p_type"),
    )


# ─────────────────────────────────────────────
# 團膳 / 萬能廚師替代模組
# 中央菜單 → 分配學校與人數 → Recipe BOM → 採購彙總
# ─────────────────────────────────────────────


class KitchenSchool(db.Model):
    __tablename__ = "kitchen_school"

    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False, unique=True)
    code = db.Column(db.String(50), unique=True)
    active = db.Column(db.Boolean, nullable=False, default=True)


class KitchenSupplier(db.Model):
    __tablename__ = "kitchen_supplier"

    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False, unique=True)
    phone = db.Column(db.String(50))
    note = db.Column(db.String(255))


class KitchenIngredient(db.Model):
    __tablename__ = "kitchen_ingredient"

    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False, unique=True)
    # 採購單價預設以每 kg 計價；每人配方固定以 g 儲存
    unit_price = db.Column(db.Numeric(12, 4), nullable=False, default=0)
    supplier_id = db.Column(
        db.Integer,
        db.ForeignKey("kitchen_supplier.id"),
        nullable=True,
    )
    note = db.Column(db.String(255))

    supplier = db.relationship("KitchenSupplier")


class KitchenRecipe(db.Model):
    __tablename__ = "kitchen_recipe"

    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(120), nullable=False, unique=True)
    category = db.Column(db.String(50))  # 主食 / 主菜 / 副菜 / 湯品 / 青菜...
    # 成品預計每人打菜量（例如 95g）；可空白，不參與採購計算
    serving_output_g = db.Column(db.Numeric(10, 2), nullable=True)
    note = db.Column(db.String(255))

    ingredients = db.relationship(
        "KitchenRecipeIngredient",
        back_populates="recipe",
        cascade="all, delete-orphan",
        order_by="KitchenRecipeIngredient.id",
    )


class KitchenRecipeIngredient(db.Model):
    __tablename__ = "kitchen_recipe_ingredient"

    id = db.Column(db.Integer, primary_key=True)
    recipe_id = db.Column(
        db.Integer,
        db.ForeignKey("kitchen_recipe.id", ondelete="CASCADE"),
        nullable=False,
    )
    ingredient_id = db.Column(
        db.Integer,
        db.ForeignKey("kitchen_ingredient.id"),
        nullable=False,
    )
    # AP 採購個人量：每人需要幾克，例如骨腿丁 88g
    grams_per_person = db.Column(db.Numeric(10, 3), nullable=False)

    recipe = db.relationship("KitchenRecipe", back_populates="ingredients")
    ingredient = db.relationship("KitchenIngredient")

    __table_args__ = (
        db.UniqueConstraint("recipe_id", "ingredient_id"),
    )


class KitchenMenuPlan(db.Model):
    __tablename__ = "kitchen_menu_plan"

    id = db.Column(db.Integer, primary_key=True)
    service_date = db.Column(db.Date, nullable=False, index=True)
    meal_type = db.Column(db.String(30), nullable=False, default="午餐")
    name = db.Column(db.String(120))
    note = db.Column(db.String(255))

    items = db.relationship(
        "KitchenMenuPlanItem",
        back_populates="plan",
        cascade="all, delete-orphan",
        order_by="KitchenMenuPlanItem.sort_order",
    )
    assignments = db.relationship(
        "KitchenMenuAssignment",
        back_populates="plan",
        cascade="all, delete-orphan",
    )

    __table_args__ = (
        db.UniqueConstraint("service_date", "meal_type", "name"),
    )


class KitchenMenuPlanItem(db.Model):
    __tablename__ = "kitchen_menu_plan_item"

    id = db.Column(db.Integer, primary_key=True)
    plan_id = db.Column(
        db.Integer,
        db.ForeignKey("kitchen_menu_plan.id", ondelete="CASCADE"),
        nullable=False,
    )
    recipe_id = db.Column(
        db.Integer,
        db.ForeignKey("kitchen_recipe.id"),
        nullable=False,
    )
    sort_order = db.Column(db.Integer, nullable=False, default=0)

    plan = db.relationship("KitchenMenuPlan", back_populates="items")
    recipe = db.relationship("KitchenRecipe")

    __table_args__ = (
        db.UniqueConstraint("plan_id", "recipe_id"),
    )


class KitchenMenuAssignment(db.Model):
    __tablename__ = "kitchen_menu_assignment"

    id = db.Column(db.Integer, primary_key=True)
    plan_id = db.Column(
        db.Integer,
        db.ForeignKey("kitchen_menu_plan.id", ondelete="CASCADE"),
        nullable=False,
    )
    school_id = db.Column(
        db.Integer,
        db.ForeignKey("kitchen_school.id"),
        nullable=False,
    )
    headcount = db.Column(db.Integer, nullable=False, default=0)

    plan = db.relationship("KitchenMenuPlan", back_populates="assignments")
    school = db.relationship("KitchenSchool")

    __table_args__ = (
        db.UniqueConstraint("plan_id", "school_id"),
    )
