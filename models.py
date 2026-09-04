# models.py
from datetime import datetime

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
# 團膳正式模組
# Recipe BOM（每人 AP 數量，可用 g 或 個）→ 菜單 → 學校人數 → 採購 snapshot
# ─────────────────────────────────────────────


class KitchenSchool(db.Model):
    __tablename__ = "kitchen_school"

    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False, unique=True)
    code = db.Column(db.String(50), unique=True)
    # 平常供餐人數；排菜時自動帶入，當天仍可另外修改。
    default_headcount = db.Column(db.Integer, nullable=False, default=0)
    # 平常素食人數；葷食仍沿用 default_headcount，確保舊資料相容。
    default_vegetarian_headcount = db.Column(db.Integer, nullable=False, default=0)
    active = db.Column(db.Boolean, nullable=False, default=True)
    created_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow, onupdate=datetime.utcnow)


class KitchenSupplier(db.Model):
    __tablename__ = "kitchen_supplier"

    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False, unique=True)
    phone = db.Column(db.String(50))
    mobile = db.Column(db.String(50))
    fax = db.Column(db.String(50))
    contact = db.Column(db.String(100))
    address = db.Column(db.String(255))
    source_file = db.Column(db.String(255))
    note = db.Column(db.String(255))
    active = db.Column(db.Boolean, nullable=False, default=True)
    created_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow, onupdate=datetime.utcnow)

    items = db.relationship(
        "KitchenSupplierItem",
        back_populates="supplier",
        cascade="all, delete-orphan",
        order_by="KitchenSupplierItem.name",
    )

    @property
    def historical_order_count(self):
        return sum(item.order_count or 0 for item in self.items if item.active)

    @property
    def last_purchase_date(self):
        dates = [item.last_purchase_date for item in self.items if item.active and item.last_purchase_date]
        return max(dates) if dates else None


class KitchenIngredient(db.Model):
    __tablename__ = "kitchen_ingredient"

    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False, unique=True)
    supplier_id = db.Column(db.Integer, db.ForeignKey("kitchen_supplier.id"), nullable=True)

    # 配方基本單位：大部分食材用 g；雞腿、蛋等可用 個。
    base_unit = db.Column(db.String(10), nullable=False, default="g")
    purchase_unit = db.Column(db.String(20), nullable=False, default="kg")
    # 欄位名稱沿用舊 schema；實際語意為「1 採購單位包含多少 base_unit」。
    # base_unit=g：1 kg = 1000；base_unit=個：1 箱 = 50（個）。
    grams_per_purchase_unit = db.Column(db.Numeric(14, 3), nullable=False, default=1000)
    unit_price = db.Column(db.Numeric(14, 4), nullable=False, default=0)
    order_increment = db.Column(db.Numeric(14, 4), nullable=False, default=0.001)

    note = db.Column(db.String(255))
    active = db.Column(db.Boolean, nullable=False, default=True)
    created_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow, onupdate=datetime.utcnow)

    supplier = db.relationship("KitchenSupplier")


class KitchenSupplierItem(db.Model):
    """Historical product catalog extracted from each supplier workbook."""

    __tablename__ = "kitchen_supplier_item"

    id = db.Column(db.Integer, primary_key=True)
    supplier_id = db.Column(db.Integer, db.ForeignKey("kitchen_supplier.id", ondelete="CASCADE"), nullable=False)
    ingredient_id = db.Column(db.Integer, db.ForeignKey("kitchen_ingredient.id"), nullable=True)
    source_key = db.Column(db.String(160), nullable=True)
    name = db.Column(db.String(120), nullable=False)
    unit = db.Column(db.String(20), nullable=False)
    package_conversion = db.Column(db.String(120), nullable=True)
    last_quantity = db.Column(db.Numeric(16, 3), nullable=True)
    last_unit_price = db.Column(db.Numeric(16, 4), nullable=True)
    last_purchase_date = db.Column(db.Date, nullable=True)
    order_count = db.Column(db.Integer, nullable=False, default=0)
    source_file = db.Column(db.String(255))
    manual_override = db.Column(db.Boolean, nullable=False, default=False)
    active = db.Column(db.Boolean, nullable=False, default=True)
    created_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow, onupdate=datetime.utcnow)

    supplier = db.relationship("KitchenSupplier", back_populates="items")
    ingredient = db.relationship("KitchenIngredient")

    __table_args__ = (
        db.UniqueConstraint("supplier_id", "name", name="uq_kitchen_supplier_item"),
    )


class KitchenRecipe(db.Model):
    __tablename__ = "kitchen_recipe"

    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(120), nullable=False, unique=True)
    category = db.Column(db.String(50))
    # 成品預計每人打菜量（g）；不參與 AP 採購計算。
    serving_output_g = db.Column(db.Numeric(10, 2), nullable=True)
    note = db.Column(db.String(255))
    active = db.Column(db.Boolean, nullable=False, default=True)
    created_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow, onupdate=datetime.utcnow)

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
    # 欄位名稱沿用舊 schema；實際語意為「每人 AP 數量」，單位由 ingredient.base_unit 決定。
    # 例：骨腿丁 88 g/人；棒棒腿 1 個/人。
    grams_per_person = db.Column(db.Numeric(12, 3), nullable=False)
    # manual：人工確認；estimated：由製造表反推；pending：已知材料但克數待確認。
    quantity_status = db.Column(db.String(20), nullable=False, default="manual")
    source_note = db.Column(db.String(255))

    recipe = db.relationship("KitchenRecipe", back_populates="ingredients")
    ingredient = db.relationship("KitchenIngredient")

    __table_args__ = (
        db.UniqueConstraint("recipe_id", "ingredient_id", name="uq_kitchen_recipe_ingredient"),
    )


class KitchenMenuPlan(db.Model):
    __tablename__ = "kitchen_menu_plan"

    id = db.Column(db.Integer, primary_key=True)
    service_date = db.Column(db.Date, nullable=False, index=True)
    meal_type = db.Column(db.String(30), nullable=False, default="午餐")
    name = db.Column(db.String(120), nullable=False, default="中央菜單")
    note = db.Column(db.String(255))
    status = db.Column(db.String(20), nullable=False, default="draft")
    created_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow, onupdate=datetime.utcnow)

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
        db.UniqueConstraint("service_date", "meal_type", "name", name="uq_kitchen_menu_plan_identity"),
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
        db.UniqueConstraint("plan_id", "recipe_id", name="uq_kitchen_menu_plan_recipe"),
    )


class KitchenDailyDishNote(db.Model):
    """每日廚房表的食材備註。

    配方食材是預設值；現場可在這裡追加「（數量／規格）」，
    不會回頭改掉共用的配方 BOM。
    """

    __tablename__ = "kitchen_daily_dish_note"

    id = db.Column(db.Integer, primary_key=True)
    service_date = db.Column(db.Date, nullable=False, index=True)
    variant = db.Column(db.String(20), nullable=False, default="regular")
    recipe_id = db.Column(
        db.Integer,
        db.ForeignKey("kitchen_recipe.id", ondelete="CASCADE"),
        nullable=False,
    )
    ingredients_text = db.Column(db.Text, nullable=False, default="")
    # 空白時使用學校人數自動計算；有值時保留現場人工調整。
    combo_count = db.Column(db.Integer, nullable=True)
    # JSON 物件：以學校 id 為 key，保存每校班級數；空白欄位不寫入。
    school_class_counts = db.Column(db.Text, nullable=False, default="{}")
    # 保留總班級數欄位以相容既有資料與舊版程式。
    class_count = db.Column(db.Integer, nullable=True)
    bento_count = db.Column(db.Integer, nullable=True)
    small_bento_count = db.Column(db.Integer, nullable=True)
    updated_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow, onupdate=datetime.utcnow)

    recipe = db.relationship("KitchenRecipe")

    __table_args__ = (
        db.UniqueConstraint(
            "service_date", "variant", "recipe_id",
            name="uq_kitchen_daily_dish_note",
        ),
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
    # serving：正常供餐；no_service：學校當日停餐，不納入採購。
    service_status = db.Column(db.String(20), nullable=False, default="serving")

    plan = db.relationship("KitchenMenuPlan", back_populates="assignments")
    school = db.relationship("KitchenSchool")

    __table_args__ = (
        db.UniqueConstraint("plan_id", "school_id", name="uq_kitchen_menu_assignment"),
    )


class KitchenPurchaseOrder(db.Model):
    __tablename__ = "kitchen_purchase_order"

    id = db.Column(db.Integer, primary_key=True)
    service_date = db.Column(db.Date, nullable=False, index=True)
    # 以下單頭廠商欄位保留以相容舊資料；新資料固定為每日採購單，
    # 實際供應廠商記在每一個採購品項上。
    supplier_id = db.Column(db.Integer, db.ForeignKey("kitchen_supplier.id"), nullable=True)
    supplier_key = db.Column(db.String(100), nullable=False)
    supplier_name_snapshot = db.Column(db.String(120), nullable=False)
    supplier_overridden = db.Column(db.Boolean, nullable=False, default=False)
    status = db.Column(db.String(20), nullable=False, default="draft")
    note = db.Column(db.String(255))
    created_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow, onupdate=datetime.utcnow)

    supplier = db.relationship("KitchenSupplier")
    items = db.relationship(
        "KitchenPurchaseOrderItem",
        back_populates="order",
        cascade="all, delete-orphan",
        order_by="KitchenPurchaseOrderItem.ingredient_name_snapshot",
    )

    __table_args__ = (
        db.UniqueConstraint("service_date", "supplier_key", name="uq_kitchen_purchase_order_supplier_day"),
    )


class KitchenPurchaseOrderItem(db.Model):
    __tablename__ = "kitchen_purchase_order_item"

    id = db.Column(db.Integer, primary_key=True)
    order_id = db.Column(
        db.Integer,
        db.ForeignKey("kitchen_purchase_order.id", ondelete="CASCADE"),
        nullable=False,
    )
    ingredient_id = db.Column(db.Integer, db.ForeignKey("kitchen_ingredient.id"), nullable=True)
    supplier_id = db.Column(db.Integer, db.ForeignKey("kitchen_supplier.id"), nullable=True)
    supplier_item_id = db.Column(db.Integer, db.ForeignKey("kitchen_supplier_item.id"), nullable=True)
    supplier_name_snapshot = db.Column(db.String(120), nullable=True)

    ingredient_name_snapshot = db.Column(db.String(120), nullable=False)
    base_unit_snapshot = db.Column(db.String(10), nullable=False, default="g")
    # 欄位名稱 required_grams 沿用舊 schema；實際為 required base-unit amount。
    required_grams = db.Column(db.Numeric(16, 3), nullable=False, default=0)
    required_qty = db.Column(db.Numeric(16, 4), nullable=False, default=0)
    purchase_unit_snapshot = db.Column(db.String(20), nullable=False)
    grams_per_purchase_unit_snapshot = db.Column(db.Numeric(16, 3), nullable=False)
    recommended_order_qty = db.Column(db.Numeric(16, 4), nullable=False, default=0)
    actual_order_qty = db.Column(db.Numeric(16, 4), nullable=False, default=0)
    # 可選包裝換算，例如 24 kg = 2 箱；兩邊數量都保留人工輸入。
    package_qty = db.Column(db.Numeric(16, 4), nullable=True)
    package_unit = db.Column(db.String(20), nullable=True)
    # 記住當時採用的廠商品項換算，避免日後廠商資料修改影響歷史採購單。
    package_conversion_snapshot = db.Column(db.String(120), nullable=True)
    unit_price_snapshot = db.Column(db.Numeric(16, 4), nullable=False, default=0)
    amount = db.Column(db.Numeric(18, 4), nullable=False, default=0)
    note = db.Column(db.String(255))
    # 簡化採購工作表上的逐項交貨資訊。
    delivery_date = db.Column(db.Date, nullable=True)
    delivery_slot = db.Column(db.String(10), nullable=True)
    # 防呆勾選：表示這個品項已實際完成叫貨。
    ordered = db.Column(db.Boolean, nullable=False, default=False)
    # 使用者手動改過數量/單價/備註後，重新產生需求時保留人工值。
    manual_override = db.Column(db.Boolean, nullable=False, default=False)

    order = db.relationship("KitchenPurchaseOrder", back_populates="items")
    ingredient = db.relationship("KitchenIngredient")
    supplier = db.relationship("KitchenSupplier")
    supplier_item = db.relationship("KitchenSupplierItem")

    __table_args__ = (
        db.UniqueConstraint("order_id", "ingredient_id", name="uq_kitchen_purchase_order_item"),
    )
