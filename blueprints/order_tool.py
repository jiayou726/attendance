"""團膳菜單 / 配方 / 採購叫貨正式模組。

核心流程：
Recipe BOM（每人 AP 克數）→ 中央菜單 → 學校人數
→ 食材需求彙總 → 供應商採購草稿 → 人工調整 → Confirm snapshot。
"""

from __future__ import annotations

import secrets
from collections import defaultdict
from datetime import date, timedelta
from decimal import Decimal, InvalidOperation, ROUND_CEILING

from flask import (
    Blueprint,
    abort,
    current_app,
    flash,
    redirect,
    render_template,
    request,
    session,
    url_for,
)
from sqlalchemy.exc import IntegrityError

from extensions import db
from models import (
    KitchenIngredient,
    KitchenMenuAssignment,
    KitchenMenuPlan,
    KitchenMenuPlanItem,
    KitchenPurchaseOrder,
    KitchenPurchaseOrderItem,
    KitchenRecipe,
    KitchenRecipeIngredient,
    KitchenSchool,
    KitchenSupplier,
)


order_bp = Blueprint("order_tool", __name__)

CATEGORIES = ("主食", "主菜", "副菜", "青菜", "湯品", "點心", "其他")
MEAL_TYPES = ("早餐", "午餐", "晚餐", "點心")
PURCHASE_UNITS = ("kg", "箱", "包", "斤", "瓶", "個", "袋", "桶")
WEEKDAY_LABELS = ("週一", "週二", "週三", "週四", "週五", "週六", "週日")


# ─────────────────────────────────────────────
# Security / shared helpers
# ─────────────────────────────────────────────


def _csrf_token() -> str:
    token = session.get("_kitchen_csrf")
    if not token:
        token = secrets.token_urlsafe(32)
        session["_kitchen_csrf"] = token
    return token


@order_bp.before_request
def _protect_kitchen():
    if not session.get("role"):
        return redirect(url_for("auth.login", next=request.full_path.rstrip("?")))

    if request.method == "POST" and current_app.config.get("KITCHEN_CSRF_ENABLED", True):
        expected = session.get("_kitchen_csrf", "")
        received = request.form.get("_csrf_token", "")
        if not expected or not received or not secrets.compare_digest(expected, received):
            abort(400, description="安全驗證已失效，請重新整理頁面再操作。")
    return None


@order_bp.errorhandler(400)
def _bad_request(error):
    flash(getattr(error, "description", None) or "輸入資料不正確。", "error")
    return redirect(request.referrer or url_for("order_tool.index"))


@order_bp.context_processor
def _template_helpers():
    return {
        "csrf_token": _csrf_token,
        "status_label": _status_label,
        "order_total": _order_total,
        "recipe_total_g": _recipe_total_g,
        "component_cost": _component_cost,
        "trim_decimal": _trim_decimal,
    }


def _status_label(status: str) -> str:
    return {
        "draft": "草稿",
        "confirmed": "已確認",
        "cancelled": "已取消",
    }.get(status, status or "-")


def _decimal(raw, *, default=None) -> Decimal | None:
    if raw is None or str(raw).strip() == "":
        return default
    try:
        value = Decimal(str(raw).strip())
        if not value.is_finite():
            return default
        return value
    except (InvalidOperation, ValueError):
        return default


def _int(raw, *, default=None) -> int | None:
    try:
        return int(str(raw).strip())
    except (TypeError, ValueError):
        return default


def _date(raw, *, default=None) -> date | None:
    if not raw:
        return default
    try:
        return date.fromisoformat(str(raw))
    except ValueError:
        return default


def _recipe_total_g(recipe: KitchenRecipe) -> Decimal:
    return sum((x.grams_per_person or Decimal("0") for x in recipe.ingredients), Decimal("0"))


def _component_cost(component: KitchenRecipeIngredient) -> Decimal:
    ing = component.ingredient
    grams_per_unit = ing.grams_per_purchase_unit or Decimal("0")
    if grams_per_unit <= 0:
        return Decimal("0")
    return (component.grams_per_person or Decimal("0")) / grams_per_unit * (ing.unit_price or Decimal("0"))


def _recipe_cost(recipe: KitchenRecipe) -> Decimal:
    return sum((_component_cost(x) for x in recipe.ingredients), Decimal("0"))


def _order_total(order: KitchenPurchaseOrder) -> Decimal:
    return sum((x.amount or Decimal("0") for x in order.items), Decimal("0"))


def _trim_decimal(value) -> str:
    d = _decimal(value, default=Decimal("0")) or Decimal("0")
    text = format(d, "f")
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return text or "0"


def _round_up_increment(value: Decimal, increment: Decimal) -> Decimal:
    if value <= 0:
        return Decimal("0")
    if increment <= 0:
        increment = Decimal("1")
    steps = (value / increment).to_integral_value(rounding=ROUND_CEILING)
    return steps * increment


def _commit(success: str, redirect_to: str, **values):
    try:
        db.session.commit()
        flash(success, "success")
    except IntegrityError:
        db.session.rollback()
        flash("資料重複或違反資料關聯，請檢查名稱與設定。", "error")
    return redirect(url_for(redirect_to, **values))


def _active_confirmed_orders(service_date: date) -> bool:
    return (
        KitchenPurchaseOrder.query.filter_by(service_date=service_date, status="confirmed").first()
        is not None
    )


def _require_draft_plan(plan: KitchenMenuPlan) -> bool:
    if plan.status != "draft":
        flash("這張菜單已確認，請先重開草稿才能修改。", "warning")
        return False
    return True


# ─────────────────────────────────────────────
# Dashboard
# ─────────────────────────────────────────────


@order_bp.get("/")
def index():
    today = date.today()
    week_end = today + timedelta(days=6)
    plans = (
        KitchenMenuPlan.query.filter(KitchenMenuPlan.service_date.between(today, week_end))
        .order_by(KitchenMenuPlan.service_date, KitchenMenuPlan.meal_type)
        .all()
    )
    recent_orders = (
        KitchenPurchaseOrder.query.order_by(KitchenPurchaseOrder.service_date.desc(), KitchenPurchaseOrder.id.desc())
        .limit(10)
        .all()
    )
    return render_template(
        "kitchen/dashboard.html",
        plans=plans,
        recent_orders=recent_orders,
        recipe_count=KitchenRecipe.query.filter_by(active=True).count(),
        ingredient_count=KitchenIngredient.query.filter_by(active=True).count(),
        school_count=KitchenSchool.query.filter_by(active=True).count(),
    )


# ─────────────────────────────────────────────
# School CRUD
# ─────────────────────────────────────────────


@order_bp.route("/schools", methods=["GET", "POST"])
def schools():
    if request.method == "POST":
        name = request.form.get("name", "").strip()
        if not name:
            flash("學校名稱不可空白。", "error")
            return redirect(url_for("order_tool.schools"))
        db.session.add(KitchenSchool(name=name, code=request.form.get("code", "").strip() or None))
        return _commit("學校已新增。", "order_tool.schools")

    rows = KitchenSchool.query.order_by(KitchenSchool.active.desc(), KitchenSchool.name).all()
    edit_row = db.session.get(KitchenSchool, _int(request.args.get("edit"), default=0)) if request.args.get("edit") else None
    return render_template("kitchen/schools.html", rows=rows, edit_row=edit_row)


@order_bp.post("/schools/<int:school_id>/update")
def school_update(school_id: int):
    row = db.session.get(KitchenSchool, school_id)
    if not row:
        abort(404)
    name = request.form.get("name", "").strip()
    if not name:
        flash("學校名稱不可空白。", "error")
        return redirect(url_for("order_tool.schools", edit=school_id))
    row.name = name
    row.code = request.form.get("code", "").strip() or None
    return _commit("學校資料已更新。", "order_tool.schools")


@order_bp.post("/schools/<int:school_id>/toggle")
def school_toggle(school_id: int):
    row = db.session.get(KitchenSchool, school_id)
    if not row:
        abort(404)
    row.active = not row.active
    return _commit("學校狀態已更新。", "order_tool.schools")


# ─────────────────────────────────────────────
# Supplier CRUD
# ─────────────────────────────────────────────


@order_bp.route("/suppliers", methods=["GET", "POST"])
def suppliers():
    if request.method == "POST":
        name = request.form.get("name", "").strip()
        if not name:
            flash("廠商名稱不可空白。", "error")
            return redirect(url_for("order_tool.suppliers"))
        db.session.add(KitchenSupplier(
            name=name,
            phone=request.form.get("phone", "").strip() or None,
            note=request.form.get("note", "").strip() or None,
        ))
        return _commit("廠商已新增。", "order_tool.suppliers")

    rows = KitchenSupplier.query.order_by(KitchenSupplier.active.desc(), KitchenSupplier.name).all()
    edit_row = db.session.get(KitchenSupplier, _int(request.args.get("edit"), default=0)) if request.args.get("edit") else None
    return render_template("kitchen/suppliers.html", rows=rows, edit_row=edit_row)


@order_bp.post("/suppliers/<int:supplier_id>/update")
def supplier_update(supplier_id: int):
    row = db.session.get(KitchenSupplier, supplier_id)
    if not row:
        abort(404)
    name = request.form.get("name", "").strip()
    if not name:
        flash("廠商名稱不可空白。", "error")
        return redirect(url_for("order_tool.suppliers", edit=supplier_id))
    row.name = name
    row.phone = request.form.get("phone", "").strip() or None
    row.note = request.form.get("note", "").strip() or None
    return _commit("廠商資料已更新。", "order_tool.suppliers")


@order_bp.post("/suppliers/<int:supplier_id>/toggle")
def supplier_toggle(supplier_id: int):
    row = db.session.get(KitchenSupplier, supplier_id)
    if not row:
        abort(404)
    row.active = not row.active
    return _commit("廠商狀態已更新。", "order_tool.suppliers")


# ─────────────────────────────────────────────
# Ingredient CRUD
# ─────────────────────────────────────────────


def _ingredient_form_values():
    name = request.form.get("name", "").strip()
    supplier_id = _int(request.form.get("supplier_id"), default=None)
    purchase_unit = request.form.get("purchase_unit", "kg").strip()
    grams_per_unit = _decimal(request.form.get("grams_per_purchase_unit"))
    unit_price = _decimal(request.form.get("unit_price"))
    increment = _decimal(request.form.get("order_increment"))
    note = request.form.get("note", "").strip() or None

    if not name:
        return None, "食材名稱不可空白。"
    if purchase_unit not in PURCHASE_UNITS:
        return None, "採購單位不正確。"
    if grams_per_unit is None or grams_per_unit <= 0:
        return None, "每採購單位克數必須大於 0。"
    if unit_price is None or unit_price < 0:
        return None, "單價不可為負數。"
    if increment is None or increment <= 0:
        return None, "最小叫貨增量必須大於 0。"
    if supplier_id is not None and not db.session.get(KitchenSupplier, supplier_id):
        return None, "找不到指定廠商。"

    return {
        "name": name,
        "supplier_id": supplier_id,
        "purchase_unit": purchase_unit,
        "grams_per_purchase_unit": grams_per_unit,
        "unit_price": unit_price,
        "order_increment": increment,
        "note": note,
    }, None


@order_bp.route("/ingredients", methods=["GET", "POST"])
def ingredients():
    if request.method == "POST":
        values, error = _ingredient_form_values()
        if error:
            flash(error, "error")
            return redirect(url_for("order_tool.ingredients"))
        db.session.add(KitchenIngredient(**values))
        return _commit("食材已新增。", "order_tool.ingredients")

    rows = KitchenIngredient.query.order_by(KitchenIngredient.active.desc(), KitchenIngredient.name).all()
    suppliers_all = KitchenSupplier.query.order_by(KitchenSupplier.active.desc(), KitchenSupplier.name).all()
    edit_row = db.session.get(KitchenIngredient, _int(request.args.get("edit"), default=0)) if request.args.get("edit") else None
    return render_template(
        "kitchen/ingredients.html",
        rows=rows,
        suppliers=suppliers_all,
        edit_row=edit_row,
        units=PURCHASE_UNITS,
    )


@order_bp.post("/ingredients/<int:ingredient_id>/update")
def ingredient_update(ingredient_id: int):
    row = db.session.get(KitchenIngredient, ingredient_id)
    if not row:
        abort(404)
    values, error = _ingredient_form_values()
    if error:
        flash(error, "error")
        return redirect(url_for("order_tool.ingredients", edit=ingredient_id))
    for key, value in values.items():
        setattr(row, key, value)
    return _commit("食材資料已更新。", "order_tool.ingredients")


@order_bp.post("/ingredients/<int:ingredient_id>/toggle")
def ingredient_toggle(ingredient_id: int):
    row = db.session.get(KitchenIngredient, ingredient_id)
    if not row:
        abort(404)
    row.active = not row.active
    return _commit("食材狀態已更新。", "order_tool.ingredients")


# ─────────────────────────────────────────────
# Recipe CRUD / BOM
# ─────────────────────────────────────────────


def _recipe_form_values():
    name = request.form.get("name", "").strip()
    category = request.form.get("category", "主菜").strip()
    output_raw = request.form.get("serving_output_g", "").strip()
    serving_output = _decimal(output_raw, default=None) if output_raw else None
    if not name:
        return None, "菜色名稱不可空白。"
    if category not in CATEGORIES:
        return None, "菜色分類不正確。"
    if serving_output is not None and serving_output < 0:
        return None, "打菜量不可為負數。"
    return {
        "name": name,
        "category": category,
        "serving_output_g": serving_output,
        "note": request.form.get("note", "").strip() or None,
    }, None


@order_bp.route("/recipes", methods=["GET", "POST"])
def recipes():
    if request.method == "POST":
        values, error = _recipe_form_values()
        if error:
            flash(error, "error")
            return redirect(url_for("order_tool.recipes"))
        recipe = KitchenRecipe(**values)
        db.session.add(recipe)
        try:
            db.session.commit()
            flash("菜色已建立，請設定每人配方。", "success")
            return redirect(url_for("order_tool.recipe_detail", recipe_id=recipe.id))
        except IntegrityError:
            db.session.rollback()
            flash("菜色名稱已存在。", "error")
            return redirect(url_for("order_tool.recipes"))

    rows = KitchenRecipe.query.order_by(KitchenRecipe.active.desc(), KitchenRecipe.category, KitchenRecipe.name).all()
    edit_row = db.session.get(KitchenRecipe, _int(request.args.get("edit"), default=0)) if request.args.get("edit") else None
    return render_template("kitchen/recipes.html", rows=rows, edit_row=edit_row, categories=CATEGORIES)


@order_bp.post("/recipes/<int:recipe_id>/update")
def recipe_update(recipe_id: int):
    row = db.session.get(KitchenRecipe, recipe_id)
    if not row:
        abort(404)
    values, error = _recipe_form_values()
    if error:
        flash(error, "error")
        return redirect(url_for("order_tool.recipes", edit=recipe_id))
    for key, value in values.items():
        setattr(row, key, value)
    return _commit("菜色資料已更新。", "order_tool.recipes")


@order_bp.post("/recipes/<int:recipe_id>/toggle")
def recipe_toggle(recipe_id: int):
    row = db.session.get(KitchenRecipe, recipe_id)
    if not row:
        abort(404)
    row.active = not row.active
    return _commit("菜色狀態已更新。", "order_tool.recipes")


@order_bp.get("/recipes/<int:recipe_id>")
def recipe_detail(recipe_id: int):
    recipe = db.session.get(KitchenRecipe, recipe_id)
    if not recipe:
        abort(404)
    ingredients_all = KitchenIngredient.query.filter_by(active=True).order_by(KitchenIngredient.name).all()
    return render_template(
        "kitchen/recipe_detail.html",
        recipe=recipe,
        ingredients=ingredients_all,
        total_g=_recipe_total_g(recipe),
        total_cost=_recipe_cost(recipe),
    )


@order_bp.post("/recipes/<int:recipe_id>/copy")
def recipe_copy(recipe_id: int):
    source = db.session.get(KitchenRecipe, recipe_id)
    if not source:
        abort(404)
    base = f"{source.name} 複製"
    name = base
    index = 2
    while KitchenRecipe.query.filter_by(name=name).first():
        name = f"{base} {index}"
        index += 1
    copy = KitchenRecipe(
        name=name,
        category=source.category,
        serving_output_g=source.serving_output_g,
        note=source.note,
    )
    db.session.add(copy)
    db.session.flush()
    for x in source.ingredients:
        db.session.add(KitchenRecipeIngredient(
            recipe_id=copy.id,
            ingredient_id=x.ingredient_id,
            grams_per_person=x.grams_per_person,
        ))
    db.session.commit()
    flash("已複製菜色，可直接修改不同食材或克數。", "success")
    return redirect(url_for("order_tool.recipe_detail", recipe_id=copy.id))


@order_bp.post("/recipes/<int:recipe_id>/ingredients")
def recipe_ingredient_add(recipe_id: int):
    recipe = db.session.get(KitchenRecipe, recipe_id)
    if not recipe:
        abort(404)
    ingredient_id = _int(request.form.get("ingredient_id"), default=0) or 0
    grams = _decimal(request.form.get("grams_per_person"))
    ingredient = db.session.get(KitchenIngredient, ingredient_id)
    if not ingredient or not ingredient.active or grams is None or grams <= 0:
        flash("食材或每人克數不正確。", "error")
        return redirect(url_for("order_tool.recipe_detail", recipe_id=recipe_id))

    existing = KitchenRecipeIngredient.query.filter_by(recipe_id=recipe_id, ingredient_id=ingredient_id).first()
    if existing:
        existing.grams_per_person = grams
        message = "配方克數已更新。"
    else:
        db.session.add(KitchenRecipeIngredient(recipe_id=recipe_id, ingredient_id=ingredient_id, grams_per_person=grams))
        message = "食材已加入配方。"
    db.session.commit()
    flash(message, "success")
    return redirect(url_for("order_tool.recipe_detail", recipe_id=recipe_id))


@order_bp.post("/recipe-ingredients/<int:row_id>/update")
def recipe_ingredient_update(row_id: int):
    row = db.session.get(KitchenRecipeIngredient, row_id)
    if not row:
        abort(404)
    grams = _decimal(request.form.get("grams_per_person"))
    if grams is None or grams <= 0:
        flash("每人克數必須大於 0。", "error")
        return redirect(url_for("order_tool.recipe_detail", recipe_id=row.recipe_id))
    row.grams_per_person = grams
    db.session.commit()
    flash("每人克數已更新。", "success")
    return redirect(url_for("order_tool.recipe_detail", recipe_id=row.recipe_id))


@order_bp.post("/recipe-ingredients/<int:row_id>/delete")
def recipe_ingredient_delete(row_id: int):
    row = db.session.get(KitchenRecipeIngredient, row_id)
    if not row:
        abort(404)
    recipe_id = row.recipe_id
    db.session.delete(row)
    db.session.commit()
    flash("食材已從配方移除。", "success")
    return redirect(url_for("order_tool.recipe_detail", recipe_id=recipe_id))


# ─────────────────────────────────────────────
# Menu plan CRUD
# ─────────────────────────────────────────────


@order_bp.route("/plans", methods=["GET", "POST"])
def plans():
    if request.method == "POST":
        service_date = _date(request.form.get("service_date"))
        meal_type = request.form.get("meal_type", "午餐").strip()
        name = request.form.get("name", "").strip()
        if not service_date or meal_type not in MEAL_TYPES or not name:
            flash("日期、餐別或菜單名稱不正確。", "error")
            return redirect(url_for("order_tool.plans"))
        plan = KitchenMenuPlan(
            service_date=service_date,
            meal_type=meal_type,
            name=name,
            note=request.form.get("note", "").strip() or None,
        )
        db.session.add(plan)
        try:
            db.session.commit()
            flash("菜單已建立。", "success")
            return redirect(url_for("order_tool.plan_detail", plan_id=plan.id))
        except IntegrityError:
            db.session.rollback()
            flash("同日、同餐別、同名稱的菜單已存在。", "error")
            return redirect(url_for("order_tool.plans"))

    start = _date(request.args.get("start"), default=date.today()) or date.today()
    end = start + timedelta(days=13)
    rows = (
        KitchenMenuPlan.query.filter(KitchenMenuPlan.service_date.between(start, end))
        .order_by(KitchenMenuPlan.service_date, KitchenMenuPlan.meal_type, KitchenMenuPlan.name)
        .all()
    )
    days = []
    for offset in range(7):
        day = start + timedelta(days=offset)
        days.append({
            "date": day,
            "weekday": WEEKDAY_LABELS[day.weekday()],
            "plans": [p for p in rows if p.service_date == day],
        })
    return render_template(
        "kitchen/plans.html",
        rows=rows,
        days=days,
        start=start,
        today=date.today().isoformat(),
        meal_types=MEAL_TYPES,
    )


@order_bp.get("/plans/<int:plan_id>")
def plan_detail(plan_id: int):
    plan = db.session.get(KitchenMenuPlan, plan_id)
    if not plan:
        abort(404)
    recipes_all = KitchenRecipe.query.filter_by(active=True).order_by(KitchenRecipe.category, KitchenRecipe.name).all()
    schools_all = KitchenSchool.query.filter_by(active=True).order_by(KitchenSchool.name).all()
    return render_template(
        "kitchen/plan_detail.html",
        plan=plan,
        recipes=recipes_all,
        schools=schools_all,
        total_people=sum(max(x.headcount, 0) for x in plan.assignments),
        has_confirmed_orders=_active_confirmed_orders(plan.service_date),
        meal_types=MEAL_TYPES,
    )


@order_bp.post("/plans/<int:plan_id>/update")
def plan_update(plan_id: int):
    plan = db.session.get(KitchenMenuPlan, plan_id)
    if not plan:
        abort(404)
    if not _require_draft_plan(plan):
        return redirect(url_for("order_tool.plan_detail", plan_id=plan_id))
    service_date = _date(request.form.get("service_date"))
    meal_type = request.form.get("meal_type", "").strip()
    name = request.form.get("name", "").strip()
    if not service_date or meal_type not in MEAL_TYPES or not name:
        flash("日期、餐別或名稱不正確。", "error")
        return redirect(url_for("order_tool.plan_detail", plan_id=plan_id))
    plan.service_date = service_date
    plan.meal_type = meal_type
    plan.name = name
    plan.note = request.form.get("note", "").strip() or None
    return _commit("菜單基本資料已更新。", "order_tool.plan_detail", plan_id=plan_id)


@order_bp.post("/plans/<int:plan_id>/items")
def plan_item_add(plan_id: int):
    plan = db.session.get(KitchenMenuPlan, plan_id)
    if not plan:
        abort(404)
    if not _require_draft_plan(plan):
        return redirect(url_for("order_tool.plan_detail", plan_id=plan_id))
    recipe_id = _int(request.form.get("recipe_id"), default=0) or 0
    recipe = db.session.get(KitchenRecipe, recipe_id)
    if not recipe or not recipe.active:
        flash("找不到可用菜色。", "error")
        return redirect(url_for("order_tool.plan_detail", plan_id=plan_id))
    if not KitchenMenuPlanItem.query.filter_by(plan_id=plan_id, recipe_id=recipe_id).first():
        max_order = max((x.sort_order for x in plan.items), default=-1)
        db.session.add(KitchenMenuPlanItem(plan_id=plan_id, recipe_id=recipe_id, sort_order=max_order + 1))
        db.session.commit()
        flash("菜色已加入。", "success")
    else:
        flash("這道菜已在菜單中。", "warning")
    return redirect(url_for("order_tool.plan_detail", plan_id=plan_id))


@order_bp.post("/plan-items/<int:row_id>/delete")
def plan_item_delete(row_id: int):
    row = db.session.get(KitchenMenuPlanItem, row_id)
    if not row:
        abort(404)
    plan_id = row.plan_id
    if not _require_draft_plan(row.plan):
        return redirect(url_for("order_tool.plan_detail", plan_id=plan_id))
    db.session.delete(row)
    db.session.commit()
    flash("菜色已移除。", "success")
    return redirect(url_for("order_tool.plan_detail", plan_id=plan_id))


@order_bp.post("/plans/<int:plan_id>/assignments")
def assignment_add(plan_id: int):
    plan = db.session.get(KitchenMenuPlan, plan_id)
    if not plan:
        abort(404)
    if not _require_draft_plan(plan):
        return redirect(url_for("order_tool.plan_detail", plan_id=plan_id))
    school_id = _int(request.form.get("school_id"), default=0) or 0
    headcount = _int(request.form.get("headcount"), default=None)
    school = db.session.get(KitchenSchool, school_id)
    if not school or not school.active or headcount is None or headcount < 0:
        flash("學校或人數不正確。", "error")
        return redirect(url_for("order_tool.plan_detail", plan_id=plan_id))
    row = KitchenMenuAssignment.query.filter_by(plan_id=plan_id, school_id=school_id).first()
    if row:
        row.headcount = headcount
        message = "學校人數已更新。"
    else:
        db.session.add(KitchenMenuAssignment(plan_id=plan_id, school_id=school_id, headcount=headcount))
        message = "學校已加入菜單。"
    db.session.commit()
    flash(message, "success")
    return redirect(url_for("order_tool.plan_detail", plan_id=plan_id))


@order_bp.post("/assignments/<int:row_id>/update")
def assignment_update(row_id: int):
    row = db.session.get(KitchenMenuAssignment, row_id)
    if not row:
        abort(404)
    if not _require_draft_plan(row.plan):
        return redirect(url_for("order_tool.plan_detail", plan_id=row.plan_id))
    headcount = _int(request.form.get("headcount"), default=None)
    if headcount is None or headcount < 0:
        flash("人數不可為負數。", "error")
        return redirect(url_for("order_tool.plan_detail", plan_id=row.plan_id))
    row.headcount = headcount
    db.session.commit()
    flash("人數已更新。", "success")
    return redirect(url_for("order_tool.plan_detail", plan_id=row.plan_id))


@order_bp.post("/assignments/<int:row_id>/delete")
def assignment_delete(row_id: int):
    row = db.session.get(KitchenMenuAssignment, row_id)
    if not row:
        abort(404)
    plan_id = row.plan_id
    if not _require_draft_plan(row.plan):
        return redirect(url_for("order_tool.plan_detail", plan_id=plan_id))
    db.session.delete(row)
    db.session.commit()
    flash("學校已從菜單移除。", "success")
    return redirect(url_for("order_tool.plan_detail", plan_id=plan_id))


@order_bp.post("/plans/<int:plan_id>/copy")
def plan_copy(plan_id: int):
    source = db.session.get(KitchenMenuPlan, plan_id)
    if not source:
        abort(404)
    target_date = _date(request.form.get("target_date"))
    if not target_date:
        flash("請選擇正確的複製日期。", "error")
        return redirect(url_for("order_tool.plan_detail", plan_id=plan_id))
    if KitchenMenuPlan.query.filter_by(service_date=target_date, meal_type=source.meal_type, name=source.name).first():
        flash("目標日期已有相同餐別與名稱的菜單。", "error")
        return redirect(url_for("order_tool.plan_detail", plan_id=plan_id))

    copy = KitchenMenuPlan(
        service_date=target_date,
        meal_type=source.meal_type,
        name=source.name,
        note=source.note,
        status="draft",
    )
    db.session.add(copy)
    db.session.flush()
    for x in source.items:
        db.session.add(KitchenMenuPlanItem(plan_id=copy.id, recipe_id=x.recipe_id, sort_order=x.sort_order))
    for x in source.assignments:
        db.session.add(KitchenMenuAssignment(plan_id=copy.id, school_id=x.school_id, headcount=x.headcount))
    db.session.commit()
    flash("菜單已複製，請確認新日期的人數。", "success")
    return redirect(url_for("order_tool.plan_detail", plan_id=copy.id))


@order_bp.post("/plans/<int:plan_id>/confirm")
def plan_confirm(plan_id: int):
    plan = db.session.get(KitchenMenuPlan, plan_id)
    if not plan:
        abort(404)
    if not plan.items or sum(x.headcount for x in plan.assignments) <= 0:
        flash("至少要有菜色與大於 0 的供餐人數才能確認。", "error")
        return redirect(url_for("order_tool.plan_detail", plan_id=plan_id))
    plan.status = "confirmed"
    db.session.commit()
    flash("菜單已確認。", "success")
    return redirect(url_for("order_tool.plan_detail", plan_id=plan_id))


@order_bp.post("/plans/<int:plan_id>/reopen")
def plan_reopen(plan_id: int):
    plan = db.session.get(KitchenMenuPlan, plan_id)
    if not plan:
        abort(404)
    if _active_confirmed_orders(plan.service_date):
        flash("這一天已有已確認採購單，請先到採購頁重開相關採購單。", "error")
        return redirect(url_for("order_tool.plan_detail", plan_id=plan_id))
    plan.status = "draft"
    db.session.commit()
    flash("菜單已重開為草稿。", "success")
    return redirect(url_for("order_tool.plan_detail", plan_id=plan_id))


@order_bp.post("/plans/<int:plan_id>/delete")
def plan_delete(plan_id: int):
    plan = db.session.get(KitchenMenuPlan, plan_id)
    if not plan:
        abort(404)
    if plan.status != "draft" or _active_confirmed_orders(plan.service_date):
        flash("只能刪除沒有正式採購紀錄的草稿菜單。", "error")
        return redirect(url_for("order_tool.plan_detail", plan_id=plan_id))
    db.session.delete(plan)
    db.session.commit()
    flash("草稿菜單已刪除。", "success")
    return redirect(url_for("order_tool.plans"))


# ─────────────────────────────────────────────
# Purchase calculation / persisted orders
# ─────────────────────────────────────────────


def _supplier_identity(ingredient: KitchenIngredient):
    if ingredient.supplier_id and ingredient.supplier:
        return f"supplier:{ingredient.supplier_id}", ingredient.supplier_id, ingredient.supplier.name
    return "unassigned", None, "⚠ 未指定供應商"


def _requirements_for_date(service_date: date):
    plans_on_day = KitchenMenuPlan.query.filter_by(service_date=service_date).all()
    # supplier_key -> ingredient_id -> aggregate row
    grouped: dict[str, dict[int, dict]] = defaultdict(dict)

    for plan in plans_on_day:
        people = sum(max(x.headcount, 0) for x in plan.assignments)
        if people <= 0:
            continue
        for menu_item in plan.items:
            for component in menu_item.recipe.ingredients:
                ing = component.ingredient
                supplier_key, supplier_id, supplier_name = _supplier_identity(ing)
                grams = (component.grams_per_person or Decimal("0")) * people
                current = grouped[supplier_key].get(ing.id)
                if current is None:
                    current = {
                        "supplier_id": supplier_id,
                        "supplier_name": supplier_name,
                        "ingredient": ing,
                        "required_grams": Decimal("0"),
                    }
                    grouped[supplier_key][ing.id] = current
                current["required_grams"] += grams
    return grouped


def _generate_date_orders(service_date: date):
    # 已有任一 confirmed 採購單時，整天都不自動重算，避免產生新供應商單造成重複叫貨。
    if _active_confirmed_orders(service_date):
        return 0, True

    requirements = _requirements_for_date(service_date)
    active_keys = set(requirements.keys())
    existing_orders = KitchenPurchaseOrder.query.filter_by(service_date=service_date).all()
    existing_by_key = {x.supplier_key: x for x in existing_orders}
    count = 0

    for supplier_key, ingredient_rows in requirements.items():
        first = next(iter(ingredient_rows.values()))
        order = existing_by_key.get(supplier_key)
        if order is None:
            order = KitchenPurchaseOrder(
                service_date=service_date,
                supplier_id=first["supplier_id"],
                supplier_key=supplier_key,
                supplier_name_snapshot=first["supplier_name"],
                status="draft",
            )
            db.session.add(order)
            db.session.flush()
        else:
            order.status = "draft"
            order.supplier_id = first["supplier_id"]
            order.supplier_name_snapshot = first["supplier_name"]
            KitchenPurchaseOrderItem.query.filter_by(order_id=order.id).delete(synchronize_session=False)
            db.session.flush()

        for data in ingredient_rows.values():
            ing = data["ingredient"]
            grams_per_unit = ing.grams_per_purchase_unit or Decimal("0")
            increment = ing.order_increment or Decimal("0")
            if grams_per_unit <= 0 or increment <= 0:
                continue
            required_grams = data["required_grams"]
            required_qty = required_grams / grams_per_unit
            recommended = _round_up_increment(required_qty, increment)
            unit_price = ing.unit_price or Decimal("0")
            item = KitchenPurchaseOrderItem(
                order_id=order.id,
                ingredient_id=ing.id,
                ingredient_name_snapshot=ing.name,
                required_grams=required_grams,
                required_qty=required_qty,
                purchase_unit_snapshot=ing.purchase_unit,
                grams_per_purchase_unit_snapshot=grams_per_unit,
                recommended_order_qty=recommended,
                actual_order_qty=recommended,
                unit_price_snapshot=unit_price,
                amount=recommended * unit_price,
            )
            db.session.add(item)
        count += 1

    # 需求已不存在的舊草稿直接移除；confirmed 在前面已整天阻擋。
    for order in existing_orders:
        if order.status == "draft" and order.supplier_key not in active_keys:
            db.session.delete(order)

    db.session.commit()
    return count, False


@order_bp.route("/purchases", methods=["GET"])
def purchases():
    start = _date(request.args.get("start"), default=date.today() - timedelta(days=7)) or date.today()
    end = _date(request.args.get("end"), default=date.today() + timedelta(days=14)) or start
    if end < start:
        start, end = end, start
    orders = (
        KitchenPurchaseOrder.query.filter(KitchenPurchaseOrder.service_date.between(start, end))
        .order_by(KitchenPurchaseOrder.service_date.desc(), KitchenPurchaseOrder.supplier_name_snapshot)
        .all()
    )
    return render_template("kitchen/purchases.html", orders=orders, start=start, end=end)


# 舊網址相容：/purchase 導向新的持久化採購頁。
@order_bp.get("/purchase")
def purchase_legacy_redirect():
    args = {}
    if request.args.get("start"):
        args["start"] = request.args["start"]
    if request.args.get("end"):
        args["end"] = request.args["end"]
    return redirect(url_for("order_tool.purchases", **args))


@order_bp.post("/purchases/generate")
def generate_purchases():
    start = _date(request.form.get("start"), default=date.today()) or date.today()
    end = _date(request.form.get("end"), default=start) or start
    if end < start:
        start, end = end, start
    if (end - start).days > 31:
        flash("一次最多產生 32 天的採購草稿。", "error")
        return redirect(url_for("order_tool.purchases", start=start, end=end))

    created = 0
    blocked_dates = []
    day = start
    while day <= end:
        count, blocked = _generate_date_orders(day)
        created += count
        if blocked:
            blocked_dates.append(str(day))
        day += timedelta(days=1)

    if created:
        flash(f"已建立 / 更新 {created} 張採購草稿。", "success")
    if blocked_dates:
        flash("以下日期已有已確認採購單，因此沒有自動重算：" + "、".join(blocked_dates), "warning")
    if not created and not blocked_dates:
        flash("此期間沒有可計算的菜單需求。請確認菜色配方與學校人數。", "warning")
    return redirect(url_for("order_tool.purchases", start=start, end=end))


@order_bp.get("/purchases/<int:order_id>")
def purchase_detail(order_id: int):
    order = db.session.get(KitchenPurchaseOrder, order_id)
    if not order:
        abort(404)
    suppliers_all = KitchenSupplier.query.order_by(KitchenSupplier.active.desc(), KitchenSupplier.name).all()
    do_print = request.args.get("print") == "1"
    return render_template(
        "kitchen/purchase_detail.html",
        order=order,
        suppliers=suppliers_all,
        total=_order_total(order),
        do_print=do_print,
    )


@order_bp.post("/purchases/<int:order_id>/update")
def purchase_update(order_id: int):
    order = db.session.get(KitchenPurchaseOrder, order_id)
    if not order:
        abort(404)
    if order.status != "draft":
        flash("只有草稿採購單可以修改。", "error")
        return redirect(url_for("order_tool.purchase_detail", order_id=order_id))

    supplier_id = _int(request.form.get("supplier_id"), default=None)
    if supplier_id is not None:
        supplier = db.session.get(KitchenSupplier, supplier_id)
        if not supplier:
            flash("找不到指定廠商。", "error")
            return redirect(url_for("order_tool.purchase_detail", order_id=order_id))
        new_key = f"supplier:{supplier.id}"
        new_name = supplier.name
    else:
        supplier = None
        new_key = "unassigned"
        new_name = "⚠ 未指定供應商"

    conflict = KitchenPurchaseOrder.query.filter(
        KitchenPurchaseOrder.service_date == order.service_date,
        KitchenPurchaseOrder.supplier_key == new_key,
        KitchenPurchaseOrder.id != order.id,
    ).first()
    if conflict:
        flash("同一天已存在這個廠商的採購單，請先處理該單，避免重複叫貨。", "error")
        return redirect(url_for("order_tool.purchase_detail", order_id=order_id))

    order.supplier_id = supplier.id if supplier else None
    order.supplier_key = new_key
    order.supplier_name_snapshot = new_name
    order.note = request.form.get("note", "").strip() or None
    return _commit("採購單資料已更新。", "order_tool.purchase_detail", order_id=order_id)


@order_bp.post("/purchase-items/<int:item_id>/update")
def purchase_item_update(item_id: int):
    item = db.session.get(KitchenPurchaseOrderItem, item_id)
    if not item:
        abort(404)
    if item.order.status != "draft":
        flash("已確認的採購單不可直接修改。", "error")
        return redirect(url_for("order_tool.purchase_detail", order_id=item.order_id))
    actual = _decimal(request.form.get("actual_order_qty"))
    price = _decimal(request.form.get("unit_price"))
    if actual is None or actual < 0 or price is None or price < 0:
        flash("實際叫貨量與單價不可為負數。", "error")
        return redirect(url_for("order_tool.purchase_detail", order_id=item.order_id))
    item.actual_order_qty = actual
    item.unit_price_snapshot = price
    item.amount = actual * price
    item.note = request.form.get("note", "").strip() or None
    db.session.commit()
    flash("採購項目已更新。", "success")
    return redirect(url_for("order_tool.purchase_detail", order_id=item.order_id))


@order_bp.post("/purchases/<int:order_id>/confirm")
def purchase_confirm(order_id: int):
    order = db.session.get(KitchenPurchaseOrder, order_id)
    if not order:
        abort(404)
    if order.status != "draft":
        flash("這張採購單目前不是草稿。", "error")
        return redirect(url_for("order_tool.purchase_detail", order_id=order_id))
    if not order.items:
        flash("空白採購單不可確認。", "error")
        return redirect(url_for("order_tool.purchase_detail", order_id=order_id))
    if not order.supplier_id:
        flash("正式確認前必須指定供應商。", "error")
        return redirect(url_for("order_tool.purchase_detail", order_id=order_id))

    order.status = "confirmed"
    # 確認任何正式採購時，同日菜單一起鎖定，防止操作人員誤以為改菜單會改掉已下單內容。
    for plan in KitchenMenuPlan.query.filter_by(service_date=order.service_date).all():
        plan.status = "confirmed"
    db.session.commit()
    flash("採購單已確認並保存歷史 snapshot。", "success")
    return redirect(url_for("order_tool.purchase_detail", order_id=order_id))


@order_bp.post("/purchases/<int:order_id>/reopen")
def purchase_reopen(order_id: int):
    order = db.session.get(KitchenPurchaseOrder, order_id)
    if not order:
        abort(404)
    order.status = "draft"
    db.session.commit()
    flash("採購單已重開草稿。既有 snapshot 數字仍保留，除非你重新按產生草稿。", "warning")
    return redirect(url_for("order_tool.purchase_detail", order_id=order_id))


@order_bp.post("/purchases/<int:order_id>/cancel")
def purchase_cancel(order_id: int):
    order = db.session.get(KitchenPurchaseOrder, order_id)
    if not order:
        abort(404)
    order.status = "cancelled"
    db.session.commit()
    flash("採購單已取消，歷史資料仍保留。", "warning")
    return redirect(url_for("order_tool.purchase_detail", order_id=order_id))
