"""團膳菜單 / 配方 / 採購叫貨 Blueprint。

流程：
中央菜單 Plan → 指派學校與人數 → Recipe BOM（每人 AP 克數）
→ 自動彙總食材 → 依供應商分組 → 列印 / 另存 PDF。
"""

from __future__ import annotations

from collections import defaultdict
from datetime import date, datetime, timedelta
from decimal import Decimal

from flask import Blueprint, abort, redirect, render_template_string, request, url_for
from sqlalchemy.exc import IntegrityError

from extensions import db
from models import (
    KitchenIngredient,
    KitchenMenuAssignment,
    KitchenMenuPlan,
    KitchenMenuPlanItem,
    KitchenRecipe,
    KitchenRecipeIngredient,
    KitchenSchool,
    KitchenSupplier,
)


order_bp = Blueprint("order_tool", __name__)


STYLE = r"""
<style>
:root{
  --bg:#f5f7fb;--card:#fff;--text:#17202a;--muted:#667085;--line:#e5e7eb;
  --brand:#1769aa;--brand2:#0f5b91;--danger:#b42318;--ok:#157347;
}
*{box-sizing:border-box} body{margin:0;background:var(--bg);color:var(--text);
font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","Noto Sans TC","Microsoft JhengHei",sans-serif}
a{color:var(--brand);text-decoration:none}.wrap{max-width:1180px;margin:auto;padding:24px}
.top{display:flex;align-items:center;justify-content:space-between;gap:12px;margin-bottom:18px;flex-wrap:wrap}
h1{font-size:26px;margin:0}h2{font-size:20px;margin:0 0 14px}h3{font-size:17px;margin:0 0 10px}
.nav{display:flex;gap:8px;flex-wrap:wrap}.nav a,.btn{display:inline-block;border:0;border-radius:10px;
padding:9px 14px;background:var(--brand);color:#fff;cursor:pointer;font-size:14px}.nav a:hover,.btn:hover{background:var(--brand2)}
.btn.secondary{background:#eef3f8;color:#24445f}.btn.danger{background:#fff0ee;color:var(--danger)}
.grid{display:grid;grid-template-columns:repeat(12,1fr);gap:16px}.span-4{grid-column:span 4}.span-5{grid-column:span 5}
.span-6{grid-column:span 6}.span-7{grid-column:span 7}.span-8{grid-column:span 8}.span-12{grid-column:span 12}
.card{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:18px;box-shadow:0 2px 8px rgba(16,24,40,.04)}
.stat{font-size:30px;font-weight:750}.muted{color:var(--muted);font-size:13px}.row{display:flex;gap:10px;align-items:end;flex-wrap:wrap}
.field{display:flex;flex-direction:column;gap:5px;min-width:150px;flex:1}label{font-size:13px;color:#475467;font-weight:600}
input,select,textarea{width:100%;padding:9px 10px;border:1px solid #cfd6df;border-radius:9px;background:#fff;font:inherit}
table{width:100%;border-collapse:collapse}th,td{padding:10px 8px;border-bottom:1px solid var(--line);text-align:left;vertical-align:top;font-size:14px}
th{font-size:12px;color:#667085;background:#fafbfc}.right{text-align:right}.pill{display:inline-block;background:#eef6ff;color:#175c91;border-radius:999px;padding:3px 8px;font-size:12px}
.empty{padding:28px;text-align:center;color:#667085}.inline{display:inline}.notice{padding:10px 12px;border-radius:10px;background:#fff8e6;color:#7a4b00;margin-bottom:14px}
@media(max-width:800px){.span-4,.span-5,.span-6,.span-7,.span-8{grid-column:span 12}.wrap{padding:14px}table{display:block;overflow-x:auto;white-space:nowrap}}
@media print{body{background:#fff}.no-print{display:none!important}.wrap{max-width:none;padding:0}.card{box-shadow:none;border:0;padding:0}table{display:table;white-space:normal}th,td{font-size:11px;padding:5px}.purchase-group{break-inside:avoid;margin-bottom:22px}h1{font-size:20px}h2{font-size:16px}}
</style>
"""


NAV = r"""
<div class="top no-print">
  <h1>團膳管理</h1>
  <div class="nav">
    <a href="{{ url_for('order_tool.index') }}">總覽</a>
    <a href="{{ url_for('order_tool.plans') }}">菜單</a>
    <a href="{{ url_for('order_tool.recipes') }}">菜色配方</a>
    <a href="{{ url_for('order_tool.ingredients') }}">食材</a>
    <a href="{{ url_for('order_tool.schools') }}">學校</a>
    <a href="{{ url_for('order_tool.suppliers') }}">廠商</a>
    <a href="{{ url_for('order_tool.purchase') }}">採購叫貨</a>
  </div>
</div>
"""


def _render(title: str, body: str, **ctx):
    template = f"""<!doctype html><html lang="zh-Hant"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>{title}</title>{STYLE}</head>
<body><main class="wrap">{NAV}{body}</main></body></html>"""
    return render_template_string(template, **ctx)


def _to_decimal(raw: str | None, default: str = "0") -> Decimal:
    try:
        return Decimal((raw or default).strip())
    except Exception:
        return Decimal(default)


def _to_int(raw: str | None, default: int = 0) -> int:
    try:
        return int(raw or default)
    except Exception:
        return default


def _parse_date(raw: str | None, fallback: date | None = None) -> date:
    if raw:
        try:
            return datetime.strptime(raw, "%Y-%m-%d").date()
        except ValueError:
            pass
    return fallback or date.today()


def _commit_or_back(endpoint: str, **values):
    try:
        db.session.commit()
    except IntegrityError:
        db.session.rollback()
        abort(409, description="資料重複，請確認名稱或同日設定是否已存在。")
    return redirect(url_for(endpoint, **values))


@order_bp.get("/")
def index():
    today = date.today()
    week_end = today + timedelta(days=6)
    plans = (
        KitchenMenuPlan.query
        .filter(KitchenMenuPlan.service_date.between(today, week_end))
        .order_by(KitchenMenuPlan.service_date, KitchenMenuPlan.meal_type)
        .all()
    )
    body = r"""
<div class="grid">
  <section class="card span-4"><div class="muted">菜色配方</div><div class="stat">{{ recipe_count }}</div></section>
  <section class="card span-4"><div class="muted">食材</div><div class="stat">{{ ingredient_count }}</div></section>
  <section class="card span-4"><div class="muted">供餐學校</div><div class="stat">{{ school_count }}</div></section>
  <section class="card span-12">
    <div class="top"><div><h2>未來 7 天菜單</h2><div class="muted">一份中央菜單可以同時指派多間學校與各自人數。</div></div>
    <a class="btn" href="{{ url_for('order_tool.plans') }}">新增菜單</a></div>
    {% if plans %}<table><thead><tr><th>日期</th><th>餐別</th><th>名稱</th><th>菜色</th><th>學校 / 人數</th><th></th></tr></thead><tbody>
    {% for p in plans %}<tr><td>{{ p.service_date }}</td><td>{{ p.meal_type }}</td><td>{{ p.name or '中央菜單' }}</td>
      <td>{{ p.items|length }} 道</td><td>{{ p.assignments|length }} 校 / {{ p.assignments|sum(attribute='headcount') }} 人</td>
      <td><a href="{{ url_for('order_tool.plan_detail', plan_id=p.id) }}">編輯</a></td></tr>{% endfor %}
    </tbody></table>{% else %}<div class="empty">這 7 天還沒有菜單。</div>{% endif %}
  </section>
</div>"""
    return _render(
        "團膳管理",
        body,
        plans=plans,
        recipe_count=KitchenRecipe.query.count(),
        ingredient_count=KitchenIngredient.query.count(),
        school_count=KitchenSchool.query.filter_by(active=True).count(),
    )


@order_bp.route("/schools", methods=["GET", "POST"])
def schools():
    if request.method == "POST":
        name = request.form.get("name", "").strip()
        if not name:
            abort(400, description="學校名稱不可空白。")
        db.session.add(KitchenSchool(name=name, code=request.form.get("code", "").strip() or None))
        return _commit_or_back("order_tool.schools")

    rows = KitchenSchool.query.order_by(KitchenSchool.name).all()
    body = r"""
<div class="grid"><section class="card span-5"><h2>新增學校 / 客戶</h2>
<form method="post"><div class="field"><label>名稱</label><input name="name" required placeholder="例如：內小"></div><br>
<div class="field"><label>代碼（可空白）</label><input name="code" placeholder="例如：A01"></div><br><button class="btn">新增</button></form></section>
<section class="card span-7"><h2>學校清單</h2>{% if rows %}<table><tr><th>代碼</th><th>名稱</th><th>狀態</th></tr>
{% for r in rows %}<tr><td>{{ r.code or '-' }}</td><td>{{ r.name }}</td><td>{{ '啟用' if r.active else '停用' }}</td></tr>{% endfor %}</table>
{% else %}<div class="empty">尚無資料</div>{% endif %}</section></div>"""
    return _render("學校", body, rows=rows)


@order_bp.route("/suppliers", methods=["GET", "POST"])
def suppliers():
    if request.method == "POST":
        name = request.form.get("name", "").strip()
        if not name:
            abort(400, description="廠商名稱不可空白。")
        db.session.add(KitchenSupplier(
            name=name,
            phone=request.form.get("phone", "").strip() or None,
            note=request.form.get("note", "").strip() or None,
        ))
        return _commit_or_back("order_tool.suppliers")

    rows = KitchenSupplier.query.order_by(KitchenSupplier.name).all()
    body = r"""
<div class="grid"><section class="card span-5"><h2>新增廠商</h2><form method="post">
<div class="field"><label>廠商名稱</label><input name="name" required></div><br><div class="field"><label>電話</label><input name="phone"></div><br>
<div class="field"><label>備註</label><input name="note"></div><br><button class="btn">新增</button></form></section>
<section class="card span-7"><h2>廠商清單</h2>{% if rows %}<table><tr><th>廠商</th><th>電話</th><th>備註</th></tr>
{% for r in rows %}<tr><td>{{ r.name }}</td><td>{{ r.phone or '-' }}</td><td>{{ r.note or '-' }}</td></tr>{% endfor %}</table>{% else %}<div class="empty">尚無資料</div>{% endif %}</section></div>"""
    return _render("廠商", body, rows=rows)


@order_bp.route("/ingredients", methods=["GET", "POST"])
def ingredients():
    suppliers_all = KitchenSupplier.query.order_by(KitchenSupplier.name).all()
    if request.method == "POST":
        name = request.form.get("name", "").strip()
        if not name:
            abort(400, description="食材名稱不可空白。")
        supplier_id = _to_int(request.form.get("supplier_id")) or None
        db.session.add(KitchenIngredient(
            name=name,
            unit_price=_to_decimal(request.form.get("unit_price")),
            supplier_id=supplier_id,
            note=request.form.get("note", "").strip() or None,
        ))
        return _commit_or_back("order_tool.ingredients")

    rows = KitchenIngredient.query.order_by(KitchenIngredient.name).all()
    body = r"""
<div class="grid"><section class="card span-5"><h2>新增食材</h2><div class="muted">配方一律以「每人幾克」儲存；單價預設為每公斤。</div><br>
<form method="post"><div class="field"><label>食材名稱</label><input name="name" required placeholder="骨腿丁"></div><br>
<div class="field"><label>單價 / kg</label><input name="unit_price" type="number" step="0.0001" value="0"></div><br>
<div class="field"><label>預設供應商</label><select name="supplier_id"><option value="">未指定</option>{% for s in suppliers %}<option value="{{ s.id }}">{{ s.name }}</option>{% endfor %}</select></div><br>
<div class="field"><label>備註</label><input name="note"></div><br><button class="btn">新增</button></form></section>
<section class="card span-7"><h2>食材清單</h2>{% if rows %}<table><tr><th>食材</th><th class="right">單價/kg</th><th>供應商</th></tr>
{% for r in rows %}<tr><td>{{ r.name }}</td><td class="right">${{ '%.2f'|format(r.unit_price|float) }}</td><td>{{ r.supplier.name if r.supplier else '⚠ 未指定' }}</td></tr>{% endfor %}</table>{% else %}<div class="empty">尚無資料</div>{% endif %}</section></div>"""
    return _render("食材", body, rows=rows, suppliers=suppliers_all)


@order_bp.route("/recipes", methods=["GET", "POST"])
def recipes():
    if request.method == "POST":
        name = request.form.get("name", "").strip()
        if not name:
            abort(400, description="菜色名稱不可空白。")
        output_raw = request.form.get("serving_output_g", "").strip()
        recipe = KitchenRecipe(
            name=name,
            category=request.form.get("category", "").strip() or None,
            serving_output_g=_to_decimal(output_raw) if output_raw else None,
            note=request.form.get("note", "").strip() or None,
        )
        db.session.add(recipe)
        try:
            db.session.commit()
        except IntegrityError:
            db.session.rollback()
            abort(409, description="菜色名稱已存在。")
        return redirect(url_for("order_tool.recipe_detail", recipe_id=recipe.id))

    rows = KitchenRecipe.query.order_by(KitchenRecipe.category, KitchenRecipe.name).all()
    body = r"""
<div class="grid"><section class="card span-5"><h2>新增菜色</h2><form method="post">
<div class="field"><label>菜色名稱</label><input name="name" required placeholder="南洋綠咖哩雞"></div><br>
<div class="field"><label>分類</label><select name="category"><option>主食</option><option selected>主菜</option><option>副菜</option><option>青菜</option><option>湯品</option><option>點心</option></select></div><br>
<div class="field"><label>預計打菜量 g / 人（可空白）</label><input name="serving_output_g" type="number" step="0.01" placeholder="95"></div><br>
<div class="field"><label>備註</label><input name="note"></div><br><button class="btn">建立後設定配方</button></form></section>
<section class="card span-7"><h2>菜色配方</h2>{% if rows %}<table><tr><th>類別</th><th>菜色</th><th class="right">生料 g/人</th><th class="right">打菜 g/人</th><th></th></tr>
{% for r in rows %}<tr><td><span class="pill">{{ r.category or '未分類' }}</span></td><td>{{ r.name }}</td><td class="right">{{ '%.1f'|format(r.ingredients|sum(attribute='grams_per_person')|float) }}</td>
<td class="right">{{ '%.1f'|format(r.serving_output_g|float) if r.serving_output_g is not none else '-' }}</td><td><a href="{{ url_for('order_tool.recipe_detail', recipe_id=r.id) }}">編輯配方</a></td></tr>{% endfor %}</table>
{% else %}<div class="empty">先建立第一道菜。</div>{% endif %}</section></div>"""
    return _render("菜色配方", body, rows=rows)


@order_bp.route("/recipes/<int:recipe_id>", methods=["GET"])
def recipe_detail(recipe_id: int):
    recipe = db.session.get(KitchenRecipe, recipe_id)
    if not recipe:
        abort(404)
    ingredients_all = KitchenIngredient.query.order_by(KitchenIngredient.name).all()
    total_g = sum((x.grams_per_person or 0) for x in recipe.ingredients)
    total_cost = sum(
        ((x.grams_per_person or 0) / Decimal("1000")) * (x.ingredient.unit_price or 0)
        for x in recipe.ingredients
    )
    body = r"""
<div class="top"><div><h1>{{ recipe.name }}</h1><div class="muted">每人配方（AP 採購量）</div></div><a class="btn secondary" href="{{ url_for('order_tool.recipes') }}">← 回菜色清單</a></div>
<div class="grid"><section class="card span-7"><h2>目前配方</h2>
{% if recipe.ingredients %}<table><tr><th>食材</th><th class="right">每人 AP</th><th class="right">單價/kg</th><th class="right">每人成本</th><th></th></tr>
{% for x in recipe.ingredients %}<tr><td>{{ x.ingredient.name }}</td><td class="right">{{ '%.3f'|format(x.grams_per_person|float) }} g</td>
<td class="right">${{ '%.2f'|format(x.ingredient.unit_price|float) }}</td><td class="right">${{ '%.3f'|format((x.grams_per_person|float/1000)*(x.ingredient.unit_price|float)) }}</td>
<td><form class="inline" method="post" action="{{ url_for('order_tool.recipe_ingredient_delete', row_id=x.id) }}"><button class="btn danger">移除</button></form></td></tr>{% endfor %}</table>
<div class="row" style="margin-top:14px"><div><b>每份生料：{{ '%.1f'|format(total_g|float) }} g</b></div><div><b>每份成本：約 ${{ '%.2f'|format(total_cost|float) }}</b></div>
{% if recipe.serving_output_g is not none %}<div><b>打菜量：{{ recipe.serving_output_g }} g</b></div>{% endif %}</div>
{% else %}<div class="empty">還沒有食材。</div>{% endif %}</section>
<section class="card span-5"><h2>加入食材</h2>{% if ingredients %}<form method="post" action="{{ url_for('order_tool.recipe_ingredient_add', recipe_id=recipe.id) }}">
<div class="field"><label>食材</label><select name="ingredient_id" required>{% for i in ingredients %}<option value="{{ i.id }}">{{ i.name }}</option>{% endfor %}</select></div><br>
<div class="field"><label>每人採購量（g）</label><input name="grams_per_person" type="number" min="0.001" step="0.001" required placeholder="例如 88"></div><br><button class="btn">加入配方</button></form>
{% else %}<div class="notice">請先到「食材」建立食材。</div><a class="btn" href="{{ url_for('order_tool.ingredients') }}">去建立食材</a>{% endif %}</section></div>"""
    return _render(
        recipe.name,
        body,
        recipe=recipe,
        ingredients=ingredients_all,
        total_g=total_g,
        total_cost=total_cost,
    )


@order_bp.post("/recipes/<int:recipe_id>/ingredients")
def recipe_ingredient_add(recipe_id: int):
    if not db.session.get(KitchenRecipe, recipe_id):
        abort(404)
    ingredient_id = _to_int(request.form.get("ingredient_id"))
    grams = _to_decimal(request.form.get("grams_per_person"))
    if ingredient_id <= 0 or grams <= 0:
        abort(400, description="食材與克數不正確。")
    existing = KitchenRecipeIngredient.query.filter_by(recipe_id=recipe_id, ingredient_id=ingredient_id).first()
    if existing:
        existing.grams_per_person = grams
    else:
        db.session.add(KitchenRecipeIngredient(
            recipe_id=recipe_id,
            ingredient_id=ingredient_id,
            grams_per_person=grams,
        ))
    db.session.commit()
    return redirect(url_for("order_tool.recipe_detail", recipe_id=recipe_id))


@order_bp.post("/recipe-ingredients/<int:row_id>/delete")
def recipe_ingredient_delete(row_id: int):
    row = db.session.get(KitchenRecipeIngredient, row_id)
    if not row:
        abort(404)
    recipe_id = row.recipe_id
    db.session.delete(row)
    db.session.commit()
    return redirect(url_for("order_tool.recipe_detail", recipe_id=recipe_id))


@order_bp.route("/plans", methods=["GET", "POST"])
def plans():
    if request.method == "POST":
        service_date = _parse_date(request.form.get("service_date"))
        plan = KitchenMenuPlan(
            service_date=service_date,
            meal_type=request.form.get("meal_type", "午餐").strip() or "午餐",
            name=request.form.get("name", "").strip() or "中央菜單",
            note=request.form.get("note", "").strip() or None,
        )
        db.session.add(plan)
        try:
            db.session.commit()
        except IntegrityError:
            db.session.rollback()
            abort(409, description="同一天、同餐別、同名稱的菜單已存在。")
        return redirect(url_for("order_tool.plan_detail", plan_id=plan.id))

    start = _parse_date(request.args.get("start"), date.today())
    end = start + timedelta(days=13)
    rows = (
        KitchenMenuPlan.query
        .filter(KitchenMenuPlan.service_date.between(start, end))
        .order_by(KitchenMenuPlan.service_date, KitchenMenuPlan.meal_type)
        .all()
    )
    body = r"""
<div class="grid"><section class="card span-5"><h2>建立中央菜單</h2><form method="post">
<div class="field"><label>供餐日期</label><input name="service_date" type="date" value="{{ today }}" required></div><br>
<div class="field"><label>餐別</label><select name="meal_type"><option selected>午餐</option><option>早餐</option><option>晚餐</option><option>點心</option></select></div><br>
<div class="field"><label>菜單名稱</label><input name="name" value="中央菜單" placeholder="例如 A 菜單"></div><br>
<div class="field"><label>備註</label><input name="note"></div><br><button class="btn">建立並排菜</button></form></section>
<section class="card span-7"><div class="top"><h2>{{ start }} 起 14 天</h2><form method="get" class="row"><input name="start" type="date" value="{{ start }}"><button class="btn secondary">查看</button></form></div>
{% if rows %}<table><tr><th>日期</th><th>餐別</th><th>菜單</th><th>菜色</th><th>供餐</th><th></th></tr>
{% for p in rows %}<tr><td>{{ p.service_date }}</td><td>{{ p.meal_type }}</td><td>{{ p.name }}</td><td>{{ p.items|length }} 道</td><td>{{ p.assignments|length }} 校 / {{ p.assignments|sum(attribute='headcount') }} 人</td>
<td><a href="{{ url_for('order_tool.plan_detail', plan_id=p.id) }}">編輯</a></td></tr>{% endfor %}</table>{% else %}<div class="empty">這段期間沒有菜單。</div>{% endif %}</section></div>"""
    return _render("菜單", body, rows=rows, start=start, today=date.today().isoformat())


@order_bp.get("/plans/<int:plan_id>")
def plan_detail(plan_id: int):
    plan = db.session.get(KitchenMenuPlan, plan_id)
    if not plan:
        abort(404)
    recipes_all = KitchenRecipe.query.order_by(KitchenRecipe.category, KitchenRecipe.name).all()
    schools_all = KitchenSchool.query.filter_by(active=True).order_by(KitchenSchool.name).all()
    total_people = sum(x.headcount for x in plan.assignments)
    body = r"""
<div class="top"><div><h1>{{ plan.service_date }}・{{ plan.meal_type }}・{{ plan.name }}</h1><div class="muted">先排菜，再指派學校與該日人數。</div></div>
<a class="btn secondary" href="{{ url_for('order_tool.plans') }}">← 回菜單</a></div>
<div class="grid"><section class="card span-6"><h2>① 今日菜色</h2>{% if plan.items %}<table><tr><th>類別</th><th>菜色</th><th>生料 g/人</th><th></th></tr>
{% for x in plan.items %}<tr><td>{{ x.recipe.category or '-' }}</td><td>{{ x.recipe.name }}</td><td>{{ '%.1f'|format(x.recipe.ingredients|sum(attribute='grams_per_person')|float) }}</td>
<td><form class="inline" method="post" action="{{ url_for('order_tool.plan_item_delete', row_id=x.id) }}"><button class="btn danger">移除</button></form></td></tr>{% endfor %}</table>{% else %}<div class="empty">尚未排菜。</div>{% endif %}
<hr style="border:0;border-top:1px solid #eee;margin:18px 0"><h3>加入菜色</h3>{% if recipes %}<form method="post" action="{{ url_for('order_tool.plan_item_add', plan_id=plan.id) }}" class="row">
<div class="field"><select name="recipe_id">{% for r in recipes %}<option value="{{ r.id }}">{{ r.category or '未分類' }}｜{{ r.name }}</option>{% endfor %}</select></div><button class="btn">加入</button></form>{% else %}<div class="notice">請先建立菜色配方。</div>{% endif %}</section>
<section class="card span-6"><h2>② 分配學校 / 人數</h2>{% if plan.assignments %}<table><tr><th>學校</th><th class="right">人數</th><th></th></tr>
{% for x in plan.assignments %}<tr><td>{{ x.school.name }}</td><td class="right">{{ x.headcount }}</td><td><form class="inline" method="post" action="{{ url_for('order_tool.assignment_delete', row_id=x.id) }}"><button class="btn danger">移除</button></form></td></tr>{% endfor %}
<tr><td><b>總人數</b></td><td class="right"><b>{{ total_people }}</b></td><td></td></tr></table>{% else %}<div class="empty">尚未分配學校。</div>{% endif %}
<hr style="border:0;border-top:1px solid #eee;margin:18px 0"><h3>加入 / 更新學校</h3>{% if schools %}<form method="post" action="{{ url_for('order_tool.assignment_add', plan_id=plan.id) }}" class="row">
<div class="field"><label>學校</label><select name="school_id">{% for s in schools %}<option value="{{ s.id }}">{{ s.name }}</option>{% endfor %}</select></div>
<div class="field"><label>當日人數</label><input name="headcount" type="number" min="0" required></div><button class="btn">儲存</button></form>{% else %}<div class="notice">請先建立學校。</div>{% endif %}</section>
<section class="card span-12"><div class="top"><div><h2>③ 試算採購</h2><div class="muted">完成菜色與學校後，可直接查看這一天要叫多少貨。</div></div>
<a class="btn" href="{{ url_for('order_tool.purchase', start=plan.service_date, end=plan.service_date) }}">產生叫貨表</a></div></section></div>"""
    return _render(
        "編輯菜單",
        body,
        plan=plan,
        recipes=recipes_all,
        schools=schools_all,
        total_people=total_people,
    )


@order_bp.post("/plans/<int:plan_id>/items")
def plan_item_add(plan_id: int):
    if not db.session.get(KitchenMenuPlan, plan_id):
        abort(404)
    recipe_id = _to_int(request.form.get("recipe_id"))
    if not db.session.get(KitchenRecipe, recipe_id):
        abort(400)
    existing = KitchenMenuPlanItem.query.filter_by(plan_id=plan_id, recipe_id=recipe_id).first()
    if not existing:
        max_order = max((x.sort_order for x in KitchenMenuPlanItem.query.filter_by(plan_id=plan_id).all()), default=-1)
        db.session.add(KitchenMenuPlanItem(plan_id=plan_id, recipe_id=recipe_id, sort_order=max_order + 1))
        db.session.commit()
    return redirect(url_for("order_tool.plan_detail", plan_id=plan_id))


@order_bp.post("/plan-items/<int:row_id>/delete")
def plan_item_delete(row_id: int):
    row = db.session.get(KitchenMenuPlanItem, row_id)
    if not row:
        abort(404)
    plan_id = row.plan_id
    db.session.delete(row)
    db.session.commit()
    return redirect(url_for("order_tool.plan_detail", plan_id=plan_id))


@order_bp.post("/plans/<int:plan_id>/assignments")
def assignment_add(plan_id: int):
    if not db.session.get(KitchenMenuPlan, plan_id):
        abort(404)
    school_id = _to_int(request.form.get("school_id"))
    headcount = max(_to_int(request.form.get("headcount")), 0)
    if not db.session.get(KitchenSchool, school_id):
        abort(400)
    row = KitchenMenuAssignment.query.filter_by(plan_id=plan_id, school_id=school_id).first()
    if row:
        row.headcount = headcount
    else:
        db.session.add(KitchenMenuAssignment(plan_id=plan_id, school_id=school_id, headcount=headcount))
    db.session.commit()
    return redirect(url_for("order_tool.plan_detail", plan_id=plan_id))


@order_bp.post("/assignments/<int:row_id>/delete")
def assignment_delete(row_id: int):
    row = db.session.get(KitchenMenuAssignment, row_id)
    if not row:
        abort(404)
    plan_id = row.plan_id
    db.session.delete(row)
    db.session.commit()
    return redirect(url_for("order_tool.plan_detail", plan_id=plan_id))


def _build_purchase(start: date, end: date):
    plans = (
        KitchenMenuPlan.query
        .filter(KitchenMenuPlan.service_date.between(start, end))
        .order_by(KitchenMenuPlan.service_date, KitchenMenuPlan.meal_type)
        .all()
    )

    # key = (date, supplier_name, ingredient_id)
    agg: dict[tuple, dict] = {}
    plan_summaries = []

    for plan in plans:
        people = sum(max(a.headcount, 0) for a in plan.assignments)
        plan_summaries.append({
            "date": plan.service_date,
            "meal_type": plan.meal_type,
            "name": plan.name,
            "people": people,
            "schools": len(plan.assignments),
        })
        if people <= 0:
            continue

        for menu_item in plan.items:
            for component in menu_item.recipe.ingredients:
                ing = component.ingredient
                supplier_name = ing.supplier.name if ing.supplier else "⚠ 未指定供應商"
                key = (plan.service_date, supplier_name, ing.id)
                grams = (component.grams_per_person or Decimal("0")) * people
                if key not in agg:
                    agg[key] = {
                        "date": plan.service_date,
                        "supplier": supplier_name,
                        "ingredient": ing.name,
                        "grams": Decimal("0"),
                        "unit_price": ing.unit_price or Decimal("0"),
                    }
                agg[key]["grams"] += grams

    grouped: dict[tuple, list] = defaultdict(list)
    for row in agg.values():
        row["kg"] = row["grams"] / Decimal("1000")
        row["cost"] = row["kg"] * row["unit_price"]
        grouped[(row["date"], row["supplier"])].append(row)

    groups = []
    for (service_date, supplier), rows in sorted(grouped.items(), key=lambda x: (x[0][0], x[0][1])):
        rows.sort(key=lambda r: r["ingredient"])
        groups.append({
            "date": service_date,
            "supplier": supplier,
            "rows": rows,
            "cost": sum((r["cost"] for r in rows), Decimal("0")),
        })

    total_cost = sum((g["cost"] for g in groups), Decimal("0"))
    missing_supplier = sum(1 for g in groups if g["supplier"].startswith("⚠"))
    return groups, plan_summaries, total_cost, missing_supplier


@order_bp.get("/purchase")
def purchase():
    start = _parse_date(request.args.get("start"), date.today())
    end = _parse_date(request.args.get("end"), start)
    if end < start:
        start, end = end, start

    groups, summaries, total_cost, missing_supplier = _build_purchase(start, end)
    do_print = request.args.get("print") == "1"
    body = r"""
<div class="top"><div><h1>採購叫貨表</h1><div class="muted">{{ start }} ～ {{ end }}</div></div>
<div class="row no-print"><form method="get" class="row"><div class="field"><label>開始</label><input name="start" type="date" value="{{ start }}"></div>
<div class="field"><label>結束</label><input name="end" type="date" value="{{ end }}"></div><button class="btn secondary">重新計算</button></form>
<a class="btn" href="{{ url_for('order_tool.purchase', start=start, end=end, print=1) }}" target="_blank">列印 / 另存 PDF</a></div></div>
{% if missing_supplier %}<div class="notice no-print">有 {{ missing_supplier }} 個日期×供應商群組包含未指定供應商食材，請先回「食材」補上廠商。</div>{% endif %}
<section class="card"><h2>供餐摘要</h2>{% if summaries %}<table><tr><th>日期</th><th>餐別</th><th>菜單</th><th class="right">學校</th><th class="right">人數</th></tr>
{% for s in summaries %}<tr><td>{{ s.date }}</td><td>{{ s.meal_type }}</td><td>{{ s.name }}</td><td class="right">{{ s.schools }}</td><td class="right">{{ s.people }}</td></tr>{% endfor %}</table>{% else %}<div class="empty">此期間沒有菜單。</div>{% endif %}</section><br>
{% if groups %}{% for g in groups %}<section class="card purchase-group"><div class="top"><div><h2>{{ g.date }}｜{{ g.supplier }}</h2></div><div><b>小計 ${{ '%.2f'|format(g.cost|float) }}</b></div></div>
<table><tr><th>食材</th><th class="right">需求量</th><th class="right">單價/kg</th><th class="right">預估金額</th></tr>
{% for r in g.rows %}<tr><td>{{ r.ingredient }}</td><td class="right"><b>{{ '%.3f'|format(r.kg|float) }} kg</b></td><td class="right">${{ '%.2f'|format(r.unit_price|float) }}</td><td class="right">${{ '%.2f'|format(r.cost|float) }}</td></tr>{% endfor %}</table></section><br>{% endfor %}
<section class="card"><div class="right"><h2>預估採購總額：${{ '%.2f'|format(total_cost|float) }}</h2></div></section>
{% elif summaries %}<section class="card"><div class="empty">有菜單，但尚未產生採購量。請確認：菜單有菜色、菜色有配方、菜單已分配學校與人數。</div></section>{% endif %}
{% if do_print %}<script>window.addEventListener('load',()=>window.print());</script>{% endif %}
"""
    return _render(
        "採購叫貨表",
        body,
        start=start,
        end=end,
        groups=groups,
        summaries=summaries,
        total_cost=total_cost,
        missing_supplier=missing_supplier,
        do_print=do_print,
    )
