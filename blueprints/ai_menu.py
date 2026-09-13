"""AI 菜單：規則式排菜、練習入口、公版 Excel 與正式套用。"""
from __future__ import annotations

import json
import os
from collections import defaultdict
from datetime import date, datetime, timedelta
from io import BytesIO
from types import SimpleNamespace

from flask import current_app, flash, redirect, render_template, request, send_file, session, url_for
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from sqlalchemy.orm import joinedload

from blueprints.order_tool import _date, _int, order_bp
from extensions import db
from models import KitchenMenuPlan, KitchenMenuPlanItem, KitchenRecipe, KitchenRecipeIngredient
from ai_models import KitchenMealProfile, KitchenMenuDraft, KitchenMenuDraftItem
from services import meal_profiles as profile_service
from services import recipe_tags as tag_service
from services.ai_menu_engine import CATEGORY_ORDER, DEFAULT_STRUCTURE, Rules, build_structure, generate
from services.nutrition import recipe_nutrition
from services.practice_database import ensure_practice_database

WEEKDAY = "一二三四五六日"
PRACTICE_USER = "practice"


def _seed_profiles():
    if profile_service.seed_profiles(db.session):
        db.session.commit()


def _profiles(read_only=False):
    rows = (KitchenMealProfile.query.filter_by(active=True)
            .order_by(KitchenMealProfile.sort_order, KitchenMealProfile.id).all())
    if rows or not read_only:
        return rows
    result=[]
    for code,name,kmin,kmax,order,doc,version,url,note in profile_service.PROFILE_SEEDS:
        result.append(SimpleNamespace(id=None, code=code, name=name, kcal_min=kmin, kcal_max=kmax,
                                      kcal_label=f"{kmin}～{kmax} kcal", source_document=doc,
                                      source_version=version, note=note))
    return result


def _profile_by_code(code, read_only=False):
    row=KitchenMealProfile.query.filter_by(code=code, active=True).one_or_none()
    if row or not read_only:
        return row
    for p in _profiles(read_only=True):
        if p.code == code:
            return p
    return None


def _defaults():
    today=date.today()
    start=(today.replace(day=1)+timedelta(days=32)).replace(day=1)
    end=(start+timedelta(days=32)).replace(day=1)-timedelta(days=1)
    return start,end


def _parse_rules(form):
    return Rules(
        repeat_days=max(0,_int(form.get("repeat_days"),default=5) or 0),
        fish_per_week_min=max(0,_int(form.get("fish_per_week_min"),default=1) or 0),
        fried_per_week_max=max(0,_int(form.get("fried_per_week_max"),default=1) or 0),
        sweet_soup_per_week_max=max(0,_int(form.get("sweet_soup_per_week_max"),default=1) or 0),
        exclude_incomplete_nutrition=form.get("exclude_incomplete_nutrition") == "1",
    )


def _parse_request(read_only=False):
    profile=_profile_by_code((request.form.get("profile_code") or "").strip(), read_only=read_only)
    start=_date(request.form.get("start_date")); end=_date(request.form.get("end_date"))
    if profile is None:
        return None,"請選擇供餐對象。"
    if start is None or end is None or end < start:
        return None,"請填寫正確的排菜日期。"
    if (end-start).days > 62:
        return None,"一次最多排 63 天。"
    structure=build_structure({c:request.form.get(f"structure_{c}") for c in CATEGORY_ORDER})
    if not sum(structure.values()):
        return None,"每天至少要有一道菜。"
    rules=_parse_rules(request.form)
    include_weekends=request.form.get("include_weekends") == "1"
    kcal_min=int(profile.kcal_min) if profile.kcal_min is not None else None
    kcal_max=int(profile.kcal_max) if profile.kcal_max is not None else None
    rows=generate(start,end,structure,rules,kcal_min,kcal_max,include_weekends)
    return dict(profile=profile,start=start,end=end,structure=structure,rules=rules,
                include_weekends=include_weekends,rows=rows),None


def _group_draft(draft):
    grouped=defaultdict(list)
    items=(KitchenMenuDraftItem.query
           .options(joinedload(KitchenMenuDraftItem.recipe).joinedload(KitchenRecipe.ingredients).joinedload(KitchenRecipeIngredient.ingredient))
           .filter_by(draft_id=draft.id)
           .order_by(KitchenMenuDraftItem.service_date,KitchenMenuDraftItem.sort_order,KitchenMenuDraftItem.id).all())
    for item in items:
        nutrition=recipe_nutrition(item.recipe) if item.recipe else None
        grouped[item.service_date].append(dict(
            category=item.category, slot_index=item.slot_index, recipe=item.recipe,
            kcal=nutrition.kcal_int if nutrition else None,
            complete=bool(nutrition and nutrition.complete),
        ))
    return grouped


def _practice_ok():
    return bool(session.get("kitchen_practice"))


@order_bp.get("/ai-menu")
def ai_menu_index():
    _seed_profiles()
    start,end=_defaults()
    drafts=KitchenMenuDraft.query.order_by(KitchenMenuDraft.id.desc()).limit(20).all()
    return render_template("kitchen/ai_menu.html", profiles=_profiles(), start_default=start,
        end_default=end, categories=CATEGORY_ORDER, defaults=DEFAULT_STRUCTURE,
        drafts=drafts, practice=_practice_ok(), practice_user=PRACTICE_USER)


@order_bp.post("/ai-menu/generate")
def ai_menu_generate():
    payload,error=_parse_request(read_only=False)
    if error:
        flash(error,"error"); return redirect(url_for("order_tool.ai_menu_index"))
    draft=KitchenMenuDraft(
        name=(request.form.get("name") or "AI 自動排菜").strip()[:120],
        profile_id=payload["profile"].id, start_date=payload["start"], end_date=payload["end"],
        include_weekends=payload["include_weekends"], structure_json=json.dumps(payload["structure"],ensure_ascii=False),
        rules_json=json.dumps(payload["rules"].__dict__,ensure_ascii=False), status="draft")
    db.session.add(draft); db.session.flush()
    for row in payload["rows"]:
        counts=defaultdict(int)
        for sort_order,candidate in enumerate(row.items,1):
            counts[candidate.category]+=1
            db.session.add(KitchenMenuDraftItem(draft_id=draft.id,service_date=row.service_date,
                category=candidate.category,slot_index=counts[candidate.category],sort_order=sort_order,
                recipe_id=candidate.id))
    db.session.commit()
    return redirect(url_for("order_tool.ai_menu_draft",draft_id=draft.id))


@order_bp.get("/ai-menu/drafts/<int:draft_id>")
def ai_menu_draft(draft_id):
    draft=db.session.get(KitchenMenuDraft,draft_id)
    if draft is None: return ("Not found",404)
    return render_template("kitchen/ai_menu_result.html", draft=draft, grouped=_group_draft(draft),
                           weekday=WEEKDAY, practice=_practice_ok())


@order_bp.get("/ai-menu/drafts/<int:draft_id>/public.xlsx")
def ai_menu_public_excel(draft_id):
    draft=db.session.get(KitchenMenuDraft,draft_id)
    if draft is None: return ("Not found",404)
    grouped=_group_draft(draft)
    wb=Workbook(); ws=wb.active; ws.title="公版菜單"
    headers=["日期","星期","類別","菜色","預估熱量(kcal/人)","營養狀態"]
    ws.append(headers)
    for c in ws[1]: c.font=Font(bold=True)
    for d in sorted(grouped):
        for item in grouped[d]:
            ws.append([d,f"週{WEEKDAY[d.weekday()]}",item["category"],item["recipe"].name if item["recipe"] else "",
                       item["kcal"] if item["kcal"] is not None else "","完整" if item["complete"] else "待確認"])
    for col,width in {"A":14,"B":10,"C":12,"D":28,"E":20,"F":14}.items(): ws.column_dimensions[col].width=width
    for row in ws.iter_rows():
        for c in row: c.alignment=Alignment(vertical="center",wrap_text=True)
    output=BytesIO(); wb.save(output); output.seek(0)
    return send_file(output,as_attachment=True,download_name=f"公版菜單_{draft.start_date}_{draft.end_date}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@order_bp.post("/ai-menu/drafts/<int:draft_id>/apply")
def ai_menu_apply(draft_id):
    if request.form.get("confirm_apply") != "YES":
        flash("未套用：請勾選確認後再執行。","error"); return redirect(url_for("order_tool.ai_menu_draft",draft_id=draft_id))
    draft=db.session.get(KitchenMenuDraft,draft_id)
    if draft is None: return ("Not found",404)
    grouped=defaultdict(list)
    for item in KitchenMenuDraftItem.query.filter_by(draft_id=draft.id).order_by(KitchenMenuDraftItem.service_date,KitchenMenuDraftItem.sort_order).all():
        grouped[item.service_date].append(item)
    for service_date,items in grouped.items():
        plan=KitchenMenuPlan.query.filter_by(service_date=service_date,meal_type=draft.meal_type,name="中央菜單").one_or_none()
        if plan is None:
            plan=KitchenMenuPlan(service_date=service_date,meal_type=draft.meal_type,name="中央菜單",status="draft")
            db.session.add(plan); db.session.flush()
        KitchenMenuPlanItem.query.filter_by(plan_id=plan.id).delete(synchronize_session=False)
        for index,item in enumerate(items,1):
            if item.recipe_id:
                db.session.add(KitchenMenuPlanItem(plan_id=plan.id,recipe_id=item.recipe_id,sort_order=index))
    draft.status="applied"; draft.applied_at=datetime.utcnow(); db.session.commit()
    flash("已套用到練習中央菜單；正式資料不受影響。" if _practice_ok() else "已套用到正式中央菜單；學校人數與採購仍沿用原本流程。","success")
    return redirect(url_for("order_tool.ai_menu_draft",draft_id=draft.id))


@order_bp.post("/ai-menu/drafts/<int:draft_id>/delete")
def ai_menu_delete(draft_id):
    draft=db.session.get(KitchenMenuDraft,draft_id)
    if draft: db.session.delete(draft); db.session.commit()
    return redirect(url_for("order_tool.ai_menu_index"))


@order_bp.route("/ai-menu/practice/login",methods=["GET","POST"])
def ai_menu_practice_login():
    configured=current_app.config.get("KITCHEN_PRACTICE_PASSWORD") or os.environ.get("KITCHEN_PRACTICE_PASSWORD","")
    error=""
    if request.method == "POST":
        if configured and request.form.get("user") == PRACTICE_USER and request.form.get("password") == configured:
            try:
                ensure_practice_database()
            except Exception:
                current_app.logger.exception("Failed to initialize kitchen practice database")
                error="練習資料庫初始化失敗，請確認 PRACTICE_DATABASE_URL 或部署儲存空間設定。"
            else:
                db.session.remove()
                session["kitchen_practice"]=True
                return redirect(url_for("order_tool.index"))
        elif not error:
            error="練習帳號或密碼錯誤。" if configured else "尚未設定 KITCHEN_PRACTICE_PASSWORD。"
    return render_template("kitchen/ai_menu_practice_login.html",error=error,configured=bool(configured),practice_user=PRACTICE_USER)


@order_bp.get("/ai-menu/practice/logout")
def ai_menu_practice_logout():
    db.session.remove()
    session.pop("kitchen_practice",None)
    return redirect(url_for("order_tool.ai_menu_practice_login"))


@order_bp.get("/ai-menu/practice")
def ai_menu_practice():
    if not _practice_ok(): return redirect(url_for("order_tool.ai_menu_practice_login"))
    return redirect(url_for("order_tool.index"))


@order_bp.post("/ai-menu/practice/generate")
def ai_menu_practice_generate():
    if not _practice_ok(): return redirect(url_for("order_tool.ai_menu_practice_login"))
    # Backward-compatible endpoint: practice now uses the exact same draft flow
    # as the normal planner, with the ORM transparently routed to the sandbox.
    return ai_menu_generate()


@order_bp.get("/recipe-tags")
def ai_recipe_tags():
    recipes=(KitchenRecipe.query.options(joinedload(KitchenRecipe.ingredients).joinedload(KitchenRecipeIngredient.ingredient),joinedload(KitchenRecipe.tags))
             .filter(KitchenRecipe.active.is_(True)).order_by(KitchenRecipe.category,KitchenRecipe.name).all())
    rows=[dict(recipe=r,suggested=tag_service.suggest_tags(r),manual={t.tag for t in r.tags if t.source=="manual"}) for r in recipes]
    return render_template("kitchen/recipe_tags.html",rows=rows,tag_defs=tag_service.TAG_DEFS)


@order_bp.post("/recipe-tags")
def ai_recipe_tags_save():
    recipe_id=_int(request.form.get("recipe_id"),default=0)
    recipe=db.session.get(KitchenRecipe,recipe_id) if recipe_id else None
    if recipe is None: return ("Not found",404)
    tag_service.set_manual_tags(db.session,recipe,request.form.getlist("tags")); db.session.commit()
    flash("菜色標記已儲存。","success")
    return redirect(url_for("order_tool.ai_recipe_tags"))
