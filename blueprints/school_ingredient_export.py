"""學校食材登錄 Excel 匯出。

不新增資料表／欄位：直接使用既有學校、菜單、配方、人數與當日採購供應商，
並以使用者提供的 schoolingredient Excel 原檔作為輸出模板。
"""

from __future__ import annotations

from copy import copy
from datetime import date
from decimal import Decimal
from io import BytesIO
from pathlib import Path

from flask import Blueprint, abort, current_app, request, send_file
from openpyxl import load_workbook
from sqlalchemy.orm import selectinload

from models import (
    KitchenMenuAssignment,
    KitchenMenuPlan,
    KitchenMenuPlanItem,
    KitchenPurchaseOrder,
    KitchenPurchaseOrderItem,
    KitchenRecipe,
    KitchenRecipeIngredient,
)

school_ingredient_export_bp = Blueprint("school_ingredient_export", __name__)

TEMPLATE_PATH = Path("static/schoolingredient_template.xlsx")
EXPECTED_HEADERS = (
    "供餐日期",
    "學校",
    "菜色名稱",
    "食材名稱",
    "進貨日期",
    "生產日期",
    "有效日期",
    "批號",
    "製造商",
    "供應商名稱",
    "食材驗證標章",
    "驗證號碼",
    "產品名稱",
    "重量(公斤)",
    "非基改玉米",
    "非基改黃豆",
    "加工品",
    "食材原產地",
)


def _parse_date(raw: str | None) -> date:
    try:
        return date.fromisoformat(raw or "")
    except ValueError:
        return date.today()


def _ingredient_weight_kg(component: KitchenRecipeIngredient, headcount: int) -> Decimal | None:
    """把此學校、此菜色、此食材需求換成公斤；沒有可靠換算時回傳 None。"""

    ingredient = component.ingredient
    per_person = component.grams_per_person or Decimal("0")
    base_amount = per_person * Decimal(max(headcount, 0))
    base_unit = (ingredient.base_unit or "g").strip().lower()
    purchase_unit = (ingredient.purchase_unit or "").strip().lower()

    if base_unit in {"g", "公克", "克"}:
        return base_amount / Decimal("1000")

    if purchase_unit in {"kg", "公斤"}:
        units_per_kg = ingredient.grams_per_purchase_unit or Decimal("0")
        if units_per_kg > 0:
            return base_amount / units_per_kg

    return None


def _supplier_name(item: KitchenPurchaseOrderItem) -> str:
    snapshot = (item.supplier_name_snapshot or "").strip()
    if snapshot and not snapshot.startswith("⚠"):
        return snapshot
    if item.supplier:
        return item.supplier.name
    return ""


def _supplier_names_for_date(service_date: date) -> dict[int, str]:
    """以採購頁當日品項的供應商為準，確保匯出與畫面一致。"""

    orders = (
        KitchenPurchaseOrder.query.filter(
            KitchenPurchaseOrder.service_date == service_date,
            KitchenPurchaseOrder.status.in_(("draft", "confirmed")),
        )
        .options(
            selectinload(KitchenPurchaseOrder.items)
            .selectinload(KitchenPurchaseOrderItem.supplier)
        )
        .order_by(KitchenPurchaseOrder.status.desc(), KitchenPurchaseOrder.id.desc())
        .all()
    )
    names = {}
    for order in orders:
        for item in order.items:
            if item.ingredient_id and item.ingredient_id not in names:
                names[item.ingredient_id] = _supplier_name(item)
    return names


def _template_values(sheet):
    """只從原模板讀取可重用的驗證資料與固定欄位，不把範例日期/產品名稱帶入。"""

    certification = {}
    fixed = {"corn": "Y", "soy": "Y", "processed": "N", "origin": "臺灣"}
    for row_number in range(2, sheet.max_row + 1):
        ingredient_name = str(sheet.cell(row_number, 4).value or "").strip()
        if ingredient_name and ingredient_name not in certification:
            mark = sheet.cell(row_number, 11).value
            number = sheet.cell(row_number, 12).value
            certification[ingredient_name] = (
                "" if mark is None else str(mark).strip(),
                "" if number is None else str(number).strip(),
            )
        if row_number == 2:
            fixed = {
                "corn": sheet.cell(row_number, 15).value or "Y",
                "soy": sheet.cell(row_number, 16).value or "Y",
                "processed": sheet.cell(row_number, 17).value or "N",
                "origin": sheet.cell(row_number, 18).value or "臺灣",
            }
    return certification, fixed


def _rows_for_date(service_date: date, certification, fixed):
    assignments = (
        KitchenMenuAssignment.query.join(KitchenMenuPlan)
        .filter(
            KitchenMenuPlan.service_date == service_date,
            KitchenMenuAssignment.service_status == "serving",
            KitchenMenuAssignment.headcount > 0,
        )
        .options(
            selectinload(KitchenMenuAssignment.school),
            selectinload(KitchenMenuAssignment.plan)
            .selectinload(KitchenMenuPlan.items)
            .selectinload(KitchenMenuPlanItem.recipe)
            .selectinload(KitchenRecipe.ingredients)
            .selectinload(KitchenRecipeIngredient.ingredient),
        )
        .all()
    )
    assignments.sort(key=lambda row: (row.school.name.casefold(), row.plan.meal_type, row.plan.name.casefold()))
    supplier_names = _supplier_names_for_date(service_date)

    rows = []
    unresolved = set()
    for assignment in assignments:
        for menu_item in sorted(assignment.plan.items, key=lambda item: item.sort_order):
            recipe = menu_item.recipe
            for component in recipe.ingredients:
                if (component.grams_per_person or Decimal("0")) <= 0:
                    continue
                ingredient = component.ingredient
                weight_kg = _ingredient_weight_kg(component, assignment.headcount)
                if weight_kg is None:
                    unresolved.add(ingredient.name)
                    continue

                supplier_name = supplier_names.get(ingredient.id, "")
                if not supplier_name and ingredient.supplier:
                    supplier_name = ingredient.supplier.name
                mark, verification_number = certification.get(ingredient.name.strip(), ("", ""))
                rows.append((
                    service_date,
                    assignment.school.name,
                    recipe.name,
                    ingredient.name,
                    service_date,
                    None,
                    None,
                    None,
                    supplier_name,
                    supplier_name,
                    mark,
                    verification_number,
                    None,
                    weight_kg,
                    fixed["corn"],
                    fixed["soy"],
                    fixed["processed"],
                    fixed["origin"],
                ))
    return rows, sorted(unresolved)


def _copy_template_row(sheet, source_row: int, target_row: int):
    for column in range(1, max(sheet.max_column, 21) + 1):
        source = sheet.cell(source_row, column)
        target = sheet.cell(target_row, column)
        target._style = copy(source._style)
        target.font = copy(source.font)
        target.fill = copy(source.fill)
        target.border = copy(source.border)
        target.alignment = copy(source.alignment)
        target.number_format = source.number_format
        target.protection = copy(source.protection)
    sheet.row_dimensions[target_row].height = sheet.row_dimensions[source_row].height


def _build_workbook(service_date: date):
    template_path = Path(current_app.root_path) / TEMPLATE_PATH
    if not template_path.is_file():
        abort(500, description="找不到學校食材登錄 Excel 模板。")

    workbook = load_workbook(template_path)
    sheet = workbook.active
    headers = tuple(sheet.cell(1, column).value for column in range(1, 19))
    if headers != EXPECTED_HEADERS:
        abort(500, description="學校食材登錄 Excel 模板欄位已被修改。")

    certification, fixed = _template_values(sheet)
    rows, unresolved = _rows_for_date(service_date, certification, fixed)
    if unresolved:
        names = "、".join(unresolved)
        return None, (
            f"以下食材目前沒有可靠的公斤換算，為避免匯出錯誤重量已停止：{names}。"
            "請先確認其基本單位／採購換算。"
        )

    first_data_row = 2
    template_last_row = max(sheet.max_row, first_data_row)
    required_last_row = max(template_last_row, first_data_row + len(rows) - 1)
    for row_number in range(template_last_row + 1, required_last_row + 1):
        _copy_template_row(sheet, first_data_row, row_number)

    for row_number in range(first_data_row, required_last_row + 1):
        for column in range(1, max(sheet.max_column, 21) + 1):
            sheet.cell(row_number, column).value = None

    for row_number, row in enumerate(rows, start=first_data_row):
        for column, value in enumerate(row, start=1):
            sheet.cell(row_number, column).value = value

    return workbook, None


@school_ingredient_export_bp.get("/summary/school-ingredient.xlsx")
def school_ingredient_export():
    service_date = _parse_date(request.args.get("date"))
    workbook, error = _build_workbook(service_date)
    if error:
        return error, 400

    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    return send_file(
        output,
        as_attachment=True,
        download_name=f"學校食材登錄-{service_date.isoformat()}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
