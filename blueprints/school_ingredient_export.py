"""學校食材登錄 Excel 匯出。

第一版刻意不新增資料表／欄位：直接使用既有學校、菜單、配方與人數計算，
並把使用者提供的 2026-06-01 登錄表中可固定的欄位留在程式常數。
"""

from __future__ import annotations

from datetime import date
from decimal import Decimal
from io import BytesIO

from flask import Blueprint, request, send_file
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from sqlalchemy.orm import selectinload

from models import (
    KitchenMenuAssignment,
    KitchenMenuPlan,
    KitchenMenuPlanItem,
    KitchenRecipe,
    KitchenRecipeIngredient,
)

school_ingredient_export_bp = Blueprint("school_ingredient_export", __name__)

DEFAULT_SUPPLIER_NAME = "廣豐食品有限公司"
DEFAULT_ORIGIN = "臺灣"

EXPORT_HEADERS = (
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

# 由使用者提供的 schoolingredient_20260601.xlsx 整理；未收錄的食材保持空白，
# 不猜測驗證標章／號碼，也不需要為此修改資料庫。
CERTIFICATION_BY_INGREDIENT = {
    "三色丁": ("CAS台灣優良農產品", "123706"),
    "帶皮雞胸肉": ("CAS台灣優良農產品", "016683"),
    "敏豆": ("產銷履歷", "2604155169113265"),
    "杏鮑菇": ("生產追溯-農產品", "1004000002"),
    "洋蔥": ("生產追溯-農產品", "1201004693"),
    "玉米粒": ("CAS台灣優良農產品", "123701"),
    "甜不辣": ("生產追溯-水產品", "0316600001"),
    "甜椒": ("生產追溯-農產品", "1101003260"),
    "粳米": ("產銷履歷", "2605190098801217"),
    "絞肉": ("生產追溯-豬肉", "LE300431"),
    "荷葉白菜": ("台灣有機農產品", "1-010-100311"),
    "蚵白菜": ("產銷履歷", "00993663601066"),
    "雞蛋(白殼)": ("雞蛋噴印-洗選鮮蛋", "D43003260530C"),
    "香菇": ("生產追溯-農產品", "1004000002"),
}


def _parse_date(raw: str | None) -> date:
    try:
        return date.fromisoformat(raw or "")
    except ValueError:
        return date.today()


def _ingredient_weight_kg(component: KitchenRecipeIngredient, headcount: int) -> Decimal | None:
    """把此學校、此菜色、此食材的需求量換算成公斤；無可靠換算時回傳 None。"""

    ingredient = component.ingredient
    per_person = component.grams_per_person or Decimal("0")
    base_amount = per_person * Decimal(max(headcount, 0))
    base_unit = (ingredient.base_unit or "g").strip().lower()
    purchase_unit = (ingredient.purchase_unit or "").strip().lower()

    if base_unit in {"g", "公克", "克"}:
        return base_amount / Decimal("1000")

    # base_unit 若是「個」，但採購單位本身就是 kg，仍可依既有換算欄位換成 kg。
    if purchase_unit in {"kg", "公斤"}:
        units_per_kg = ingredient.grams_per_purchase_unit or Decimal("0")
        if units_per_kg > 0:
            return base_amount / units_per_kg

    return None


def _rows_for_date(service_date: date):
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
                certification, verification_number = CERTIFICATION_BY_INGREDIENT.get(
                    ingredient.name.strip(), ("", "")
                )
                rows.append((
                    service_date,
                    assignment.school.name,
                    recipe.name,
                    ingredient.name,
                    service_date,
                    None,  # 生產日期：依需求留空
                    None,  # 有效日期：依需求留空
                    None,  # 批號：既有範例亦留空
                    DEFAULT_SUPPLIER_NAME,  # 製造商 = 供應廠商
                    DEFAULT_SUPPLIER_NAME,
                    certification,
                    verification_number,
                    None,  # 產品名稱：沒有可靠來源時不猜
                    weight_kg,
                    "Y",
                    "Y",
                    "N",
                    DEFAULT_ORIGIN,
                ))
    return rows, sorted(unresolved)


def _build_workbook(service_date: date, rows):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Sheet1"
    sheet.append(list(EXPORT_HEADERS))
    for row in rows:
        sheet.append(list(row))

    header_fill = PatternFill("solid", fgColor="EAF2F8")
    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for row_number in range(2, sheet.max_row + 1):
        sheet.cell(row_number, 1).number_format = "yyyy/mm/dd"
        sheet.cell(row_number, 5).number_format = "yyyy/mm/dd"
        sheet.cell(row_number, 14).number_format = "0.###"

    widths = (12, 28, 22, 18, 12, 12, 12, 14, 24, 24, 24, 22, 18, 14, 12, 12, 10, 12)
    for index, width in enumerate(widths, start=1):
        sheet.column_dimensions[sheet.cell(1, index).column_letter].width = width
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:R{max(sheet.max_row, 1)}"
    sheet.sheet_view.showGridLines = True
    return workbook


@school_ingredient_export_bp.get("/summary/school-ingredient.xlsx")
def school_ingredient_export():
    service_date = _parse_date(request.args.get("date"))
    rows, unresolved = _rows_for_date(service_date)
    if unresolved:
        names = "、".join(unresolved)
        return (
            f"以下食材目前沒有可靠的公斤換算，為避免匯出錯誤重量已停止：{names}。"
            "請先確認其基本單位／採購換算。",
            400,
        )

    workbook = _build_workbook(service_date, rows)
    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    return send_file(
        output,
        as_attachment=True,
        download_name=f"學校食材登錄-{service_date.isoformat()}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
