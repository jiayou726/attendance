"""Read-only drafts and strict exports for the three meal-industry imports.

No ORM writes: edits belong to the export, never to menus or procurement.
Schemas transcribed from the official templates downloaded 2026-11-05.
"""
from datetime import date
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from io import BytesIO
import json
import os
from pathlib import Path

import xlwt
from flask import current_app
from sqlalchemy import or_, and_
from sqlalchemy.orm import selectinload

from models import (KitchenMenuPlan, KitchenMenuAssignment, KitchenMenuPlanItem,
                    KitchenRecipe, KitchenRecipeIngredient, KitchenPurchaseOrder,
                    KitchenPurchaseOrderItem)

ROOT = Path(__file__).resolve().parents[1]
SCHEMAS = json.loads((ROOT / "data/ftrace/schemas.json").read_text())
KINDS = (("receiving", "收貨"), ("manufacture", "製造"), ("delivery", "交貨"))

def configured_products():
    """Business codes live in private service configuration, never source code."""
    raw = current_app.config.get("FTRACE_PRODUCTS_JSON",
                                 os.environ.get("FTRACE_PRODUCTS_JSON", "[]"))
    try:
        products = json.loads(raw) if isinstance(raw, str) else raw
        if not isinstance(products, list):
            raise ValueError
        result = {}
        for product in products:
            name = product["name"]
            if not isinstance(name, str) or name in result:
                raise ValueError
            if not all(isinstance(product.get(key, ""), str)
                       for key in ("code", "customer_code")):
                raise ValueError
            grams = product.get("grams")
            if grams is not None and (not Decimal(str(grams)).is_finite()
                                      or Decimal(str(grams)) <= 0):
                raise ValueError
            result[name] = product
        return result
    except (ValueError, TypeError, KeyError, InvalidOperation):
        current_app.logger.warning("Invalid FTRACE_PRODUCTS_JSON; codes require manual confirmation.")
        return {}


def school_group(name):
    """User-confirmed scope: small bentos belong to Pingzhen's meal product."""
    normalized = "".join(name.split())
    if (normalized == "明原" or "廣豐" in normalized
            or ("中央大學" in normalized and "幼兒園" in normalized)):
        return None
    if "平鎮高中" in normalized or "鎮高" in normalized:
        return "平鎮高中"
    if "信義" in normalized:
        return "信義國小"
    if "新勢" in normalized:
        return "新勢國小"
    for school in ("中平國小", "新明國中"):
        if school in normalized:
            return school
    return normalized


def text_number(value, places=3):
    if value is None:
        return ""
    number = Decimal(str(value)).quantize(Decimal(10) ** -places, rounding=ROUND_HALF_UP)
    return format(number, "f").rstrip("0").rstrip(".") if number else "0"


def _kg(item):
    # Quantity × conversion is a weight only when the snapshot base unit is g.
    if item.base_unit_snapshot != "g":
        return ""
    return text_number(Decimal(item.actual_order_qty or 0) *
                       Decimal(item.grams_per_purchase_unit_snapshot or 0) / 1000)


def draft(day):
    from blueprints.order_tool import _menu_item_components
    products = configured_products()
    receipts = []
    items = (KitchenPurchaseOrderItem.query.join(KitchenPurchaseOrder)
             .filter(KitchenPurchaseOrder.status != "cancelled", or_(
                 KitchenPurchaseOrderItem.delivery_date == day,
                 and_(KitchenPurchaseOrderItem.delivery_date.is_(None),
                      KitchenPurchaseOrder.service_date == day)))
             .options(selectinload(KitchenPurchaseOrderItem.ingredient),
                      selectinload(KitchenPurchaseOrderItem.supplier),
                      selectinload(KitchenPurchaseOrderItem.order))
             .order_by(KitchenPurchaseOrderItem.id).all())
    for item in items:
        receipts.append(dict(key=str(item.id), include=True,
            material_key=str(item.ingredient_id or f"purchase-{item.id}"),
            supplier_key=str(item.supplier_id or item.order.supplier_id or ""),
            name=item.ingredient_name_snapshot,
            supplier_name=item.supplier_name_snapshot or item.order.supplier_name_snapshot,
            source=f"採購單 #{item.order_id}：{item.actual_order_qty} {item.purchase_unit_snapshot}",
            material_code="", spec_code="", supplier_code="", kg=_kg(item),
            received_date=(item.delivery_date or day).isoformat()))

    plans = (KitchenMenuPlan.query.filter_by(service_date=day)
             .options(selectinload(KitchenMenuPlan.assignments).selectinload(KitchenMenuAssignment.school),
                      selectinload(KitchenMenuPlan.items).selectinload(KitchenMenuPlanItem.recipe)
                      .selectinload(KitchenRecipe.ingredients)
                      .selectinload(KitchenRecipeIngredient.ingredient))
             .order_by(KitchenMenuPlan.id).all())
    meals = {}
    skipped = set()
    for plan in plans:
        if plan.status == "cancelled":
            continue
        components = [c for item in plan.items for c in _menu_item_components(item)]
        for assignment in plan.assignments:
            if assignment.service_status != "serving" or assignment.headcount <= 0:
                continue
            school = assignment.school
            group = school_group(school.name)
            if group is None:
                skipped.add(school.name)
                continue
            key = f"{group}:{plan.meal_type}"
            meal = meals.setdefault(key, dict(key=key, school_key=group,
                name=("平鎮高中（含小便當）" if group == "平鎮高中" else
                      group + "（含幼兒園）" if group in {"新勢國小", "信義國小"} else school.name),
                meal_type=plan.meal_type, quantity=0, kg="",
                product_code="", customer_code="", include=True, materials={}, source=[]))
            meal["quantity"] += assignment.headcount
            meal["source"].append(plan.name)
            for component in components:
                ingredient = component.ingredient
                if ingredient is not None:
                    meal["materials"][str(ingredient.id)] = dict(
                        key=str(ingredient.id), name=ingredient.name, code="")
    for meal in meals.values():
        meal["quantity"] = str(meal["quantity"])
        meal["materials"] = list(meal["materials"].values())
        meal["source"] = "、".join(meal["source"])
        product = products.get(meal["school_key"])
        if product:
            meal["product_code"] = product.get("code", "")
            meal["customer_code"] = product.get("customer_code", "")
            grams = product.get("grams")
            if grams:
                meal["kg"] = text_number(Decimal(meal["quantity"]) * Decimal(str(grams)) / 1000)
                meal["source"] += f"；按平台產品規格 {grams} g／份預填成品重量，請核對實際出餐。"
    return dict(date=day.isoformat(), receipts=receipts, meals=list(meals.values()), skipped=sorted(skipped))


class ExportError(ValueError):
    pass


def _value(row, field, label, required=True):
    value = str(row.get(field, "")).strip()
    if required and not value:
        raise ExportError(f"{row.get('name', '資料')}：請填寫{label}。")
    if len(value) > 160 or value.startswith(("=", "+", "@")):
        raise ExportError(f"{label}格式不正確。")
    return value


def _number(row, field, label, places, positive=True):
    value = _value(row, field, label)
    try:
        number = Decimal(value)
    except InvalidOperation:
        raise ExportError(f"{row.get('name', '資料')}：{label}必須為數字。") from None
    if not number.is_finite() or number < 0 or (positive and number == 0) or number > Decimal("999999999"):
        raise ExportError(f"{label}必須是有效的正數。")
    if number != number.quantize(Decimal(10) ** -places):
        raise ExportError(f"{label}最多 {places} 位小數。")
    return text_number(number, places)


def build_sheets(kind, day, payload):
    if kind not in SCHEMAS:
        raise ExportError("請選擇收貨、製造或交貨。")
    sheets = {name: [] for name in SCHEMAS[kind]}
    report_date = day.strftime("%Y/%m/%d")
    field = "receipts" if kind == "receiving" else "meals"
    source = payload.get(field)
    if not isinstance(source, list) or len(source) > 99 or not all(isinstance(r, dict) for r in source):
        raise ExportError("資料格式錯誤，每日最多匯出 99 筆。")
    rows = [r for r in source if r.get("include") is True]
    if not rows:
        raise ExportError("這一天沒有勾選可匯出的資料。")
    for index, row in enumerate(rows, 1):
        seq = f"{index:02}"
        code = _value(row, "material_code" if kind == "receiving" else "product_code", "平台自編碼")
        key = ["2", report_date, seq, code]
        if kind == "receiving":
            spec = _value(row, "spec_code", "原材料廠商自編碼（原料規格碼）")
            supplier = _value(row, "supplier_code", "供應商自編碼")
            kg = _number(row, "kg", "實收淨重（公斤）", 3)
            try:
                received = date.fromisoformat(_value(row, "received_date", "收貨日期"))
            except ValueError:
                raise ExportError("收貨日期格式錯誤。") from None
            if received > day:
                raise ExportError("收貨日期不可晚於申報日期。")
            sheets["收貨資料"].append(key + [spec, kg, "", "", "", ""])
            sheets["收貨資料收貨日期"].append(key + [spec, received.strftime("%Y/%m/%d")])
            sheets["收貨資料供應商"].append(key + [spec, supplier, ""])
        elif kind == "manufacture":
            kg = _number(row, "kg", "製造成品總淨重（公斤）", 3)
            qty = _number(row, "quantity", "製造份數", 2)
            materials = row.get("materials")
            if not isinstance(materials, list) or not materials or len(materials) > 1000:
                raise ExportError(f"{row.get('name', '')}：請填寫當日使用的原材料。")
            codes = set()
            for material in materials:
                if not isinstance(material, dict):
                    raise ExportError("原材料格式錯誤。")
                codes.add(_value(material, "code", "原材料自編碼"))
            sheets["製造資料"].append(key + [kg, qty])
            sheets["製造餐食原材料"].extend(key + [c] for c in sorted(codes))
        else:
            kg = _number(row, "kg", "交貨總淨重（公斤）", 3)
            qty = _number(row, "quantity", "交貨份數", 2)
            customer = _value(row, "customer_code", "下游業者自編碼")
            sheets["交貨資料"].append(key + [kg, qty, "0", "1", ""])
            sheets["交貨資料下游業者"].append(key + [customer, ""])
            sheets["交貨資料其他"].append(key + ["0", "", "0", "", "0", "", ""])
    return sheets


def export_workbook(kind, sheets):
    # The live uploader rejects OOXML (.xlsx). Use real BIFF8/OLE .xls,
    # not a renamed .xlsx, matching the official templates.
    workbook = xlwt.Workbook(encoding="utf-8")
    body_style = xlwt.easyxf("align: vert centre, wrap on;", num_format_str="@")
    header_style = xlwt.easyxf("font: bold on; align: vert centre, wrap on;", num_format_str="@")
    for name, headers in SCHEMAS[kind].items():
        sheet = workbook.add_sheet(name)
        for r, row in enumerate([headers, *sheets[name]]):
            for c, value in enumerate(row):
                sheet.write(r, c, str(value), header_style if r == 0 else body_style)
        sheet.row(0).height = 720
        for c in range(len(headers)):
            sheet.col(c).width = 24 * 256
        sheet.set_panes_frozen(True)
        sheet.set_horz_split_pos(1)
    stream = BytesIO()
    workbook.save(stream)
    from services.ftrace_xls import excel_envelope
    return BytesIO(excel_envelope(stream.getvalue()))
