from datetime import date
from decimal import Decimal
from io import BytesIO
import json

from openpyxl import load_workbook
import pytest

from app import create_app
from extensions import db
from models import (KitchenIngredient, KitchenMenuPlan, KitchenMenuPlanItem,
                    KitchenMenuAssignment, KitchenPurchaseOrder,
                    KitchenPurchaseOrderItem, KitchenRecipe,
                    KitchenRecipeIngredient, KitchenSchool)
from services.ftrace import SCHEMAS, ExportError, build_sheets, draft, export_workbook

DAY = date(2026, 11, 5)


@pytest.fixture()
def app(tmp_path):
    app = create_app(dict(TESTING=True, SECRET_KEY="ftrace-test-only",
        SQLALCHEMY_DATABASE_URI=f"sqlite:///{tmp_path / 'ftrace.db'}",
        AUTO_CREATE_DB=True, KITCHEN_CSRF_ENABLED=False, PRODUCTION=False))
    with app.app_context():
        ingredient = KitchenIngredient(name="高麗菜", base_unit="g", purchase_unit="kg")
        recipe = KitchenRecipe(name="炒高麗菜")
        school = KitchenSchool(name="測試國小", default_headcount=10)
        db.session.add_all([ingredient, recipe, school])
        db.session.flush()
        db.session.add(KitchenRecipeIngredient(recipe_id=recipe.id,
            ingredient_id=ingredient.id, grams_per_person=80))
        for suffix, count in (("菜單", 10), ("素食菜單", 2)):
            plan = KitchenMenuPlan(service_date=DAY, name=f"測試國小{suffix}")
            db.session.add(plan); db.session.flush()
            db.session.add(KitchenMenuPlanItem(plan_id=plan.id, recipe_id=recipe.id))
            db.session.add(KitchenMenuAssignment(plan_id=plan.id, school_id=school.id,
                headcount=count, service_status="serving"))
        order = KitchenPurchaseOrder(service_date=DAY, supplier_key="daily",
            supplier_name_snapshot="每日採購單", status="draft")
        db.session.add(order); db.session.flush()
        db.session.add(KitchenPurchaseOrderItem(order_id=order.id,
            ingredient_id=ingredient.id, ingredient_name_snapshot="高麗菜",
            base_unit_snapshot="g", purchase_unit_snapshot="kg",
            grams_per_purchase_unit_snapshot=1000, actual_order_qty=Decimal("1.2345")))
        db.session.commit()
    yield app
    with app.app_context():
        db.session.remove()
        db.drop_all()


def payload():
    return {"receipts": [dict(name="高麗菜", include=True, material_code="B001",
        spec_code="B001A025", supplier_code="A025", kg="1.235",
        received_date="2026-11-05")], "meals": [dict(name="測試國小", include=True,
        product_code="AB005", customer_code="C001", kg="7.2", quantity="12",
        materials=[dict(name="高麗菜", code="B001")])]}


def test_drafts_are_read_only_and_do_not_treat_raw_weight_as_finished_weight(app):
    with app.app_context():
        values = draft(DAY)
        assert values["receipts"][0]["kg"] == "1.235"
        assert len(values["meals"]) == 1
        assert values["meals"][0]["quantity"] == "12"
        assert values["meals"][0]["kg"] == ""
        assert len(values["meals"][0]["materials"]) == 1
        assert not db.session.new and not db.session.dirty and not db.session.deleted
        item = KitchenPurchaseOrderItem.query.one()
        item.base_unit_snapshot = "個"
        db.session.flush()
        assert draft(DAY)["receipts"][0]["kg"] == ""


def test_receipt_selection_uses_delivery_date_and_ignores_cancelled_orders(app):
    with app.app_context():
        item = KitchenPurchaseOrderItem.query.one()
        item.delivery_date = date(2026, 11, 4)
        db.session.flush()
        assert draft(DAY)["receipts"] == []
        assert len(draft(date(2026, 11, 4))["receipts"]) == 1
        item.order.status = "cancelled"
        db.session.flush()
        assert draft(date(2026, 11, 4))["receipts"] == []


def test_small_bentos_merge_into_pingzhen_and_unneeded_schools_are_excluded(app):
    with app.app_context():
        recipe = KitchenRecipe.query.one()
        for name, count in (("桃園市平鎮區平鎮高中", 100), ("平鎮高中 (小便當)", 20),
                            ("明原", 10), ("廣豐食品股份有限公司", 10),
                            ("桃園市中壢區國立中央大學附設私立幼兒園", 10),
                            ("桃園市平鎮區新勢國小附設幼兒園", 10)):
            school = KitchenSchool(name=name)
            plan = KitchenMenuPlan(service_date=DAY, name=name+"菜單")
            db.session.add_all([school, plan]); db.session.flush()
            db.session.add(KitchenMenuPlanItem(plan_id=plan.id, recipe_id=recipe.id))
            db.session.add(KitchenMenuAssignment(plan_id=plan.id, school_id=school.id,
                headcount=count, service_status="serving"))
        db.session.flush()
        values = draft(DAY)
        pingzhen = next(r for r in values["meals"] if "平鎮" in r["name"])
        assert pingzhen["quantity"] == "120"
        assert pingzhen["product_code"] == "AA002"
        assert pingzhen["customer_code"] == "AA002"
        assert pingzhen["kg"] == "72"
        assert len(pingzhen["materials"]) == 1
        assert len(values["meals"]) == 2  # existing test school + Pingzhen
        assert len(values["skipped"]) == 4


@pytest.mark.parametrize("kind", SCHEMAS)
def test_workbook_roundtrip_matches_official_sheets_headers_and_text_cells(kind):
    sheets = build_sheets(kind, DAY, payload())
    book = load_workbook(export_workbook(kind, sheets), data_only=False)
    assert book.sheetnames == list(SCHEMAS[kind])
    for sheet in book:
        assert list(next(sheet.values)) == SCHEMAS[kind][sheet.title]
        assert sheet.max_row == len(sheets[sheet.title]) + 1
        for row in list(sheet)[1:]:
            for cell in row:
                if cell.value is not None:
                    assert cell.data_type == "s"
                    assert cell.number_format == "@"
        if sheet.max_row > 1:
            assert sheet["A2"].value == "2"
            assert sheet["B2"].value == "2026/11/05"
            assert sheet["C2"].value == "01"


@pytest.mark.parametrize("field,value", [("kg","NaN"),("kg","Infinity"),
    ("kg","-1"),("kg","0"),("kg","1.0001"),("material_code",""),
    ("spec_code",""),("supplier_code",""),("material_code","=1+1"),
    ("received_date","2026-11-06"),("received_date","bad")])
def test_invalid_receipt_values_fail_closed(field, value):
    data = payload(); data["receipts"][0][field] = value
    with pytest.raises(ExportError): build_sheets("receiving", DAY, data)


def test_excluded_receipts_and_duplicate_ingredients():
    data = payload()
    data["receipts"].append(dict(name="skip", include=False))
    assert len(build_sheets("receiving", DAY, data)["收貨資料"]) == 1
    data["meals"][0]["materials"] *= 2
    assert len(build_sheets("manufacture", DAY, data)["製造餐盒原材料"]) == 1
    data["meals"][0]["materials"][0]["code"] = ""
    with pytest.raises(ExportError): build_sheets("manufacture", DAY, data)


def test_routes_render_and_export_without_updating_source_records(app):
    client = app.test_client()
    for kind in SCHEMAS:
        response = client.get(f"/admin/order-tool/ftrace?date={DAY}&kind={kind}")
        assert response.status_code == 200
        assert "非追不可匯入檔" in response.get_data(as_text=True)
        response = client.post("/admin/order-tool/ftrace/export", data={
            "date":str(DAY), "kind":kind, "payload":json.dumps(payload()), "confirmed":"1"})
        assert response.status_code == 200
        assert response.headers["Cache-Control"] == "no-store"
        assert load_workbook(BytesIO(response.data)).sheetnames == list(SCHEMAS[kind])
    with app.app_context():
        assert KitchenPurchaseOrderItem.query.one().actual_order_qty == Decimal("1.2345")
        assert KitchenMenuAssignment.query.count() == 2
    assert client.get("/admin/order-tool/ftrace?date=bad").status_code == 400
    assert client.get("/admin/order-tool/ftrace?kind=bad").status_code == 400
    assert client.post("/admin/order-tool/ftrace/export", data={"payload":"{}"}).status_code == 400


def test_export_requires_csrf_with_production_setting(app):
    app.config["KITCHEN_CSRF_ENABLED"] = True
    client = app.test_client()
    response = client.post("/admin/order-tool/ftrace/export", data={
        "date":str(DAY), "kind":"receiving", "payload":json.dumps(payload()), "confirmed":"1"})
    assert "spreadsheetml" not in response.content_type
    client.get(f"/admin/order-tool/ftrace?date={DAY}")
    with client.session_transaction() as session:
        token = session["_kitchen_csrf"]
    response = client.post("/admin/order-tool/ftrace/export", data={
        "date":str(DAY), "kind":"receiving", "payload":json.dumps(payload()),
        "confirmed":"1", "_csrf_token":token})
    assert response.status_code == 200
