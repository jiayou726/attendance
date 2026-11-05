from datetime import date
from decimal import Decimal
from io import BytesIO
import json

import xlrd
import xlwt
from xlrd.compdoc import CompDoc
import pytest

from app import create_app
from extensions import db
from models import (KitchenIngredient, KitchenMenuPlan, KitchenMenuPlanItem,
                    KitchenMenuAssignment, KitchenPurchaseOrder,
                    KitchenPurchaseOrderItem, KitchenRecipe,
                    KitchenRecipeIngredient, KitchenSchool)
from services.ftrace import SCHEMAS, ExportError, build_sheets, draft, export_workbook
from services.ftrace_xls import excel_envelope

DAY = date(2026, 11, 5)


@pytest.fixture()
def app(tmp_path):
    app = create_app(dict(TESTING=True, SECRET_KEY="ftrace-test-only",
        SQLALCHEMY_DATABASE_URI=f"sqlite:///{tmp_path / 'ftrace.db'}",
        AUTO_CREATE_DB=True, KITCHEN_CSRF_ENABLED=False, PRODUCTION=False))
    app.config["FTRACE_PRODUCTS_JSON"] = [
        dict(name="平鎮高中", code="TEST-P1", customer_code="TEST-C1", grams=600),
        dict(name="新勢國小", code="TEST-P2", customer_code="TEST-C2"),
        dict(name="信義國小", code="TEST-P3", customer_code="TEST-C3", grams=600)]
    app.config["FTRACE_REFERENCES_JSON"] = dict(materials=[], products=[], parties=[])
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
        product_code="TEST-P0", customer_code="TEST-C0", kg="7.2", quantity="12",
        materials=[dict(name="高麗菜", code="B001")])]}


def private_references():
    return dict(materials=[dict(name="高麗菜", code="TEST-M",
        specs=[dict(code="TEST-S", supplier_name="原料製造商")])], products=[],
        parties=[dict(name="測試採購公司", code="TEST-V", aliases=["每日採購單"])])


def test_system_prefills_all_three_receipt_codes_without_browser_import_or_database_writes(app):
    app.config["FTRACE_REFERENCES_JSON"] = private_references()
    with app.app_context():
        row = draft(DAY)["receipts"][0]
        assert (row["material_code"], row["spec_code"], row["supplier_code"]) == ("TEST-M", "TEST-S", "TEST-V")
        assert draft(DAY)["meals"][0]["materials"][0]["code"] == "TEST-M"
        assert not db.session.new and not db.session.dirty and not db.session.deleted
    response = app.test_client().get("/admin/order-tool/ftrace?date=2026-11-05")
    assert response.status_code == 200
    assert response.headers["Cache-Control"] == "no-store"
    assert b'ftrace-references' in response.data and b'TEST-S' in response.data


def test_ambiguous_references_are_not_guessed_and_supplier_spec_is_selected_only_when_unique(app):
    refs = private_references()
    refs["materials"][0]["specs"].append(dict(code="TEST-S2", supplier_name="另一製造商"))
    refs["parties"].append(dict(name="另一交易對象", code="TEST-V2", aliases=["每日採購單"]))
    app.config["FTRACE_REFERENCES_JSON"] = refs
    with app.app_context():
        row = draft(DAY)["receipts"][0]
        assert row["material_code"] == "TEST-M" and row["spec_code"] == row["supplier_code"] == ""
        refs["materials"][0]["specs"][1]["supplier_name"] = "每日採購單"
        assert draft(DAY)["receipts"][0]["spec_code"] == "TEST-S2"
        refs["materials"].append(dict(name="高麗菜", code="OTHER-M"))
        row = draft(DAY)["receipts"][0]
        assert row["material_code"] == row["spec_code"] == ""


def test_compressed_private_references_load_without_exposing_unrelated_fields(app, monkeypatch):
    import base64, zlib
    from services.ftrace_references import reference_data
    refs = private_references()
    refs["materials"][0]["specs"][0].update(product_name="測試食材", package_label="1公斤／袋", manufacturer_name="測試製造廠")
    refs["parties"][0]["contact"] = "MUST-NOT-RENDER"
    app.config.pop("FTRACE_REFERENCES_JSON")
    packed = base64.b64encode(zlib.compress(json.dumps(refs).encode())).decode()
    monkeypatch.setenv("FTRACE_REFERENCES_ZLIB_BASE64", packed)
    with app.app_context():
        assert reference_data()["parties"][0]["code"] == "TEST-V"
        assert "contact" not in reference_data()["parties"][0]
        assert reference_data()["materials"][0]["specs"][0]["package_label"] == "1公斤／袋"
        assert draft(DAY)["receipts"][0]["supplier_code"] == "TEST-V"
        monkeypatch.setenv("FTRACE_REFERENCES_ZLIB_BASE64", "invalid")
        assert draft(DAY)["receipts"][0]["material_code"] == ""


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
                            ("桃園市平鎮區新勢國小附設幼兒園", 10),
                            ("桃園市平鎮區新勢國小", 100),
                            ("信義幼兒園", 5), ("桃園市中壢區信義國小", 50)):
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
        assert pingzhen["product_code"] == "TEST-P1"
        assert pingzhen["customer_code"] == "TEST-C1"
        assert pingzhen["kg"] == "72"
        assert len(pingzhen["materials"]) == 1
        assert len(values["meals"]) == 4  # test school + Pingzhen + 2 primaries
        assert len(values["skipped"]) == 3
        xinshi = next(r for r in values["meals"] if "新勢" in r["name"])
        xinyi = next(r for r in values["meals"] if "信義" in r["name"])
        assert (xinshi["quantity"], xinshi["product_code"]) == ("110", "TEST-P2")
        assert (xinyi["quantity"], xinyi["product_code"]) == ("55", "TEST-P3")


def test_product_codes_come_from_private_environment_and_invalid_settings_fail_closed(app, monkeypatch):
    app.config.pop("FTRACE_PRODUCTS_JSON")
    monkeypatch.setenv("FTRACE_PRODUCTS_JSON", json.dumps([
        dict(name="測試國小", code="PRIVATE-P", customer_code="PRIVATE-C", grams=600)]))
    with app.app_context():
        meal = draft(DAY)["meals"][0]
        assert (meal["product_code"], meal["customer_code"], meal["kg"]) == ("PRIVATE-P", "PRIVATE-C", "7.2")
        assert meal["grams_per_portion"] == "600"
        monkeypatch.setenv("FTRACE_PRODUCTS_JSON", "invalid")
        meal = draft(DAY)["meals"][0]
        assert (meal["product_code"], meal["customer_code"], meal["kg"]) == ("", "", "")


@pytest.mark.parametrize("kind", SCHEMAS)
def test_workbook_roundtrip_matches_official_sheets_headers_and_text_cells(kind):
    sheets = build_sheets(kind, DAY, payload())
    content = export_workbook(kind, sheets).getvalue()
    assert content[:8] == bytes.fromhex("d0cf11e0a1b11ae1")
    compound = CompDoc(content)
    assert compound.get_named_stream("\x01CompObj")
    assert compound.get_named_stream("\x01Ole")
    book = xlrd.open_workbook(file_contents=content, formatting_info=True)
    assert book.sheet_names() == list(SCHEMAS[kind])
    for sheet in book.sheets():
        assert sheet.row_values(0) == SCHEMAS[kind][sheet.name]
        assert sheet.nrows == len(sheets[sheet.name]) + 1
        for r in range(1, sheet.nrows):
            for cell in sheet.row(r):
                if cell.value:
                    assert cell.ctype == xlrd.XL_CELL_TEXT
                    xf = book.xf_list[cell.xf_index]
                    assert book.format_map[xf.format_key].format_str == "@"
        if sheet.nrows > 1:
            assert sheet.row_values(1)[:3] == ["2", "2026/11/05", "01"]


def test_excel_envelope_repairs_write_access_and_preserves_large_workbook_cells():
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet("原料資料")
    for index in range(1500):
        sheet.write(index, 0, f"{index:04}-" + "中文原料品名" * 10)
    raw = BytesIO(); workbook.save(raw)
    content = excel_envelope(raw.getvalue())
    before = CompDoc(raw.getvalue()).get_named_stream("Workbook")
    after = CompDoc(content).get_named_stream("Workbook")
    import struct
    offset = 0
    while True:
        opcode, size = struct.unpack_from("<HH", before, offset)
        if opcode == 0x005C:
            break
        offset += 4 + size
    assert size == 112
    assert before[:offset + 4] == after[:offset + 4]
    assert before[offset + 4 + size:] == after[offset + 4 + size:]
    # MS-XLS WriteAccess contains XLUnicodeString, not a bare padded name.
    assert struct.unpack_from("<HB", after, offset + 4) == (2, 0)
    assert after[offset + 7:offset + 9] == b"  "
    read = xlrd.open_workbook(file_contents=content).sheet_by_index(0)
    assert read.nrows == 1500
    assert read.cell_value(1499, 0) == "1499-" + "中文原料品名" * 10


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
    assert len(build_sheets("manufacture", DAY, data)["製造餐食原材料"]) == 1
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
        assert xlrd.open_workbook(file_contents=response.data).sheet_names() == list(SCHEMAS[kind])
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
    assert "ms-excel" not in response.content_type
    client.get(f"/admin/order-tool/ftrace?date={DAY}")
    with client.session_transaction() as session:
        token = session["_kitchen_csrf"]
    response = client.post("/admin/order-tool/ftrace/export", data={
        "date":str(DAY), "kind":"receiving", "payload":json.dumps(payload()),
        "confirmed":"1", "_csrf_token":token})
    assert response.status_code == 200
