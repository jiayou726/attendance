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
                    KitchenRecipeIngredient, KitchenSchool, KitchenSupplier)
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
        ingredient = KitchenIngredient(name="高麗菜", base_unit="g", purchase_unit="kg",
                                       supplier=KitchenSupplier(name="每日採購單"))
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
        assert values["receipts"][0]["kg"] == "0.96"  # (10 + 2) servings * 80 g
        assert len(values["meals"]) == 1
        assert values["meals"][0]["quantity"] == "12"
        assert values["meals"][0]["kg"] == ""
        assert len(values["meals"][0]["materials"]) == 1
        assert not db.session.new and not db.session.dirty and not db.session.deleted
        # The base unit of the menu ingredient, not the purchase snapshot,
        # determines whether the system can estimate kilograms.
        ingredient = KitchenIngredient.query.one()
        ingredient.base_unit = "個"
        db.session.flush()
        assert draft(DAY)["receipts"][0]["kg"] == ""


def test_receipt_uses_served_menu_date_not_purchase_delivery_or_order_state(app):
    with app.app_context():
        item = KitchenPurchaseOrderItem.query.one()
        item.delivery_date = date(2026, 11, 4)
        item.ordered = False
        item.actual_order_qty = Decimal("99")
        item.order.status = "cancelled"
        db.session.flush()
        receipts = draft(DAY)["receipts"]
        assert len(receipts) == 1
        assert receipts[0]["key"] == f"menu:{KitchenIngredient.query.one().id}"
        assert receipts[0]["kg"] == "0.96"
        assert draft(date(2026, 11, 4))["receipts"] == []


def test_receipt_and_manufacturing_skip_unserved_and_cancelled_plans(app):
    with app.app_context():
        plans = KitchenMenuPlan.query.order_by(KitchenMenuPlan.id).all()
        assignment = plans[0].assignments[0]
        assignment.service_status = "no_service"
        db.session.flush()
        assert draft(DAY)["receipts"][0]["kg"] == "0.16"
        assert draft(DAY)["meals"][0]["quantity"] == "2"
        assignment.service_status = "serving"
        plans[1].status = "cancelled"
        db.session.flush()
        assert draft(DAY)["receipts"][0]["kg"] == "0.8"
        assert draft(DAY)["meals"][0]["quantity"] == "10"


def test_receipt_matches_manufacturing_menu_specific_recipe_overrides(app):
    with app.app_context():
        ingredient = KitchenIngredient.query.one()
        first_plan = KitchenMenuPlan.query.order_by(KitchenMenuPlan.id).first()
        first_plan.items[0].ingredients_override_json = json.dumps([
            {"ingredient_id": ingredient.id, "grams_per_person": 150}
        ])
        db.session.flush()
        values = draft(DAY)
        assert values["receipts"][0]["kg"] == "1.66"
        assert values["meals"][0]["materials"][0]["estimated_kg"] == "1.66"


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


def batch_payload():
    value = payload()
    value['batch_mode'] = True
    value['receipts'][0].update(seq='91', batch='TEST-RAW-LOT', manufactured_date='', expiry_date='')
    value['meals'][0].update(seq='92', batch='2026/11/05', manufactured_date='2026-11-05', expiry_date='2026-11-05')
    value['meals'][0]['materials'][0]['lots'] = [dict(report_date='2026-11-04', seq='91',
        spec_code='TEST-RAW-SPEC', batch='TEST-RAW-LOT', received_date='', kg='0.96', receipt_kg='1.235')]
    return value


def test_batch_receipt_and_manufacture_and_delivery_keep_real_links_and_biff(app):
    from services.ftrace_batch import BATCH_SCHEMAS
    value = batch_payload()
    receipt = build_sheets('receiving', DAY, value)
    assert receipt['收貨資料批號資訊'][0] == ['2','2026/11/05','91','B001','B001A025','TEST-RAW-LOT','','','2026/11/05','1.235','','']
    made = build_sheets('manufacture', DAY, value)
    assert made['製造資料'][0][4:7] == ['2026/11/05','2026/11/05','2026/11/05']
    link = made['製造餐食資料原材料供應商'][0]
    assert link[5:] == ['2026','11','04','91','B001','TEST-RAW-SPEC','TEST-RAW-LOT','','0.96']
    delivered = build_sheets('delivery', DAY, value)
    assert delivered['交貨資料'][0][4] == made['製造資料'][0][4]
    assert delivered['交貨資料下游業者'][0][-4:] == ['TEST-C0','7.2','12','']
    for kind,sheets in [('receiving',receipt),('manufacture',made),('delivery',delivered)]:
        book=xlrd.open_workbook(file_contents=export_workbook(kind,sheets,batch_mode=True).getvalue())
        assert book.sheet_names() == list(BATCH_SCHEMAS[kind])
        for sheet in book.sheets():
            assert sheet.row_values(0) == BATCH_SCHEMAS[kind][sheet.name]
            assert all(sheet.ncols == len(row) for row in sheets[sheet.name])
    response=app.test_client().post('/admin/order-tool/ftrace/export',data=dict(date=DAY.isoformat(),kind='manufacture',confirmed='1',payload=json.dumps(value)))
    assert response.status_code == 200
    assert xlrd.open_workbook(file_contents=response.data).sheet_by_name('製造資料').cell_value(1,4)=='2026/11/05'


@pytest.mark.parametrize('case', ['missing_lot','future_lot','invalid_month','two_identifiers','too_much','aggregate_overuse','bad_expiry','receipt_day'])
def test_batch_invalid_or_untraceable_data_is_rejected(case):
    value=batch_payload(); meal=value['meals'][0]; material=meal['materials'][0]; lot=material['lots'][0]
    kind='manufacture'
    if case=='missing_lot': material['lots']=[]
    elif case=='future_lot': lot['report_date']='2026-11-06'
    elif case=='invalid_month': lot['report_date']='2026-13'
    elif case=='two_identifiers': lot['received_date']='2026-11-04'
    elif case=='too_much': lot['kg']='2'
    elif case=='aggregate_overuse':
        import copy
        other=copy.deepcopy(meal);other['seq']='93';other['product_code']='TEST-P9';value['meals'].append(other)
    elif case=='bad_expiry': meal['expiry_date']='2026-11-04'
    elif case=='receipt_day': kind='receiving';value['receipts'][0]['received_date']='2026-11-04'
    with pytest.raises(ExportError):build_sheets(kind,DAY,value)


def test_batch_manufacture_can_link_legacy_stock_or_multiple_receipt_lots():
    value=batch_payload();lot=value['meals'][0]['materials'][0]['lots'][0]
    lot.update(report_date='2026-10',seq='',batch='',received_date='2026-10-30',kg='0.46')
    value['meals'][0]['materials'][0]['lots'].append(dict(report_date='2026-11-05',seq='01',spec_code='TEST-S',batch='TEST-LOT',received_date='',kg='0.5'))
    rows=build_sheets('manufacture',DAY,value)['製造餐食資料原材料供應商']
    assert rows[0][5:9]==['2026','10','',''] and rows[0][-2:]==['2026/10/30','0.46']
    assert len(rows)==2 and rows[1][-1]=='0.5'


def test_recipe_usage_estimate_is_weight_only_and_export_does_not_write(app):
    with app.app_context():
        assert draft(DAY)['meals'][0]['materials'][0]['estimated_kg']=='0.96'
        build_sheets('manufacture',DAY,batch_payload())
        assert not db.session.new and not db.session.dirty


def test_batch_suggestions_follow_existing_material_or_product_code_and_roc_date(app):
    from services.ftrace_batch import suggested_batch
    assert suggested_batch('TEST-M', date(2026,10,8)) == 'TEST-M1151008'
    assert suggested_batch('', DAY) == ''
    assert suggested_batch('X'*40, DAY) == ''
    app.config['FTRACE_REFERENCES_JSON']=private_references()
    with app.app_context():
        value=draft(DAY)
        assert value['receipts'][0]['batch']=='TEST-M1151105'
        assert value['receipts'][0]['seq']=='01'
        assert value['meals'][0]['manufactured_date']==DAY.isoformat()


def test_shared_party_registry_preserves_private_aliases_and_role_scoping(app, monkeypatch):
    from services import ftrace_references as refs
    parties = [
        dict(name="正確供應商有限公司", code="A001", relation="010000", aliases=["每日採購單"]),
        dict(name="平鎮高中", code="AA002", relation="100000", aliases=[]),
        dict(name="琮元寶實業有限公司", code="A021", relation="010000", aliases=[]),
        dict(name="琮元寶實業有限公司", code="A038", relation="010000", aliases=[]),
    ]
    monkeypatch.setattr(refs, "_db_parties", lambda: parties)
    with app.app_context():
        app.config["FTRACE_REFERENCES_JSON"] = dict(materials=[], products=[],
            parties=[dict(name="舊供應商名稱", code="A001", aliases=["舊別名"])])
        references = refs.reference_data()
        assert len(references["parties"]) == 4
        matched = [p for p in references["parties"] if p["code"] == "A001"]
        assert len(matched) == 1 and "舊別名" in matched[0]["aliases"]
        data = dict(receipts=[
            dict(name="食材", supplier_name="每日採購單", supplier_code=""),
            dict(name="食材", supplier_name="平鎮高中", supplier_code=""),
            dict(name="食材", supplier_name="琮元寶實業有限公司", supplier_code=""),
        ], meals=[dict(name="平鎮高中（含小便當）", school_key="平鎮高中",
                      customer_code="", materials=[])])
        refs.apply_reference_defaults(data, references)
        assert [r["supplier_code"] for r in data["receipts"]] == ["A001", "", ""]
        assert data["meals"][0]["customer_code"] == "AA002"
        data["meals"][0]["customer_code"] = "已人工確認的客戶碼"
        refs.apply_reference_defaults(data, references)
        assert data["meals"][0]["customer_code"] == "已人工確認的客戶碼"
