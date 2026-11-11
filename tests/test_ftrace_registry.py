"""Read-only unified receipt/manufacturing/delivery records."""
import json
from datetime import date
from types import SimpleNamespace

from services.ftrace_registry import receipt_registry, meal_registry
from tests.test_ftrace_export import app


def imported(id=1, kg="20", issues=None):
    return SimpleNamespace(id=id, report_date="2026-10-05", material_code="R01",
        name="白米", disabled=False, payload_json=json.dumps({
            "name":"白米", "report_date":"2026-10-05", "seq":"01",
            "material_code":"R01","spec_code":"S01", "batch":"LOT-1",
            "received_date":"","receipt_kg":kg,"issues":issues or []
        }))


def document(kind="receiving", status="exported", kg="20", batch_mode=True, day=date(2026,10,5), id=8):
    receipt={"name":"白米","include":True,"seq":"01","material_code":"R01",
             "spec_code":"S01","batch":"LOT-1","received_date":"2026-10-05","kg":kg}
    meal={"include":True,"key":"school-1","name":"學校午餐","product_code":"P01",
          "batch":"FOOD-1","seq":"01","quantity":"100","kg":"50",
          "manufactured_date":"2026-10-05","expiry_date":"2026-10-05",
          "customer_code":"C01","materials":[{"name":"白米","code":"R01","lots":[{
              "report_date":"2026-10-05","spec_code":"S01","batch":"LOT-1","kg":"20"
          }]}]}
    return SimpleNamespace(id=id, kind=kind, status=status, day=day,
                           payload_json=json.dumps({"batch_mode":batch_mode,
                                                   "receipts":[receipt],"meals":[meal]}))


def test_receipt_registry_merges_same_identity_but_preserves_both_provenances():
    results=receipt_registry([imported()], [document()])
    assert len(results)==1
    row=results[0]
    assert row["duplicate"]
    assert row["document_ids"]==[8] and row["imported_ids"]==[1]
    assert row["source_label"]=="歷史 XLS、共用收貨"
    assert row["weight"]=="20"
    assert not row["conflict"]
    assert "平台未確認" in row["state"]


def test_conflicting_weights_always_visible_not_overwritten():
    row=receipt_registry([imported(kg="20")],[document(kg="25")])[0]
    assert row["conflict"] and row["level"]=="warn"
    assert "重量不一致" in row["state"]
    assert row["weight"]=="20／25"


def test_incomplete_receipt_identities_do_not_get_silently_combined():
    a,b=imported(1),imported(2)
    for record in (a,b):
        lot=json.loads(record.payload_json)
        lot["batch"]=""
        lot["received_date"]=""
        record.payload_json=json.dumps(lot)
    assert len(receipt_registry([a,b],[]))==2


def test_imported_history_does_not_imply_confirmed_platform_submission():
    row=receipt_registry([imported()],[])[0]
    assert "平台未查證" in row["state"]
    submitted=receipt_registry([],[document(status="submitted")])[0]
    assert "人工確認" in submitted["state"]


def test_manufacture_and_delivery_are_separate_and_delivery_has_comparison():
    make=document(kind="manufacture",id=4)
    ship=document(kind="delivery",id=5)
    assert len(meal_registry([make,ship],"manufacture"))==1
    passed=meal_registry([make,ship],"delivery")
    assert len(passed)==1 and passed[0]["level"]=="ok"
    assert passed[0]["document_id"]==5


def test_delivery_has_visible_alert_when_manufacture_missing_or_mode_differs():
    ship=document(kind="delivery")
    issues=meal_registry([ship],"delivery")
    assert issues[0]["level"]=="warn"
    assert "無同日製造" in issues[0]["state"]
    no_batch=document(kind="manufacture",batch_mode=False)
    issues=meal_registry([no_batch,ship],"delivery")
    assert "批號模式不同" in issues[0]["state"]


def test_delivery_detects_over_ship_without_changing_payload():
    make=document(kind="manufacture")
    ship=document(kind="delivery")
    payload=json.loads(ship.payload_json)
    payload["meals"][0]["quantity"]="120"
    ship.payload_json=json.dumps(payload)
    issues=meal_registry([make,ship],"delivery")
    assert "交貨份數大於製造份數" in issues[0]["state"]
    assert json.loads(ship.payload_json)["meals"][0]["quantity"]=="120"


def test_registry_page_groups_tabs_and_imports_without_mutating_data(app):
    client=app.test_client()
    for kind in ("receiving","manufacture","delivery"):
        response=client.get("/admin/order-tool/ftrace/records?kind="+kind)
        assert response.status_code==200
        html=response.get_data(as_text=True)
        assert "收貨、製造、交貨紀錄" in html
        assert "本系統的申報紀錄" in html
        if kind == "receiving":
            assert "匯入以前的收貨 XLS" in html
        else:
            assert "交貨" in html
    assert client.get("/admin/order-tool/ftrace/records?kind=receiving&q=UNKNOWN").status_code==200
