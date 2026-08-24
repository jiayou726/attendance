from datetime import date

import pytest
from openpyxl import Workbook

from app import create_app
from extensions import db
from models import KitchenSupplier, KitchenSupplierItem
from scripts.import_kitchen_history import (
    clean_supplier_product_name,
    extract_package_conversion,
    read_supplier_orders,
)


def test_historical_xlsx_rows_extract_item_unit_price_and_date(tmp_path):
    path = tmp_path / "龍軒行.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["114年", "品名", "數量", "單位", "單價", "金額"])
    sheet.append(["9/2", "芋大丁", 250, "斤", 23, 5750])
    sheet.append([None, "芋絲", 30, "斤", 23, 690])
    sheet.append(["9/12", "洋芋", 200, "斤", 21, 4200])
    workbook.save(path)

    rows = read_supplier_orders(path)
    assert [(row["name"], row["unit"]) for row in rows] == [("芋大丁", "斤"), ("芋絲", "斤"), ("洋芋", "斤")]
    assert rows[0]["unit_price"] == 23
    assert rows[0]["purchase_date"] == date(2025, 9, 2)
    assert rows[1]["purchase_date"] == date(2025, 9, 2)


@pytest.mark.parametrize(("raw", "expected"), [
    ("棒棒腿 5 A 15kg IQF", "棒棒腿"),
    ("棒棒腿 518kg", "棒棒腿"),
    ("TS4 18KG", "棒棒腿"),
    ("雞排4", "雞排"),
    ("火腿丁-東豪", "火腿丁"),
    ("棒5", "棒棒腿"),
    ("5棒進口", "棒棒腿"),
    ("檸檬雞翅 W7", "檸檬雞翅"),
    ("雞胸丁 3包/件", "雞胸丁"),
    ("菲力雞排80\"", "菲力雞排"),
    ("照燒里肌8", "照燒里肌"),
    ("日式8", "日式豬排"),
])
def test_supplier_product_name_removes_package_noise(raw, expected):
    assert clean_supplier_product_name(raw) == expected


@pytest.mark.parametrize(("raw", "unit", "expected"), [
    ("火腿丁(15kg/箱)", "箱", "1箱＝15kg"),
    ("香讚雞堡(200片/箱)", "箱", "1箱＝200片"),
    ("香讚雞塊(4包/件)", "箱", "1件＝4包"),
    ("卡啦丁1kg*12包", "箱", "1箱＝12包（每包1kg）"),
    ("卡啦棒腿 9.2元*125支", "箱", "1箱＝125支"),
    ("烤麩(1包約130粒)", "包", "1包≈130粒"),
    ("仙草(5KG)", "桶", "1桶＝5kg"),
    ("普通雞塊", "箱", ""),
])
def test_extract_package_conversion(raw, unit, expected):
    assert extract_package_conversion(raw, unit) == expected


@pytest.fixture()
def app(tmp_path):
    app = create_app({
        "TESTING": True,
        "SECRET_KEY": "supplier-catalog-test",
        "SQLALCHEMY_DATABASE_URI": f"sqlite:///{tmp_path / 'supplier.db'}",
        "AUTO_CREATE_DB": True,
        "KITCHEN_CSRF_ENABLED": False,
        "PRODUCTION": False,
    })
    yield app
    with app.app_context():
        db.drop_all()


def test_supplier_detail_shows_catalog_and_historical_count(app):
    client = app.test_client()
    with client.session_transaction() as session:
        session["role"] = "mgr"
    with app.app_context():
        supplier = KitchenSupplier(name="龍軒行")
        db.session.add(supplier)
        db.session.flush()
        db.session.add(KitchenSupplierItem(
            supplier_id=supplier.id,
            name="洋芋",
            unit="斤",
            package_conversion="1件＝30斤",
            last_quantity=200,
            last_unit_price=21,
            last_purchase_date=date(2025, 9, 12),
            order_count=5,
        ))
        db.session.commit()
        supplier_id = supplier.id
        item_id = supplier.items[0].id

    response = client.get(f"/admin/order-tool/suppliers/{supplier_id}")
    assert response.status_code == 200
    assert "洋芋".encode() in response.data
    assert "包裝換算".encode() in response.data
    assert "1件＝30斤".encode() in response.data
    assert "2025-09-12".encode() in response.data
    assert ">5<".encode() in response.data

    updated = client.post(f"/admin/order-tool/supplier-items/{item_id}/update", data={
        "name": "洋芋大丁", "unit": "斤", "package_conversion": "1件＝25斤", "last_unit_price": "24",
    })
    assert updated.status_code == 302
    with app.app_context():
        item = db.session.get(KitchenSupplierItem, item_id)
        assert item.name == "洋芋大丁"
        assert item.package_conversion == "1件＝25斤"
        assert float(item.last_unit_price) == 24

    deleted = client.post(f"/admin/order-tool/supplier-items/{item_id}/delete")
    assert deleted.status_code == 302
    with app.app_context():
        item = db.session.get(KitchenSupplierItem, item_id)
        assert item is not None
        assert item.active is False
        assert item.manual_override is True

    created = client.post(f"/admin/order-tool/suppliers/{supplier_id}/items", data={
        "name": "紅蘿蔔", "unit": "斤", "package_conversion": "1件＝30斤", "last_unit_price": "18.5",
    })
    assert created.status_code == 302
    with app.app_context():
        item = KitchenSupplierItem.query.filter_by(supplier_id=supplier_id, name="紅蘿蔔").one()
        assert item.unit == "斤"
        assert item.package_conversion == "1件＝30斤"
        assert float(item.last_unit_price) == 18.5
        assert item.active is True
        assert item.manual_override is True
