import io
import re

import pytest
from openpyxl import Workbook

from app import create_app
from extensions import db
from models import KitchenRecipe, KitchenRecipeIngredient


@pytest.fixture()
def app(tmp_path):
    app = create_app({
        "TESTING": True,
        "SECRET_KEY": "catalog-import-test",
        "SQLALCHEMY_DATABASE_URI": f"sqlite:///{tmp_path / 'catalog.db'}",
        "AUTO_CREATE_DB": True,
        "KITCHEN_CSRF_ENABLED": False,
        "PRODUCTION": False,
    })
    yield app
    with app.app_context():
        db.drop_all()


@pytest.fixture()
def client(app):
    client = app.test_client()
    with client.session_transaction() as session:
        session["role"] = "mgr"
    return client


def _xlsx():
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "匯入資料"
    sheet.append(["菜色名稱", "分類", "材料名稱", "每人用量", "基本單位", "採購單位", "1採購單位換算", "單價", "廠商", "備註"])
    sheet.append(["已存在菜", "主菜", "舊材料", 99, "g", "kg", 1000, 10, "舊廠商", "應跳過"])
    sheet.append(["新菜", "主菜", "雞丁", 45, "g", "kg", 1000, 82, "測試肉品", "第一筆"])
    sheet.append(["新菜", "主菜", "洋蔥", 12, "g", "kg", 1000, 25, "測試蔬菜", "第二筆"])
    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    return output


def test_catalog_template_and_preview_then_apply_skip_duplicates(app, client):
    assert client.get("/admin/order-tool/catalog-template.xlsx").status_code == 200
    with app.app_context():
        db.session.add(KitchenRecipe(name="已存在菜", category="主菜", note="原資料"))
        db.session.commit()

    response = client.post(
        "/admin/order-tool/catalog-import",
        data={"file": (_xlsx(), "總表.xlsx")},
        content_type="multipart/form-data",
    )
    assert response.status_code == 200
    assert "可以新增".encode() in response.data
    assert "新菜".encode() in response.data
    assert "重複跳過".encode() in response.data
    match = re.search(rb'name="token" value="([^"]+)"', response.data)
    assert match

    applied = client.post("/admin/order-tool/catalog-import/apply", data={"token": match.group(1).decode()})
    assert applied.status_code == 302
    with app.app_context():
        assert KitchenRecipe.query.filter_by(name="已存在菜").one().note == "原資料"
        new_recipe = KitchenRecipe.query.filter_by(name="新菜").one()
        assert len(new_recipe.ingredients) == 2
        assert {row.ingredient.name for row in new_recipe.ingredients} == {"雞丁", "洋蔥"}
        assert KitchenRecipeIngredient.query.filter_by(recipe_id=new_recipe.id, quantity_status="manual").count() == 2
