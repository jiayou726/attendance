from datetime import date
from decimal import Decimal

import pytest
from sqlalchemy import inspect, text

from app import create_app
from extensions import db
from models import (
    KitchenIngredient,
    KitchenMenuAssignment,
    KitchenMenuPlan,
    KitchenPurchaseOrder,
    KitchenRecipe,
    KitchenRecipeIngredient,
    KitchenSchool,
    KitchenSupplier,
)


@pytest.fixture()
def app(tmp_path):
    app = create_app({
        "TESTING": True,
        "SECRET_KEY": "guardrail-test",
        "SQLALCHEMY_DATABASE_URI": f"sqlite:///{tmp_path / 'guardrail.db'}",
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
    with client.session_transaction() as sess:
        sess["role"] = "mgr"
    return client


def test_app_adds_missing_school_service_status_without_full_migration(tmp_path):
    import sqlite3

    db_path = tmp_path / "legacy_kitchen.db"
    connection = sqlite3.connect(db_path)
    connection.execute("""
        CREATE TABLE kitchen_menu_assignment (
            id INTEGER PRIMARY KEY,
            plan_id INTEGER NOT NULL,
            school_id INTEGER NOT NULL,
            headcount INTEGER NOT NULL DEFAULT 0
        )
    """)
    connection.execute(
        "INSERT INTO kitchen_menu_assignment (id, plan_id, school_id, headcount) VALUES (1, 1, 1, 100)"
    )
    connection.commit()
    connection.close()

    legacy_app = create_app({
        "TESTING": True,
        "SECRET_KEY": "schema-compatibility-test",
        "SQLALCHEMY_DATABASE_URI": f"sqlite:///{db_path}",
        "AUTO_CREATE_DB": False,
        "KITCHEN_CSRF_ENABLED": False,
        "PRODUCTION": False,
    })
    with legacy_app.app_context():
        columns = {
            column["name"]
            for column in inspect(db.engine).get_columns("kitchen_menu_assignment")
        }
        status = db.session.execute(text(
            "SELECT service_status FROM kitchen_menu_assignment WHERE id = 1"
        )).scalar_one()
    assert "service_status" in columns
    assert status == "serving"


def _seed_g_recipe(app, client, service_date="2026-08-13"):
    client.post("/admin/order-tool/schools", data={"name": "內小"})
    client.post("/admin/order-tool/suppliers", data={"name": "測試肉品"})
    with app.app_context():
        school = KitchenSchool.query.filter_by(name="內小").one()
        supplier = KitchenSupplier.query.filter_by(name="測試肉品").one()
        school_id, supplier_id = school.id, supplier.id

    client.post("/admin/order-tool/ingredients", data={
        "name": "骨腿丁", "supplier_id": supplier_id, "base_unit": "g",
        "purchase_unit": "kg", "grams_per_purchase_unit": "1000",
        "unit_price": "82", "order_increment": "0.001",
    })
    client.post("/admin/order-tool/recipes", data={"name": "南洋綠咖哩雞", "category": "主菜", "serving_output_g": "95"})
    with app.app_context():
        ingredient = KitchenIngredient.query.filter_by(name="骨腿丁").one()
        recipe = KitchenRecipe.query.filter_by(name="南洋綠咖哩雞").one()
        ingredient_id, recipe_id = ingredient.id, recipe.id

    client.post(f"/admin/order-tool/recipes/{recipe_id}/ingredients", data={"ingredient_id": ingredient_id, "grams_per_person": "88"})
    client.post("/admin/order-tool/plans", data={"service_date": service_date, "meal_type": "午餐", "name": "中央菜單"})
    with app.app_context():
        plan = KitchenMenuPlan.query.filter_by(service_date=date.fromisoformat(service_date), name="中央菜單").one()
        plan_id = plan.id
    client.post(f"/admin/order-tool/plans/{plan_id}/items", data={"recipe_id": recipe_id})
    client.post(f"/admin/order-tool/plans/{plan_id}/assignments", data={"school_id": school_id, "headcount": "801"})
    return {"school": school_id, "supplier": supplier_id, "ingredient": ingredient_id, "recipe": recipe_id, "plan": plan_id}


def test_piece_based_recipe_calculates_boxes(app, client):
    client.post("/admin/order-tool/schools", data={"name": "內小"})
    client.post("/admin/order-tool/suppliers", data={"name": "雞肉商"})
    with app.app_context():
        school_id = KitchenSchool.query.filter_by(name="內小").one().id
        supplier_id = KitchenSupplier.query.filter_by(name="雞肉商").one().id

    client.post("/admin/order-tool/ingredients", data={
        "name": "棒棒腿", "supplier_id": supplier_id, "base_unit": "個",
        "purchase_unit": "箱", "grams_per_purchase_unit": "50",
        "unit_price": "1000", "order_increment": "1",
    })
    client.post("/admin/order-tool/recipes", data={"name": "炸棒棒腿", "category": "主菜"})
    with app.app_context():
        ingredient_id = KitchenIngredient.query.filter_by(name="棒棒腿").one().id
        recipe_id = KitchenRecipe.query.filter_by(name="炸棒棒腿").one().id

    client.post(f"/admin/order-tool/recipes/{recipe_id}/ingredients", data={"ingredient_id": ingredient_id, "grams_per_person": "1"})
    client.post("/admin/order-tool/plans", data={"service_date": "2026-08-14", "meal_type": "午餐", "name": "中央菜單"})
    with app.app_context():
        plan_id = KitchenMenuPlan.query.filter_by(service_date=date(2026, 8, 14)).one().id
    client.post(f"/admin/order-tool/plans/{plan_id}/items", data={"recipe_id": recipe_id})
    client.post(f"/admin/order-tool/plans/{plan_id}/assignments", data={"school_id": school_id, "headcount": "801"})
    client.post("/admin/order-tool/purchases/generate", data={"start": "2026-08-14", "end": "2026-08-14"})

    with app.app_context():
        item = KitchenPurchaseOrder.query.filter_by(service_date=date(2026, 8, 14)).one().items[0]
        assert item.base_unit_snapshot == "個"
        assert item.required_grams == Decimal("801.000")
        assert item.required_qty == Decimal("16.0200")
        assert item.recommended_order_qty == Decimal("17.0000")
        assert item.actual_order_qty == Decimal("17.0000")
        assert item.amount == Decimal("17000.0000")


def test_same_school_same_day_same_meal_cannot_be_double_assigned(app, client):
    ids = _seed_g_recipe(app, client)
    client.post("/admin/order-tool/plans", data={"service_date": "2026-08-13", "meal_type": "午餐", "name": "第二張菜單"})
    with app.app_context():
        second = KitchenMenuPlan.query.filter_by(service_date=date(2026, 8, 13), name="第二張菜單").one()
        second_id = second.id
    client.post(f"/admin/order-tool/plans/{second_id}/assignments", data={"school_id": ids["school"], "headcount": "801"})

    with app.app_context():
        count = (
            KitchenMenuAssignment.query.join(KitchenMenuPlan)
            .filter(
                KitchenMenuAssignment.school_id == ids["school"],
                KitchenMenuPlan.service_date == date(2026, 8, 13),
                KitchenMenuPlan.meal_type == "午餐",
            ).count()
        )
        assert count == 1


def test_manual_purchase_override_survives_regeneration(app, client):
    _seed_g_recipe(app, client)
    client.post("/admin/order-tool/purchases/generate", data={"start": "2026-08-13", "end": "2026-08-13"})
    with app.app_context():
        order = KitchenPurchaseOrder.query.one()
        item = order.items[0]
        item_id = item.id

    client.post(f"/admin/order-tool/purchase-items/{item_id}/update", data={
        "actual_order_qty": "72", "unit_price": "83.5", "note": "人工調整",
    })
    client.post("/admin/order-tool/purchases/generate", data={"start": "2026-08-13", "end": "2026-08-13"})

    with app.app_context():
        order = KitchenPurchaseOrder.query.one()
        item = order.items[0]
        assert item.id == item_id
        assert item.required_qty == Decimal("70.4880")
        assert item.recommended_order_qty == Decimal("70.4880")
        assert item.actual_order_qty == Decimal("72.0000")
        assert item.unit_price_snapshot == Decimal("83.5000")
        assert item.amount == Decimal("6012.0000")
        assert item.note == "人工調整"
        assert item.manual_override is True
