from datetime import date
from decimal import Decimal

import pytest

from app import create_app
from extensions import db
from models import (
    Checkin,
    Employee,
    KitchenIngredient,
    KitchenMenuAssignment,
    KitchenMenuPlan,
    KitchenPurchaseOrder,
    KitchenPurchaseOrderItem,
    KitchenRecipe,
    KitchenRecipeIngredient,
    KitchenSchool,
    KitchenSupplier,
)

TEST_DAY = date(2026, 8, 13)


@pytest.fixture()
def app(tmp_path):
    db_path = tmp_path / "kitchen_test.db"
    app = create_app({
        "TESTING": True,
        "SECRET_KEY": "test-only-secret",
        "SQLALCHEMY_DATABASE_URI": f"sqlite:///{db_path}",
        "AUTO_CREATE_DB": True,
        "KITCHEN_CSRF_ENABLED": False,
        "ADMIN_HR_PASSWORD": "test-hr-password",
        "ADMIN_MGR_PASSWORD": "test-mgr-password",
        "PRODUCTION": False,
    })
    yield app
    with app.app_context():
        db.session.remove()
        db.drop_all()


@pytest.fixture()
def client(app):
    return app.test_client()


@pytest.fixture()
def authed_client(client):
    with client.session_transaction() as sess:
        sess["role"] = "mgr"
    return client


def _ids(app):
    with app.app_context():
        return {
            "school": KitchenSchool.query.filter_by(name="內小").one().id,
            "supplier": KitchenSupplier.query.filter_by(name="測試肉品").one().id,
            "ingredient": KitchenIngredient.query.filter_by(name="骨腿丁").one().id,
            "recipe": KitchenRecipe.query.filter_by(name="南洋綠咖哩雞").one().id,
        }


def _ids_partial(app):
    with app.app_context():
        school = KitchenSchool.query.filter_by(name="內小").one_or_none()
        supplier = KitchenSupplier.query.filter_by(name="測試肉品").one_or_none()
        return {
            "school": school.id if school else None,
            "supplier": supplier.id if supplier else None,
        }


def _seed_core_via_routes(app, client):
    assert client.post("/admin/order-tool/schools", data={"name": "內小", "code": "1-08"}).status_code == 302
    assert client.post("/admin/order-tool/suppliers", data={"name": "測試肉品", "phone": "02-12345678"}).status_code == 302
    ids = _ids_partial(app)

    assert client.post("/admin/order-tool/ingredients", data={
        "name": "骨腿丁",
        "supplier_id": str(ids["supplier"]),
        "purchase_unit": "kg",
        "grams_per_purchase_unit": "1000",
        "unit_price": "82",
        "order_increment": "0.001",
        "note": "",
    }).status_code == 302

    assert client.post("/admin/order-tool/recipes", data={
        "name": "南洋綠咖哩雞",
        "category": "主菜",
        "serving_output_g": "95",
        "note": "",
    }).status_code == 302

    ids = _ids(app)
    assert client.post(f"/admin/order-tool/recipes/{ids['recipe']}/ingredients", data={
        "ingredient_id": str(ids["ingredient"]),
        "grams_per_person": "88",
    }).status_code == 302
    return ids


def _create_plan(app, client, ids, service_date="2026-08-13", headcount=801):
    assert client.post("/admin/order-tool/plans", data={
        "service_date": service_date,
        "meal_type": "午餐",
        "name": "中央菜單",
        "note": "",
    }).status_code == 302
    service_day = date.fromisoformat(service_date)
    with app.app_context():
        plan = KitchenMenuPlan.query.filter_by(service_date=service_day, meal_type="午餐", name="中央菜單").one()
        plan_id = plan.id
    assert client.post(f"/admin/order-tool/plans/{plan_id}/items", data={"recipe_id": str(ids["recipe"])}).status_code == 302
    assert client.post(f"/admin/order-tool/plans/{plan_id}/assignments", data={
        "school_id": str(ids["school"]),
        "headcount": str(headcount),
    }).status_code == 302
    return plan_id


def test_kitchen_requires_login(client):
    response = client.get("/admin/order-tool/", follow_redirects=False)
    assert response.status_code == 302
    assert "/admin/login" in response.headers["Location"]


def test_static_pages_render_and_attendance_models_survive(app, authed_client):
    with app.app_context():
        db.session.add(Employee(id=9001, name="原打卡員工", area="A", default_break=0.5))
        db.session.add(Checkin(employee_id=9001, work_date="2026-08-13", p_type="in", ts="2026-08-13 08:00:00"))
        db.session.commit()

    for path in (
        "/admin/order-tool/",
        "/admin/order-tool/schools",
        "/admin/order-tool/suppliers",
        "/admin/order-tool/ingredients",
        "/admin/order-tool/recipes",
        "/admin/order-tool/plans",
        "/admin/order-tool/purchases",
        "/admin/order-tool/purchase",
    ):
        response = authed_client.get(path)
        assert response.status_code in (200, 302), path

    response = authed_client.get("/admin/order-tool/")
    assert b"kitchen_mobile.css" in response.data

    with app.app_context():
        assert db.session.get(Employee, 9001).name == "原打卡員工"
        assert Checkin.query.filter_by(employee_id=9001).count() == 1


def test_full_801_person_purchase_calculation_and_snapshot(app, authed_client):
    ids = _seed_core_via_routes(app, authed_client)
    plan_id = _create_plan(app, authed_client, ids, headcount=801)

    response = authed_client.post("/admin/order-tool/purchases/generate", data={
        "start": "2026-08-13",
        "end": "2026-08-13",
    })
    assert response.status_code == 302

    with app.app_context():
        order = KitchenPurchaseOrder.query.filter_by(service_date=TEST_DAY).one()
        item = order.items[0]
        assert item.required_grams == Decimal("70488.000")
        assert item.required_qty == Decimal("70.4880")
        assert item.recommended_order_qty == Decimal("70.4880")
        assert item.actual_order_qty == Decimal("70.4880")
        assert item.unit_price_snapshot == Decimal("82.0000")
        assert item.amount == Decimal("5780.0160")
        order_id = order.id

    detail = authed_client.get(f"/admin/order-tool/purchases/{order_id}")
    assert detail.status_code == 200
    assert b"70.488" in detail.data
    assert b"5780.02" in detail.data

    assert authed_client.post(f"/admin/order-tool/purchases/{order_id}/confirm").status_code == 302

    # 修改 master data 後，已確認歷史採購單必須保持原 snapshot。
    with app.app_context():
        ingredient = db.session.get(KitchenIngredient, ids["ingredient"])
        ingredient.unit_price = Decimal("90")
        component = KitchenRecipeIngredient.query.filter_by(recipe_id=ids["recipe"], ingredient_id=ids["ingredient"]).one()
        component.grams_per_person = Decimal("90")
        db.session.commit()

        order = db.session.get(KitchenPurchaseOrder, order_id)
        item = order.items[0]
        assert order.status == "confirmed"
        assert item.required_grams == Decimal("70488.000")
        assert item.unit_price_snapshot == Decimal("82.0000")
        assert item.amount == Decimal("5780.0160")
        assert db.session.get(KitchenMenuPlan, plan_id).status == "confirmed"

    authed_client.post("/admin/order-tool/purchases/generate", data={"start": "2026-08-13", "end": "2026-08-13"})
    with app.app_context():
        item = KitchenPurchaseOrder.query.filter_by(service_date=TEST_DAY).one().items[0]
        assert item.required_grams == Decimal("70488.000")
        assert item.unit_price_snapshot == Decimal("82.0000")


def test_cross_school_aggregation(app, authed_client):
    ids = _seed_core_via_routes(app, authed_client)
    plan_id = _create_plan(app, authed_client, ids, headcount=801)

    authed_client.post("/admin/order-tool/schools", data={"name": "文山", "code": "1-21"})
    with app.app_context():
        second_school_id = KitchenSchool.query.filter_by(name="文山").one().id
    authed_client.post(f"/admin/order-tool/plans/{plan_id}/assignments", data={
        "school_id": str(second_school_id),
        "headcount": "586",
    })
    authed_client.post("/admin/order-tool/purchases/generate", data={"start": "2026-08-13", "end": "2026-08-13"})

    with app.app_context():
        item = KitchenPurchaseOrder.query.filter_by(service_date=TEST_DAY).one().items[0]
        assert item.required_grams == Decimal("122056.000")
        assert item.required_qty == Decimal("122.0560")


def test_supplier_grouping(app, authed_client):
    ids = _seed_core_via_routes(app, authed_client)
    _create_plan(app, authed_client, ids, headcount=801)

    authed_client.post("/admin/order-tool/suppliers", data={"name": "測試蔬菜"})
    with app.app_context():
        veg_supplier = KitchenSupplier.query.filter_by(name="測試蔬菜").one()
        veg_supplier_id = veg_supplier.id
    authed_client.post("/admin/order-tool/ingredients", data={
        "name": "洋芋",
        "supplier_id": str(veg_supplier_id),
        "purchase_unit": "kg",
        "grams_per_purchase_unit": "1000",
        "unit_price": "20",
        "order_increment": "0.001",
    })
    with app.app_context():
        potato_id = KitchenIngredient.query.filter_by(name="洋芋").one().id
    authed_client.post(f"/admin/order-tool/recipes/{ids['recipe']}/ingredients", data={
        "ingredient_id": str(potato_id),
        "grams_per_person": "8",
    })
    authed_client.post("/admin/order-tool/purchases/generate", data={"start": "2026-08-13", "end": "2026-08-13"})

    with app.app_context():
        orders = KitchenPurchaseOrder.query.filter_by(service_date=TEST_DAY).all()
        assert len(orders) == 2
        assert {o.supplier_name_snapshot for o in orders} == {"測試肉品", "測試蔬菜"}
        potato_order = next(o for o in orders if o.supplier_name_snapshot == "測試蔬菜")
        assert potato_order.items[0].required_grams == Decimal("6408.000")


def test_actual_purchase_qty_and_price_are_persisted(app, authed_client):
    ids = _seed_core_via_routes(app, authed_client)
    _create_plan(app, authed_client, ids)
    authed_client.post("/admin/order-tool/purchases/generate", data={"start": "2026-08-13", "end": "2026-08-13"})

    with app.app_context():
        order = KitchenPurchaseOrder.query.one()
        item = order.items[0]
        order_id, item_id = order.id, item.id

    authed_client.post(f"/admin/order-tool/purchase-items/{item_id}/update", data={
        "actual_order_qty": "72",
        "unit_price": "83.5",
        "note": "人工調整",
    })

    with app.app_context():
        item = db.session.get(KitchenPurchaseOrderItem, item_id)
        assert item.actual_order_qty == Decimal("72.0000")
        assert item.unit_price_snapshot == Decimal("83.5000")
        assert item.amount == Decimal("6012.0000")
        assert item.note == "人工調整"
        assert db.session.get(KitchenPurchaseOrder, order_id).status == "draft"


def test_invalid_negative_values_do_not_mutate(app, authed_client):
    ids = _seed_core_via_routes(app, authed_client)
    plan_id = _create_plan(app, authed_client, ids, headcount=801)
    with app.app_context():
        assignment = KitchenMenuAssignment.query.filter_by(plan_id=plan_id, school_id=ids["school"]).one()
        assignment_id = assignment.id

    authed_client.post(f"/admin/order-tool/assignments/{assignment_id}/update", data={"headcount": "-5"})
    with app.app_context():
        assert db.session.get(KitchenMenuAssignment, assignment_id).headcount == 801

    response = authed_client.post("/admin/order-tool/ingredients", data={
        "name": "錯誤食材",
        "purchase_unit": "kg",
        "grams_per_purchase_unit": "1000",
        "unit_price": "-1",
        "order_increment": "0.001",
    })
    assert response.status_code == 302
    with app.app_context():
        assert KitchenIngredient.query.filter_by(name="錯誤食材").first() is None


def test_csrf_protects_kitchen_post(tmp_path):
    db_path = tmp_path / "csrf.db"
    app = create_app({
        "TESTING": True,
        "SECRET_KEY": "csrf-test-secret",
        "SQLALCHEMY_DATABASE_URI": f"sqlite:///{db_path}",
        "AUTO_CREATE_DB": True,
        "KITCHEN_CSRF_ENABLED": True,
        "PRODUCTION": False,
    })
    client = app.test_client()
    with client.session_transaction() as sess:
        sess["role"] = "mgr"

    client.post("/admin/order-tool/schools", data={"name": "不該成功"})
    with app.app_context():
        assert KitchenSchool.query.filter_by(name="不該成功").first() is None

    client.get("/admin/order-tool/schools")
    with client.session_transaction() as sess:
        token = sess["_kitchen_csrf"]
    client.post("/admin/order-tool/schools", data={"name": "合法學校", "_csrf_token": token})
    with app.app_context():
        assert KitchenSchool.query.filter_by(name="合法學校").one()
        db.drop_all()


def test_file_sqlite_persists_across_app_restart(tmp_path):
    db_path = tmp_path / "restart.db"
    uri = f"sqlite:///{db_path}"
    config = {
        "TESTING": True,
        "SECRET_KEY": "restart-test",
        "SQLALCHEMY_DATABASE_URI": uri,
        "AUTO_CREATE_DB": True,
        "KITCHEN_CSRF_ENABLED": False,
        "PRODUCTION": False,
    }
    app1 = create_app(config)
    with app1.app_context():
        db.session.add(KitchenSchool(name="重啟後還在"))
        db.session.commit()
        db.session.remove()

    app2 = create_app({**config, "AUTO_CREATE_DB": False})
    with app2.app_context():
        assert KitchenSchool.query.filter_by(name="重啟後還在").one()
        db.drop_all()
