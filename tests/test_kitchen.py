from datetime import date, timedelta
from decimal import Decimal
from io import BytesIO

import pytest
from openpyxl import Workbook

from app import create_app
from extensions import db
from models import (
    Checkin,
    Employee,
    KitchenIngredient,
    KitchenMenuAssignment,
    KitchenMenuPlan,
    KitchenMenuPlanItem,
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


def test_kitchen_does_not_require_login_but_other_admin_pages_do(client):
    response = client.get("/admin/order-tool/", follow_redirects=False)
    assert response.status_code == 200

    protected = client.get("/admin/not-a-kitchen-page", follow_redirects=False)
    assert protected.status_code == 302
    assert "/admin/login" in protected.headers["Location"]


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


def test_summary_is_a_monday_to_sunday_grid_and_can_add_a_dish(app, authed_client):
    ids = _seed_core_via_routes(app, authed_client)

    response = authed_client.get("/admin/order-tool/summary?week=2026-08-13")
    assert response.status_code == 200
    page = response.get_data(as_text=True)
    assert "每週菜單總表" in page
    assert all(label in page for label in ("週一", "週二", "週三", "週四", "週五", "週六", "週日"))
    assert "08/10" in page and "08/16" in page
    assert "搜尋或輸入新菜色" in page

    response = authed_client.post("/admin/order-tool/summary/dishes", data={
        "service_date": "2026-08-13",
        "week": "2026-08-10",
        "recipe_id": str(ids["recipe"]),
    })
    assert response.status_code == 302
    assert "week=2026-08-10" in response.headers["Location"]

    with app.app_context():
        plan = KitchenMenuPlan.query.filter_by(
            service_date=TEST_DAY, meal_type="午餐", name="中央菜單"
        ).one()
        item = KitchenMenuPlanItem.query.filter_by(plan_id=plan.id, recipe_id=ids["recipe"]).one()
        assert item.recipe.name == "南洋綠咖哩雞"

    page = authed_client.get("/admin/order-tool/summary?week=2026-08-10").get_data(as_text=True)
    assert "南洋綠咖哩雞" in page

    response = authed_client.post("/admin/order-tool/summary/dishes", data={
        "service_date": "2026-08-14",
        "week": "2026-08-10",
        "dish_name": "香煎鯖魚",
        "category": "主菜",
    })
    assert response.status_code == 302
    with app.app_context():
        new_recipe = KitchenRecipe.query.filter_by(name="香煎鯖魚", category="主菜").one()
        friday_plan = KitchenMenuPlan.query.filter_by(service_date=date(2026, 8, 14)).one()
        assert KitchenMenuPlanItem.query.filter_by(
            plan_id=friday_plan.id, recipe_id=new_recipe.id
        ).count() == 1


def test_summary_next_step_builds_school_week_menu_with_headcount(app, authed_client):
    ids = _seed_core_via_routes(app, authed_client)
    authed_client.post("/admin/order-tool/summary/dishes", data={
        "service_date": "2026-08-13",
        "week": "2026-08-10",
        "recipe_id": str(ids["recipe"]),
    })

    page = authed_client.get(
        f"/admin/order-tool/summary/schools?week=2026-08-10&school_id={ids['school']}"
    ).get_data(as_text=True)
    assert "各家學校菜單" in page
    assert "南洋綠咖哩雞" in page
    assert "供餐人數" in page

    payload = {"school_id": str(ids["school"]), "week": "2026-08-10"}
    for offset in range(7):
        day = date(2026, 8, 10) + timedelta(days=offset)
        payload[f"headcount_{day.isoformat()}"] = "0"
    payload["headcount_2026-08-13"] = "586"
    payload["recipes_2026-08-13"] = str(ids["recipe"])
    response = authed_client.post("/admin/order-tool/summary/schools/save", data=payload)
    assert response.status_code == 302

    with app.app_context():
        plan = KitchenMenuPlan.query.filter_by(
            service_date=TEST_DAY, meal_type="午餐", name="內小菜單"
        ).one()
        assignment = KitchenMenuAssignment.query.filter_by(plan_id=plan.id, school_id=ids["school"]).one()
        assert assignment.headcount == 586
        assert [item.recipe_id for item in plan.items] == [ids["recipe"]]

    plans_page = authed_client.get("/admin/order-tool/plans", follow_redirects=False)
    assert plans_page.status_code == 302
    assert "/summary" in plans_page.headers["Location"]


def test_school_menu_can_import_school_excel_directly(app, authed_client):
    ids = _seed_core_via_routes(app, authed_client)
    response = authed_client.post(
        "/admin/order-tool/summary/schools/import",
        data={
            "school_id": str(ids["school"]),
            "week": "2026-06-01",
            "menu_file": (_menu_upload_file(), "中平國小115年6月菜單.xlsx"),
        },
        content_type="multipart/form-data",
        follow_redirects=True,
    )
    assert response.status_code == 200
    assert "已匯入 內小" in response.get_data(as_text=True)
    with app.app_context():
        plan = KitchenMenuPlan.query.filter_by(
            service_date=date(2026, 6, 1), meal_type="午餐", name="內小菜單"
        ).one()
        assert KitchenMenuAssignment.query.filter_by(plan_id=plan.id, school_id=ids["school"]).one().headcount == 0
        assert "沙茶雞肉" in [item.recipe.name for item in plan.items]


def test_single_day_procurement_has_simple_fields_and_searchable_supplier(app, authed_client):
    ids = _seed_core_via_routes(app, authed_client)
    authed_client.post("/admin/order-tool/summary/dishes", data={
        "service_date": "2026-08-13",
        "week": "2026-08-10",
        "recipe_id": str(ids["recipe"]),
    })
    payload = {"school_id": str(ids["school"]), "week": "2026-08-10"}
    for offset in range(7):
        day = date(2026, 8, 10) + timedelta(days=offset)
        payload[f"headcount_{day.isoformat()}"] = "0"
    payload["headcount_2026-08-13"] = "801"
    payload["recipes_2026-08-13"] = str(ids["recipe"])
    authed_client.post("/admin/order-tool/summary/schools/save", data=payload)

    response = authed_client.post(
        "/admin/order-tool/summary/procurement/generate",
        data={"date": "2026-08-13"},
        follow_redirects=True,
    )
    page = response.get_data(as_text=True)
    assert all(label in page for label in (
        "食材名稱", "總供餐人次", "系統需求量", "實際採購量", "交貨日期／時段", "供應廠商"
    ))
    assert 'list="supplier-search-options"' in page
    assert ">801</b> 人次" in page

    with app.app_context():
        item = KitchenPurchaseOrderItem.query.one()
        item_id = item.id
    saved = authed_client.post("/admin/order-tool/summary/procurement/save", data={
        "date": "2026-08-13",
        "item_ids": str(item_id),
        f"actual_{item_id}": "71",
        f"delivery_date_{item_id}": "2026-08-12",
        f"delivery_slot_{item_id}": "下午",
        f"supplier_{item_id}": "測試肉品",
    })
    assert saved.status_code == 302
    with app.app_context():
        item = db.session.get(KitchenPurchaseOrderItem, item_id)
        assert item.actual_order_qty == Decimal("71")
        assert item.delivery_date == date(2026, 8, 12)
        assert item.delivery_slot == "下午"
        assert item.order.supplier.name == "測試肉品"


def _menu_upload_file():
    workbook = Workbook()
    hidden = workbook.active
    hidden.title = "舊菜單"
    hidden.sheet_state = "hidden"
    hidden.append(["日期", "星期", "主食", "主菜"])
    hidden.append(["6/1", "一", "不應匯入的舊菜", "舊主菜"])

    sheet = workbook.create_sheet("6月", 0)
    sheet["A1"] = "中平國小115年6月菜單"
    sheet.append(["日期", "星期", "主食", "主菜", "副菜", "", "蔬菜", "湯品", "全穀雜糧類(份)"])
    sheet.append(["6/1", "一", "白米飯", "沙茶雞肉", "麻婆豆腐", "麻婆豆腐", "有機蔬菜", "玉米蛋花湯", 5])
    sheet.append([None, None, "白米/煮", "雞肉/燒", "豆腐/煮", "豆腐/煮", "青菜/炒", "玉米/煮", None])
    sheet.append([date(2026, 6, 2), "二", "糙米飯", "日式炸豬排", "日式蒸蛋", "金茸蒲瓜", "有機蔬菜", "肉骨茶湯", 5])
    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    return output


def test_summary_import_detects_template_dates_and_deduplicates(app, authed_client):
    response = authed_client.post(
        "/admin/order-tool/summary/import",
        data={"menu_file": (_menu_upload_file(), "中平國小115年6月菜單.xlsx")},
        content_type="multipart/form-data",
        follow_redirects=True,
    )
    assert response.status_code == 200
    page = response.get_data(as_text=True)
    assert "匯入完成" in page
    assert "2026/06/01 — 2026/06/07" in page

    with app.app_context():
        monday = KitchenMenuPlan.query.filter_by(service_date=date(2026, 6, 1)).one()
        tuesday = KitchenMenuPlan.query.filter_by(service_date=date(2026, 6, 2)).one()
        assert [item.recipe.name for item in monday.items] == [
            "白米飯", "沙茶雞肉", "麻婆豆腐", "有機蔬菜", "玉米蛋花湯"
        ]
        assert len(tuesday.items) == 6
        assert KitchenRecipe.query.filter_by(name="有機蔬菜").one().category == "青菜"
        assert KitchenRecipe.query.filter_by(name="不應匯入的舊菜").count() == 0
        original_item_count = KitchenMenuPlanItem.query.count()

    second = authed_client.post(
        "/admin/order-tool/summary/import",
        data={"menu_file": (_menu_upload_file(), "中平國小115年6月菜單.xlsx")},
        content_type="multipart/form-data",
        follow_redirects=True,
    )
    assert "略過" in second.get_data(as_text=True)
    with app.app_context():
        assert KitchenMenuPlanItem.query.count() == original_item_count


def test_summary_import_rejects_unknown_template_without_writing(app, authed_client):
    workbook = Workbook()
    workbook.active.append(["這是一個還沒支援的格式", "內容"])
    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    response = authed_client.post(
        "/admin/order-tool/summary/import",
        data={"menu_file": (output, "陌生模板.xlsx")},
        content_type="multipart/form-data",
        follow_redirects=True,
    )
    assert "請提供這種格式作為新模板" in response.get_data(as_text=True)
    with app.app_context():
        assert KitchenMenuPlan.query.count() == 0


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
