from io import BytesIO
import os

import pytest

from openpyxl import Workbook

from datetime import datetime, timedelta, timezone

from sqlalchemy import func, select, text

from app import create_app
from extensions import db
from models import KitchenMenuPlan, KitchenRecipe, KitchenSchool
from services.practice_database import cleanup_expired_accounts, list_accounts, workspace_engine


def _app(tmp_path, *, limit=20):
    main_path = tmp_path / "formal.db"
    practice_path = tmp_path / "practice-control.db"
    return create_app({
        "TESTING": True,
        "SECRET_KEY": "practice-test-secret",
        "SQLALCHEMY_DATABASE_URI": f"sqlite:///{main_path}",
        "SQLALCHEMY_BINDS": {"practice": f"sqlite:///{practice_path}"},
        "AUTO_CREATE_DB": True,
        "KITCHEN_CSRF_ENABLED": False,
        "PRACTICE_ACCOUNT_LIMIT": limit,
        "PRACTICE_ACCOUNT_TTL_DAYS": 10,
        "PRACTICE_ACCOUNTS_ENABLED": True,
        "PRODUCTION": False,
    })


def _login(client, name):
    response = client.post(
        "/admin/order-tool/ai-menu/practice/login",
        data={"name": name},
    )
    assert response.status_code == 302
    assert response.headers["Location"].endswith("/admin/order-tool/")


def _create_plan(client, name="練習週菜單"):
    return client.post("/admin/order-tool/plans", data={
        "service_date": "2026-10-05",
        "meal_type": "午餐",
        "name": name,
    })


def test_named_accounts_resume_across_computers_and_remain_isolated(tmp_path):
    app = _app(tmp_path)
    with app.app_context():
        db.session.add(KitchenSchool(name="正式學校"))
        db.session.commit()

    first_computer = app.test_client()
    _login(first_computer, "王小明")
    assert _create_plan(first_computer).status_code == 302

    second_computer = app.test_client()
    _login(second_computer, "王小明")
    response = second_computer.get("/admin/order-tool/summary?week=2026-10-05")
    assert response.status_code == 200
    assert "練習週菜單" in response.get_data(as_text=True)

    different_person = app.test_client()
    _login(different_person, "李小華")
    response = different_person.get("/admin/order-tool/summary?week=2026-10-05")
    assert response.status_code == 200
    assert "練習週菜單" not in response.get_data(as_text=True)

    with app.app_context():
        assert KitchenSchool.query.filter_by(name="正式學校").one()
        assert KitchenMenuPlan.query.count() == 0
        assert len(list_accounts()) == 2


def test_practice_master_data_is_read_only(tmp_path):
    app = _app(tmp_path)
    with app.app_context():
        db.session.add(KitchenSchool(name="正式學校"))
        db.session.commit()
    client = app.test_client()
    _login(client, "練習者")

    response = client.post("/admin/order-tool/schools", data={"name": "不可新增"})
    assert response.status_code == 302
    with app.app_context():
        assert KitchenSchool.query.filter_by(name="不可新增").first() is None


def test_capacity_and_expiry_release_slots(tmp_path):
    app = _app(tmp_path, limit=2)
    one = app.test_client()
    two = app.test_client()
    three = app.test_client()
    _login(one, "一號")
    _login(two, "二號")

    response = three.post("/admin/order-tool/ai-menu/practice/login", data={"name": "三號"})
    assert response.status_code == 200
    assert "已達 2 人上限" in response.get_data(as_text=True)

    with app.app_context():
        engine = db.engines["practice"]
        stale = datetime.now(timezone.utc) - timedelta(days=11)
        with engine.begin() as connection:
            connection.execute(text(
                "UPDATE practice_account SET last_active_at=:stale WHERE normalized_name=:name"
            ), {"stale": stale, "name": "一號"})
        assert cleanup_expired_accounts() == 1

    _login(three, "三號")
    with app.app_context():
        assert {row["display_name"] for row in list_accounts()} == {"二號", "三號"}


def test_admin_link_can_reset_and_delete_account(tmp_path):
    app = _app(tmp_path)
    learner = app.test_client()
    _login(learner, "管理測試")
    _create_plan(learner)

    admin = app.test_client()
    with admin.session_transaction() as admin_session:
        admin_session["role"] = "hr"
        admin_session["_practice_admin_csrf"] = "admin-test-token"

    employee_page = admin.get("/admin/").get_data(as_text=True)
    assert "出勤卡查詢" in employee_page
    assert "練習人員管理" in employee_page
    management_page = admin.get("/admin/practice-users/")
    assert management_page.status_code == 200
    assert "管理測試" in management_page.get_data(as_text=True)

    with app.app_context():
        account = list_accounts()[0]
        account_id = account["id"]
        assert account["plan_count"] == 1
    response = admin.post(
        f"/admin/practice-users/{account_id}/reset",
        data={"_csrf_token": "admin-test-token"},
    )
    assert response.status_code == 302
    with app.app_context():
        assert list_accounts()[0]["plan_count"] == 0

    response = admin.post(
        f"/admin/practice-users/{account_id}/delete",
        data={"_csrf_token": "admin-test-token"},
    )
    assert response.status_code == 302
    with app.app_context():
        assert list_accounts() == []


def test_practice_excel_import_new_recipe_stays_private(tmp_path):
    """Excel imports may introduce dishes absent from the formal recipe catalog."""
    app = _app(tmp_path)
    with app.app_context():
        db.session.add(KitchenRecipe(name="正式既有菜", category="主菜"))
        db.session.commit()

    client = app.test_client()
    _login(client, "菜單匯入練習者")
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "一般菜單"
    sheet.append(["日期", "星期", "主菜", "湯品"])
    sheet.append(["2026/10/22", "四", "正式既有菜", "練習新湯品"])
    output = BytesIO()
    workbook.save(output)
    output.seek(0)

    imported = client.post(
        "/admin/order-tool/summary/import",
        data={"menu_file": (output, "2026年10月練習菜單.xlsx")},
        content_type="multipart/form-data",
        follow_redirects=True,
    )
    assert imported.status_code == 200
    assert "匯入完成" in imported.get_data(as_text=True)
    assert "練習新湯品" in imported.get_data(as_text=True)

    with app.app_context():
        # The formal master catalog has not been modified by a practice user.
        assert KitchenRecipe.query.filter_by(name="練習新湯品").count() == 0
        assert KitchenRecipe.query.filter_by(name="正式既有菜").count() == 1
        engine = workspace_engine(list_accounts()[0])
        with engine.connect() as connection:
            private_recipes = connection.execute(
                select(KitchenRecipe.__table__)
            ).mappings().all()
            private_plans = connection.execute(
                select(KitchenMenuPlan.__table__)
            ).mappings().all()
        assert {row["name"] for row in private_recipes} == {"正式既有菜", "練習新湯品"}
        assert len(private_plans) == 1

    # The same account can safely import the same workbook again: no duplicates.
    repeated = client.post(
        "/admin/order-tool/summary/import",
        data={"menu_file": (_practice_menu_file(), "2026年10月練習菜單.xlsx")},
        content_type="multipart/form-data",
        follow_redirects=True,
    )
    assert repeated.status_code == 200
    assert "略過" in repeated.get_data(as_text=True)
    with app.app_context():
        engine = workspace_engine(list_accounts()[0])
        with engine.connect() as connection:
            assert len(connection.execute(select(KitchenRecipe.__table__)).all()) == 2


def _practice_menu_file():
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "一般菜單"
    sheet.append(["日期", "星期", "主菜", "湯品"])
    sheet.append(["2026/10/22", "四", "正式既有菜", "練習新湯品"])
    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    return output


def _realistic_november_menu():
    """Match the uploaded 中平 November template: offset date column,
    six dish columns, and a separate ingredient-description row per day.
    """
    from datetime import date

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "中平115.11月"
    sheet.append([None] * 16)
    sheet.append([None, "日期", "星期", "主食", "主菜", "副菜", "副菜", "青菜", "湯品",
                  "全穀雜糧類(份)", "豆魚蛋肉類(份)", "蔬菜類(份)", "油脂類(份)",
                  "其他類(份)", "熱量(大卡)", "三章1Q"])
    for day in range(1, 31):
        d = date(2026, 11, day)
        if d.weekday() >= 5:
            continue
        sheet.append([
            None, datetime(2026, 11, day), "一" if d.weekday() == 0 else "大塊肉",
            f"飯{day}", f"主菜{day}", f"副菜甲{day}", f"副菜乙{day}",
            f"青菜{day}", f"湯{day}", 5.6, 2.2, 2, 2.2, None, 706, "V",
        ])
        sheet.append([None, None, None, None, "豬肉Q.洋蔥Q/燒",
                      "豆腐.菇Q/煮", "玉米粒S/煮", None, "雞蛋Q"])
    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    return output


def test_practice_import_month_template_is_complete_and_idempotent(tmp_path):
    app = _app(tmp_path)
    client = app.test_client()
    _login(client, "完整月份匯入")
    response = client.post(
        "/admin/order-tool/summary/import",
        data={"menu_file": (_realistic_november_menu(), "中平11月菜單.xlsx")},
        content_type="multipart/form-data",
        follow_redirects=True,
    )
    page = response.get_data(as_text=True)
    assert response.status_code == 200
    assert "匯入完成" in page
    assert "21 天新增 126 道菜" in page

    with app.app_context():
        assert KitchenMenuPlan.query.count() == 0  # Formal database stays untouched.
        assert KitchenRecipe.query.count() == 0
        engine = workspace_engine(list_accounts()[0])
        with engine.connect() as connection:
            assert connection.execute(select(func.count()).select_from(
                KitchenMenuPlan.__table__
            )).scalar_one() == 21
            assert connection.execute(select(func.count()).select_from(
                KitchenRecipe.__table__
            )).scalar_one() == 126

    again = client.post(
        "/admin/order-tool/summary/import",
        data={"menu_file": (_realistic_november_menu(), "中平11月菜單.xlsx")},
        content_type="multipart/form-data",
        follow_redirects=True,
    )
    assert "0 天新增 0 道菜" in again.get_data(as_text=True)


@pytest.mark.skipif(not os.environ.get("PRACTICE_POSTGRES_TEST_URL"),
                    reason="Requires dedicated disposable PostgreSQL for integration tests")
def test_practice_postgres_excel_import_uses_private_recipe_overlay():
    # Never point this test at a shared or production database.
    postgres_url = os.environ["PRACTICE_POSTGRES_TEST_URL"]
    app = create_app({
        "TESTING": True,
        "SECRET_KEY": "postgres-practice-test-secret",
        "SQLALCHEMY_DATABASE_URI": postgres_url,
        "SQLALCHEMY_BINDS": {"practice": postgres_url},
        "AUTO_CREATE_DB": True,
        "KITCHEN_CSRF_ENABLED": False,
        "PRACTICE_ACCOUNT_LIMIT": 20,
        "PRACTICE_ACCOUNT_TTL_DAYS": 10,
        "PRACTICE_ACCOUNTS_ENABLED": True,
        "PRODUCTION": False,
    })
    with app.app_context():
        db.session.add(KitchenRecipe(name="正式既有菜", category="主菜"))
        db.session.commit()
    client = app.test_client()
    _login(client, "PostgreSQL 練習匯入")
    response = client.post(
        "/admin/order-tool/summary/import",
        data={"menu_file": (_realistic_november_menu(), "中平11月菜單.xlsx")},
        content_type="multipart/form-data",
        follow_redirects=True,
    )
    page = response.get_data(as_text=True)
    assert response.status_code == 200
    assert "21 天新增 126 道菜" in page

    with app.app_context():
        assert KitchenRecipe.query.count() == 1  # No formal recipes created.
        assert KitchenMenuPlan.query.count() == 0
        account = list_accounts()[0]
        engine = workspace_engine(account)
        with engine.connect() as connection:
            private_recipes = connection.execute(select(KitchenRecipe.__table__)).mappings().all()
            assert len(private_recipes) == 127
            assert sum(row["id"] < 0 for row in private_recipes) == 126
            assert connection.execute(select(func.count()).select_from(
                KitchenMenuPlan.__table__
            )).scalar_one() == 21
            assert connection.execute(select(func.count()).select_from(
                db.metadata.tables["kitchen_menu_plan_item"]
            )).scalar_one() == 126
            plan_ids = connection.execute(select(
                db.metadata.tables["kitchen_menu_plan_item"].c.plan_id
            )).scalars().all()
            assert len(set(plan_ids)) == 21, {
                plan_id: plan_ids.count(plan_id) for plan_id in set(plan_ids)
            }
            plan = KitchenMenuPlan.__table__
            item = db.metadata.tables["kitchen_menu_plan_item"]
            recipe = KitchenRecipe.__table__
            links = connection.execute(
                select(plan.c.service_date, recipe.c.name)
                .select_from(plan.join(item, plan.c.id == item.c.plan_id)
                                 .join(recipe, recipe.c.id == item.c.recipe_id))
            ).all()
            mismatched = [(str(day), name) for day, name in links
                          if not name.endswith(str(day.day))]
            assert not mismatched, mismatched[:30]

    repeated = client.post(
        "/admin/order-tool/summary/import",
        data={"menu_file": (_realistic_november_menu(), "中平11月菜單.xlsx")},
        content_type="multipart/form-data",
        follow_redirects=True,
    )
    repeated_page = repeated.get_data(as_text=True)
    assert "0 天新增 0 道菜" in repeated_page, repeated_page.partition('class="flash')[2][:1000]
