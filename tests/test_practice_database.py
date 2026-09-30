from datetime import datetime, timedelta, timezone

from sqlalchemy import func, select, text

from app import create_app
from extensions import db
from models import KitchenMenuPlan, KitchenSchool
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
