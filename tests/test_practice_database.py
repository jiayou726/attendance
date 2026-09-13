from sqlalchemy import select

from app import create_app
from extensions import db
from models import KitchenSchool


def test_practice_workspace_isolated_from_formal_database(tmp_path):
    main_path = tmp_path / "formal.db"
    practice_path = tmp_path / "practice.db"
    app = create_app({
        "TESTING": True,
        "SECRET_KEY": "practice-test-secret",
        "SQLALCHEMY_DATABASE_URI": f"sqlite:///{main_path}",
        "SQLALCHEMY_BINDS": {"practice": f"sqlite:///{practice_path}"},
        "AUTO_CREATE_DB": True,
        "KITCHEN_CSRF_ENABLED": False,
        "KITCHEN_PRACTICE_PASSWORD": "sandbox-pass",
        "PRODUCTION": False,
    })

    with app.app_context():
        db.session.add(KitchenSchool(name="正式學校"))
        db.session.commit()

    client = app.test_client()
    response = client.post("/admin/order-tool/ai-menu/practice/login", data={
        "user": "practice",
        "password": "sandbox-pass",
    })
    assert response.status_code == 302
    assert response.headers["Location"].endswith("/admin/order-tool/")

    # The normal kitchen CRUD route remains fully usable in practice mode.
    response = client.post("/admin/order-tool/schools", data={"name": "練習新增學校"})
    assert response.status_code == 302

    with app.app_context():
        # Formal DB is untouched by the practice write.
        assert KitchenSchool.query.filter_by(name="正式學校").one()
        assert KitchenSchool.query.filter_by(name="練習新增學校").first() is None

        # Practice DB received a one-time copy of reusable master data plus the
        # new sandbox-only write.
        practice_engine = db.engines["practice"]
        with practice_engine.connect() as connection:
            names = set(connection.execute(select(KitchenSchool.__table__.c.name)).scalars())
        assert names == {"正式學校", "練習新增學校"}

    client.get("/admin/order-tool/ai-menu/practice/logout")
    client.post("/admin/order-tool/schools", data={"name": "正式新增學校"})
    with app.app_context():
        assert KitchenSchool.query.filter_by(name="正式新增學校").one()
        with db.engines["practice"].connect() as connection:
            names = set(connection.execute(select(KitchenSchool.__table__.c.name)).scalars())
        assert "正式新增學校" not in names
