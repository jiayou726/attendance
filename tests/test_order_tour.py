"""Order-tool guided tour is loaded as an engine only.

Pages must not define real steps here. The flow-reorder branch plugs those in.
"""

import pytest

from app import create_app
from extensions import db


@pytest.fixture()
def app(tmp_path):
    db_path = tmp_path / "order_tour_test.db"
    application = create_app({
        "TESTING": True,
        "SECRET_KEY": "test-only-secret",
        "SQLALCHEMY_DATABASE_URI": f"sqlite:///{db_path}",
        "AUTO_CREATE_DB": True,
        "KITCHEN_CSRF_ENABLED": False,
        "ADMIN_HR_PASSWORD": "test-hr-password",
        "ADMIN_MGR_PASSWORD": "test-mgr-password",
        "PRODUCTION": False,
    })
    yield application
    with application.app_context():
        db.session.remove()
        db.drop_all()


@pytest.fixture()
def client(app):
    return app.test_client()


def test_order_tool_loads_tour_engine_without_steps(client):
    page = client.get("/admin/order-tool/").get_data(as_text=True)
    assert "order_tour.css" in page
    assert "order_tour.js" in page
    assert "startTour(" not in page
    assert "mountHelpButton" not in page
    assert "mountFirstVisitBanner" not in page

    script = client.get("/static/order_tour.js").get_data(as_text=True)
    for snippet in (
        "function startTour",
        "function attach",
        "mountHelpButton",
        "mountFirstVisitBanner",
        "ORDER_FLOW_TOUR_STEPS",
        "order-tool-tour-progress",
        "advanceOn",
        "localStorage",
        "繼續導覽",
        "暫停",
        "結束導覽",
        "上一步",
        "下一步",
        "關閉",
        "ArrowRight",
        "ArrowLeft",
        "Escape",
        "scrollIntoView",
        'class="tour-spot',
        'class="tour-tip',
        "fill-rule",
    ):
        assert snippet in script, snippet

    css = client.get("/static/order_tour.css").get_data(as_text=True)
    for snippet in (
        ".tour-bg",
        ".tour-spot",
        ".tour-tip",
        "order-tour-help",
        "order-tour-banner",
        "font-size: 22px",
        "font-size: 18px",
    ):
        assert snippet in css, snippet

    demo = client.get("/static/order_tour_demo.html").get_data(as_text=True)
    assert "order_tour.js" in demo
    assert "mountHelpButton" in demo
    assert "mountFirstVisitBanner" in demo
    assert "菜色用量表" in demo
    assert "採購叫貨" in demo
    assert "advanceOn: 'click'" in demo
    assert "繼續導覽" in demo or "pause" in demo
    assert "?page=" in demo
    assert "pageUrl('usage')" in demo
    assert "pageUrl('order')" in demo
