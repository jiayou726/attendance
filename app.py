# --- 強制先載入官方 Blueprint 定義（防止覆寫）---
import flask.blueprints

import os
from datetime import date
from decimal import Decimal
from flask import Flask, redirect, request, session, url_for
from sqlalchemy import inspect, text
from werkzeug.middleware.proxy_fix import ProxyFix

from config import Config
from extensions import db, migrate

# 藍圖
from blueprints.auth import auth_bp
from blueprints.punch import punch_bp
from blueprints.employees import emp_bp
from blueprints.records import rec_bp
from blueprints.export import exp_bp
from blueprints.import_employees import import_bp
from blueprints.order_tool import order_bp
from blueprints.recipe_performance import install_recipe_performance_views
from blueprints.school_ingredient_export import school_ingredient_export_bp
from blueprints.nonregistered_menu_format import install_nonregistered_menu_export_format_fix


def _ensure_kitchen_schema_compatibility(app: Flask):
    """Apply only the additive kitchen fix needed when deploys skip Alembic."""

    with app.app_context():
        from models import KitchenDailyDishNote

        inspector = inspect(db.engine)
        with db.engine.begin() as connection:
            table_name = "kitchen_menu_assignment"
            if inspector.has_table(table_name):
                columns = {column["name"] for column in inspector.get_columns(table_name)}
                if "service_status" not in columns:
                    connection.execute(text(
                        "ALTER TABLE kitchen_menu_assignment "
                        "ADD COLUMN service_status VARCHAR(20) NOT NULL DEFAULT 'serving'"
                    ))
            school_table = "kitchen_school"
            if inspector.has_table(school_table):
                columns = {column["name"] for column in inspector.get_columns(school_table)}
                if "default_vegetarian_headcount" not in columns:
                    connection.execute(text(
                        "ALTER TABLE kitchen_school ADD COLUMN "
                        "default_vegetarian_headcount INTEGER NOT NULL DEFAULT 0"
                    ))

        # 日常表格的人工食材備註是空白可建的附加資料；
        # 部署若尚未執行 Alembic，仍可安全、重複地補上新表。
        KitchenDailyDishNote.__table__.create(bind=db.engine, checkfirst=True)

        # Historical supplier prices already live in production.  Apply the
        # conservative, idempotent unit conversion after each deploy so a
        # pushed release updates zero-priced ingredient master rows as well.
        required_tables = {
            "kitchen_ingredient", "kitchen_supplier", "kitchen_supplier_item",
        }
        if all(inspector.has_table(table) for table in required_tables):
            from scripts.backfill_kitchen_unit_prices import backfill_prices

            selected = backfill_prices(
                date(2012, 1, 1),
                date(2025, 12, 31),
                Decimal("5"),
                Decimal("2000"),
                overwrite=False,
            )
            if selected:
                db.session.commit()


def create_app(config_overrides=None) -> Flask:
    app = Flask(__name__)
    app.config.from_object(Config)
    if config_overrides:
        app.config.update(config_overrides)

    # 正式環境 fail closed：少了秘密或管理密碼就不要假裝安全上線。
    if app.config.get("PRODUCTION", False):
        if not app.config.get("SECRET_KEY_CONFIGURED", False):
            raise RuntimeError("PRODUCTION=1 時必須設定 SECRET_KEY。")
        if not (app.config.get("ADMIN_HR_PASSWORD") or app.config.get("ADMIN_MGR_PASSWORD")):
            raise RuntimeError("PRODUCTION=1 時至少要設定一組管理後台密碼。")
        if not app.config.get("KITCHEN_CSRF_ENABLED", True):
            raise RuntimeError("PRODUCTION=1 時不可關閉 KITCHEN_CSRF_ENABLED。")

    app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1)

    db.init_app(app)
    migrate.init_app(app, db)

    app.register_blueprint(auth_bp, url_prefix="/admin")
    app.register_blueprint(emp_bp, url_prefix="/admin")
    app.register_blueprint(rec_bp, url_prefix="/admin")
    app.register_blueprint(exp_bp, url_prefix="/admin")
    app.register_blueprint(import_bp, url_prefix="/admin")
    app.register_blueprint(order_bp, url_prefix="/admin/order-tool")
    app.register_blueprint(school_ingredient_export_bp, url_prefix="/admin/order-tool")
    app.register_blueprint(punch_bp)

    # 非登合菜名範本本身有幾列格式不同；匯出前把所有資料列
    # 統一成第 2 列格式，避免日期、底色、框線在中間幾列跑掉。
    install_nonregistered_menu_export_format_fix(app)

    # 菜色配方清單會逐列計算 AP 生料；先把 BOM 與食材一次批次載入，
    # 避免正式站對遠端 PostgreSQL 產生每道菜一次的 N+1 查詢。
    install_recipe_performance_views(app)

    @app.before_request
    def protect_admin_pages():
        # 團膳菜單是內部作業工具，依需求可直接使用；其餘管理後台仍需登入。
        kitchen_public = request.path == "/admin/order-tool" or request.path.startswith("/admin/order-tool/")
        if (request.path.startswith("/admin") and not kitchen_public
                and request.path not in {"/admin/login", "/admin/logout"}):
            if not session.get("role"):
                return redirect(url_for("auth.login", next=request.full_path.rstrip("?")))
        return None

    @app.route("/")
    def home():
        return redirect(url_for("punch.qrcode_view"))

    # 本機 SQLite / 明確指定的測試環境可自動建表。
    # 正式 Supabase/PostgreSQL 預設關閉，應由 migration 管理 schema。
    if app.config.get("AUTO_CREATE_DB", False):
        with app.app_context():
            db.create_all()
    _ensure_kitchen_schema_compatibility(app)

    return app


if __name__ == "__main__":
    app = create_app()
    port = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("FLASK_DEBUG", "0") == "1" and not app.config.get("PRODUCTION", False)
    app.run(host="0.0.0.0", port=port, debug=debug)
