# --- 強制先載入官方 Blueprint 定義（防止覆寫）---
import flask.blueprints

import os
from flask import Flask, redirect, url_for
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


def create_app() -> Flask:
    app = Flask(__name__)
    app.config.from_object(Config)
    app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1)

    db.init_app(app)
    migrate.init_app(app, db)

    app.register_blueprint(auth_bp, url_prefix="/admin")
    app.register_blueprint(emp_bp, url_prefix="/admin")
    app.register_blueprint(rec_bp, url_prefix="/admin")
    app.register_blueprint(exp_bp, url_prefix="/admin")
    app.register_blueprint(import_bp, url_prefix="/admin")
    app.register_blueprint(order_bp, url_prefix="/admin/order-tool")
    app.register_blueprint(punch_bp)

    @app.route("/")
    def home():
        return redirect(url_for("punch.qrcode_view"))

    # 本機 SQLite / 明確指定的測試環境可自動建表。
    # 正式 Supabase/PostgreSQL 預設關閉，應由 migration 管理 schema。
    if app.config.get("AUTO_CREATE_DB", False):
        with app.app_context():
            db.create_all()

    return app


if __name__ == "__main__":
    app = create_app()
    port = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("FLASK_DEBUG", "0") == "1" and not app.config.get("PRODUCTION", False)
    app.run(host="0.0.0.0", port=port, debug=debug)
