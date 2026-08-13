# --- 強制先載入官方 Blueprint 定義（防止覆寫）---
import flask.blueprints   # 這行一定放最上面

import os
from flask import Flask, redirect, url_for, request
from werkzeug.middleware.proxy_fix import ProxyFix
from config     import Config
from extensions import db, migrate

# 藍圖
from blueprints.auth      import auth_bp
from blueprints.punch     import punch_bp
from blueprints.employees       import emp_bp
from blueprints.records         import rec_bp
from blueprints.export          import exp_bp
from blueprints.import_employees import import_bp
from blueprints.order_tool import order_bp


def create_app() -> Flask:
    app = Flask(__name__)
    app.config.from_object(Config)
    app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1)

    # ── 初始化 ORM / Migrate ──
    db.init_app(app)
    migrate.init_app(app, db)

    # ── 註冊藍圖 ──
    app.register_blueprint(auth_bp, url_prefix="/admin")
    app.register_blueprint(emp_bp,  url_prefix="/admin")
    app.register_blueprint(rec_bp,  url_prefix="/admin")
    app.register_blueprint(exp_bp,  url_prefix="/admin")
    app.register_blueprint(import_bp, url_prefix="/admin")
    app.register_blueprint(order_bp, url_prefix="/admin/order-tool")
    app.register_blueprint(punch_bp)              # /punch

    # ── 團膳頁專用 responsive UI ──
    # 只對 /admin/order-tool 注入 CSS，不影響既有打卡、薪資等頁面。
    @app.after_request
    def inject_kitchen_responsive_ui(response):
        if (
            request.path.startswith("/admin/order-tool")
            and response.mimetype == "text/html"
            and response.status_code < 400
        ):
            html = response.get_data(as_text=True)
            css_tag = (
                '<link rel="stylesheet" '
                'href="/static/kitchen_mobile.css?v=2">'
            )
            if "</head>" in html and "kitchen_mobile.css" not in html:
                html = html.replace("</head>", f"{css_tag}</head>", 1)
                response.set_data(html)
                response.headers["Content-Length"] = len(response.get_data())
        return response

    # ── 首頁導向 ──
    @app.route("/")
    def home():
        return redirect(url_for("punch.qrcode_view"))

    # ── ★ 第一次啟動自動建立所有資料表 ──
    with app.app_context():
        db.create_all()          # 如果已存在資料表則忽略，不會覆寫
        
    return app


# ────────────────────────── 本機 / 雲端啟動點 ──────────────────────────
if __name__ == "__main__":
    # 雲端平台（Render、Railway…）會把埠號放在 PORT 環境變數
    port = int(os.environ.get("PORT", 5000))
    # 正式環境建議把 debug 關掉，以免洩漏 Stack Trace
    create_app().run(host="0.0.0.0", port=port, debug=True)
