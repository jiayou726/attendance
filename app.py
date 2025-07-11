# --- 強制先載入官方 Blueprint 定義（防止覆寫）---
import flask.blueprints   # 這行一定放最上面

import os
from flask import Flask, redirect, url_for
from config     import Config
from extensions import db, migrate
from sqlalchemy import text

# 藍圖
from blueprints.auth      import auth_bp
from blueprints.punch     import punch_bp
from blueprints.employees import emp_bp
from blueprints.records   import rec_bp
from blueprints.export    import exp_bp


def create_app() -> Flask:
    app = Flask(__name__)
    app.config.from_object(Config)

    # ── 初始化 ORM / Migrate ──
    db.init_app(app)
    migrate.init_app(app, db)

    # ── 註冊藍圖 ──
    app.register_blueprint(auth_bp, url_prefix="/admin")
    app.register_blueprint(emp_bp,  url_prefix="/admin")
    app.register_blueprint(rec_bp,  url_prefix="/admin")
    app.register_blueprint(exp_bp,  url_prefix="/admin")
    app.register_blueprint(punch_bp)              # /punch

    # ── 首頁導向 ──
    @app.route("/")
    def home():
        return redirect(url_for("punch.form"))

    # ── ★ 第一次啟動自動建立所有資料表 ──
    with app.app_context():
        db.create_all()          # 如果已存在資料表則忽略，不會覆寫
        @app.route("/dbping")
        def dbping():
            try:
                from extensions import db
                # SELECT 1 只是輕量驗證連線
                db.session.execute("SELECT 1")
                return "OK"           # 連線成功
            except Exception as e:
            # 回傳錯誤訊息便於debug（完成功能後可移除）
                  return f"ERR: {e}"
    return app


# ────────────────────────── 本機 / 雲端啟動點 ──────────────────────────
if __name__ == "__main__":
    # 雲端平台（Render、Railway…）會把埠號放在 PORT 環境變數
    port = int(os.environ.get("PORT", 5000))
    # 正式環境建議把 debug 關掉，以免洩漏 Stack Trace
    create_app().run(host="0.0.0.0", port=port, debug=False)
