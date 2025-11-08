"""單檔 Flask 叫貨工具（引用專案 Blueprint）。

使用方式：
1. pip install flask pandas openpyxl
2. (選擇) 設定 FLASK_SECRET_KEY
3. python 叫貨.py
4. 瀏覽 http://127.0.0.1:5000
"""

from __future__ import annotations

import os

from flask import Flask

from blueprints.order_tool import order_bp


def create_app() -> Flask:
    app = Flask(__name__)
    app.secret_key = os.environ.get("FLASK_SECRET_KEY", "dev-secret-change-me")
    app.register_blueprint(order_bp)
    return app


if __name__ == "__main__":
    create_app().run(debug=True)
