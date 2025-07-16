import os
BASE = os.path.abspath(os.path.dirname(__file__))

class Config:
    SECRET_KEY = "change-me"

    # 優先使用雲端 Postgres，否則退回本地 SQLite
    SQLALCHEMY_DATABASE_URI = os.getenv(
        "DATABASE_URL",
        f"sqlite:///{os.path.join(BASE, 'attendance.db')}"
    )

    SQLALCHEMY_TRACK_MODIFICATIONS = False

    # ── SQLAlchemy 連線池設定 ──
    #   pool_pre_ping: 送出簡單查詢檢查連線是否存活，可避免用到失效連線
    #   pool_recycle:  若主機會在閒置一段時間後關閉連線，設定秒數可自動重連
    SQLALCHEMY_ENGINE_OPTIONS = {
        "pool_pre_ping": True,
        "pool_recycle": 280,
    }
