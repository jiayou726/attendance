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
    SQLALCHEMY_ENGINE_OPTIONS = {
        "pool_pre_ping": True,
        "pool_recycle": 280,
    }

    # ─────────────────────────────────────────────
    # 打卡頁「短效 gate / token」設定（IP/UA 綁定）
    # ─────────────────────────────────────────────
    # gate：視為「仍在現場」的有效視窗（秒）
    PUNCH_GATE_TTL_SEC = int(os.getenv("PUNCH_GATE_TTL_SEC", "120"))
    # 一次性 token 的時效（秒）
    PUNCH_TOKEN_TTL_SEC = int(os.getenv("PUNCH_TOKEN_TTL_SEC", "120"))
    # 是否把 IP/UA 摻入驗證（降低轉傳風險）
    PUNCH_BIND_IP = os.getenv("PUNCH_BIND_IP", "1") == "1"
    PUNCH_BIND_UA = os.getenv("PUNCH_BIND_UA", "1") == "1"
