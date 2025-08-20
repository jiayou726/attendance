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

    # ─────────────────────────────────────────────
    # 打卡頁「短效 gate / token」設定（IP 綁定）
    # ─────────────────────────────────────────────
    # gate 時效（秒）：頁面生效視窗，過期後即使 F5 也不再發新 token
    PUNCH_GATE_TTL_SEC = int(os.getenv("PUNCH_GATE_TTL_SEC", "120"))
    # 提交用一次性 token 的時效（秒），通常與 gate 相同即可
    PUNCH_TOKEN_TTL_SEC = int(os.getenv("PUNCH_TOKEN_TTL_SEC", "120"))
    # 是否把 IP/UA 摻入 token 驗證（降低轉傳風險；可關）
    PUNCH_BIND_IP = os.getenv("PUNCH_BIND_IP", "1") == "1"
    PUNCH_BIND_UA = os.getenv("PUNCH_BIND_UA", "1") == "1"
