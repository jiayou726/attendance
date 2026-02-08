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

    # 打卡定位圍欄（任一座標點半徑內可打卡）
    PUNCH_GEOFENCE_ENABLED = os.getenv("PUNCH_GEOFENCE_ENABLED", "1") == "1"
    PUNCH_ALLOW_RADIUS_M = float(os.getenv("PUNCH_ALLOW_RADIUS_M", "100"))
    PUNCH_REQUIRE_ACCURACY_M = float(os.getenv("PUNCH_REQUIRE_ACCURACY_M", "250"))
    PUNCH_GEOFENCE_POINTS = [
        (25.01694367507761, 121.31324835609023),
        (24.864282032572383, 121.2034893872944),
        (25.00698765042551, 121.31874151968856),
        (24.891353522196585, 121.18859964870184),
        (24.967540156446894, 121.33181626685203),
        (24.957170435658362, 121.33947399543568),
        (24.959181153914287, 121.33274972452216),
        (24.10406010679953, 120.70080375767031),
        (24.855102775902832, 121.21613202921003),
        (24.842342354709256, 121.21078526496831),
    ]
