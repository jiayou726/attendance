import os
from datetime import timedelta

BASE = os.path.abspath(os.path.dirname(__file__))


def _bool_env(name: str, default: bool = False) -> bool:
    raw = os.getenv(name)
    if raw is None:
        return default
    return raw.strip().lower() in {"1", "true", "yes", "on"}


def _database_url() -> str:
    url = os.getenv("DATABASE_URL", "").strip()
    # 部分平台仍可能給 postgres://；SQLAlchemy/psycopg2 使用 postgresql://。
    if url.startswith("postgres://"):
        url = "postgresql://" + url[len("postgres://"):]
    return url or f"sqlite:///{os.path.join(BASE, 'attendance.db')}"


class Config:
    # 正式環境必須在部署平台設定 SECRET_KEY。
    # 本機未設定時使用隨機值，因此重啟後 session 會失效，但不會把固定秘密放進 repo。
    SECRET_KEY_CONFIGURED = bool(os.getenv("SECRET_KEY"))
    SECRET_KEY = os.getenv("SECRET_KEY") or os.urandom(32)

    SQLALCHEMY_DATABASE_URI = _database_url()
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    SQLALCHEMY_ENGINE_OPTIONS = {
        "pool_pre_ping": True,
        "pool_recycle": int(os.getenv("DB_POOL_RECYCLE_SEC", "280")),
    }

    PRODUCTION = _bool_env("PRODUCTION", False)
    SESSION_COOKIE_HTTPONLY = True
    SESSION_COOKIE_SECURE = PRODUCTION
    SESSION_COOKIE_SAMESITE = "Lax"
    PERMANENT_SESSION_LIFETIME = timedelta(hours=int(os.getenv("ADMIN_SESSION_HOURS", "12")))

    # 管理後台帳密不可寫死在 public repo。
    ADMIN_HR_PASSWORD = os.getenv("ADMIN_HR_PASSWORD", "")
    ADMIN_MGR_PASSWORD = os.getenv("ADMIN_MGR_PASSWORD", "")

    # 新機器/測試環境可選擇自動建表；正式 Supabase 建議使用 migration。
    AUTO_CREATE_DB = _bool_env("AUTO_CREATE_DB", not bool(os.getenv("DATABASE_URL")))

    # 團膳後台 POST 專用 CSRF。測試可明確關閉，production 不應關閉。
    KITCHEN_CSRF_ENABLED = _bool_env("KITCHEN_CSRF_ENABLED", True)

    # ─────────────────────────────────────────────
    # 打卡頁「短效 gate / token」設定（IP/UA 綁定）
    # ─────────────────────────────────────────────
    PUNCH_GATE_TTL_SEC = int(os.getenv("PUNCH_GATE_TTL_SEC", "120"))
    PUNCH_TOKEN_TTL_SEC = int(os.getenv("PUNCH_TOKEN_TTL_SEC", "120"))
    PUNCH_BIND_IP = os.getenv("PUNCH_BIND_IP", "0") == "1"
    PUNCH_BIND_UA = os.getenv("PUNCH_BIND_UA", "1") == "1"

    PUNCH_GEOFENCE_ENABLED = os.getenv("PUNCH_GEOFENCE_ENABLED", "1") == "1"
    PUNCH_ALLOW_RADIUS_M = float(os.getenv("PUNCH_ALLOW_RADIUS_M", "500"))
    PUNCH_REQUIRE_ACCURACY_M = float(os.getenv("PUNCH_REQUIRE_ACCURACY_M", "250"))
    PUNCH_GEOFENCE_POINTS = [
        (24.842556724831017, 121.2107761047848),
        (24.960056999676954, 121.30991556472662),
        (24.96919144649612, 121.33483066999854),
        (24.880557272740955, 121.18754959707424),
        (24.10393862990907, 120.70077911082637),
    ]
