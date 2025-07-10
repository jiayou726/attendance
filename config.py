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
