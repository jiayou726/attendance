import os
BASE = os.path.abspath(os.path.dirname(__file__))

class Config:
    SECRET_KEY = "change-me"
    SQLALCHEMY_DATABASE_URI = f"sqlite:///{os.path.join(BASE, 'attendance.db')}"
    SQLALCHEMY_TRACK_MODIFICATIONS = False
