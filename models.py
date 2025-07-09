# models.py
from extensions import db

class Employee(db.Model):
    __tablename__ = 'employee'   # 若你資料庫中表名為 employees，請改成 'employees'
    id            = db.Column(db.Integer, primary_key=True)
    name          = db.Column(db.String(50), nullable=False)
    area          = db.Column(db.String(50))
    default_break = db.Column(db.Float, nullable=False, default=0.0)

class Checkin(db.Model):
    __tablename__ = 'checkin'    # 若你原本叫 checkins，請更新為一致
    id          = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey('employee.id'), nullable=False)
    work_date   = db.Column(db.String(10), nullable=False)     
    p_type      = db.Column(db.String(3), nullable=False)      
    ts          = db.Column(db.String(19), nullable=False)   
    note = db.Column(db.String(50))  # 允許輸入中文備註
  
    __table_args__ = (
        db.UniqueConstraint('employee_id', 'work_date', 'p_type'),
    )
