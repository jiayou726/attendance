# blueprints/employees.py
# -*- coding: utf-8 -*-

from flask.blueprints import Blueprint
from flask import render_template_string, request, redirect, url_for, abort
from extensions import db
from models     import Employee
from .          import CSS          # 刪除 AREAS

emp_bp = Blueprint("emp", __name__, url_prefix="/admin")

@emp_bp.route("/")
def list_employees():
    employees = Employee.query.order_by(Employee.id).all()
    rows = ""
    for e in employees:
        rows += (
            f"<tr>"
            f"<td>{e.id}</td>"
            f"<td>{e.name}</td>"
            f"<td>{e.area}</td>"
            f"<td>{e.default_break or 0}</td>"
            f"<td>"
            f"<a href=\"{url_for('emp.edit_employee', eid=e.id)}\">編輯</a> "
            f"<form method='post' action=\"{url_for('emp.delete_employee', eid=e.id)}\" "
            f"style='display:inline' "
            f"onsubmit=\"return confirm('刪除 {e.id}-{e.name}？')\">"
            f"<button type='submit'>刪除</button></form>"
            f"</td>"
            f"</tr>"
        )

    return render_template_string(f"""
    <!doctype html>
    <html><head>{CSS}</head><body>
      <h2>員工名單</h2>
      <table>
        <tr><th>ID</th><th>姓名</th><th>區域</th><th>預設午休(小時)</th><th>操作</th></tr>
        {rows}
      </table>
      <p>
        <a href="{url_for('emp.add_employee')}">新增員工</a> |
        <a href="{url_for('imp.import_employees')}">批次匯入</a> |
        <a href="{url_for('rec.show_records')}">出勤卡查詢</a>
      </p>
    </body></html>
    """)

@emp_bp.route("/add", methods=["GET", "POST"])
def add_employee():
    if request.method == "POST":
        eid  = request.form["eid"].strip()
        name = request.form["name"].strip()
        area = request.form["area"].strip()
        try:
            default_break = float(request.form.get("default_break", 0))
        except ValueError:
            default_break = 0.0

        db.session.add(Employee(id=eid, name=name, area=area, default_break=default_break))
        db.session.commit()
        return redirect(url_for("emp.list_employees"))

    return render_template_string(f"""
    <!doctype html>
    <html><head>{CSS}</head><body>
      <h2>新增員工</h2>
      <form method="post">
        員工編號：<input name="eid" required><br>
        姓名：<input name="name" required><br>
        區域：<input name="area" required><br>   <!-- 改成自由輸入 -->
        預設午休(小時)：<select name="default_break">
          <option value="0">0</option>
          <option value="0.5">0.5</option>
          <option value="1">1</option>
        </select><br>
        <button type="submit">儲存</button>
        <a href="{url_for('emp.list_employees')}">返回</a>
      </form>
    </body></html>
    """)

@emp_bp.route("/edit/<eid>", methods=["GET", "POST"])
def edit_employee(eid):
    emp = Employee.query.get_or_404(eid)
    if request.method == "POST":
        emp.name = request.form["name"].strip()
        emp.area = request.form["area"].strip()
        try:
            emp.default_break = float(request.form.get("default_break", 0))
        except ValueError:
            emp.default_break = 0.0
        db.session.commit()
        return redirect(url_for("emp.list_employees"))

    return render_template_string(f"""
    <!doctype html>
    <html><head>{CSS}</head><body>
      <h2>編輯員工：{eid}</h2>
      <form method="post">
        員工編號：<input value="{eid}" readonly><br>
        姓名：<input name="name" value="{emp.name}" required><br>
        區域：<input name="area" value="{emp.area}" required><br>  <!-- 改成自由輸入 -->
        預設午休(小時)：<select name="default_break">
          <option value="0" {'selected' if emp.default_break==0 else ''}>0</option>
          <option value="0.5" {'selected' if emp.default_break==0.5 else ''}>0.5</option>
          <option value="1" {'selected' if emp.default_break==1 else ''}>1</option>
        </select><br>
        <button type="submit">更新</button>
        <a href="{url_for('emp.list_employees')}">返回</a>
      </form>
    </body></html>
    """)

@emp_bp.route("/delete/<eid>", methods=["POST"])
def delete_employee(eid):
    Employee.query.filter_by(id=eid).delete()
    db.session.commit()
    return redirect(url_for("emp.list_employees"))
