from flask import Blueprint, render_template_string, request, session, redirect, url_for, abort
from . import CSS

PASSWORDS = {"hr": "hr1234", "mgr": "mgr1234"}

auth_bp = Blueprint("auth", __name__)

@auth_bp.route("/login", methods=["GET","POST"])
def login():
    err=""
    if request.method=="POST":
        role = request.form["role"]
        pw   = request.form["pw"]
        if role in PASSWORDS and pw == PASSWORDS[role]:
            session["role"]=role
            return redirect("/admin/")
        err="<p style='color:red'>錯誤</p>"
    opt = "".join(f"<option value={r}>{'人資' if r=='hr' else '主管'}</option>"
                  for r in PASSWORDS)
    return render_template_string(f"""<!doctype html><html><head>{CSS}</head><body>
    <h2>管理登入</h2>{err}
    <form method=post>
      <select name=role>{opt}</select><br>
      <input type=password name=pw required><br>
      <button>登入</button>
    </form>
    <p><a href="/">回首頁</a></p></body></html>""")

def require(role):
    r=session.get("role")
    if not r:             # 未登入
        return redirect("/admin/login")
    if role=="hr" and r!="hr":
        abort(403)
