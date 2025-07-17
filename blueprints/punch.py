# blueprints/punch.py
# --------------------------------------------------------
#靜態「打卡」二維碼
# 「返回」鍵修正 → 回到打卡表單 (/punch/)
# --------------------------------------------------------

from flask import (
    Blueprint, render_template_string, request,
    redirect, url_for,
)
from datetime     import datetime, date, timedelta
from calendar     import monthrange
import  re, io, base64, qrcode

from extensions import db
from models      import Employee, Checkin
from .           import CSS, NIGHT_END, merge_night

punch_bp = Blueprint("punch", __name__, url_prefix="/punch")


# ──────────────────────────────── 5 分鐘 QR 產生 ────────────────────────────────
@punch_bp.route("/qrcode")
def qrcode_view():
    url = url_for("punch.form", _external=True)

    qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_M)
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image()
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    b64 = base64.b64encode(buf.getvalue()).decode()

    return render_template_string(
        f"""<!doctype html><html><head>{CSS}</head><body>
  <h3>打卡 QR Code</h3>
  <img src=\"data:image/png;base64,{b64}\" alt=\"QR Code\">
</body></html>"""
    )

# ──────────────────────────────── 打卡表單 ────────────────────────────────

@punch_bp.route("/", methods=["GET"])
def form():
    err = request.args.get("err")


    return render_template_string(f"""<!doctype html><html><head>{CSS}</head><body>
  {f'<p class=error>{err}</p>' if err else ''}
  <h2>員工打卡</h2>
  <form method=post>
    <input name=eid placeholder="員工編號" required autofocus>
    <select name=type>
      <option value=in>1.上班</option><option value=out>2.下班</option>
    </select>
    <button>打卡</button>
  </form>
  <p><a href="/admin/login">管理</a></p></body></html>""")


@punch_bp.route("/", methods=["POST"])
def punch():
    
    eid = request.form["eid"].strip()
    typ = request.form["type"]
    now = datetime.now()
    wd  = now.date().isoformat()
    ts  = now.isoformat(timespec="seconds")

    emp = Employee.query.get(eid)
    if not emp:
        return redirect(url_for(".card", eid=eid, st="error", msg="員工不存在"))

    dup = Checkin.query.filter_by(employee_id=eid, work_date=wd, p_type=typ).first()
    if dup:
        msg, st = "已打過卡", "warn"
    else:
        db.session.add(Checkin(employee_id=eid, work_date=wd, p_type=typ, ts=ts))
        db.session.commit()
        msg, st = "打卡成功", "success"

    return redirect(url_for(".card", eid=eid, st=st, msg=msg))


# ──────────────────────────────── 員工月卡頁 ────────────────────────────────
@punch_bp.route("/result/<eid>")
def card(eid: str):
    ym_param = request.args.get("ym")
    today    = date.today()
    if ym_param and re.fullmatch(r"\d{4}-\d{2}", ym_param):
        y, m = map(int, ym_param.split("-"))
        ym   = ym_param
    else:
        y, m = today.year, today.month
        ym   = f"{y}-{m:02d}"

    ym_opts, cursor = [], today.replace(day=1)
    for _ in range(6):
        ym_val = cursor.strftime("%Y-%m")
        ym_lab = cursor.strftime("%Y / %m")
        sel    = "selected" if ym_val == ym else ""
        ym_opts.append(f'<option value="{ym_val}" {sel}>{ym_lab}</option>')
        cursor = (cursor - timedelta(days=1)).replace(day=1)

    emp = Employee.query.get_or_404(eid)
    days_in_month = monthrange(y, m)[1]
    next_m = (date(y, m, 1) + timedelta(days=32)).replace(day=1)

    rows = (
        Checkin.query
        .with_entities(Checkin.work_date, Checkin.p_type, Checkin.ts)
        .filter(Checkin.employee_id == eid)
        .filter(
            (Checkin.work_date.like(f"{y}-{m:02d}%"))
            | (
                (Checkin.work_date == next_m.isoformat())
                & (Checkin.p_type == "out")
                & (Checkin.ts < f"{next_m}T{NIGHT_END}:00")
            )
        )
        .order_by(Checkin.work_date, Checkin.p_type)
        .all()
    )

    recs = merge_night([(r.work_date, r.p_type, r.ts[11:16]) for r in rows])
    body = "".join(
        f"<tr><td>{m:02d}-{d:02d}</td>"
        f"<td>{recs.get((f'{d:02d}','in'), '-')}</td>"
        f"<td>{recs.get((f'{d:02d}','out'), '-')}</td></tr>"
        for d in range(1, days_in_month + 1)
    )

    st, msg = request.args.get("st"), request.args.get("msg")
    return render_template_string(f"""<!doctype html><html><head>{CSS}</head><body>
  {f'<h3 class={st}>{msg}</h3>' if st else ''}
  <h3>{emp.name}（{eid}） 區域：{emp.area}　{y}/{m}</h3>

  <form method="get" id="ymForm">
    <input type="hidden" name="eid" value="{eid}">
    月份：<select name="ym" onchange="ymForm.submit()">
      {''.join(ym_opts)}
    </select>
  </form>

  <table>
    <tr><th>日期</th><th>上班</th><th>下班</th></tr>{body}
  </table>
  <p><a href="{url_for('.form')}">返回打卡</a></p></body></html>""")
