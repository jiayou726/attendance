# blueprints/punch.py
# --------------------------------------------------------
# 5 分鐘換 QR Code + 自動重整頁面　（完整可用版本）
# --------------------------------------------------------

from flask import (
    Blueprint, render_template_string, request,
    redirect, url_for, session, current_app,
)
from itsdangerous import URLSafeTimedSerializer, BadSignature, SignatureExpired
from datetime     import datetime, date, timedelta
from calendar     import monthrange
import time, re, io, base64, qrcode          # ← 新增 time；qrcode 已列
# --------------------------------------------------------

from extensions import db
from models      import Employee, Checkin
from .           import CSS, NIGHT_END, merge_night

punch_bp = Blueprint("punch", __name__, url_prefix="/punch")

# ──────────────────────────────── 5 分鐘 QR Code 產生 ────────────────────────────────
@punch_bp.route("/qrcode")
def qrcode_view():
    """
    產生只在 5 分鐘內有效的 QR Code，前端亦每 300 秒自動重整。
    """
    now_ts = int(time.time())
    s      = URLSafeTimedSerializer(current_app.config["SECRET_KEY"])
    token  = s.dumps({"ts": now_ts})                         # 只簽時間戳
    url    = url_for("punch.form", t=token, _external=True)

    # 產生 QR → Base64 嵌入
    qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_M)
    qr.add_data(url); qr.make(fit=True)
    img = qr.make_image()
    buf = io.BytesIO(); img.save(buf, format="PNG")
    b64 = base64.b64encode(buf.getvalue()).decode()

    gen_time = datetime.fromtimestamp(now_ts).strftime("%H:%M:%S")
    return render_template_string(f"""<!doctype html><html><head>{CSS}
<meta http-equiv="refresh" content="300">  <!-- 5 分鐘自動刷新 -->
</head><body>
  <h3>即時 QR Code（生成時間 {gen_time}，每 5 分鐘更新）</h3>
  <img src="data:image/png;base64,{b64}" alt="QR Code">
</body></html>""")


# ──────────────────────────────── 打卡表單 ────────────────────────────────
@punch_bp.route("/", methods=["GET"])
def form():
    err = request.args.get("err")

    if not session.get("token_ok"):
        token = request.args.get("t")
        if not token:
            if not err:
                return redirect(url_for(".form", err="請重新掃描 QR Code"))
        else:
            s = URLSafeTimedSerializer(current_app.config["SECRET_KEY"])
            try:
                # ★ 只接受 300 秒內簽章
                s.loads(token, max_age=300)
            except (BadSignature, SignatureExpired):
                return redirect(url_for(".form", err="請重新掃描 QR Code"))
            session["token_ok"] = True

    return render_template_string(f"""<!doctype html><html><head>{CSS}</head><body>
  {f'<p class=error>{err}</p>' if err else ''}
  <h2>員工打卡</h2>
  <form method=post>
    <input name=eid placeholder="員工編號" required autofocus>
    <select name=type>
      <option value=in>上班</option><option value=out>下班</option>
    </select>
    <button>打卡</button>
  </form>
  <p><a href="/admin/login">管理</a></p></body></html>""")


@punch_bp.route("/", methods=["POST"])
def punch():
    if not session.get("token_ok"):
        return redirect(url_for(".form", err="請重新掃描 QR Code"))

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

    session.pop("token_ok", None)
    return redirect(url_for(".card", eid=eid, st=st, msg=msg))


# ──────────────────────────────── 員工月卡頁 ────────────────────────────────
@punch_bp.route("/result/<eid>")
def card(eid: str):
    # 1️⃣ 取得查詢年月；預設當月
    ym_param = request.args.get("ym")
    today    = date.today()
    if ym_param and re.fullmatch(r"\d{4}-\d{2}", ym_param):
        y, m = map(int, ym_param.split("-"))
        ym   = ym_param
    else:
        y, m = today.year, today.month
        ym   = f"{y}-{m:02d}"

    # 2️⃣ 月份下拉（最近 6 個月）
    ym_opts, cursor = [], today.replace(day=1)
    for _ in range(6):
        ym_val = cursor.strftime("%Y-%m")
        ym_lab = cursor.strftime("%Y / %m")
        sel    = "selected" if ym_val == ym else ""
        ym_opts.append(f'<option value="{ym_val}" {sel}>{ym_lab}</option>')
        cursor = (cursor - timedelta(days=1)).replace(day=1)

    # 3️⃣ 其他資料
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
  <p><a href="/">返回</a></p></body></html>""")
