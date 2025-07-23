# blueprints/records.py
# -*- coding: utf-8 -*-
"""
行政後台：出勤月表與單筆編輯（六段獨立上下班）
| 日期 | 上午上 | 上午下 | 下午上 | 下午下 | 加班上 | 加班下 | 備註 | 正班 | 加班≤2 | 加班>2 | 假日 |
"""
from flask import Blueprint, render_template_string, request, redirect, url_for, abort
from datetime import date, timedelta
import re

from extensions import db
from models import Employee, Checkin
from . import CSS, merge_night, calc_hours, NIGHT_END

rec_bp = Blueprint("rec", __name__, url_prefix="/admin")
LEAVE_PTYPE = "lv"          # 備註／假別


# ────────── 權限（示範） ──────────
def require(role: str):
    return None              # Demo：總是通過


# ────────── 出勤月表 ──────────
@rec_bp.route("/records")
def show_records():
    if require("mgr"):
        return abort(403)

    # ───── 參數 ─────
    eid   = request.args.get("eid", "")
    area  = request.args.get("area", "")
    ym_p  = request.args.get("ym")
    today = date.today()

    if ym_p and re.fullmatch(r"\d{4}-\d{2}", ym_p):
        y, m = map(int, ym_p.split("-")); ym = ym_p
    else:
        y, m = today.year, today.month; ym = f"{y}-{m:02d}"

    # ───── 下拉資料 ─────
    areas = db.session.query(Employee.area).distinct().order_by(Employee.area).all()
    area_opts = "".join(
        f'<option value="{a.area}" {"selected" if a.area==area else ""}>{a.area}</option>'
        for a in areas)

    emp_q  = Employee.query.filter_by(area=area) if area else Employee.query
    emp_ls = emp_q.order_by(Employee.id).all()
    emp_opts = "".join(
        f'<option value="{e.id}" {"selected" if str(e.id)==eid else ""}>{e.id}-{e.name}</option>'
        for e in emp_ls)

    ym_opts, cur = [], today.replace(day=1)
    for _ in range(12):
        v = cur.strftime("%Y-%m"); l = cur.strftime("%Y / %m")
        ym_opts.append(f'<option value="{v}" {"selected" if v==ym else ""}>{l}</option>')
        cur = (cur - timedelta(days=1)).replace(day=1)

    form_html = (f'<form id="recForm">'
                 f'區域：<select name="area" onchange="recForm.submit()"><option></option>{area_opts}</select>　'
                 f'員工：<select name="eid"  onchange="recForm.submit()"><option></option>{emp_opts}</select>　'
                 f'月份：<select name="ym"   onchange="recForm.submit()">{"".join(ym_opts)}</select></form>')

    if not area and not eid:
        return render_template_string(
            f'<!doctype html><html><head>{CSS}</head><body>'
            f'<h2>出勤卡查詢</h2>{form_html}</body></html>')

    targets    = emp_ls if (area and not eid) else [Employee.query.get_or_404(eid)]
    html_parts = [form_html]

    # ───── 逐員工產表 ─────
    for emp in targets:
        brk       = emp.default_break or 0.0
        first_day = date(y, m, 1)
        days_in_m = ((first_day.replace(day=28) + timedelta(days=4)).replace(day=1) - first_day).days
        next_m    = (first_day + timedelta(days=32)).replace(day=1)

        raws = (Checkin.query
                .with_entities(Checkin.work_date, Checkin.p_type, Checkin.ts, Checkin.note)
                .filter(Checkin.employee_id == emp.id)
                .filter((Checkin.work_date >= first_day.isoformat()) &
                        (Checkin.work_date <= (next_m + timedelta(days=1)).isoformat()))
                .order_by(Checkin.work_date, Checkin.p_type)
                .all())

        # 凌晨下班併前日
        recs  = merge_night([(r.work_date, r.p_type, r.ts[11:16]) for r in raws])
        notes = {r.work_date[8:10]: (r.note or "請假") for r in raws if r.p_type == LEAVE_PTYPE}

        cols = ["am-in", "am-out", "pm-in", "pm-out", "ot-in", "ot-out"]

        # 累計
        wday = reg_sum = ot2_sum = otx_sum = hol_sum = 0
        rows_html = ""
        back = url_for("rec.show_records", eid=emp.id, ym=ym)

        for d in range(1, days_in_m + 1):
            dd      = f"{d:02d}"
            is_hol  = date(y, m, d).weekday() >= 5
            reg = ot2 = otx = 0.0

            # 三段成對（早 / 午 / 加班）
            for pin, pout, is_ot in [("am-in", "am-out", False),
                                     ("pm-in", "pm-out", False),
                                     ("ot-in", "ot-out", True)]:
                s, e = recs.get((dd, pin)), recs.get((dd, pout))
                if not (s and e):
                    continue

                hours = sum(calc_hours(s, e, brk))

                if is_ot:               # —— 加班段：全部進加班欄
                    cap = max(0, 2 - ot2)
                    take2       = min(hours, cap)
                    ot2        += take2
                    otx        += (hours - take2)
                else:                   # —— 日班段：先填正班，再流入加班
                    cap = max(0, 8 - reg)
                    take_reg    = min(hours, cap)
                    reg        += take_reg
                    remain      = hours - take_reg
                    cap2        = max(0, 2 - ot2)
                    take2       = min(remain, cap2)
                    ot2        += take2
                    otx        += (remain - take2)

            hol = (reg + ot2 + otx) if is_hol else 0
            if is_hol:
                hol_sum += hol
            else:
                reg_sum += reg; ot2_sum += ot2; otx_sum += otx
            if reg or ot2 or otx:
                wday += 1

            note  = notes.get(dd, "")
            style = (' style="background:#FFF2CC"' if note else
                     ' style="background:#DDDDDD"' if is_hol else '')

            def link(pt, val):
                day = f"{y}-{m:02d}-{dd}"
                # ↓ 只有一個 f-string，把整個 <a …>{val}</a> 包起來
                return f'<a href="{url_for("rec.edit_record", emp=emp.id, date=day, typ=pt, back=back)}">{val}</a>'

            cells = "".join(
                f'<td>{link(pt, recs.get((dd, pt)) or "-")}</td>' for pt in cols)
            cells = "".join(
                f'<td>{link(pt, recs.get((dd, pt)) or "-")}</td>' for pt in cols)

            rows_html += (f'<tr{style}><td>{m:02d}-{dd}</td>{cells}'
                          f'<td>{link(LEAVE_PTYPE, note or "-")}</td>'
                          f'<td>{reg or ""}</td><td>{ot2 or ""}</td>'
                          f'<td>{otx or ""}</td><td>{hol or ""}</td></tr>')

        total_row = (f'<tr><th>總計</th><th colspan="7"></th>'
                     f'<th>{reg_sum}</th><th>{ot2_sum}</th>'
                     f'<th>{otx_sum}</th><th>{hol_sum}</th></tr>')

        html_parts.append(f"""
<h2>{emp.name}（{emp.id}） 區域：{emp.area}　{y}/{m}</h2>
<h3>出勤天數：{wday}</h3>
<table>
<tr><th>日期</th><th>上午上</th><th>上午下</th><th>下午上</th><th>下午下</th>
    <th>加班上</th><th>加班下</th><th>備註</th>
    <th>正班</th><th>加班≤2</th><th>加班&gt;2</th><th>假日</th></tr>
{rows_html}{total_row}</table>
<p><a href="{url_for('exp.export', eid=emp.id, ym=ym)}">匯出 Excel</a> |
   <a href="{url_for('emp.list_employees')}">返回員工管理</a></p><hr>""")

    return render_template_string(
        f'<!doctype html><html><head>{CSS}</head><body>{"".join(html_parts)}</body></html>')


# ────────── 單筆編輯 ──────────
@rec_bp.route("/edit_rec", methods=["GET", "POST"])
def edit_record():
    if require("mgr"):
        return abort(403)

    emp_id = request.args.get("emp")
    dt     = request.args.get("date")      # YYYY‑MM‑DD
    typ    = request.args.get("typ")       # am‑in / … / lv
    back   = request.args.get("back")

    rec = (Checkin.query
           .filter_by(employee_id=emp_id, work_date=dt, p_type=typ)
           .first())

    init_val = (rec.note if typ == LEAVE_PTYPE else rec.ts[11:16]) if rec else ""

    # —— POST
    if request.method == "POST":
        if request.form.get("clear"):      # 刪除
            if rec:
                db.session.delete(rec); db.session.commit()
            return redirect(back)

        val = request.form.get("val", "").strip()
        if typ == LEAVE_PTYPE:             # 備註 / 假別
            if not val:
                return abort(400, "假別不可空白")
            if rec: rec.note = val
            else:  db.session.add(Checkin(
                    employee_id=emp_id, work_date=dt, p_type=LEAVE_PTYPE,
                    ts=f"{dt}T00:00:00", note=val))
        else:                              # 六段時間
            if not re.fullmatch(r"\d{2}:\d{2}", val):
                return abort(400, "需輸入 HH:MM")
            ts = f"{dt}T{val}:00"
            if rec: rec.ts = ts
            else:  db.session.add(Checkin(
                    employee_id=emp_id, work_date=dt, p_type=typ, ts=ts))
        db.session.commit()
        return redirect(back)

    title = {"am-in":"上午上班", "am-out":"上午下班",
             "pm-in":"下午上班", "pm-out":"下午下班",
             "ot-in":"加班上班", "ot-out":"加班下班",
             LEAVE_PTYPE:"備註 / 假別"}.get(typ, "未知")

    return render_template_string(f"""<!doctype html><html><head>{CSS}</head><body>
<h2>{emp_id}　{dt}　{title}</h2>
<form method="post">
<input name="val" value="{init_val}" placeholder="HH:MM 或假別" style="width:180px"><br>
<button type="submit">儲存</button>
<button type="submit" name="clear" value="1"
        style="background:red;color:#fff;margin-left:10px">清除</button>
</form>
<p><a href="{back}">返回</a></p></body></html>""")
