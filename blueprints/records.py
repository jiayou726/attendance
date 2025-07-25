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
from . import CSS, merge_night, calc_hours   # ← 已改好的 calc_hours
rec_bp = Blueprint("rec", __name__, url_prefix="/admin")

LEAVE_PTYPE = "lv"          # 備註／假別
def require(role: str):     # Demo 權限
    return None

# ────────────────────── 出勤月表 ──────────────────────
@rec_bp.route("/records")
def show_records():
    if require("mgr"):
        return abort(403)

    # 1. 解析參數
    eid  = request.args.get("eid", "")
    area = request.args.get("area", "")
    ym_p = request.args.get("ym")
    today = date.today()
    if ym_p and re.fullmatch(r"\d{4}-\d{2}", ym_p):
        y, m = map(int, ym_p.split("-")); ym = ym_p
    else:
        y, m = today.year, today.month; ym = f"{y}-{m:02d}"

    # 2. 下拉選單
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
        sel = "selected" if cur.year==y and cur.month==m else ""
        ym_opts.append(f'<option value="{cur:%Y-%m}" {sel}>{cur:%Y / %m}</option>')
        cur = (cur - timedelta(days=1)).replace(day=1)

    form_html = (
        '<form id="recForm">'
        f'區域：<select name="area" onchange="recForm.submit()"><option></option>{area_opts}</select>　'
        f'員工：<select name="eid"  onchange="recForm.submit()"><option></option>{emp_opts}</select>　'
        f'月份：<select name="ym"   onchange="recForm.submit()">{"".join(ym_opts)}</select>'
        '</form>'
    )

    if not area and not eid:
        return render_template_string(f'<!doctype html><html><head>{CSS}</head><body>'
                                      f'<h2>出勤卡查詢</h2>{form_html}</body></html>')

    targets    = emp_ls if (area and not eid) else [Employee.query.get_or_404(eid)]
    html_parts = [form_html]

    # 3. 逐員工產表
    for emp in targets:
        brk = emp.default_break or 0.0
        first = date(y, m, 1)
        days_in_m = ((first.replace(day=28)+timedelta(days=4)).replace(day=1) - first).days
        next_m = (first + timedelta(days=32)).replace(day=1)

        raws = (Checkin.query
                .with_entities(Checkin.work_date, Checkin.p_type, Checkin.ts, Checkin.note)
                .filter_by(employee_id=emp.id)
                .filter(Checkin.work_date >= first.isoformat(),
                        Checkin.work_date <= (next_m + timedelta(days=1)).isoformat())
                .order_by(Checkin.work_date, Checkin.p_type).all())

        recs  = merge_night([(r.work_date, r.p_type, r.ts[11:16]) for r in raws])
        notes = {r.work_date[8:10]: (r.note or "請假") for r in raws if r.p_type == LEAVE_PTYPE}
        cols  = ["am-in","am-out","pm-in","pm-out","ot-in","ot-out"]

        wday = reg_sum = ot2_sum = otx_sum = hol_sum = 0
        rows_html = ""
        back = url_for("rec.show_records", eid=emp.id, ym=ym)

        for d in range(1, days_in_m+1):
            dd = f"{d:02d}"
            is_hol = date(y, m, d).weekday() >= 5
            reg = ot2 = otx = 0.0

            # 早班；若缺 am‑out 則用 pm‑out；若有 pm‑in 就 skip 午休
            s_am = recs.get((dd, "am-in"))
            e_am = recs.get((dd, "am-out")) or recs.get((dd, "pm-out"))
            has_pm_in = recs.get((dd, "pm-in")) is not None
            if s_am and e_am:
                h = sum(calc_hours(s_am, e_am, brk, skip_break=has_pm_in))
                take_reg = min(h, 8-reg); reg += take_reg
                remain   = h - take_reg
                take2    = min(remain, 2-ot2); ot2 += take2; otx += remain - take2

            # 午班 (一定不扣午休)
            s_pm = recs.get((dd, "pm-in")); e_pm = recs.get((dd, "pm-out"))
            if s_pm and e_pm:
                h = sum(calc_hours(s_pm, e_pm, brk, skip_break=True))
                take_reg = min(h, 8-reg); reg += take_reg
                remain   = h - take_reg
                take2    = min(remain, 2-ot2); ot2 += take2; otx += remain - take2

            # 加班段
            s_ot = recs.get((dd, "ot-in")); e_ot = recs.get((dd, "ot-out"))
            if s_ot and e_ot:
                h = sum(calc_hours(s_ot, e_ot, brk, skip_break=True))
                take2 = min(h, 2-ot2); ot2 += take2; otx += h - take2

            hol = reg + ot2 + otx if is_hol else 0
            if is_hol: hol_sum += hol
            else:      reg_sum += reg; ot2_sum += ot2; otx_sum += otx
            if reg or ot2 or otx: wday += 1

            note  = notes.get(dd, "")
            style = (' style="background:#FFF2CC"' if note else
                     ' style="background:#DDDDDD"' if is_hol else '')

            def link(pt, val):
                day = f"{y}-{m:02d}-{dd}"                           # 先算出日期字串
                url = url_for("rec.edit_record",                    # 再呼叫 url_for
                               emp=emp.id,
                               date=day,
                               typ=pt,
                               back=back)
                return f'<a href="{url}">{val}</a>'     
            
            cells = "".join(f'<td>{link(pt, recs.get((dd, pt)) or "-")}</td>' for pt in cols)
            rows_html += (f'<tr{style}><td>{m:02d}-{dd}</td>{cells}'
                          f'<td>{link(LEAVE_PTYPE, note or "-")}</td>'
                          f'<td>{reg or ""}</td><td>{ot2 or ""}</td><td>{otx or ""}</td><td>{hol or ""}</td></tr>')

        total_row = (f'<tr><th>總計</th><th colspan="7"></th>'
                     f'<th>{reg_sum}</th><th>{ot2_sum}</th><th>{otx_sum}</th><th>{hol_sum}</th></tr>')

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

    return render_template_string(f'<!doctype html><html><head>{CSS}</head><body>'
                                  f'{"".join(html_parts)}</body></html>')

# ────────────────────── 單筆編輯 ──────────────────────
@rec_bp.route("/edit_rec", methods=["GET", "POST"])
def edit_record():
    if require("mgr"):
        return abort(403)

    emp_id = request.args.get("emp")
    dt     = request.args.get("date")
    typ    = request.args.get("typ")
    back   = request.args.get("back")

    rec = Checkin.query.filter_by(employee_id=emp_id, work_date=dt, p_type=typ).first()
    init_val = (rec.note if typ == LEAVE_PTYPE else rec.ts[11:16]) if rec else ""

    if request.method == "POST":
        if request.form.get("clear"):
            if rec:
                db.session.delete(rec); db.session.commit()
            return redirect(back)

        val = request.form.get("val", "").strip()
        if typ == LEAVE_PTYPE:
            if not val:
                return abort(400, "假別不可空白")
            if rec: rec.note = val
            else: db.session.add(Checkin(employee_id=emp_id, work_date=dt,
                                         p_type=LEAVE_PTYPE, ts=f"{dt}T00:00:00", note=val))
        else:
            if not re.fullmatch(r"\d{2}:\d{2}", val):
                return abort(400, "需輸入 HH:MM")
            ts = f"{dt}T{val}:00"
            if rec: rec.ts = ts
            else: db.session.add(Checkin(employee_id=emp_id, work_date=dt,
                                         p_type=typ, ts=ts))
        db.session.commit()
        return redirect(back)

    title_map = {"am-in":"上午上班","am-out":"上午下班","pm-in":"下午上班","pm-out":"下午下班",
                 "ot-in":"加班上班","ot-out":"加班下班", LEAVE_PTYPE:"備註 / 假別"}
    title = title_map.get(typ, "未知")

    return render_template_string(f"""<!doctype html><html><head>{CSS}</head><body>
<h2>{emp_id}　{dt}  {title}</h2>
<form method="post">
<input name="val" value="{init_val}" placeholder="HH:MM 或假別" style="width:180px"><br>
<button type="submit">儲存</button>
<button type="submit" name="clear" value="1"
        style="background:red;color:#fff;margin-left:10px">清除</button>
</form>
<p><a href="{back}">返回</a></p></body></html>""")
