# blueprints/records.py
# -*- coding: utf-8 -*-
"""
行政後台：出勤月表與單筆編輯（支援六段打卡）

表格欄位
──────────────
| 日期 | 上午上 | 上午下 | 下午上 | 下午下 | 加班上 | 加班下 | 備註 | 正班 | 加班≤2 | 加班>2 | 假日 |
"""
from flask import Blueprint, render_template_string, request, redirect, url_for, abort
from datetime import date, timedelta
import re

from extensions import db
from models import Employee, Checkin
from . import CSS, merge_night, calc_hours, NIGHT_END

rec_bp = Blueprint("rec", __name__, url_prefix="/admin")

LEAVE_PTYPE = "lv"        # 備註／假別用

# ────────────────────────────────── 權限檢查（示範） ──────────────────────────────────
def require(role: str):
    """範例用：永遠通過。實際應檢查登入者角色。"""
    return None

# ──────────────────────────────── 出勤月表 ────────────────────────────────
@rec_bp.route("/records")
def show_records():
    if require("mgr"):
        return abort(403)

    # 參數
    eid      = request.args.get("eid", "")
    ym_param = request.args.get("ym")
    today    = date.today()

    # 決定年月
    if ym_param and re.fullmatch(r"\d{4}-\d{2}", ym_param):
        y, m = map(int, ym_param.split("-"))
        ym   = ym_param
    else:
        y, m = today.year, today.month
        ym   = f"{y}-{m:02d}"

    # 取得員工
    emp = Employee.query.get(eid) if eid else None
    brk = emp.default_break if emp else 0.0

    # 員工下拉
    emp_opts = "".join(
        f'<option value="{e.id}" {"selected" if str(e.id)==eid else ""}>{e.id}-{e.name}</option>'
        for e in Employee.query.order_by(Employee.id)
    )

    # 月份下拉（近 12 個月）
    ym_opts, cursor = [], today.replace(day=1)
    for _ in range(12):
        v = cursor.strftime("%Y-%m")
        l = cursor.strftime("%Y / %m")
        sel = "selected" if v == ym else ""
        ym_opts.append(f'<option value="{v}" {sel}>{l}</option>')
        cursor = (cursor - timedelta(days=1)).replace(day=1)

    form_html = (
        '<form method="get" id="recForm">'
        '員工：<select name="eid" onchange="recForm.submit()">'
        '<option></option>' + emp_opts + '</select>'
        '　月份：<select name="ym" onchange="recForm.submit()">' + "".join(ym_opts) + '</select>'
        '</form>'
    )

    # 未選員工 → 僅顯示查詢欄
    if not emp:
        return render_template_string(
            f'<!doctype html><html><head>{CSS}</head><body>'
            f'<h2>出勤卡查詢</h2>{form_html}</body></html>'
        )

    # 當月日期範圍
    first_day = date(y, m, 1)
    num_days  = ((first_day.replace(day=28) + timedelta(days=4)).replace(day=1) - first_day).days
    next_m    = (first_day + timedelta(days=32)).replace(day=1)

    # 讀取打卡紀錄（含跨日下班）
    raws = (
        Checkin.query
        .with_entities(Checkin.work_date, Checkin.p_type, Checkin.ts, Checkin.note)
        .filter(Checkin.employee_id == eid)
        .filter(
            (Checkin.work_date.like(f"{y}-{m:02d}%")) |
            (
                (Checkin.work_date == next_m.isoformat()) &
                (Checkin.p_type.in_(["am-out", "pm-out", "ot-out"])) &
                (Checkin.ts < f"{next_m}T{NIGHT_END}:00")
            )
        )
        .order_by(Checkin.work_date, Checkin.p_type)
        .all()
    )

    # dict {(dd, p_type): "HH:MM"}
    recs  = merge_night([(r.work_date, r.p_type, r.ts[11:16]) for r in raws])
    notes = {r.work_date[8:10]: (r.note or "請假") for r in raws if r.p_type == LEAVE_PTYPE}

    punch_cols = ["am-in", "am-out", "pm-in", "pm-out", "ot-in", "ot-out"]

    rows_html = ""
    wday = reg_sum = ot2_sum = otx_sum = hol_sum = 0
    back_url = url_for("rec.show_records", eid=eid, ym=ym)

    for d in range(1, num_days + 1):
        curr   = date(y, m, d)
        is_hol = curr.weekday() >= 5
        dd     = f"{d:02d}"

        # ---------- 工時計算 ----------
        reg = ot2 = otx = 0.0
        # 1) 正常三組配對
        for pin, pout in [("am-in","am-out"),("pm-in","pm-out"),("ot-in","ot-out")]:
            s = recs.get((dd, pin))
            e = recs.get((dd, pout))
            if s and e:
                r, o2, ox = calc_hours(s, e, brk)
                reg += r; ot2 += o2; otx += ox

        # 2) Fallback：若完全沒成對，用最早 in + 最晚 out
        if (reg + ot2 + otx) == 0:
            first_in = next((recs.get((dd, t))
                             for t in ["am-in","pm-in","ot-in"] if recs.get((dd, t))), None)
            last_out = next((recs.get((dd, t))
                             for t in ["ot-out","pm-out","am-out"] if recs.get((dd, t))), None)
            if first_in and last_out:
                r, o2, ox = calc_hours(first_in, last_out, brk)
                reg += r; ot2 += o2; otx += ox
                
        hol_hours = (reg + ot2 + otx) if is_hol else 0

        # 累計
        if is_hol:
            hol_sum += hol_hours
        else:
            reg_sum += reg; ot2_sum += ot2; otx_sum += otx
        if reg or ot2 or otx:
            wday += 1

        # 顯示樣式
        note = notes.get(dd, "")
        if note:
            tr_style = ' style="background:#FFF2CC"'
        elif is_hol:
            tr_style = ' style="background:#DDDDDD"'
        else:
            tr_style = ''

        # 產生可編輯連結
        def link(ptype: str, label: str):
            day = f"{y}-{m:02d}-{dd}"
            return (
                f'<a href="{url_for("rec.edit_record", emp=eid, date=day, typ=ptype, back=back_url)}">'
                f'{label}</a>'
            )

        cells = "".join(
            f'<td>{link(pt, recs.get((dd, pt), "") or "-")}</td>' for pt in punch_cols
        )

        if is_hol:
            reg_cell = ot2_cell = otx_cell = ''
            hol_cell = hol_hours or ''
        else:
            reg_cell, ot2_cell, otx_cell = reg or '', ot2 or '', otx or ''
            hol_cell = ''

        rows_html += (
            f'<tr{tr_style}><td>{m:02d}-{dd}</td>{cells}'
            f'<td>{link(LEAVE_PTYPE, note or "-")}</td>'
            f'<td>{reg_cell}</td><td>{ot2_cell}</td><td>{otx_cell}</td><td>{hol_cell}</td></tr>'
        )

    total_row = (
        f'<tr><th>總計</th><th colspan="7"></th>'
        f'<th>{reg_sum}</th><th>{ot2_sum}</th><th>{otx_sum}</th><th>{hol_sum}</th></tr>'
    )

    return render_template_string(f"""<!doctype html><html><head>{CSS}</head><body>
<h2>{emp.name}（{emp.id}） 區域：{emp.area}　{y}/{m}</h2>
<h3>出勤天數：{wday}</h3>{form_html}
<table>
<tr><th>日期</th><th>上午上</th><th>上午下</th><th>下午上</th><th>下午下</th><th>加班上</th><th>加班下</th><th>備註</th>
    <th>正班</th><th>加班≤2</th><th>加班&gt;2</th><th>假日</th></tr>
{rows_html}{total_row}</table>
<p><a href="{url_for('exp.export', eid=eid, ym=ym)}">匯出 Excel</a> |
   <a href="{url_for('emp.list_employees')}">返回員工管理</a></p>
</body></html>""")

# ──────────────────────────────── 編輯單筆 ────────────────────────────────
@rec_bp.route("/edit_rec", methods=["GET", "POST"])
def edit_record():
    if require("mgr"):
        return abort(403)

    emp_id = request.args.get("emp")
    dt     = request.args.get("date")     # YYYY-MM-DD
    typ    = request.args.get("typ")     # am-in / am-out / … / lv
    back   = request.args.get("back")

    rec = (
        Checkin.query
        .filter_by(employee_id=emp_id, work_date=dt, p_type=typ)
        .first()
    )

    init_val = (rec.note if typ == LEAVE_PTYPE else rec.ts[11:16]) if rec else ""

    # ---------------- POST ----------------
    if request.method == "POST":
        if request.form.get("clear"):          # 清除
            if rec:
                db.session.delete(rec); db.session.commit()
            return redirect(back)

        val = request.form.get("val", "").strip()

        if typ == LEAVE_PTYPE:                 # 備註 / 假別
            if not val:
                return abort(400, "假別不可空白")
            if rec:
                rec.note = val
            else:
                db.session.add(
                    Checkin(
                        employee_id=emp_id,
                        work_date=dt,
                        p_type=LEAVE_PTYPE,
                        ts=f"{dt}T00:00:00",
                        note=val
                    )
                )
        else:                                  # 六段打卡時間
            if not re.fullmatch(r"[0-9]{2}:[0-9]{2}", val):
                return abort(400, "請輸入 HH:MM 格式")
            full_ts = f"{dt}T{val}:00"
            if rec:
                rec.ts = full_ts
            else:
                db.session.add(
                    Checkin(
                        employee_id=emp_id,
                        work_date=dt,
                        p_type=typ,
                        ts=full_ts
                    )
                )

        db.session.commit()
        return redirect(back)

    # ---------------- GET：顯示表單 ----------------
    title_map = {
        "am-in":"上午上班","am-out":"上午下班",
        "pm-in":"下午上班","pm-out":"下午下班",
        "ot-in":"加班上班","ot-out":"加班下班",
        LEAVE_PTYPE:"備註 / 假別"
    }
    title = title_map.get(typ, "未知")

    return render_template_string(f"""<!doctype html><html><head>{CSS}</head><body>
<h2>{emp_id}　{dt}　{title}</h2>
<form method="post">
<input name="val" value="{init_val}" placeholder="HH:MM 或假別文字" style="width:180px"><br>
<button type="submit">儲存</button>
<button type="submit" name="clear" value="1"
        style="background:red;color:#fff;margin-left:10px">清除</button>
</form>
<p><a href="{back}">返回</a></p></body></html>""")
