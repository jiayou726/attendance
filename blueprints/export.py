# blueprints/export.py
# -*- coding: utf-8 -*-
"""
將指定月份的各區域員工出勤資料匯出為 Excel。
‧ 夜班併入規則、工時計算等邏輯依 common 工具函式。
‧ 區域工作表不再使用靜態 AREAS，改為動態抓取資料庫現有區域。
"""
from flask import Blueprint, send_file, request
from datetime import date, timedelta
from extensions import db
from models      import Employee, Checkin
from .           import merge_night, calc_hours, NIGHT_END   # ← 已移除 AREAS
import pandas as pd, io, calendar, re

exp_bp = Blueprint("exp", __name__, url_prefix="/admin")


@exp_bp.route("/export")
def exp():
    # ── 解析 ym 參數：預設當月 ──
    ym_param = request.args.get("ym")
    today    = date.today()
    if ym_param and re.fullmatch(r"\d{4}-\d{2}", ym_param):
        y, m = map(int, ym_param.split("-"))
    else:
        y, m = today.year, today.month

    days = calendar.monthrange(y, m)[1]
    nxt  = (date(y, m, 1) + timedelta(days=32)).replace(day=1)

    # ── 建立 Excel Writer ──
    buf    = io.BytesIO()
    writer = pd.ExcelWriter(buf, engine="xlsxwriter")
    book   = writer.book

    # ── 格式 ──
    hdr_fmt   = book.add_format({'bold': True, 'border': 1, 'align': 'center', 'bg_color': '#D3D3D3'})
    cell_fmt  = book.add_format({'border': 1, 'align': 'center'})
    hol_fmt   = book.add_format({'border': 1, 'align': 'center', 'bg_color': '#FFF2CC'})
    leave_fmt = book.add_format({'border': 1, 'align': 'center', 'bg_color': '#FFFF00'})
    total_fmt = book.add_format({'bold': True, 'border': 1, 'align': 'center'})
    note_fmt  = book.add_format({'border': 1, 'align': 'left'})

    fields = ['上', '下', '正班', '加班≤2', '加班>2', '假日', '出勤天數']

    # ── 依實際區域逐張工作表 ──
    areas = [
        a[0] for a in
        db.session.query(Employee.area).filter(Employee.area.isnot(None)).distinct().order_by(Employee.area).all()
    ]

    for area in areas:
        emps = Employee.query.filter_by(area=area).order_by(Employee.id).all()
        if not emps:
            continue

        ws = book.add_worksheet(area)

        # ── 標題列 ──
        ws.write(0, 0, '日期', hdr_fmt)
        for idx, emp in enumerate(emps):
            start = 1 + idx * len(fields)
            end   = start + len(fields) - 1
            ws.merge_range(0, start, 0, end, f"{emp.id}-{emp.name}", hdr_fmt)
            for j, f in enumerate(fields):
                ws.write(1, start + j, f, hdr_fmt)

        # 欄寬與凍結
        ws.set_column(0, 0, 10)
        for idx in range(len(emps)):
            ws.set_column(1 + idx*len(fields), 1 + idx*len(fields)+len(fields)-1, 12)
        ws.freeze_panes(2, 1)

        # ── 取資料 ──
        emp_recs, emp_notes = {}, {}
        for emp in emps:
            raws = (
                Checkin.query
                .with_entities(Checkin.work_date, Checkin.p_type, Checkin.ts, Checkin.note)
                .filter(Checkin.employee_id == emp.id)
                .filter(
                    (Checkin.work_date.like(f"{y}-{m:02d}%")) |
                    ((Checkin.work_date == nxt.isoformat()) &
                     (Checkin.p_type == "out") &
                     (Checkin.ts < f"{nxt}T{NIGHT_END}:00"))
                )
                .order_by(Checkin.work_date, Checkin.p_type)
                .all()
            )

            emp_recs[emp.id] = merge_night([(r.work_date, r.p_type, r.ts[11:16]) for r in raws])
            emp_notes[emp.id] = {r.work_date[8:10]: (r.note or '請假')
                                 for r in raws if r.p_type == 'leave'}

        totals = {emp.id: {'reg':0,'ot2':0,'otx':0,'hol':0,'wday':0} for emp in emps}

        # ── 日明細 ──
        for d in range(1, days + 1):
            dd   = f"{d:02d}"
            row  = 2 + d - 1
            curr = date(y, m, d)
            is_hol = curr.weekday() >= 5
            ws.write(row, 0, f"{m:02d}-{dd}", cell_fmt)

            for idx, emp in enumerate(emps):
                recs  = emp_recs[emp.id]
                notes = emp_notes[emp.id]

                inn   = recs.get((dd, 'in'), "")
                out   = recs.get((dd, 'out'), "")
                reg, ot2, otx = calc_hours(inn, out, emp.default_break)
                hol_hours = reg + ot2 + otx if is_hol else 0

                if reg or ot2 or otx or hol_hours:
                    totals[emp.id]['wday'] += 1
                totals[emp.id]['reg'] += reg
                totals[emp.id]['ot2'] += ot2
                totals[emp.id]['otx'] += otx
                totals[emp.id]['hol'] += hol_hours

                note_txt  = notes.get(dd, '')
                has_leave = bool(note_txt)
                fmt_main  = leave_fmt if has_leave else (hol_fmt if is_hol else cell_fmt)

                base = 1 + idx * len(fields)
                ws.write(row, base+0, inn, fmt_main)
                ws.write(row, base+1, out, fmt_main)

                if has_leave and not inn and not out:
                    ws.write(row, base+2, note_txt, fmt_main)
                    ws.write(row, base+3, '', fmt_main)
                    ws.write(row, base+4, '', fmt_main)
                else:
                    ws.write(row, base+2, "" if is_hol else reg, fmt_main)
                    ws.write(row, base+3, "" if is_hol else ot2, fmt_main)
                    ws.write(row, base+4, "" if is_hol else otx, fmt_main)

                ws.write(row, base+5, hol_hours or "", fmt_main)
                ws.write(row, base+6, "", fmt_main)

        # ── 區域總計 ──
        total_row = 2 + days
        ws.write(total_row, 0, "總計", total_fmt)
        for idx, emp in enumerate(emps):
            t    = totals[emp.id]
            base = 1 + idx * len(fields)
            ws.write(total_row, base+2, t['reg'],  total_fmt)
            ws.write(total_row, base+3, t['ot2'],  total_fmt)
            ws.write(total_row, base+4, t['otx'],  total_fmt)
            ws.write(total_row, base+5, t['hol'],  total_fmt)
            ws.write(total_row, base+6, t['wday'], total_fmt)

        # ── 備註列 ──
        note_row = total_row + 1
        ws.write(note_row, 0, "備註", hdr_fmt)
        for idx, emp in enumerate(emps):
            notes = emp_notes[emp.id]
            if not notes:
                continue
            items = [f"{m:02d}-{k} {v}" for k, v in sorted(notes.items())]
            txt   = "；".join(items)
            start = 1 + idx * len(fields)
            end   = start + len(fields) - 1
            ws.merge_range(note_row, start, note_row, end, txt, note_fmt)

    writer.close()
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name=f"{y}-{m:02d}_出勤報表.xlsx",
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
