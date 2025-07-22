# blueprints/export.py
# -*- coding: utf-8 -*-
"""
匯出各區域出勤表並合併至薪資範本
流程：
① 以 xlsxwriter 產生出勤工作表（記憶體）
② 讀取 app/static/薪資計算範本.xlsx
③ 將工作表插入範本（保留欄寬、合併格、樣式）
④ 回傳合併後檔案
※ 範本需事先移除不需要的工作表與圖形
"""

from flask import Blueprint, send_file, request, current_app
from datetime import date, timedelta
from extensions import db
from models import Employee, Checkin
from . import merge_night, calc_hours, NIGHT_END

import pandas as pd, io, calendar, re, openpyxl
from copy import copy
from pathlib import Path

exp_bp = Blueprint("exp", __name__, url_prefix="/admin")

LEAVE_PTYPE = 'lv'          # 與 records.py 一致的請假類型

# ─────────────── Sheet 複製工具 ────────────────
def clone_sheet(src_ws, tgt_wb, new_title=None):
    tgt_ws = tgt_wb.create_sheet(new_title or src_ws.title)

    # 欄寬
    for key, dim in src_ws.column_dimensions.items():
        if dim.width:
            tgt_ws.column_dimensions[key].width = dim.width

    # 合併格
    for rng in src_ws.merged_cells.ranges:
        tgt_ws.merge_cells(str(rng))

    # 值與樣式
    for row in src_ws.iter_rows():
        for cell in row:
            tgt = tgt_ws.cell(cell.row, cell.column, cell.value)
            if cell.has_style:
                if cell.font:          tgt.font = copy(cell.font)
                if cell.border:        tgt.border = copy(cell.border)
                if cell.fill:          tgt.fill = copy(cell.fill)
                if cell.number_format: tgt.number_format = cell.number_format
                if cell.alignment:     tgt.alignment = copy(cell.alignment)
    return tgt_ws
# ──────────────────────────────────────────────

@exp_bp.route("/export")
def export():
    # ── 解析 ym 參數 ──
    ym = request.args.get("ym")
    today = date.today()
    y, m = (map(int, ym.split("-"))
            if ym and re.fullmatch(r"\d{4}-\d{2}", ym)
            else (today.year, today.month))
    days = calendar.monthrange(y, m)[1]
    nxt  = (date(y, m, 1) + timedelta(days=32)).replace(day=1)

    # ── 產出出勤工作簿（xlsxwriter） ──
    buf_exp = io.BytesIO()
    writer  = pd.ExcelWriter(buf_exp, engine="xlsxwriter")
    book    = writer.book

    hdr_fmt   = book.add_format({'bold': True, 'border': 1, 'align': 'center', 'bg_color': '#D3D3D3'})
    cell_fmt  = book.add_format({'border': 1, 'align': 'center'})
    hol_fmt   = book.add_format({'border': 1, 'align': 'center', 'bg_color': '#FFF2CC'})
    leave_fmt = book.add_format({'border': 1, 'align': 'center', 'bg_color': '#FFFF00'})
    total_fmt = book.add_format({'bold': True, 'border': 1, 'align': 'center'})
    note_fmt  = book.add_format({'border': 1, 'align': 'left'})

    fields = ['上', '下', '備註 / 假別', '正班', '加班≤2', '加班>2', '假日', '出勤天數']

    areas = [a[0] for a in db.session.query(Employee.area)
             .filter(Employee.area.isnot(None)).distinct().order_by(Employee.area).all()]

    for area in areas:
        emps = Employee.query.filter_by(area=area).order_by(Employee.id).all()
        if not emps:
            continue
        ws = book.add_worksheet(area)

        # 標題列
        ws.write(0, 0, '日期', hdr_fmt)
        for idx, emp in enumerate(emps):
            s = 1 + idx * len(fields)
            e = s + len(fields) - 1
            ws.merge_range(0, s, 0, e, f"{emp.id}-{emp.name}", hdr_fmt)
            for j, f in enumerate(fields):
                ws.write(1, s + j, f, hdr_fmt)

        ws.set_column(0, 0, 10)
        for idx in range(len(emps)):
            ws.set_column(1 + idx*len(fields), 1 + idx*len(fields)+len(fields)-1, 12)
        ws.freeze_panes(2, 1)

        # 取資料
        emp_recs, emp_notes = {}, {}
        for emp in emps:
            raws = (Checkin.query
                    .with_entities(Checkin.work_date, Checkin.p_type, Checkin.ts, Checkin.note)
                    .filter(Checkin.employee_id == emp.id)
                    .filter((Checkin.work_date.like(f"{y}-{m:02d}%")) |
                            ((Checkin.work_date == nxt.isoformat()) &
                             (Checkin.p_type == 'out') &
                             (Checkin.ts < f"{nxt}T{NIGHT_END}:00")))
                    .order_by(Checkin.work_date, Checkin.p_type).all())
            emp_recs[emp.id] = merge_night([(r.work_date, r.p_type, r.ts[11:16]) for r in raws])
            emp_notes[emp.id] = {r.work_date[8:10]: (r.note or '請假')
                                 for r in raws if r.p_type == LEAVE_PTYPE}

        totals = {emp.id: {'reg':0,'ot2':0,'otx':0,'hol':0,'wday':0} for emp in emps}

        # 日明細
        for d in range(1, days + 1):
            dd = f"{d:02d}"
            row = 2 + d - 1
            curr = date(y, m, d)
            is_hol = curr.weekday() >= 5
            ws.write(row, 0, f"{m:02d}-{dd}", cell_fmt)

            for idx, emp in enumerate(emps):
                recs  = emp_recs[emp.id]
                notes = emp_notes[emp.id]
                inn = recs.get((dd, 'in'), "")
                out = recs.get((dd, 'out'), "")
                reg, ot2, otx = calc_hours(inn, out, emp.default_break)
                hol_hours = reg + ot2 + otx if is_hol else 0

                if reg or ot2 or otx or hol_hours:
                    totals[emp.id]['wday'] += 1
                totals[emp.id]['reg'] += reg
                totals[emp.id]['ot2'] += ot2
                totals[emp.id]['otx'] += otx
                totals[emp.id]['hol'] += hol_hours

                note_txt = notes.get(dd, '')
                leave_word = f"請{note_txt}" if note_txt else ''
                fmt = leave_fmt if note_txt else (hol_fmt if is_hol else cell_fmt)

                base = 1 + idx * len(fields)
                # 上、下
                ws.write(row, base,   inn, fmt)
                ws.write(row, base+1, out, fmt)
                # 備註 / 假別
                ws.write(row, base+2, leave_word, fmt)

                if note_txt and not inn and not out:
                    # 整天請假 → 工時留空
                    ws.write(row, base+3, '', fmt)
                    ws.write(row, base+4, '', fmt)
                    ws.write(row, base+5, '', fmt)
                    ws.write(row, base+6, '', fmt)
                else:
                    ws.write(row, base+3, '' if is_hol else reg, fmt)
                    ws.write(row, base+4, '' if is_hol else ot2, fmt)
                    ws.write(row, base+5, '' if is_hol else otx, fmt)
                    ws.write(row, base+6, hol_hours or '', fmt)

                ws.write(row, base+7, '', fmt)  # 出勤天數欄暫留空

        # 區域總計
        tr = 2 + days
        ws.write(tr, 0, "總計", total_fmt)
        for idx, emp in enumerate(emps):
            t = totals[emp.id]
            b = 1 + idx * len(fields)
            ws.write(tr, b+3, t['reg'],  total_fmt)
            ws.write(tr, b+4, t['ot2'],  total_fmt)
            ws.write(tr, b+5, t['otx'],  total_fmt)
            ws.write(tr, b+6, t['hol'],  total_fmt)
            ws.write(tr, b+7, t['wday'], total_fmt)

        # 備註彙總
        nr = tr + 1
        ws.write(nr, 0, "備註", hdr_fmt)
        for idx, emp in enumerate(emps):
            notes = emp_notes[emp.id]
            if not notes:
                continue
            items = [f"{m:02d}-{k} {v}" for k, v in sorted(notes.items())]
            txt = "；".join(items)
            s = 1 + idx * len(fields)
            e = s + len(fields) - 1
            ws.merge_range(nr, s, nr, e, txt, note_fmt)

    writer.close()
    buf_exp.seek(0)
    wb_exp = openpyxl.load_workbook(buf_exp)

    # ── 合併至薪資範本 ──
    tpl_path = Path(current_app.root_path) / "static" / "薪資計算範本.xlsx"
    wb_tpl = openpyxl.load_workbook(tpl_path)

    for ws in wb_exp.worksheets:
        if ws.title in wb_tpl.sheetnames:
            del wb_tpl[ws.title]
        clone_sheet(ws, wb_tpl)

    final_buf = io.BytesIO()
    wb_tpl.save(final_buf)
    final_buf.seek(0)

    return send_file(
        final_buf,
        as_attachment=True,
        download_name=f"{y}-{m:02d}_薪資報表.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
