# blueprints/export.py
# -*- coding: utf-8 -*-
"""
匯出兩種報表：

1) 薪資報表（合併範本）：/admin/export?ym=YYYY-MM
   - 各區一張 Sheet、每位員工 6 欄：
     『正班、加班≤2、加班>2、假日、出勤天數、備註／假別』
     備註／假別放在最後一欄。
   - 以 xlsxwriter 產生後，合併到 static/薪資計算範本.xlsx，保留欄寬/樣式/合併格。

2) 工時卡片總檔（每區每人各一張 Sheet）：/admin/export/punch_all?ym=YYYY-MM
   - 欄位：『日期、上午上、上午下、下午上、下午下、加班上、加班下』
   - 00:00 ~ NIGHT_END 的下班(out)歸前一日的「加班下」。
"""

from flask import Blueprint, send_file, request, current_app, abort
from datetime import date, timedelta, datetime
from extensions import db
from models import Employee, Checkin
from . import merge_night, calc_hours, NIGHT_END

import pandas as pd, io, calendar, re, openpyxl
from copy import copy
from pathlib import Path

exp_bp = Blueprint("exp", __name__, url_prefix="/admin")

LEAVE_PTYPE = 'lv'   # 與 records.py 一致的請假類型


# ─────────────── Sheet 複製工具（保留欄寬/樣式/合併格） ────────────────
def clone_sheet(src_ws, tgt_wb, new_title=None):
    tgt_ws = tgt_wb.create_sheet(new_title or src_ws.title)

    # 欄寬
    for key, dim in src_ws.column_dimensions.items():
        if dim.width is not None:
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
# ────────────────────────────────────────────────────────────────────────────


# ============================================================================
# 路由一：薪資報表（合併範本），各區多員工、每人 6 欄
# 欄位順序：正班, 加班≤2, 加班>2, 假日, 出勤天數, 備註／假別（備註最後）
# ============================================================================
@exp_bp.route("/export")
def export():
    ym = request.args.get("ym")
    today = date.today()
    y, m = (map(int, ym.split("-"))
            if ym and re.fullmatch(r"\d{4}-\d{2}", ym)
            else (today.year, today.month))
    days = calendar.monthrange(y, m)[1]
    nxt  = (date(y, m, 1) + timedelta(days=32)).replace(day=1)

    # 以 xlsxwriter 先產生各區工作簿
    buf_exp = io.BytesIO()
    writer  = pd.ExcelWriter(buf_exp, engine="xlsxwriter")
    book    = writer.book

    hdr_fmt   = book.add_format({'bold': True, 'border': 1, 'align': 'center', 'bg_color': '#D3D3D3'})
    cell_fmt  = book.add_format({'border': 1, 'align': 'center'})
    hol_fmt   = book.add_format({'border': 1, 'align': 'center', 'bg_color': '#FFF2CC'})
    leave_fmt = book.add_format({'border': 1, 'align': 'center', 'bg_color': '#FFFF00'})
    total_fmt = book.add_format({'bold': True, 'border': 1, 'align': 'center'})
    note_fmt  = book.add_format({'border': 1, 'align': 'left'})

    fields = ['正班', '加班≤2', '加班>2', '假日', '出勤天數', '備註 / 假別']

    # 依區域列出
    areas = [a[0] for a in (db.session.query(Employee.area)
                            .filter(Employee.area.isnot(None))
                            .distinct().order_by(Employee.area).all())]

    for area in areas:
        emps = (Employee.query.filter_by(area=area)
                .order_by(Employee.id).all())
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
                    .order_by(Checkin.work_date, Checkin.p_type, Checkin.ts).all())

            # 合併跨夜，得出 (dd,'in') / (dd,'out')
            emp_recs[emp.id] = merge_night([(r.work_date, r.p_type, r.ts[11:16]) for r in raws])

            # 備註/假別（以 dd -> 文字）
            emp_notes[emp.id] = {r.work_date[8:10]: (r.note or '請假')
                                 for r in raws if r.p_type == LEAVE_PTYPE}

        totals = {emp.id: {'reg':0.0,'ot2':0.0,'otx':0.0,'hol':0.0,'wday':0} for emp in emps}

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
                hol_hours = reg + ot2 + otx if is_hol else 0.0

                worked_today = bool(reg or ot2 or otx or hol_hours)
                if worked_today:
                    totals[emp.id]['wday'] += 1
                totals[emp.id]['reg'] += reg
                totals[emp.id]['ot2'] += ot2
                totals[emp.id]['otx'] += otx
                totals[emp.id]['hol'] += hol_hours

                note_txt = notes.get(dd, '')
                leave_word = f"請{note_txt}" if note_txt else ''
                fmt = leave_fmt if note_txt else (hol_fmt if is_hol else cell_fmt)

                base = 1 + idx * len(fields)

                # 正班, 加班≤2, 加班>2, 假日, 出勤天數(每日留空), 備註/假別
                if note_txt and not inn and not out:
                    ws.write(row, base+0, '', fmt)
                    ws.write(row, base+1, '', fmt)
                    ws.write(row, base+2, '', fmt)
                    ws.write(row, base+3, '', fmt)
                    ws.write(row, base+4, '', fmt)
                    ws.write(row, base+5, leave_word, fmt)
                else:
                    ws.write(row, base+0, '' if is_hol else reg, fmt)
                    ws.write(row, base+1, '' if is_hol else ot2, fmt)
                    ws.write(row, base+2, '' if is_hol else otx, fmt)
                    ws.write(row, base+3, hol_hours or '', fmt)
                    ws.write(row, base+4, '', fmt)
                    ws.write(row, base+5, leave_word, fmt)

        # 區域總計列
        tr = 2 + days
        ws.write(tr, 0, "總計", total_fmt)
        for idx, emp in enumerate(emps):
            t = totals[emp.id]; b = 1 + idx * len(fields)
            ws.write(tr, b+0, t['reg'],  total_fmt)
            ws.write(tr, b+1, t['ot2'],  total_fmt)
            ws.write(tr, b+2, t['otx'],  total_fmt)
            ws.write(tr, b+3, t['hol'],  total_fmt)
            ws.write(tr, b+4, t['wday'], total_fmt)
            ws.write(tr, b+5, '',        total_fmt)

        # 備註彙總列
        nr = tr + 1
        ws.write(nr, 0, "備註", hdr_fmt)
        for idx, emp in enumerate(emps):
            notes = emp_notes[emp.id]
            if not notes:
                continue
            items = [f"{m:02d}-{k} {v}" for k, v in sorted(notes.items())]
            txt = "；".join(items)
            s = 1 + idx * len(fields); e = s + len(fields) - 1
            ws.merge_range(nr, s, nr, e, txt, note_fmt)

    writer.close()
    buf_exp.seek(0)
    wb_exp = openpyxl.load_workbook(buf_exp)

    # 合併到薪資範本
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


# ============================================================================
# 路由二：工時卡片總檔（每區每人一張 Sheet）
# 欄位：日期, 上午上, 上午下, 下午上, 下午下, 加班上, 加班下
# 00:00 ~ NIGHT_END 的 out 歸前一日的加班下
# ============================================================================
def _hm(ts: str) -> str:
    """將 'YYYY-MM-DDTHH:MM:SS' 取 'HH:MM'。若已是 'HH:MM' 則原樣回傳。"""
    if len(ts) >= 16 and 'T' in ts:
        return ts[11:16]
    if len(ts) == 5 and ts[2] == ':':
        return ts
    try:
        return datetime.fromisoformat(ts).strftime("%H:%M")
    except Exception:
        return ts

def _assign_pairs_for_day(records):
    """
    將當日 records（已依時間排序且僅含 in/out/lv）盡量配對成三組：
      (m_in, m_out), (a_in, a_out), (ot_in, ot_out)
    規則：依出現順序填入三個 pair；out 會填入最後一個尚未完成 out 的 pair。
    """
    pairs = [{'in': '', 'out': ''}, {'in': '', 'out': ''}, {'in': '', 'out': ''}]

    def find_place_for_out():
        for i in (2, 1, 0):
            if pairs[i]['in'] and not pairs[i]['out']:
                return i
        return None

    for r in records:
        if r['p_type'] == 'lv':
            continue
        if r['p_type'] == 'in':
            placed = False
            for i in range(3):
                if not pairs[i]['in']:
                    pairs[i]['in'] = r['ts']; placed = True; break
                if pairs[i]['in'] and pairs[i]['out']:
                    continue
            # 多餘的 in 忽略
        elif r['p_type'] == 'out':
            i = find_place_for_out()
            if i is not None:
                pairs[i]['out'] = r['ts']
            # 孤立 out 忽略

    return (pairs[0]['in'], pairs[0]['out'],
            pairs[1]['in'], pairs[1]['out'],
            pairs[2]['in'], pairs[2]['out'])

@exp_bp.route("/export/punch_all")
def export_punch_all():
    ym = request.args.get("ym")
    today = date.today()
    y, m = (map(int, ym.split("-"))
            if ym and re.fullmatch(r"\d{4}-\d{2}", ym)
            else (today.year, today.month))
    days = calendar.monthrange(y, m)[1]
    month_start = date(y, m, 1)
    nxt  = (month_start + timedelta(days=32)).replace(day=1)

    # 取所有員工，依區域、員工編號排序
    emps = Employee.query.order_by(Employee.area, Employee.id).all()

    # 取所有原始打卡
    all_raws = (Checkin.query
                .with_entities(Checkin.employee_id, Checkin.work_date, Checkin.p_type, Checkin.ts, Checkin.note)
                .filter((Checkin.work_date.like(f"{y}-{m:02d}%")) |
                        ((Checkin.work_date == nxt.isoformat()) &
                         (Checkin.p_type == 'out') &
                         (Checkin.ts < f"{nxt}T{NIGHT_END}:00")))
                .order_by(Checkin.employee_id, Checkin.work_date, Checkin.ts).all())

    # 依員工分組
    by_emp = {}
    for r in all_raws:
        by_emp.setdefault(r.employee_id, []).append(r)

    night_bar = f"{NIGHT_END}:00" if len(NIGHT_END) == 2 else NIGHT_END

    # 產出 Excel（每位員工一張表）
    buf = io.BytesIO()
    writer = pd.ExcelWriter(buf, engine="xlsxwriter")
    book = writer.book

    title_fmt = book.add_format({'bold': True, 'align': 'center', 'font_size': 14})
    sub_fmt   = book.add_format({'align': 'center'})
    hdr_fmt   = book.add_format({'bold': True, 'border': 1, 'align': 'center', 'bg_color': '#D3D3D3'})
    cell_fmt  = book.add_format({'border': 1, 'align': 'center'})
    hol_fmt   = book.add_format({'border': 1, 'align': 'center', 'bg_color': '#EDEDED'})
    total_fmt = book.add_format({'bold': True, 'border': 1, 'align': 'center'})

    for emp in emps:
        sheet_name = f"{emp.area or ''}-{emp.id}-{emp.name}"
        ws = book.add_worksheet(sheet_name[:31])  # Excel 限 31 字

        # 標題與表頭
        ws.merge_range(0, 0, 0, 6, f"{emp.name}（{emp.id}） 區域：{emp.area or ''}  {y}/{m:02d}", title_fmt)
        ws.merge_range(1, 0, 1, 6, "出勤天數：", sub_fmt)  # 先合併一次，稍後只寫入不再合併

        headers = ['日期', '上午上', '上午下', '下午上', '下午下', '加班上', '加班下']
        for c, h in enumerate(headers):
            ws.write(2, c, h, hdr_fmt)

        ws.set_column(0, 0, 10)
        ws.set_column(1, 6, 12)
        ws.freeze_panes(3, 1)

        # 整理該員工的每日紀錄（處理跨夜 out -> 前一日）
        rec_by_day = {(date(y, m, d)).isoformat(): [] for d in range(1, days+1)}
        raws = by_emp.get(emp.id, [])
        for r in raws:
            hm = _hm(r.ts)
            if r.p_type == 'out' and hm < night_bar:
                od = datetime.fromisoformat(r.work_date).date()
                target = (od - timedelta(days=1)).isoformat()
            else:
                target = r.work_date

            if target[:7] != f"{y}-{m:02d}":
                continue

            rec_by_day[target].append({'p_type': r.p_type, 'ts': hm})

        for k in rec_by_day:
            rec_by_day[k].sort(key=lambda x: x['ts'])

        attend_days = 0
        for d in range(1, days+1):
            the_date = date(y, m, d)
            key = the_date.isoformat()
            row = 3 + (d - 1)

            fmt = hol_fmt if the_date.weekday() >= 5 else cell_fmt
            m_in, m_out, a_in, a_out, ot_in, ot_out = _assign_pairs_for_day(rec_by_day.get(key, []))

            if any([m_in, m_out, a_in, a_out, ot_in, ot_out]):
                attend_days += 1

            ws.write(row, 0, f"{m:02d}-{d:02d}", fmt)
            ws.write(row, 1, m_in or '', fmt)
            ws.write(row, 2, m_out or '', fmt)
            ws.write(row, 3, a_in or '', fmt)
            ws.write(row, 4, a_out or '', fmt)
            ws.write(row, 5, ot_in or '', fmt)
            ws.write(row, 6, ot_out or '', fmt)

        tr = 3 + days
        ws.write(tr, 0, "總計", total_fmt)
        for c in range(1, 7):
            ws.write(tr, c, "", total_fmt)

        # 回填第二行的出勤天數（已經 merge 過，這裡只寫左上角 A2）
        ws.write(1, 0, f"出勤天數：{attend_days}", sub_fmt)

    writer.close()
    buf.seek(0)

    fname = f"{y}-{m:02d}_工時卡片_全員.xlsx"
    return send_file(
        buf,
        as_attachment=True,
        download_name=fname,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

