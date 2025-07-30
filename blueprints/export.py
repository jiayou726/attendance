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


# =====================================================================
# 路由一：薪資報表（合併範本）── 與月表相同的分段算法
# =====================================================================
@exp_bp.route("/export")
def export():
    ym = request.args.get("ym")
    today = date.today()
    y, m = (map(int, ym.split("-"))
            if ym and re.fullmatch(r"\d{4}-\d{2}", ym)
            else (today.year, today.month))
    days = calendar.monthrange(y, m)[1]
    nxt  = (date(y, m, 1) + timedelta(days=32)).replace(day=1)

    # ── 建立活頁簿 ──
    buf = io.BytesIO()
    writer = pd.ExcelWriter(buf, engine="xlsxwriter")
    book = writer.book

    hdr_fmt   = book.add_format({'bold': True, 'border': 1, 'align': 'center', 'bg_color': '#D3D3D3'})
    cell_fmt  = book.add_format({'border': 1, 'align': 'center'})
    hol_fmt   = book.add_format({'border': 1, 'align': 'center', 'bg_color': '#FFF2CC'})
    leave_fmt = book.add_format({'border': 1, 'align': 'center', 'bg_color': '#FFFF00'})
    total_fmt = book.add_format({'bold': True, 'border': 1, 'align': 'center'})
    note_fmt  = book.add_format({'border': 1, 'align': 'left'})

    fields = ['正班', '加班≤2', '加班>2', '假日', '出勤天數', '備註 / 假別']

    areas = [a[0] for a in (db.session.query(Employee.area)
                            .filter(Employee.area.isnot(None))
                            .distinct().order_by(Employee.area).all())]

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
            ws.set_column(1+idx*len(fields), 1+idx*len(fields)+len(fields)-1, 12)
        ws.freeze_panes(2, 1)

        # 取打卡資料
        emp_raws, emp_notes = {}, {}
        for emp in emps:
            raws = (Checkin.query
                    .with_entities(Checkin.work_date, Checkin.p_type, Checkin.ts, Checkin.note)
                    .filter(Checkin.employee_id == emp.id)
                    .filter((Checkin.work_date.like(f"{y}-{m:02d}%")) |
                            ((Checkin.work_date == nxt.isoformat()) &
                             (Checkin.p_type.like('%-out')) &
                             (Checkin.ts < f"{nxt}T{NIGHT_END}:00")))
                    .order_by(Checkin.work_date, Checkin.p_type, Checkin.ts)
                    .all())

            # recs：(dd, p_type) -> HH:MM，00:00–02:59 out 歸前日
            recs = {}
            for r in raws:
                hm = r.ts[11:16]
                hour = int(hm[:2])
                adj_date = (datetime.fromisoformat(r.work_date).date() - timedelta(days=1)).isoformat() \
                           if r.p_type.endswith('-out') and hour <= 2 else r.work_date
                if not adj_date.startswith(f"{y}-{m:02d}"):
                    continue
                dd = adj_date[8:10]
                recs[(dd, r.p_type)] = hm
            emp_raws[emp.id] = recs

            emp_notes[emp.id] = {r.work_date[8:10]: (r.note or '請假')
                                 for r in raws if r.p_type == LEAVE_PTYPE}

        totals = {emp.id: {'reg':0.0,'ot2':0.0,'otx':0.0,'hol':0.0,'wday':0} for emp in emps}

        # 逐日寫入
        for d in range(1, days+1):
            dd = f"{d:02d}"
            row = 2 + d - 1
            curr = date(y, m, d)
            is_hol = curr.weekday() >= 5
            ws.write(row, 0, f"{m:02d}-{dd}", cell_fmt)

            for idx, emp in enumerate(emps):
                recs  = emp_raws[emp.id]
                notes = emp_notes[emp.id]

                s_am = recs.get((dd, 'am-in'))
                e_am = recs.get((dd, 'am-out')) or recs.get((dd, 'pm-out'))
                has_pm_in = recs.get((dd, 'pm-in')) is not None

                reg = ot2 = otx = 0.0
                if s_am and e_am:
                    h = sum(calc_hours(s_am, e_am, emp.default_break, skip_break=has_pm_in))
                    take_reg = min(h, 8 - reg); reg += take_reg
                    remain   = h - take_reg
                    take2    = min(remain, 2 - ot2); ot2 += take2; otx += remain - take2

                s_pm = recs.get((dd, 'pm-in')); e_pm = recs.get((dd, 'pm-out'))
                if s_pm and e_pm:
                    h = sum(calc_hours(s_pm, e_pm, emp.default_break, skip_break=True))
                    take_reg = min(h, 8 - reg); reg += take_reg
                    remain   = h - take_reg
                    take2    = min(remain, 2 - ot2); ot2 += take2; otx += remain - take2

                s_ot = recs.get((dd, 'ot-in')); e_ot = recs.get((dd, 'ot-out'))
                if s_ot and e_ot:
                    h = sum(calc_hours(s_ot, e_ot, emp.default_break, skip_break=True))
                    take2 = min(h, 2 - ot2); ot2 += take2; otx += h - take2

                hol_hours = reg + ot2 + otx if is_hol else 0.0

                if reg or ot2 or otx or hol_hours:
                    totals[emp.id]['wday'] += 1
                totals[emp.id]['reg'] += reg
                totals[emp.id]['ot2'] += ot2
                totals[emp.id]['otx'] += otx
                totals[emp.id]['hol'] += hol_hours

                note_txt = notes.get(dd, '')
                leave_word = f"請{note_txt}" if note_txt else ''
                fmt = leave_fmt if note_txt else (hol_fmt if is_hol else cell_fmt)

                base = 1 + idx*len(fields)
                if note_txt and not (reg or ot2 or otx or hol_hours):
                    ws.write_row(row, base, ['']*5 + [leave_word], fmt)
                else:
                    ws.write_row(row, base, [
                        '' if is_hol else reg,
                        '' if is_hol else ot2,
                        '' if is_hol else otx,
                        hol_hours or '',
                        '',
                        leave_word
                    ], fmt)

        # 區域總計
        tr = 2 + days
        ws.write(tr, 0, "總計", total_fmt)
        for idx, emp in enumerate(emps):
            t = totals[emp.id]; b = 1 + idx*len(fields)
            ws.write_row(tr, b, [t['reg'], t['ot2'], t['otx'], t['hol'], t['wday'], ''], total_fmt)

        # 備註彙總
        nr = tr + 1
        ws.write(nr, 0, "備註", hdr_fmt)
        for idx, emp in enumerate(emps):
            notes = emp_notes[emp.id]
            if not notes:
                continue
            items = [f"{m:02d}-{k} {v}" for k, v in sorted(notes.items())]
            txt = '；'.join(items)
            s = 1 + idx*len(fields); e = s + len(fields)-1
            ws.merge_range(nr, s, nr, e, txt, note_fmt)

    writer.close()
    buf.seek(0)
    wb_new = openpyxl.load_workbook(buf)

    # 合併到範本
    tpl = Path(current_app.root_path) / "static" / "薪資計算範本.xlsx"
    wb_tpl = openpyxl.load_workbook(tpl)
    for s in wb_new.worksheets:
        if s.title in wb_tpl.sheetnames:
            del wb_tpl[s.title]
        clone_sheet(s, wb_tpl)

    out_buf = io.BytesIO()
    wb_tpl.save(out_buf)
    out_buf.seek(0)

    return send_file(
        out_buf,
        as_attachment=True,
        download_name=f"{y}-{m:02d}_薪資報表.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ======================================================================
# 路由二：工時卡片總檔（每區每人各一張 Sheet）
# 欄位：日期, 上午上, 上午下, 下午上, 下午下, 加班上, 加班下
# 00:00 ~ NIGHT_END 的 out 歸前一日的加班下
# ======================================================================

def _hm(ts: str) -> str:
    """'YYYY-MM-DDTHH:MM:SS' → 'HH:MM'；若已是 'HH:MM' 直接回傳。"""
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
    將已排序的 records（僅含 in/out/lv）配對成三組：
      (m_in, m_out), (a_in, a_out), (ot_in, ot_out)
    規則：依出現順序填入三 pair；out 填入最後一個未完成 out 的 pair。
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
            for i in range(3):
                if not pairs[i]['in']:
                    pairs[i]['in'] = r['ts']
                    break
        elif r['p_type'] == 'out':
            idx = find_place_for_out()
            if idx is not None:
                pairs[idx]['out'] = r['ts']

    return (pairs[0]['in'], pairs[0]['out'],
            pairs[1]['in'], pairs[1]['out'],
            pairs[2]['in'], pairs[2]['out'])


# ============================================================================
# 路由二：工時卡片總檔（每區每人各一張 Sheet）
# 6 段打卡；00:00–02:59 的 *-out 歸前一天
# ============================================================================

def _hm(ts: str) -> str:
    """'YYYY-MM-DDTHH:MM:SS' → 'HH:MM'；若已是 'HH:MM' 直接回傳。"""
    if len(ts) >= 16 and 'T' in ts:
        return ts[11:16]
    if len(ts) == 5 and ts[2] == ':':
        return ts
    try:
        return datetime.fromisoformat(ts).strftime("%H:%M")
    except Exception:
        return ts                         # 未查證

# 六段類型 → 欄位序
_COL_MAP = {
    'am-in': 1, 'am-out': 2,
    'pm-in': 3, 'pm-out': 4,
    'ot-in': 5, 'ot-out': 6,
}

@exp_bp.route("/export/punch_all")
def export_punch_all():
    ym = request.args.get("ym")
    today = date.today()
    y, m = (map(int, ym.split("-"))
            if ym and re.fullmatch(r"\d{4}-\d{2}", ym)
            else (today.year, today.month))
    days = calendar.monthrange(y, m)[1]

    emps = Employee.query.order_by(Employee.area, Employee.id).all()

    all_raws = (Checkin.query
                .with_entities(Checkin.employee_id, Checkin.work_date,
                               Checkin.p_type, Checkin.ts)
                .filter(Checkin.work_date.like(f"{y}-{m:02d}%"))
                .order_by(Checkin.employee_id, Checkin.work_date, Checkin.ts)
                .all())

    # 依員工分組
    by_emp = {}
    for r in all_raws:
        by_emp.setdefault(r.employee_id, []).append(r)

    # ---------------- Excel 建立 ----------------
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
        ws = book.add_worksheet(f"{emp.area or ''}-{emp.id}-{emp.name}"[:31])

        # 標題
        ws.merge_range(0, 0, 0, 6,
                       f"{emp.name}（{emp.id}） 區域：{emp.area or ''}  {y}/{m:02d}",
                       title_fmt)
        ws.merge_range(1, 0, 1, 6, "出勤天數：", sub_fmt)

        headers = ['日期', '上午上', '上午下', '下午上', '下午下', '加班上', '加班下']
        for c, h in enumerate(headers):
            ws.write(2, c, h, hdr_fmt)

        ws.set_column(0, 0, 10)
        ws.set_column(1, 6, 12)
        ws.freeze_panes(3, 1)

        # 為每一天預留 6 欄
        rec_by_day = {date(y, m, d).isoformat(): [''] * 6
                      for d in range(1, days + 1)}

        for r in by_emp.get(emp.id, []):
            ts_dt = datetime.fromisoformat(r.ts)
            hm = ts_dt.strftime("%H:%M")

            # 00:00–02:59 的 out 歸前一天
            if r.p_type.endswith('-out') and ts_dt.hour <= 2:
                tgt_date = (ts_dt.date() - timedelta(days=1)).isoformat()
            else:
                tgt_date = r.work_date

            # 只處理本月範圍
            if tgt_date[:7] != f"{y}-{m:02d}":
                continue

            col = _COL_MAP.get(r.p_type)
            if col:
                rec_by_day[tgt_date][col - 1] = hm

        # 寫入表格
        attend_days = 0
        for d in range(1, days + 1):
            the_date = date(y, m, d)
            row = 3 + d - 1
            fmt = hol_fmt if the_date.weekday() >= 5 else cell_fmt
            cols = rec_by_day[the_date.isoformat()]

            if any(cols):
                attend_days += 1

            ws.write(row, 0, f"{m:02d}-{d:02d}", fmt)
            for i, v in enumerate(cols, start=1):
                ws.write(row, i, v, fmt)

        # 總計列（只畫框）
        tr = 3 + days
        ws.write(tr, 0, "總計", total_fmt)
        for c in range(1, 7):
            ws.write(tr, c, '', total_fmt)

        ws.write(1, 0, f"出勤天數：{attend_days}", sub_fmt)

    writer.close()
    buf.seek(0)

    return send_file(
        buf,
        as_attachment=True,
        download_name=f"{y}-{m:02d}_工時卡片_全員.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
