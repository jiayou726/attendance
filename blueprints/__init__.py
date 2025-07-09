"""
共用設定與工具函式
"""

# 夜班 (凌晨前) 合併判斷點：04:00 以前下班併入前一日
NIGHT_END = "04:00"

# 午休檢查點：班別若橫跨 13:00 才扣預設午休
LUNCH_POINT = 13  # 13:00



# 共用 CSS 樣式
CSS = """<style>
body{font-family:Arial,"Noto Sans CJK TC",sans-serif;text-align:center;padding-top:30px;font-size:32px}
table{margin:auto;border-collapse:collapse;font-size:28px}
th,td{border:1px solid #444;padding:10px 16px}
input,select,button,a{font-size:28px;margin:6px;padding:10px 20px;text-decoration:none}
button{cursor:pointer}
.success{color:green}
.warn{color:orange}
.error{color:red}
</style>"""

from datetime import datetime, date, timedelta
import re  # 若其他檔案需要用到 re.fullmatch

# ────────────────────────── 時間進位 ──────────────────────────
def roundup(time_obj, is_in):
    """
    時間進位：
      - 上班 (is_in=True)：>45 分 → +1H；25–45 分 → +0.5H；其餘 → 捨去
      - 下班 (is_in=False)：>55 分 → +1H；25–55 分 → +0.5H；其餘 → 捨去
    """
    h, m = time_obj.hour, time_obj.minute
    if is_in:
        if m > 45:
            return h + 1
        if m >= 25:
            return h + 0.5
        return h
    else:
        if m > 55:
            return h + 1
        if m >= 25:
            return h + 0.5
        return h

# ────────────────────────── 工時計算 ──────────────────────────
def calc_hours(start, end, brk):
    """
    計算工時 (正班 ≤8h, 加班≤2h, 加班>2h)
    只有當「進位後上班時數 < 13」且「進位後下班時數 > 13」時才扣 brk。
    參數：
      start, end : "HH:MM" 字串
      brk        : 預設午休小時 (0, 0.5, 1)
    回傳：
      (reg, ot2, otx)  # 正班、加班≤2h、加班>2h
    """
    if not start or not end:
        return 0, 0, 0

    t1 = datetime.strptime(start, "%H:%M").time()
    t2 = datetime.strptime(end,   "%H:%M").time()
    ih = roundup(t1, True)
    oh = roundup(t2, False)

    # 跨日處理
    if oh <= ih:
        oh += 24

    # 是否跨過午休檢查點 (13:00)
    take_break = (ih < LUNCH_POINT) and (oh > LUNCH_POINT)

    total = oh - ih - (brk if take_break else 0)
    total = max(total, 0)

    reg  = min(total, 8)
    ot2  = min(max(total - 8, 0), 2)
    otx  = max(total - 10, 0)
    return round(reg, 1), round(ot2, 1), round(otx, 1)

# ────────────────────────── 夜班下班合併 ──────────────────────────
def merge_night(rows):
    """
    將凌晨 (小於 NIGHT_END) 的下班記錄歸併到前一天。
    rows: list[tuple] → (work_date, p_type, "HH:MM")
    回傳 dict：{ (day_str, p_type): "HH:MM" }
      - day_str 為兩位日字串 01..31
    """
    data = {}
    for wd, typ, tm in rows:
        dt = date.fromisoformat(wd)
        if typ == "out" and tm < NIGHT_END:  # 凌晨下班
            prev = dt - timedelta(days=1)
            day_key = f"{prev.day:02d}"
            data[(day_key, "out")] = tm
        else:
            day_key = f"{dt.day:02d}"
            data[(day_key, typ)] = tm
    return data
