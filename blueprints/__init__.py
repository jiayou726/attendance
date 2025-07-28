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
    依新規則進位：
      ‑ 上班  (is_in=True) ──────────────
          00 分   → 原整點
          01–44  → 整點 + 0.5
          45–59  → 下一整點
      ‑ 下班  (is_in=False) ─────────────
          00–24  → 原整點
          25–54  → 整點 + 0.5
          55–59  → 下一整點
    """
    h, m = time_obj.hour, time_obj.minute

    if is_in:          # 上班
        if m == 0:
            return h
        elif m >= 40:
            return h + 1
        else:          # 1‑44
            return h + 0.5
    else:              # 下班
        if m < 25:     # 0‑24
            return h
        elif m < 55:   # 25‑54
            return h + 0.5
        else:          # 55‑59
            return h + 1

# ────────────────────────── 工時計算 ──────────────────────────
def calc_hours(start: str, end: str, brk: float, *, skip_break: bool = False):
    """
    計算工時:
      - 正班 ≤ 8h
      - 加班≤2h (ot2)
      - 加班>2h (otx)

    只有「跨過 13:00」且 skip_break=False 才會扣預設午休 brk。
    """
    if not start or not end:
        return 0.0, 0.0, 0.0

    t1 = datetime.strptime(start, "%H:%M").time()
    t2 = datetime.strptime(end,   "%H:%M").time()
    ih = roundup(t1, True)
    oh = roundup(t2, False)

    # 跨日
    if oh <= ih:
        oh += 24

    take_break = (ih < LUNCH_POINT) and (oh > LUNCH_POINT) and (not skip_break)

    total = max(oh - ih - (brk if take_break else 0), 0)

    reg = min(total, 8)
    ot2 = min(max(total - 8, 0), 2)
    otx = max(total - 10, 0)
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
