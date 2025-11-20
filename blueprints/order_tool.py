"""叫貨 Excel 篩選工具 Blueprint。"""

from __future__ import annotations

import io
import re
import uuid
from datetime import datetime
from typing import List, Sequence, Tuple, Optional

import pandas as pd
from flask import Blueprint, abort, render_template_string, request, send_file

# Excel 欄位群組起點，對應原本的 1,7,13,19,25...
GROUP_STARTS: Sequence[int] = (1, 7, 13, 19, 25)
HEADER_SCAN_ROWS = 6
HEADER_SKIP_VALUES = {"廠商"}
DATE_FORMATS = ("%Y-%m-%d", "%Y/%m/%d", "%m-%d", "%m/%d")
WEEKDAY_LABELS = ("一", "二", "三", "四", "五", "六", "日")
CURRENT_YEAR = datetime.now().year

# 暫存轉出的 Excel，讓使用者可以下載。
RESULT_CACHE: dict[str, bytes] = {}

order_bp = Blueprint("order_tool", __name__)

HTML_TEMPLATE = """
<!doctype html>
<html lang="zh-Hant">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>叫貨 Excel 篩選器</title>
    <style>
      body { font-family: "Segoe UI","Microsoft JhengHei",sans-serif; background:#f5f6fa; margin:0; color:#1f2328; }
      .container { max-width:760px; margin:0 auto; padding:32px 20px 48px; }
      form, .card { background:#fff; border-radius:12px; padding:24px; box-shadow:0 2px 8px rgba(31,35,40,.08); margin-bottom:24px; }
      label { font-weight:600; display:block; margin-bottom:6px; }
      input[type=file], input[type=text] { width:100%; padding:10px 12px; border:1px solid #d0d7de; border-radius:8px; margin-bottom:18px; box-sizing:border-box; }
      button { background:#1f883d; color:#fff; border:none; border-radius:999px; padding:12px 20px; cursor:pointer; }
      button:hover { background:#1a7f37; }
      .alert { padding:14px 16px; border-radius:10px; margin-bottom:18px; }
      .alert.error { background:#ffeef0; color:#9a1c1c; border:1px solid #ffc2c7; }
      .table-wrapper { overflow-x:auto; }
      table { width:100%; border-collapse:collapse; }
      th, td { padding:10px; border-bottom:1px solid #d8dee4; text-align:left; white-space:nowrap; }
      th { background:#f8faff; }
      a.download { display:inline-block; margin:12px 0 18px; color:#0969da; font-weight:600; text-decoration:none; }
    </style>
  </head>
  <body>
    <main class="container">
      <h1>叫貨 Excel 篩選器（網頁版）</h1>

      <form method="post" enctype="multipart/form-data">
        <label for="excel_file">1. 上傳 Excel 檔案</label>
        <input id="excel_file" name="excel_file" type="file" accept=".xlsx,.xls" required />

        <label for="keywords">2. 輸入關鍵字（逗號分隔）</label>
        <input id="keywords" name="keywords" type="text" placeholder="例如：巨城, 士多啤梨" required />

        <button type="submit">3. 篩選並輸出</button>
      </form>

      {% if error %}
      <div class="alert error">{{ error }}</div>
      {% endif %}

      {% if result_rows %}
      <section class="card">
        <h2>共找到 {{ result_rows|length }} 筆資料</h2>
        <p>關鍵字：{{ keywords|join("、") }}</p>
        <a class="download" href="{{ url_for('order_tool.download', token=download_id) }}">下載結果 Excel</a>

        <div class="table-wrapper">
          <table>
            <thead>
              <tr>
                <th>工作表</th>
                <th>日期</th>
                <th>菜名</th>
                <th>廠商</th>
                <th>品名</th>
                <th>1箱g數</th>
                <th>數量</th>
                <th>單位</th>
              </tr>
            </thead>
            <tbody>
              {% for row in result_rows %}
              <tr>
                <td>{{ row["工作表"] }}</td>
                <td>{{ row["日期"] }}</td>
                <td>{{ row["菜名"] }}</td>
                <td>{{ row["廠商"] }}</td>
                <td>{{ row["品名"] }}</td>
                <td>{{ row["1箱g"] }}</td>
                <td>{{ row["數量"] }}</td>
                <td>{{ row["單位"] }}</td>
              </tr>
              {% endfor %}
            </tbody>
          </table>
        </div>
      </section>
      {% endif %}
    </main>
  </body>
</html>
"""


def _detect_group_date(df: pd.DataFrame, start: int) -> str:
    """Return the label/date that sits on top of a 5-column group."""

    if start >= df.shape[1]:
        return ""

    limit = min(HEADER_SCAN_ROWS, len(df))
    candidates: list[str] = []

    for cell in df.iloc[:limit, start]:
        if pd.isna(cell):
            continue

        if isinstance(cell, datetime):
            text = cell.strftime("%Y-%m-%d")
        else:
            text = str(cell).strip()

        if not text:
            continue

        if any(marker in text for marker in ("月", "日", "/", "-")):
            return text

        if text in HEADER_SKIP_VALUES or text.startswith("用餐"):
            continue

        candidates.append(text)

    return candidates[0] if candidates else ""


def _parse_date_components(value: str) -> Optional[Tuple[int, int, int]]:
    if not value:
        return None

    text = str(value).strip()
    if not text:
        return None

    for fmt in DATE_FORMATS:
        try:
            dt = datetime.strptime(text, fmt)
        except ValueError:
            continue

        year = dt.year if "%Y" in fmt else CURRENT_YEAR
        return year, dt.month, dt.day

    match = re.search(r"(\d{4})\D+(\d{1,2})\D+(\d{1,2})", text)
    if match:
        year, month, day = map(int, match.groups())
        return year, month, day

    match = re.search(r"(\d{1,2})\s*月\s*(\d{1,2})", text)
    if match:
        month, day = map(int, match.groups())
        return CURRENT_YEAR, month, day

    numbers = [int(n) for n in re.findall(r"\d+", text)]
    if len(numbers) >= 3:
        year, month, day = numbers[:3]
        return year, month, day
    if len(numbers) == 2:
        month, day = numbers
        return CURRENT_YEAR, month, day
    if len(numbers) == 1:
        day = numbers[0]
        return CURRENT_YEAR, 1, day

    return None


def _format_date_display(value: str) -> Tuple[str, Optional[Tuple[int, int, int]]]:
    components = _parse_date_components(value)
    text = (value or "").strip()

    if not components:
        return text, None

    year, month, day = components
    try:
        dt = datetime(year, month, day)
    except ValueError:
        return text or f"{year:04d}-{month:02d}-{day:02d}", components

    has_week = bool(re.search(r"(週|星期|禮拜|\([一二三四五六日]\))", text))
    base = text or dt.strftime("%Y-%m-%d")
    if not has_week:
        base = f"{base} ({WEEKDAY_LABELS[dt.weekday()]})"

    return base, components


def _row_sort_key(row: dict) -> Tuple[int, int, int, str, str, str]:
    token = row.get("__date_token")
    if token:
        year, month, day = token
    else:
        year, month, day = (9999, 99, 99)

    return (
        year,
        month,
        day,
        row.get("工作表", ""),
        str(row.get("廠商", "")),
        str(row.get("品名", "")),
    )


def parse_keywords(raw: str) -> List[str]:
    return [kw.strip() for kw in raw.split(",") if kw.strip()]


def filter_workbook(xls: pd.ExcelFile, keywords: Sequence[str]) -> list[dict]:
    lowered = [kw.lower() for kw in keywords]
    rows: list[dict] = []

    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet)
        if df.empty:
            continue

        group_dates = {start: _detect_group_date(df, start) for start in GROUP_STARTS}

        dish_history: dict[int, str] = {}

        for _, row in df.iterrows():
            for start in GROUP_STARTS:
                if start + 4 >= len(row):
                    continue

                dish_name = ""
                if start - 1 >= 0:
                    cell = row.iloc[start - 1]
                    if cell is not None and not (isinstance(cell, float) and pd.isna(cell)):
                        candidate = str(cell).strip()
                        if candidate:
                            dish_history[start] = candidate
                if dish_history.get(start):
                    dish_name = dish_history[start]

                vendor = row.iloc[start]
                item = row.iloc[start + 1]
                g1 = row.iloc[start + 2]
                qty = row.iloc[start + 3]
                unit = row.iloc[start + 4]

                if pd.isna(vendor) and pd.isna(item):
                    continue

                text = f"{vendor} {item}".lower()
                if any(kw in text for kw in lowered):
                    raw_date = group_dates.get(start, "")
                    display_date, token = _format_date_display(raw_date)
                    rows.append(
                        {
                            "工作表": sheet,
                            "日期": display_date,
                            "菜名": dish_name,
                            "廠商": vendor,
                            "品名": item,
                            "1箱g": g1,
                            "數量": qty,
                            "單位": unit,
                            "__date_token": token,
                        }
                    )

    rows.sort(key=_row_sort_key)
    for row in rows:
        row.pop("__date_token", None)
    return rows


@order_bp.route("/", methods=["GET", "POST"])
def index():
    context = {
        "result_rows": None,
        "keywords": [],
        "download_id": None,
        "error": None,
    }

    if request.method == "POST":
        upload = request.files.get("excel_file")
        raw_keywords = request.form.get("keywords", "")

        if not upload or upload.filename == "":
            context["error"] = "請選擇要處理的 Excel 檔案。"
            return render_template_string(HTML_TEMPLATE, **context)

        keywords = parse_keywords(raw_keywords)
        if not keywords:
            context["error"] = "請至少輸入一個關鍵字，並以逗號分隔。"
            return render_template_string(HTML_TEMPLATE, **context)

        try:
            data = upload.read()
            xls = pd.ExcelFile(io.BytesIO(data))
            rows = filter_workbook(xls, keywords)

            if not rows:
                context["error"] = f"沒有找到含有 {', '.join(keywords)} 的資料。"
                return render_template_string(HTML_TEMPLATE, **context)

            result_df = pd.DataFrame(rows)
            output = io.BytesIO()
            result_df.to_excel(output, index=False)
            output.seek(0)

            token = uuid.uuid4().hex
            RESULT_CACHE[token] = output.getvalue()

            context.update(
                {"result_rows": rows, "keywords": keywords, "download_id": token}
            )
            return render_template_string(HTML_TEMPLATE, **context)

        except Exception as exc:  # noqa: BLE001
            context["error"] = f"處理檔案時發生錯誤：{exc}"
            return render_template_string(HTML_TEMPLATE, **context)

    return render_template_string(HTML_TEMPLATE, **context)


@order_bp.get("/download/<token>")
def download(token: str):
    data = RESULT_CACHE.pop(token, None)
    if data is None:
        abort(404)

    buf = io.BytesIO(data)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name="叫貨結果.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
