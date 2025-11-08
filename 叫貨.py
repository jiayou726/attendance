"""單檔 Flask 版本叫貨工具。

使用方式：
1. pip install flask pandas openpyxl
2. (選擇) 設定環境變數 FLASK_SECRET_KEY
3. python 叫貨.py
4. 瀏覽 http://127.0.0.1:5000
"""

from __future__ import annotations

import io
import os
import uuid
from typing import List, Sequence

import pandas as pd
from flask import Flask, abort, render_template_string, request, send_file

# Excel 欄位群組起點，對應原本的 1,7,13,19,25...
GROUP_STARTS: Sequence[int] = (1, 7, 13, 19, 25)

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
        <a class="download" href="/download/{{ download_id }}">下載結果 Excel</a>

        <div class="table-wrapper">
          <table>
            <thead>
              <tr>
                <th>工作表</th>
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

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "dev-secret-change-me")

# 暫存轉出的 Excel，讓使用者可以下載。
RESULT_CACHE: dict[str, bytes] = {}


def parse_keywords(raw: str) -> List[str]:
    return [kw.strip() for kw in raw.split(",") if kw.strip()]


def filter_workbook(xls: pd.ExcelFile, keywords: Sequence[str]) -> list[dict]:
    lowered = [kw.lower() for kw in keywords]
    rows: list[dict] = []

    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet)
        if df.empty:
            continue

        for _, row in df.iterrows():
            for start in GROUP_STARTS:
                if start + 4 >= len(row):
                    continue

                vendor = row.iloc[start]
                item = row.iloc[start + 1]
                g1 = row.iloc[start + 2]
                qty = row.iloc[start + 3]
                unit = row.iloc[start + 4]

                if pd.isna(vendor) and pd.isna(item):
                    continue

                text = f"{vendor} {item}".lower()
                if any(kw in text for kw in lowered):
                    rows.append(
                        {
                            "工作表": sheet,
                            "廠商": vendor,
                            "品名": item,
                            "1箱g": g1,
                            "數量": qty,
                            "單位": unit,
                        }
                    )
    return rows


@app.route("/", methods=["GET", "POST"])
def index():
    context = {"result_rows": None, "keywords": [], "download_id": None, "error": None}

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

            context.update({"result_rows": rows, "keywords": keywords, "download_id": token})
            return render_template_string(HTML_TEMPLATE, **context)

        except Exception as exc:  # noqa: BLE001
            context["error"] = f"處理檔案時發生錯誤：{exc}"
            return render_template_string(HTML_TEMPLATE, **context)

    return render_template_string(HTML_TEMPLATE, **context)


@app.get("/download/<token>")
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


if __name__ == "__main__":
    app.run(debug=True)
