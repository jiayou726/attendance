# -*- coding: utf-8 -*-
# blueprints/punch.py — 自助打卡藍圖（含手機友善 CSS）
# ------------------------------------------------------
# • 固定 QRCode 導向 /punch
# • ★ 短效 gate（記錄首次載入之 IP + 到期）＋ 一次性 token（單次使用）
# • 帶回家後 F5：因 IP 改變或 gate 到期，不再發新 token → 顯示「頁面失效」
# • 員工打卡／月卡查詢

from flask import Blueprint, render_template_string, request, redirect, url_for, current_app, session
from calendar import monthrange
import re, io, base64, qrcode
from datetime import datetime, date, timedelta

from extensions import db
from models import Employee, Checkin
from . import NIGHT_END, merge_night

import time, secrets, hmac, hashlib

QR_VER_KEY = "QR_VER_DELTA"   # 儲存目前使用的版本偏移（整站共用）
QR_VER_SPAN = 6               # 變化範圍：最小可行版本 ~ 最小+5（最多 40）

# ──────────────────────── 手機友善 CSS ────────────────────────
CSS = r"""
<style>
  :root {
    font-size: 20px;
    --fg:#000; --bg:#fff;
    --accent:#005BBB; --warn:#B8860B;
    --error:#B22222; --success:#006400;
    --radius:10px; --shadow:0 2px 6px rgba(0,0,0,.15);
  }
  html,body{margin:0;padding:0;background:var(--bg);color:var(--fg);
            font-family:"Noto Sans TC",sans-serif;line-height:1.6;}
  h2,h3{margin:0 0 .8em;font-weight:700;text-align:center;}
  h3.success{color:var(--success);}h3.warn{color:var(--warn);}
  h3.error{color:var(--error);}
  form{max-width:480px;margin:0 auto 1.2em;display:flex;
       flex-direction:column;gap:1rem;}
  input,select,button{font-size:1rem;padding:.6em .8em;
       border:2px solid var(--fg);border-radius:var(--radius);}
  button{background:var(--accent);color:#fff;cursor:pointer;
         box-shadow:var(--shadow);}button:active{transform:scale(.97);}
  a{color:var(--accent);}
  table{width:100%;border-collapse:collapse;margin:0 auto 1.2em;}
  th,td{border:1px solid var(--fg);padding:.6em .4em;text-align:center;
        font-size:.9rem;}th{background:#f0f0f0;font-weight:700;}
  img{max-width:260px;display:block;margin:0 auto 1em;}
  @media (max-width:420px){:root{font-size:22px;}th,td{font-size:1rem;}}
</style>
"""

HEAD = (
    '<meta charset="utf-8">'
    '<meta name="viewport" content="width=device-width, initial-scale=1.0">'
    f"{CSS}"
)

punch_bp = Blueprint("punch", __name__, url_prefix="/punch")

ADMIN_QR_PWD = "hr1234"      # 更新密碼
QR_CONFIG_KEY = "QR_TEXT"    # 目前 QR 內容暫存於 app.config

# ──────────────────────── 小工具 ────────────────────────
def _client_ip() -> str:
    """
    取得用戶端 IP。若在反向代理（如 Render）後，建議在 app factory 使用 ProxyFix；
    這裡僅取 remote_addr，不自行解析 X-Forwarded-For（安全考量）。
    """
    return request.remote_addr or ""

def _bind_fingerprint() -> str:
    """依設定把 IP/UA 摻入指紋（用於 token/驗證）"""
    parts = []
    if current_app.config.get("PUNCH_BIND_IP", True):
        parts.append(_client_ip())
    if current_app.config.get("PUNCH_BIND_UA", True):
        parts.append((request.headers.get("User-Agent") or "")[:120])
    base = "|".join(parts)
    return hashlib.sha1(base.encode("utf-8")).hexdigest() if base else ""

def _issue_gate_if_needed() -> dict:
    """
    若 session 尚無 gate，建立：
      gate = {"ip": 當下 IP, "exp": now+TTL}
    若已有 gate：
      - 若 IP 改變或已過期 → 回傳 {"invalid": True, "reason": ...}
      - 否則回傳 gate 本身
    """
    now = int(time.time())
    ttl = int(current_app.config.get("PUNCH_GATE_TTL_SEC", 120))
    gate = session.get("punch_gate")

    if not gate:
        gate = {"ip": _client_ip(), "exp": now + ttl}
        session["punch_gate"] = gate
        return gate

    # gate 已存在：檢查是否仍在有效視窗且 IP 未改變
    if gate.get("ip") != _client_ip():
        return {"invalid": True, "reason": "IP 改變"}
    if now > int(gate.get("exp", 0)):
        return {"invalid": True, "reason": "已過期"}
    return gate

def _new_token() -> dict:
    """
    建立一次性 token，寫入 session：
      {"value": "...", "exp": ts, "fp": 綁定指紋}
    """
    now = int(time.time())
    ttl = int(current_app.config.get("PUNCH_TOKEN_TTL_SEC", 120))
    tok = {
        "value": secrets.token_urlsafe(12),
        "exp": now + ttl,
        "fp": _bind_fingerprint(),
    }
    session["punch_token"] = tok
    return tok

def _consume_token(token_from_form: str) -> bool:
    """
    驗證並消費一次性 token：
      - 存在於 session
      - 未過期
      - 值相等
      - 指紋一致（IP/UA）
    驗證後無論成功或失敗都移除 token（單次使用）
    """
    try:
        tok = session.get("punch_token") or {}
        ok = (
            token_from_form
            and tok.get("value") == token_from_form
            and int(time.time()) <= int(tok.get("exp", 0))
            and tok.get("fp") == _bind_fingerprint()
        )
        return bool(ok)
    finally:
        session.pop("punch_token", None)

# ────────────────────────  QR Code 產生  ────────────────────────
@punch_bp.route("/qrcode", methods=["GET", "POST"])
def qrcode_view():
    # 1) 固定 QR 內容（網址不可編輯）
    qr_text = url_for("punch.form", _external=True)

    # 2) 先用自動模式找出「最小可行版本」
    qr_auto = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_H,
        box_size=10,
        border=2,
    )
    qr_auto.add_data(qr_text)
    qr_auto.make(fit=True)
    min_ver = qr_auto.version  # 1~40 的其中一個（依內容與糾錯等級決定）

    # 3) 讀取目前站台共用的版本偏移量（影響未驗證時顯示的預設圖）
    delta = int(current_app.config.get(QR_VER_KEY, 0))

    verified = False
    msg = ""

    if request.method == "POST":
        if request.form.get("verified") == "1":
            # 同頁後續操作（已驗證）
            verified = True
        else:
            # 第一次需驗證密碼
            pwd = request.form.get("pwd", "")
            if pwd == ADMIN_QR_PWD:
                verified = True
                msg = "✅ 已驗證，可重新產生或下載圖片。"
            else:
                msg = "❌ 密碼錯誤。"

        # 已驗證才允許更新：按「重新產生圖片」就把 delta+1 並寫回 config
        if verified and request.form.get("action") == "regen":
            delta += 1
            current_app.config[QR_VER_KEY] = delta
            msg = "✅ 已更新預設圖樣。"

    # 4) 計算本次要用的版本號：
    #    在最小可行版本 ~ 最小+QR_VER_SPAN-1 之間循環，但不超過 v40
    max_ver = min(min_ver + QR_VER_SPAN - 1, 40)
    span = max_ver - min_ver + 1
    if span <= 0:
        span = 1
    use_ver = min_ver + (delta % span)

    # 5) 以指定版本產生圖片（同網址、不同版本 → 圖樣會不同）
    qr = qrcode.QRCode(
        version=use_ver,
        error_correction=qrcode.constants.ERROR_CORRECT_H,
        box_size=10,
        border=2,
    )
    qr.add_data(qr_text)
    qr.make(fit=False)  # 已指定 version，不再自動放大
    img = qr.make_image(fill_color="black", back_color="white")

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    b64 = base64.b64encode(buf.getvalue()).decode()

    # 6) 頁面（未驗證：只看 QR + 密碼；已驗證：可重新產生與下載）
    tpl = """
<!doctype html>
<html lang="zh-Hant">
<head>
  {{ HEAD|safe }}
  <title>打卡 QR Code</title>
  <style>
    .wrap{max-width:720px;margin:40px auto;padding:0 16px;}
    .card{border:1px solid #e5e7eb;border-radius:12px;padding:24px;box-shadow:0 2px 8px rgba(0,0,0,.06);}
    .qr{display:grid;place-items:center;margin:16px 0 8px;}
    .qr img{width:100%;max-width:360px;height:auto;}
    .ver{color:#6b7280;font-size:.85rem;text-align:center;margin-bottom:16px;}
    .row{display:flex;gap:12px;flex-wrap:wrap;align-items:center;justify-content:center;}
    .btn{display:inline-block;padding:10px 16px;border-radius:10px;border:1px solid #d1d5db;text-decoration:none;color:#111827;background:#fff;}
    .btn.primary{background:#111827;color:#fff;border-color:#111827;}
    input[type=password]{width:100%;max-width:360px;padding:10px 12px;border:1px solid #d1d5db;border-radius:10px;}
    .label{font-size:.9rem;color:#374151;margin:.2rem 0 .5rem;text-align:center;}
    .msg{margin-top:.75rem;color:#B22222;text-align:center;}
    .hint{color:#6b7280;font-size:.85rem;margin-top:.5rem;text-align:center;}
  </style>
</head>
<body>
<div class="wrap">
  <h2>打卡 QR Code</h2>
  <div class="card">
    <div class="qr">
      <img src="data:image/png;base64,{{ b64 }}" alt="QR Code">
    </div>

    <div class="row" style="margin-bottom:16px;">
      <a class="btn primary" href="{{ qr_text }}" target="_blank" rel="noopener">前往打卡表單</a>
    </div>

    {% if not verified %}
      <form method="post" autocomplete="off" class="row" style="flex-direction:column;align-items:center;">
        <div class="label">管理密碼：</div>
        <input type="password" name="pwd" placeholder="請輸入密碼">
        <div class="row" style="margin-top:12px;">
          <button class="btn primary" type="submit">驗證</button>
        </div>
      </form>
    {% else %}
      <form method="post" class="row" style="margin-top:4px;">
        <input type="hidden" name="verified" value="1">
        <button class="btn primary" type="submit" name="action" value="regen">重新產生圖片</button>
        <a class="btn" download="punch_qr.png" href="data:image/png;base64,{{ b64 }}">下載 QR 圖片</a>
      </form>
      <div class="hint">提示：重新整理或再次進入本頁時，會重新要求密碼。預設圖樣已更新為目前版本。</div>
    {% endif %}

    {% if msg %}<div class="msg">{{ msg }}</div>{% endif %}
  </div>
</div>
</body>
</html>
"""
    return render_template_string(
        tpl,
        HEAD=HEAD,
        b64=b64,
        qr_text=qr_text,
        verified=verified,
        msg=msg,
        min_ver=min_ver,
        use_ver=use_ver,
    )

# ────────────────────────  打卡表單（GET）  ────────────────────────
@punch_bp.route("/", methods=["GET"])
def form():
    err = request.args.get("err")

    # 1) gate 檢查／建立（只要 gate 存在且 IP 未變、未過期，就視為在場）
    gate = _issue_gate_if_needed()
    if gate.get("invalid"):
        # 不再發 token，提示失效（需要回現場重新掃描）
        reason = gate.get("reason", "")
        return render_template_string(
            f"<!doctype html><html><head>{HEAD}</head><body>"
            f"{'<p class=error>'+('頁面失效：'+reason)+'</p>' if reason else '<p class=error>頁面失效</p>'}"
            "<h2>員工打卡</h2>"
            "<p>請回現場重新掃描 QR Code。</p>"
            "<p><a href='/admin/login'>管理</a></p>"
            "</body></html>"
        )

    # 2) gate 有效：發一次性 token＋顯示倒數
    tok = _new_token()
    ttl = max(0, tok["exp"] - int(time.time()))

    return render_template_string(
        f"<!doctype html><html><head>{HEAD}</head><body>"
        f"{'<p class=error>'+err+'</p>' if err else ''}"
        "<h2>員工打卡</h2>"
        "<form method=post id='pform'>"
        "<input name=eid placeholder='員工編號' required autofocus>"
        "<select name=type>"
        "<option value=am-in>1.上午上班</option>"
        "<option value=am-out>2.上午下班</option>"
        "<option value=pm-in>3.下午上班</option>"
        "<option value=pm-out>4.下午下班</option>"
        "<option value=ot-in>5.加班上班</option>"
        "<option value=ot-out>6.加班下班</option>"
        "</select>"
        f"<input type='hidden' name='token' value='{tok['value']}'>"
        f"<div class='ttl'>本頁有效倒數：<span id='sec'>{ttl}</span> 秒，逾時請重新掃描。</div>"
        "<button id='submitBtn'>打卡</button>"
        "</form>"
        "<p><a href='/admin/login'>管理</a></p>"
        "<script>"
        "  (function(){"
        f"    var sec={ttl};"
        "    var s=document.getElementById('sec');"
        "    var btn=document.getElementById('submitBtn');"
        "    var t=setInterval(function(){"
        "      sec=Math.max(0,sec-1); s.textContent=sec;"
        "      if(sec<=0){clearInterval(t); btn.setAttribute('disabled','disabled'); btn.classList.add('disabled');}"
        "    },1000);"
        "  })();"
        "</script>"
        "</body></html>"
    )

# ────────────────────────  打卡處理（POST）  ────────────────────────
@punch_bp.route("/", methods=["POST"])
def punch():
    eid = request.form["eid"].strip()
    typ = request.form["type"]
    token = (request.form.get("token") or "").strip()

    # 1) gate 必須仍有效（IP 未變 & 未過期）
    gate = session.get("punch_gate")
    now = int(time.time())
    if not gate or gate.get("ip") != _client_ip() or now > int(gate.get("exp", 0)):
        # gate 失效後不再自動刷新，必須回現場重掃
        return redirect(url_for(".form", err="頁面失效，請回現場重新掃描"))

    # 2) 驗證並消費一次性 token（單次使用、未過期、指紋一致）
    if not _consume_token(token):
        return redirect(url_for(".form", err="頁面已過期或無效，請重新掃描"))

    # 3) 原有打卡流程
    now_dt = datetime.now()
    wd  = now_dt.date().isoformat()
    ts  = now_dt.isoformat(timespec="seconds")

    emp = Employee.query.get(eid)
    if not emp:
        return redirect(url_for(".card", eid=eid, st="error", msg="員工不存在"))

    dup = Checkin.query.filter_by(
        employee_id=eid, work_date=wd, p_type=typ
    ).first()
    if dup:
        msg, st = "已打過卡", "warn"
    else:
        db.session.add(Checkin(
            employee_id=eid, work_date=wd, p_type=typ, ts=ts
        ))
        db.session.commit()
        msg, st = "打卡成功", "success"

    return redirect(url_for(".card", eid=eid, st=st, msg=msg))

# ────────────────────────  員工月卡頁  ────────────────────────
@punch_bp.route("/result/<eid>")
def card(eid: str):
    ym_param = request.args.get("ym")
    today = date.today()
    if ym_param and re.fullmatch(r"\d{4}-\d{2}", ym_param):
        y, m = map(int, ym_param.split("-"))
        ym = ym_param
    else:
        y, m = today.year, today.month
        ym = f"{y}-{m:02d}"

    # 月份下拉清單（近 6 個月）
    ym_opts, cursor = [], today.replace(day=1)
    for _ in range(6):
        ym_val = cursor.strftime("%Y-%m")
        ym_lab = cursor.strftime("%Y / %m")
        sel = "selected" if ym_val == ym else ""
        ym_opts.append(f'<option value="{ym_val}" {sel}>{ym_lab}</option>')
        cursor = (cursor - timedelta(days=1)).replace(day=1)

    emp = Employee.query.get_or_404(eid)
    days_in_month = monthrange(y, m)[1]
    next_m = (date(y, m, 1) + timedelta(days=32)).replace(day=1)

    rows = (
        Checkin.query
        .with_entities(Checkin.work_date, Checkin.p_type, Checkin.ts)
        .filter(Checkin.employee_id == eid)
        .filter(
            (Checkin.work_date.like(f"{y}-{m:02d}%")) |
            (
                (Checkin.work_date == next_m.isoformat()) &
                (Checkin.p_type.in_({"am-out", "pm-out", "ot-out"})) &
                (Checkin.ts < f"{next_m}T{NIGHT_END}:00")
            )
        )
        .order_by(Checkin.work_date, Checkin.p_type)
        .all()
    )

    # 合併跨夜記錄、取 HH:MM
    recs = merge_night([
        (r.work_date, r.p_type, r.ts[11:16]) for r in rows
    ])

    body = "".join(
        f"<tr><td>{m:02d}-{d:02d}</td>"
        f"<td>{recs.get((f'{d:02d}','am-in'),  '-') }</td>"
        f"<td>{recs.get((f'{d:02d}','am-out'), '-') }</td>"
        f"<td>{recs.get((f'{d:02d}','pm-in'),  '-') }</td>"
        f"<td>{recs.get((f'{d:02d}','pm-out'), '-') }</td>"
        f"<td>{recs.get((f'{d:02d}','ot-in'),  '-') }</td>"
        f"<td>{recs.get((f'{d:02d}','ot-out'), '-') }</td></tr>"
        for d in range(1, days_in_month + 1)
    )

    st, msg = request.args.get("st"), request.args.get("msg")
    return render_template_string(
        f"<!doctype html><html><head>{HEAD}</head><body>"
        f"{f'<h3 class={st}>{msg}</h3>' if st else ''}"
        f"<h3>{emp.name}（{eid}） 區域：{emp.area}　{y}/{m}</h3>"
        "<form method='get' id='ymForm'>"
        f"<input type='hidden' name='eid' value='{eid}'>"
        "月份：<select name='ym' onchange='ymForm.submit()'>"
        f"{''.join(ym_opts)}</select></form>"
        "<table>"
        "<tr><th>日期</th><th>上午上</th><th>上午下</th>"
        "<th>下午上</th><th>下午下</th>"
        "<th>加班上</th><th>加班下</th></tr>"
        f"{body}"
        "</table>"
        f"<p><a href='{url_for('.form')}'>返回打卡</a></p>"
        "</body></html>"
    )
