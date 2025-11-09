# -*- coding: utf-8 -*-
# blueprints/punch.py ???芸?嚗摰?QR + 敺垢?芸??Ｙ?銝甈⊥?token嚗?
# ------------------------------------------------------
# ???箏? QRCode ?? /punch
# ??瘥活??/punch ?賣??具? IP ????gate 閬??扼? token
# ??蝡 redirect ??/punch/use?tk=... 憿舐內銵典
# ??token 蝬? IP/UA + ?唳?嚗?鈭文雿誥嚗5 ?芣??瑟?園?嚗??? token
# ???亙葆?振 F5嚗P ?寡???gate ??嚗?/punch/use ?＊蝷箝??Ｗ仃??
# ???喳??踵 token 敹???游 /punch嚗? IP嚗ate ?芸?蝥?

from flask import Blueprint, render_template_string, request, redirect, url_for, current_app, session
from calendar import monthrange
import re, io, base64, qrcode
from datetime import datetime, date, timedelta

from extensions import db
from models import Employee, Checkin
from . import NIGHT_END, merge_night

import time, secrets, hashlib

QR_VER_KEY  = "QR_VER_DELTA"  # QR ?身?見??宏
QR_VER_SPAN = 6

# ???????????????????????? ???? CSS ????????????????????????
CSS = r"""
<style>
  :root { font-size: 20px; --fg:#000; --bg:#fff; --accent:#005BBB; --warn:#B8860B;
          --error:#B22222; --success:#006400; --radius:10px; --shadow:0 2px 6px rgba(0,0,0,.15); }
  html,body{margin:0;padding:0;background:var(--bg);color:var(--fg);
            font-family:"Noto Sans TC",sans-serif;line-height:1.6;}
  h2,h3{margin:0 0 .8em;font-weight:700;text-align:center;}
  h3.success{color:var(--success);}h3.warn{color:var(--warn);}h3.error{color:var(--error);}
  form{max-width:480px;margin:0 auto 1.2em;display:flex;flex-direction:column;gap:1rem;}
  input,select,button{font-size:1rem;padding:.6em .8em;border:2px solid var(--fg);border-radius:var(--radius);}
  button{background:var(--accent);color:#fff;cursor:pointer;box-shadow:var(--shadow);}
  button:active{transform:scale(.97);}
  a{color:var(--accent);}
  table{width:100%;border-collapse:collapse;margin:0 auto 1.2em;}
  th,td{border:1px solid var(--fg);padding:.6em .4em;text-align:center;font-size:.9rem;}
  th{background:#f0f0f0;font-weight:700;}
  img{max-width:260px;display:block;margin:0 auto 1em;}
  .ttl{color:#444;font-size:.9rem;margin:.4rem 0 1rem;}
  .disabled{opacity:.5;pointer-events:none;}
  @media (max-width:420px){:root{font-size:22px;}th,td{font-size:1rem;}}
</style>
"""

HEAD = (
    '<meta charset="utf-8">'
    '<meta name="viewport" content="width=device-width, initial-scale=1.0">'
    f"{CSS}"
)

punch_bp = Blueprint("punch", __name__, url_prefix="/punch")

ADMIN_QR_PWD = "hr1234"      # ?湔撖Ⅳ嚗? QR ?嚗?
QR_CONFIG_KEY = "QR_TEXT"

# ???????????????????????? 撌亙 ????????????????????????
def _client_ip() -> str:
    # ?亙??隞??敺遣霅啣 app factory ??ProxyFix嚗ㄐ?? remote_addr嚗??剁?
    return request.remote_addr or ""

def _bind_fingerprint() -> str:
    parts = []
    if current_app.config.get("PUNCH_BIND_IP", True):
        parts.append(_client_ip())
    if current_app.config.get("PUNCH_BIND_UA", True):
        parts.append((request.headers.get("User-Agent") or "")[:120])
    base = "|".join(parts)
    return hashlib.sha1(base.encode("utf-8")).hexdigest() if base else ""

def _issue_or_refresh_gate_same_ip() -> dict:
    """Return gate info tied to current client IP, issuing or refreshing as needed."""

    now = int(time.time())
    ttl = int(current_app.config.get("PUNCH_GATE_TTL_SEC", 120))
    cur_ip = _client_ip()
    gate = session.get("punch_gate")

    if not gate:
        gate = {"ip": cur_ip, "exp": now + ttl}
        session["punch_gate"] = gate
        return gate

    if gate.get("ip") != cur_ip:
        return {"invalid": True, "reason": "IP ?寡?"}

    if now > int(gate.get("exp", 0)):
        gate = {"ip": cur_ip, "exp": now + ttl}
        session["punch_gate"] = gate
        return gate

    return gate

def _new_token() -> dict:
    """Issue a new single-use token bound to the current client fingerprint."""

    now = int(time.time())
    ttl = int(current_app.config.get("PUNCH_TOKEN_TTL_SEC", 120))
    tok = {
        "value": secrets.token_urlsafe(12),
        "exp": now + ttl,
        "fp": _bind_fingerprint(),
    }
    session["punch_token"] = tok
    return tok

def _check_token_alive() -> tuple[bool, int]:
    """Return (is_valid, seconds_left) for the session token."""

    tok = session.get("punch_token") or {}
    left = int(tok.get("exp", 0)) - int(time.time())
    return (bool(tok) and left > 0 and tok.get("fp") == _bind_fingerprint(), max(0, left))

def _consume_token(token_from_form: str) -> bool:
    """Validate and consume the pending session token regardless of outcome."""

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

# ????????????????????????  QR Code ?Ｙ??? ????????????????????????
@punch_bp.route("/qrcode", methods=["GET", "POST"])
def qrcode_view():
    qr_text = url_for('punch.form', _external=True)
    order_tool_url = url_for('order_tool.index')

    qr_auto = qrcode.QRCode(version=None, error_correction=qrcode.constants.ERROR_CORRECT_H,
                            box_size=10, border=2)
    qr_auto.add_data(qr_text)
    qr_auto.make(fit=True)
    min_ver = qr_auto.version

    delta = int(current_app.config.get(QR_VER_KEY, 0))
    verified = False
    msg = ""

    if request.method == "POST":
        if request.form.get("verified") == "1":
            verified = True
        else:
            if request.form.get("pwd", "") == ADMIN_QR_PWD:
                verified = True
                msg = "管理密碼驗證成功，現在可以重新產生或下載 QR 圖片。"
            else:
                msg = "管理密碼錯誤，請再試一次。"

        if verified and request.form.get("action") == "regen":
            delta += 1
            current_app.config[QR_VER_KEY] = delta
            msg = "已更新 QR 版本並重新產生圖片，請重新下載。"

    max_ver = min(min_ver + QR_VER_SPAN - 1, 40)
    span = max_ver - min_ver + 1
    if span <= 0: span = 1
    use_ver = min_ver + (delta % span)

    qr = qrcode.QRCode(version=use_ver, error_correction=qrcode.constants.ERROR_CORRECT_H,
                       box_size=10, border=2)
    qr.add_data(qr_text)
    qr.make(fit=False)
    img = qr.make_image(fill_color="black", back_color="white")

    buf = io.BytesIO(); img.save(buf, format="PNG")
    b64 = base64.b64encode(buf.getvalue()).decode()

    tpl = "\n".join([
        '<!doctype html>',
        '<html lang="zh-Hant">',
        '  <head>',
        '    {{ HEAD|safe }}',
        '    <title>打卡 QR Code</title>',
        '    <style>',
        '      .wrap{max-width:720px;margin:40px auto;padding:0 16px}',
        '      .card{border:1px solid #e5e7eb;border-radius:12px;padding:24px;box-shadow:0 2px 8px rgba(0,0,0,.06)}',
        '      .qr{display:grid;place-items:center;margin:16px 0 8px}',
        '      .qr img{width:100%;max-width:360px;height:auto}',
        '      .row{display:flex;gap:12px;flex-wrap:wrap;align-items:center;justify-content:center}',
        '      .btn{display:inline-block;padding:10px 16px;border-radius:10px;border:1px solid #d1d5db;text-decoration:none;color:#111827;background:#fff}',
        '      .btn.primary{background:#111827;color:#fff;border-color:#111827}',
        '      .label{font-size:.9rem;color:#374151;margin:.2rem 0 .5rem;text-align:center}',
        '      .msg{margin-top:.75rem;color:#B22222;text-align:center}',
        '    </style>',
        '  </head>',
        '  <body>',
        '    <div class="wrap">',
        '      <h2>打卡 QR Code</h2>',
        '      <div class="card">',
        '        <div class="qr"><img src="data:image/png;base64,{{ b64 }}" alt="QR Code"></div>',
        '        <div class="row" style="margin-bottom:16px;">',
        '          <a class="btn primary" href="{{ qr_text }}" target="_blank" rel="noopener">前往線上打卡表單</a>',
        '          <a class="btn" href="{{ order_tool_url }}" target="_blank" rel="noopener">前往叫貨專區</a>',
        '        </div>',
        '        {% if not verified %}',
        '        <form method="post" autocomplete="off" class="row" style="flex-direction:column;align-items:center;">',
        '          <div class="label">管理密碼：</div>',
        '          <input type="password" name="pwd" placeholder="請輸入密碼">',
        '          <div class="row" style="margin-top:12px;">',
        '            <button class="btn primary" type="submit">驗證</button>',
        '          </div>',
        '        </form>',
        '        {% else %}',
        '        <form method="post" class="row" style="margin-top:4px;">',
        '          <input type="hidden" name="verified" value="1">',
        '          <button class="btn primary" type="submit" name="action" value="regen">產生新 QR 圖片</button>',
        '          <a class="btn" download="punch_qr.png" href="data:image/png;base64,{{ b64 }}">下載 QR 圖片</a>',
        '        </form>',
        '        {% endif %}',
        '        {% if msg %}<div class="msg">{{ msg }}</div>{% endif %}',
        '      </div>',
        '    </div>',
        '  </body>',
        '</html>'
    ])


    return render_template_string(tpl, HEAD=HEAD, b64=b64, qr_text=qr_text, order_tool_url=order_tool_url)

# ????????????????????????  ?亙嚗? QR ???圈ㄐ嚗?甈∠? token 銝血???/use嚗?????????????????????????
@punch_bp.route("/", methods=["GET"])
def form():
    gate = _issue_or_refresh_gate_same_ip()
    if gate.get("invalid"):
        reason = gate.get("reason", "Gate expired")
        return render_template_string(
            f"<!doctype html><html><head>{HEAD}</head><body>"
            f"<p class='error'>無法使用打卡頁面：{reason}</p>"
            "<h2>線上打卡</h2><p>請回到原裝置重新掃描 QR Code。</p>"
            "<p><a href='/admin/login'>管理登入</a></p></body></html>"
        )

    tok = _new_token()  # 瘥活???賜? token
    return redirect(url_for(".use", tk=tok["value"]))

# ????????????????????????  憿舐內銵典嚗?亙????Ｙ???token嚗?????????????????????????
@punch_bp.route("/use", methods=["GET"])
def use():
    tk = request.args.get("tk", "")
    tok = session.get("punch_token") or {}
    if not tk or tk != tok.get("value"):
        return render_template_string(
            f"<!doctype html><html><head>{HEAD}</head><body>"
            "<p class='error'>無效的憑證，請重新掃描 QR Code。</p>"
            "<h2>線上打卡</h2><p>請回到原裝置重新掃描 QR Code。</p>"
            "<p><a href='/admin/login'>管理登入</a></p></body></html>"
        )

    # gate 敹?隞???& IP ?芾?嚗oken 敹?隞???
    gate = session.get("punch_gate")
    now = int(time.time())
    ok_gate = gate and gate.get("ip") == _client_ip() and now <= int(gate.get("exp", 0))
    ok_tok, left = _check_token_alive()

    if not ok_gate or not ok_tok:
        return render_template_string(
            f"<!doctype html><html><head>{HEAD}</head><body>"
            "<p class='error'>此連線已失效，請重新掃描 QR Code。</p>"
            "<h2>線上打卡</h2><p>請回到原裝置重新掃描 QR Code。</p>"
            "<p><a href='/admin/login'>管理登入</a></p></body></html>"
        )

    # 憿舐內銵典嚗idden 撣?token嚗?? left 蝘??寧 f-string嚗?? % ?澆???
    return render_template_string(
        f"<!doctype html><html><head>{HEAD}</head><body>"
        "<h2>線上打卡</h2>"
        "<form method='post' action='/punch/'>"
        "<input name='eid' placeholder='員工編號' required autofocus>"
        "<select name='type'>"
        "<option value='am-in'>1. 上午上班</option>"
        "<option value='am-out'>2. 上午下班</option>"
        "<option value='pm-in'>3. 下午上班</option>"
        "<option value='pm-out'>4. 下午下班</option>"
        "<option value='ot-in'>5. 加班上班</option>"
        "<option value='ot-out'>6. 加班下班</option>"
        "</select>"
        f"<input type='hidden' name='token' value='{tok['value']}'>"
        f"<div class='ttl'>此頁面將在 <span id='sec'>{left}</span> 秒後失效，請儘速提交。</div>"
        "<button id='submitBtn'>送出打卡</button>"
        "</form>"
        "<p><a href='/admin/login'>管理登入</a></p>"
        "<script>(function(){"
        f"var sec={left};"
        "var s=document.getElementById('sec');"
        "var btn=document.getElementById('submitBtn');"
        "var t=setInterval(function(){"
        "  sec=Math.max(0,sec-1); s.textContent=sec;"
        "  if(sec<=0){clearInterval(t); btn.setAttribute('disabled','disabled'); btn.classList.add('disabled');}"
        "},1000);"
        "})();</script>"
        "</body></html>"
    )

# ????????????????????????  ?漱嚗OST嚗?????????????????????????
@punch_bp.route("/", methods=["POST"])
def punch():
    eid = request.form["eid"].strip()
    typ = request.form["type"]
    token = (request.form.get("token") or "").strip()

    # gate 敹?隞???IP ?芾?銝??嚗?
    gate = session.get("punch_gate")
    now = int(time.time())
    if not gate or gate.get("ip") != _client_ip() or now > int(gate.get("exp", 0)):
        return redirect(url_for(".form", err="連線已失效，請重新掃描 QR Code。"))

    # token ?格活撽?
    if not _consume_token(token):
        return redirect(url_for(".form", err="Token 已失效，請重新掃描 QR Code。"))

    # ???瘚?
    now_dt = datetime.now()
    wd = now_dt.date().isoformat()
    ts = now_dt.isoformat(timespec="seconds")

    emp = Employee.query.get(eid)
    if not emp:
        return redirect(url_for(".card", eid=eid, st="error", msg="查無此員工。"))

    dup = Checkin.query.filter_by(employee_id=eid, work_date=wd, p_type=typ).first()
    if dup:
        msg, st = "本時段已打卡，請勿重複。", "warn"
    else:
        db.session.add(Checkin(employee_id=eid, work_date=wd, p_type=typ, ts=ts))
        db.session.commit()
        msg, st = "打卡完成。", "success"

    return redirect(url_for(".card", eid=eid, st=st, msg=msg))

# ????????????????????????  ?∪極???蝬剜??見嚗?????????????????????????
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

    recs = merge_night([(r.work_date, r.p_type, r.ts[11:16]) for r in rows])

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
    status_html = f"<h3 class='{st}'>{msg}</h3>" if st else ""
    return render_template_string(
        f"<!doctype html><html><head>{HEAD}</head><body>"
        f"{status_html}"
        f"<h3>{emp.name}（員工編號：{eid}，單位：{emp.area}） {y}/{m}</h3>"
        "<form method='get' id='ymForm'>"
        f"<input type='hidden' name='eid' value='{eid}'>"
        "<label for='ymSelect'>選擇月份：</label>"
        "<select id='ymSelect' name='ym' onchange='ymForm.submit()'>"
        f"{''.join(ym_opts)}</select></form>"
        "<table>"
        "<tr><th>日期</th><th>上午上班</th><th>上午下班</th>"
        "<th>下午上班</th><th>下午下班</th>"
        "<th>加班上班</th><th>加班下班</th></tr>"
        f"{body}"
        "</table>"
        f"<p><a href='{url_for('.form')}'>返回線上打卡</a></p>"
        "</body></html>"
    )

