from secrets import compare_digest
from urllib.parse import urlparse

from flask import (
    Blueprint,
    abort,
    current_app,
    redirect,
    render_template_string,
    request,
    session,
    url_for,
)

from . import CSS


auth_bp = Blueprint("auth", __name__)


def _passwords() -> dict[str, str]:
    return {
        "hr": current_app.config.get("ADMIN_HR_PASSWORD", ""),
        "mgr": current_app.config.get("ADMIN_MGR_PASSWORD", ""),
    }


def _safe_next(value: str | None) -> str | None:
    if not value:
        return None
    parsed = urlparse(value)
    if parsed.scheme or parsed.netloc or not value.startswith("/"):
        return None
    return value


@auth_bp.route("/login", methods=["GET", "POST"])
def login():
    error = ""
    passwords = _passwords()
    configured_roles = [(role, "人資" if role == "hr" else "主管") for role, pw in passwords.items() if pw]
    configured = bool(configured_roles)
    next_url = _safe_next(request.args.get("next") or request.form.get("next"))

    if request.method == "POST":
        role = request.form.get("role", "")
        pw = request.form.get("pw", "")
        expected = passwords.get(role, "")

        if expected and compare_digest(pw, expected):
            session.clear()
            session["role"] = role
            session.permanent = True
            return redirect(next_url or "/admin/")

        error = "帳號角色或密碼錯誤。"

    if not configured:
        error = "管理密碼尚未設定。請在部署環境設定 ADMIN_HR_PASSWORD 或 ADMIN_MGR_PASSWORD。"

    template = f"""<!doctype html><html lang="zh-Hant"><head><meta charset="utf-8">
    <meta name="viewport" content="width=device-width,initial-scale=1">{CSS}</head><body>
    <h2>管理登入</h2>
    {{% if error %}}<p style="color:#b42318">{{{{ error }}}}</p>{{% endif %}}
    <form method="post">
      {{% if next_url %}}<input type="hidden" name="next" value="{{{{ next_url }}}}">{{% endif %}}
      <select name="role" {{% if not configured %}}disabled{{% endif %}}>
        {{% for value, label in roles %}}<option value="{{{{ value }}}}">{{{{ label }}}}</option>{{% endfor %}}
      </select><br>
      <input type="password" name="pw" autocomplete="current-password" required {{% if not configured %}}disabled{{% endif %}}><br>
      <button {{% if not configured %}}disabled{{% endif %}}>登入</button>
    </form>
    <p><a href="/">回首頁</a></p></body></html>"""
    return render_template_string(
        template,
        error=error,
        next_url=next_url,
        roles=configured_roles,
        configured=configured,
    )


@auth_bp.route("/logout", methods=["GET", "POST"])
def logout():
    session.clear()
    return redirect(url_for("auth.login"))


def require(role=None):
    """相容既有呼叫方式：無登入回登入頁；role='hr' 時只允許 HR。"""
    current_role = session.get("role")
    if not current_role:
        return redirect(url_for("auth.login", next=request.path))
    if role == "hr" and current_role != "hr":
        abort(403)
    return None
