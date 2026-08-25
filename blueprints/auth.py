from hashlib import sha256
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

# 舊版曾直接把管理密碼寫在程式碼裡。現在保留舊密碼的 SHA-256 雜湊作為 fallback，
# 讓既有登入方式可繼續使用，同時避免把明文密碼重新放回 public repo。
LEGACY_PASSWORD_HASHES = {
    "hr": "bcb70742aad2b11dddb9cc1708e1b918199f8484f3d4f30aeece88d140cfd04a",
    "mgr": "4d926562dafec6a110dd71004c0a3f949c31533e0a36cb4de0078e6705949a80",
}


def _passwords() -> dict[str, str]:
    return {
        "hr": current_app.config.get("ADMIN_HR_PASSWORD", ""),
        "mgr": current_app.config.get("ADMIN_MGR_PASSWORD", ""),
    }


def _role_configured(role: str, passwords: dict[str, str]) -> bool:
    return bool(passwords.get(role) or LEGACY_PASSWORD_HASHES.get(role))


def _password_matches(role: str, password: str, passwords: dict[str, str]) -> bool:
    configured_password = passwords.get(role, "")
    if configured_password:
        return compare_digest(password, configured_password)

    legacy_hash = LEGACY_PASSWORD_HASHES.get(role, "")
    if not legacy_hash:
        return False
    candidate_hash = sha256(password.encode("utf-8")).hexdigest()
    return compare_digest(candidate_hash, legacy_hash)


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
    configured_roles = [
        (role, "人資" if role == "hr" else "主管")
        for role in ("hr", "mgr")
        if _role_configured(role, passwords)
    ]
    configured = bool(configured_roles)
    next_url = _safe_next(request.args.get("next") or request.form.get("next"))

    if request.method == "POST":
        role = request.form.get("role", "")
        pw = request.form.get("pw", "")

        if _password_matches(role, pw, passwords):
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
