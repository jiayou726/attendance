"""Administrator controls for named kitchen practice accounts."""
from __future__ import annotations

import secrets

from flask import Blueprint, abort, flash, redirect, render_template, request, session, url_for

from services.practice_database import delete_account, list_accounts, reset_account


practice_admin_bp = Blueprint("practice_admin", __name__)


def _csrf_token() -> str:
    token = session.get("_practice_admin_csrf")
    if not token:
        token = secrets.token_urlsafe(32)
        session["_practice_admin_csrf"] = token
    return token


def _check_csrf() -> None:
    expected = session.get("_practice_admin_csrf", "")
    received = request.form.get("_csrf_token", "")
    if not expected or not received or not secrets.compare_digest(expected, received):
        abort(400)


@practice_admin_bp.get("/")
def index():
    return render_template(
        "practice_admin/index.html",
        accounts=list_accounts(),
        csrf_token=_csrf_token,
    )


@practice_admin_bp.post("/<account_id>/reset")
def reset(account_id: str):
    _check_csrf()
    flash("練習帳號已重置。" if reset_account(account_id) else "找不到練習帳號。", "success")
    return redirect(url_for("practice_admin.index"))


@practice_admin_bp.post("/<account_id>/delete")
def delete(account_id: str):
    _check_csrf()
    flash("練習帳號已刪除並釋出名額。" if delete_account(account_id) else "找不到練習帳號。", "success")
    return redirect(url_for("practice_admin.index"))
