"""Restore the selected practice week; operational records stay in the account workspace."""
from __future__ import annotations

from datetime import date, timedelta

from flask import session

from extensions import db
from models import KitchenMenuPlan, KitchenPurchaseOrder


def _week_start(raw) -> date | None:
    try:
        day = raw if isinstance(raw, date) else date.fromisoformat(str(raw))
    except (TypeError, ValueError):
        return None
    return day - timedelta(days=day.weekday())


def preferred_practice_week(raw=None) -> date:
    """Explicit week > browser selection > most recently edited private record > today.

    This is used only on order-tool routes which bind the ORM to practice space.
    """
    if not session.get("kitchen_practice"):
        return _week_start(raw) or _week_start(date.today())
    if raw:
        selected = _week_start(raw)
        if selected is not None:
            session["practice_selected_week"] = selected.isoformat()
            return selected
    saved = _week_start(session.get("practice_selected_week"))
    if saved is not None:
        return saved
    plan = (
        db.session.query(KitchenMenuPlan.service_date, KitchenMenuPlan.updated_at)
        .order_by(KitchenMenuPlan.updated_at.desc(), KitchenMenuPlan.id.desc())
        .first()
    )
    order = (
        db.session.query(KitchenPurchaseOrder.service_date, KitchenPurchaseOrder.updated_at)
        .order_by(KitchenPurchaseOrder.updated_at.desc(), KitchenPurchaseOrder.id.desc())
        .first()
    )
    records = [record for record in (plan, order) if record is not None]
    recent = max(records, key=lambda item: item.updated_at) if records else None
    selected = _week_start(recent.service_date if recent else date.today())
    session["practice_selected_week"] = selected.isoformat()
    return selected
