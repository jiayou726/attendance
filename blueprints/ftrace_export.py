"""Additive routes on the existing, CSRF-protected kitchen blueprint."""
from datetime import date, datetime
import json
from zoneinfo import ZoneInfo

from flask import jsonify, render_template, request, send_file

from blueprints.order_tool import order_bp
from services.ftrace import KINDS, SCHEMAS, ExportError, build_sheets, draft, export_workbook
from services.ftrace_references import reference_data


@order_bp.get("/ftrace")
def ftrace():
    raw_date = request.args.get("date")
    try:
        day = date.fromisoformat(raw_date) if raw_date else datetime.now(ZoneInfo("Asia/Taipei")).date()
    except ValueError:
        return "日期格式錯誤，請使用 YYYY-MM-DD。", 400
    kind = request.args.get("kind", "receiving")
    if kind not in SCHEMAS:
        return "資料類型錯誤。", 400
    response = render_template("kitchen/ftrace.html", data=draft(day), kind=kind,
                               kinds=KINDS, references=reference_data())
    return response, 200, {"Cache-Control": "no-store"}


@order_bp.post("/ftrace/export")
def ftrace_export():
    try:
        if len(request.form.get("payload", "")) > 2_000_000:
            raise ExportError("匯出資料過大。")
        payload = json.loads(request.form.get("payload", ""))
        if not isinstance(payload, dict):
            raise ExportError("匯出資料格式錯誤。")
        if request.form.get("confirmed") != "1":
            raise ExportError("請先核對實際數量並勾選確認。")
        kind = request.form.get("kind")
        day = date.fromisoformat(request.form.get("date", ""))
        sheets = build_sheets(kind, day, payload)
        stream = export_workbook(kind, sheets)
    except (ValueError, TypeError, KeyError) as error:
        message = str(error) if isinstance(error, ExportError) else "日期或資料格式錯誤，請重新檢查。"
        return jsonify(error=message), 400
    label = dict(KINDS)[kind]
    response = send_file(stream, as_attachment=True,
        download_name=f"非追不可_{label}_{day.isoformat()}.xls",
        mimetype="application/vnd.ms-excel")
    response.headers["Cache-Control"] = "no-store"
    return response
