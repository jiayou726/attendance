"""團膳排菜相關商業邏輯與 QR 練習入口相容鉤子。"""
from __future__ import annotations
import re
from urllib.parse import quote
from flask import request
from blueprints.order_tool import order_bp

_PRACTICE_LINK=(
    '<a class="btn" href="/admin/order-tool/ai-menu/practice" '
    'target="_blank" rel="noopener">前往練習專區</a>'
)

def _safe_download_name(value:str)->str:
    text=re.sub(r'[\\/:*?"<>|\x00-\x1f]+',"_",(value or "").strip())
    text=re.sub(r"\s+"," ",text).strip(" ._")
    if not text:return ""
    text=text[:80].rstrip(" ._")
    if not text.lower().endswith(".xlsx"):text+=".xlsx"
    return text

@order_bp.after_app_request
def _ai_menu_ui_compatibility(response):
    if request.path=="/punch/qrcode" and response.status_code==200 and "text/html" in (response.content_type or "").lower():
        html=response.get_data(as_text=True); marker="前往叫貨專區</a>"
        if marker in html and "前往練習專區</a>" not in html:
            html=html.replace(marker,marker+_PRACTICE_LINK,1);response.set_data(html)
    if request.path.endswith("/public.xlsx") and response.status_code==200:
        filename=_safe_download_name(request.args.get("filename",""))
        if filename:response.headers["Content-Disposition"]="attachment; filename*=UTF-8''"+quote(filename)
    return response
