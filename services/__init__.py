"""團膳排菜相關的商業邏輯（營養計算、供餐對象基準、自動排菜引擎）。

這裡也放兩個不碰正式資料的 UI 相容鉤子：
1. 打卡 QR Code 頁在「叫貨專區」右側顯示 AI 菜單練習入口。
2. 公版 Excel 允許用 query string 的 ``filename`` 自訂下載檔名。

之所以用 after-app-response 鉤子，是因為舊打卡頁仍由 punch.py 內嵌 HTML 產生；
這樣不用重寫舊打卡流程，也不會影響 QR / 打卡驗證邏輯。
"""

from __future__ import annotations

import re
from urllib.parse import quote

from flask import request

from blueprints.order_tool import order_bp


_PRACTICE_LINK = (
    '<a class="btn" href="/admin/order-tool/ai-menu/practice/login" '
    'target="_blank" rel="noopener">前往練習專區</a>'
)


def _safe_download_name(value: str) -> str:
    """Return a browser-safe xlsx file name without path characters."""
    text = re.sub(r'[\\/:*?"<>|\x00-\x1f]+', "_", (value or "").strip())
    text = re.sub(r"\s+", " ", text).strip(" ._")
    if not text:
        return ""
    text = text[:80].rstrip(" ._")
    if not text.lower().endswith(".xlsx"):
        text += ".xlsx"
    return text


@order_bp.after_app_request
def _ai_menu_ui_compatibility(response):
    """Add the requested QR shortcut and optional public-export filename."""
    if request.path == "/punch/qrcode" and response.status_code == 200:
        content_type = (response.content_type or "").lower()
        if "text/html" in content_type:
            html = response.get_data(as_text=True)
            marker = "前往叫貨專區</a>"
            if marker in html and "前往練習專區</a>" not in html:
                html = html.replace(marker, marker + _PRACTICE_LINK, 1)
                response.set_data(html)

    if request.path.endswith("/public.xlsx") and response.status_code == 200:
        filename = _safe_download_name(request.args.get("filename", ""))
        if filename:
            response.headers["Content-Disposition"] = (
                "attachment; filename*=UTF-8''" + quote(filename)
            )

    return response
