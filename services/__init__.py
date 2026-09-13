"""團膳共用相容鉤子：QR 練習入口、公版菜單 Excel、完整九功能練習模式。"""
from __future__ import annotations
from io import BytesIO
import re
from urllib.parse import quote
from flask import request, render_template
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

@order_bp.app_context_processor
def _practice_context():
    active=(request.args.get("practice")=="1" or request.path.startswith("/admin/order-tool/ai-menu/practice"))
    return {"practice_mode":active}

@order_bp.before_app_request
def _practice_router():
    if request.method=="GET" and request.path=="/admin/order-tool/ai-menu/practice" and request.args.get("feature")!="ai-menu":
        return render_template("kitchen/practice_dashboard.html",practice_mode=True)
    return None

def _rewrite_public_menu_xlsx(response):
    matched=re.search(r"/ai-menu/drafts/(\d+)/public\.xlsx$",request.path)
    if not matched:return response
    from openpyxl import Workbook
    from openpyxl.styles import Alignment,Border,Font,PatternFill,Side
    from sqlalchemy.orm import joinedload
    from ai_models import KitchenMenuDraft
    from extensions import db
    from models import KitchenRecipe,KitchenRecipeIngredient
    from blueprints import ai_menu as ai_menu_module
    draft=db.session.get(KitchenMenuDraft,int(matched.group(1)))
    if draft is None:return response
    result=ai_menu_module._draft_result(draft);rows=ai_menu_module._view_rows(result,draft)
    recipe_ids={cell["recipe_id"] for row in rows for cell in row["cells"] if cell.get("recipe_id")}
    recipes=(KitchenRecipe.query.options(joinedload(KitchenRecipe.ingredients).joinedload(KitchenRecipeIngredient.ingredient)).filter(KitchenRecipe.id.in_(recipe_ids)).all()) if recipe_ids else []
    ingredient_names={recipe.id:"、".join(component.ingredient.name for component in recipe.ingredients if component.ingredient) for recipe in recipes}
    wb=Workbook();ws=wb.active;ws.title="公版菜單";ws.merge_cells("A1:P1");ws["A1"]=f"{draft.name}　{draft.start_date.year}年{draft.start_date.month}月菜單";ws["A1"].font=Font(size=18,bold=True);ws["A1"].alignment=Alignment(horizontal="center",vertical="center");ws.row_dimensions[1].height=30
    headers=["日期","星期","主食","主菜","副菜","","蔬菜","湯品","全穀雜糧類(份)","豆魚蛋肉類(份)","蔬菜類(份)","油脂與堅果種子類(份)","水果類","乳品類","熱量(大卡)","三章1Q"]
    ws.append(headers);thin=Side(style="thin",color="666666");border=Border(left=thin,right=thin,top=thin,bottom=thin);fill=PatternFill("solid",fgColor="E8EEF5")
    for c in ws[2]:c.font=Font(bold=True);c.fill=fill;c.border=border;c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
    def pick(row,cat,slot=1):return next((x for x in row["cells"] if x["category"]==cat and x["slot_index"]==slot),None)
    for row in rows:
        dishes=[pick(row,"主食"),pick(row,"主菜"),pick(row,"副菜",1),pick(row,"副菜",2),pick(row,"青菜"),pick(row,"湯品")]
        rep=result.days[row["service_date"]];sv=rep.servings or {}
        top=[row["service_date"],row["weekday"]]+[(x.get("name","") if x else "") for x in dishes]+[sv.get("grains") or "",sv.get("protein") or "",sv.get("vegetable") or "",sv.get("oil") or "",sv.get("fruit") or "",sv.get("dairy") or "",row["kcal"] if row["kcal"] is not None else "","✓"]
        bottom=["",""]+[(ingredient_names.get(x.get("recipe_id"),"") if x else "") for x in dishes]+["", "", "", "", "", "", "", ""]
        ws.append(top);ws.append(bottom);tr=ws.max_row-1;br=ws.max_row
        for col in range(1,17):
            for rr in (tr,br):ws.cell(rr,col).border=border;ws.cell(rr,col).alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
        for col in range(9,17):ws.merge_cells(start_row=tr,start_column=col,end_row=br,end_column=col)
        for col in range(3,9):ws.cell(tr,col).font=Font(size=14,bold=True);ws.cell(br,col).font=Font(size=9,color="555555")
        ws.row_dimensions[tr].height=34;ws.row_dimensions[br].height=26
    widths={"A":12,"B":7,"C":21,"D":21,"E":19,"F":19,"G":17,"H":21,"I":10,"J":10,"K":9,"L":11,"M":8,"N":8,"O":11,"P":10}
    for col,w in widths.items():ws.column_dimensions[col].width=w
    ws.freeze_panes="C3";ws.sheet_view.showGridLines=False
    out=BytesIO();wb.save(out);body=out.getvalue();response.set_data(body);response.headers["Content-Length"]=str(len(body));response.headers["Content-Type"]="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";return response

@order_bp.after_app_request
def _ai_menu_ui_compatibility(response):
    if request.path=="/punch/qrcode" and response.status_code==200 and "text/html" in (response.content_type or "").lower():
        html=response.get_data(as_text=True);marker="前往叫貨專區</a>"
        if marker in html and "前往練習專區</a>" not in html:html=html.replace(marker,marker+_PRACTICE_LINK,1);response.set_data(html)
    if request.path.endswith("/public.xlsx") and response.status_code==200:
        response=_rewrite_public_menu_xlsx(response);filename=_safe_download_name(request.args.get("filename",""))
        if filename:response.headers["Content-Disposition"]="attachment; filename*=UTF-8''"+quote(filename)
    return response
