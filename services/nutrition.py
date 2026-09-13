"""菜色熱量計算。

每人熱量 = Σ（每人用量換算成公克 ÷ 100 × 食材 kcal/100g）。
優先人工值，其次 TFDA mapping；再找不到就用明確標示的估算值。
個/人食材若沒設定 edible_grams_per_unit，也會用保守單顆重量估值，
避免整道菜因少數資料缺口無法計算。
"""
from __future__ import annotations
from dataclasses import dataclass,field
from decimal import Decimal,ROUND_HALF_UP
from services.nutrition_sources import estimated_unit_grams,match_ingredient_name
HUNDRED=Decimal("100");KCAL_QUANTUM=Decimal("1")

@dataclass
class RecipeNutrition:
    recipe_id:int;kcal_per_person:Decimal|None;missing_ingredients:list[str]=field(default_factory=list);has_ingredients:bool=True
    @property
    def complete(self):return self.kcal_per_person is not None
    @property
    def kcal_int(self):return None if self.kcal_per_person is None else int(self.kcal_per_person.quantize(KCAL_QUANTUM,rounding=ROUND_HALF_UP))
    @property
    def reason(self):
        if self.complete:return ""
        if not self.has_ingredients:return "尚未建立配方原料"
        if self.missing_ingredients:return "缺少："+"、".join(self.missing_ingredients)
        return "營養資料不完整"
    @property
    def label(self):return f"預估熱量：{self.kcal_int} kcal / 人" if self.complete else "營養資料不完整"

def ingredient_grams(ingredient,amount_per_person):
    if amount_per_person is None:return None
    amount=Decimal(str(amount_per_person));base=(ingredient.base_unit or "g").strip() or "g"
    if base=="g":return amount
    raw=getattr(ingredient,"edible_grams_per_unit",None)
    per_unit=Decimal(str(raw)) if raw is not None else estimated_unit_grams(ingredient.name)
    if per_unit<=0:return None
    return amount*per_unit

def _ingredient_kcal(ingredient):
    raw=getattr(ingredient,"kcal_per_100g",None)
    return Decimal(str(raw)) if raw is not None else match_ingredient_name(ingredient.name).kcal_per_100g

def recipe_nutrition(recipe):
    rows=list(recipe.ingredients or ())
    if not rows:return RecipeNutrition(recipe.id,None,has_ingredients=False)
    total=Decimal("0");missing=[]
    for row in rows:
        ing=row.ingredient
        if ing is None:missing.append("未知食材");continue
        grams=ingredient_grams(ing,row.grams_per_person);kcal=_ingredient_kcal(ing)
        if grams is None or kcal is None:missing.append(ing.name);continue
        total+=grams/HUNDRED*kcal
    if missing:return RecipeNutrition(recipe.id,None,missing)
    return RecipeNutrition(recipe.id,total)

def bulk_recipe_nutrition(recipes):return {r.id:recipe_nutrition(r) for r in recipes}
def day_kcal(nutritions):
    total=Decimal("0");incomplete=[]
    for name,n in nutritions:
        if n is None or not n.complete:incomplete.append(name)
        else:total+=n.kcal_per_person
    return (None,incomplete) if incomplete else (int(total.quantize(KCAL_QUANTUM,rounding=ROUND_HALF_UP)),[])
