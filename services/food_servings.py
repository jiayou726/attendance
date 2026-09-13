"""六大類食物代換份數估算（學校午餐 AI 菜單用）。

官方依據：
- 教育部「學校午餐食物內容及營養基準」109.12.28 修訂：
  公告菜單應以六大類食物份量呈現，份量定義依國民健康署食物代換表。
- 國民健康署「每日飲食指南／食物代換表」：
  全穀雜糧 1 份約 70 kcal、豆魚蛋肉 1 份約 75 kcal、蔬菜 1 份約
  100g 可食生重、水果 1 份約 100g 可食部、油脂與堅果種子 1 份約
  45 kcal（5g 脂肪）、乳品 1 份約 240mL 牛奶等值。

本模組不是把菜名硬套官方數字，而是用既有 Recipe BOM 的「每人用量」估算。
對不同豆製品、蛋、乳品等使用食物代換表常用等值重量；全穀雜糧以其
實際熱量換算 70 kcal/份。油脂份數以當日總熱量扣除其他五類的標準熱量後
反推，這能把烹調油與食材本身脂肪一起納入，但仍屬「估算值」，不是營養師
簽證值。
"""
from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal, ROUND_HALF_UP
import re

from services.nutrition import ingredient_grams
from services.nutrition_sources import match_ingredient_name

GROUP_ORDER = ("grains", "protein", "vegetable", "oil", "fruit", "dairy")
GROUP_LABELS = {
    "grains": "全穀雜糧類",
    "protein": "豆魚蛋肉類",
    "vegetable": "蔬菜類",
    "oil": "油脂與堅果種子類",
    "fruit": "水果類",
    "dairy": "乳品類",
}
KCAL_PER_SERVING = {
    "grains": Decimal("70"),
    "protein": Decimal("75"),
    "vegetable": Decimal("25"),
    "oil": Decimal("45"),
    "fruit": Decimal("60"),
    "dairy": Decimal("150"),
}
_TARGETS = {
    (620, 720): {"grains": (3.5, 4.5), "protein": (2.0, 2.0), "vegetable": (1.5, 1.5), "oil": (2.0, 2.0), "fruit": (1.0, 1.0)},
    (720, 830): {"grains": (4.5, 5.5), "protein": (2.0, 2.0), "vegetable": (2.0, 2.0), "oil": (2.5, 2.5), "fruit": (1.0, 1.0)},
    (800, 930): {"grains": (5.0, 6.5), "protein": (2.5, 2.5), "vegetable": (2.0, 2.0), "oil": (2.5, 2.5), "fruit": (1.0, 1.0)},
    (900, 1050): {"grains": (6.0, 7.5), "protein": (3.0, 3.0), "vegetable": (2.0, 2.0), "oil": (3.0, 3.0), "fruit": (1.0, 1.0)},
    (680, 810): {"grains": (4.0, 5.5), "protein": (2.0, 2.0), "vegetable": (2.0, 2.0), "oil": (2.5, 2.5), "fruit": (1.0, 1.0)},
    (680, 1050): {"grains": (4.0, 7.5), "protein": (2.0, 3.0), "vegetable": (2.0, 2.0), "oil": (2.5, 3.0), "fruit": (1.0, 1.0)},
}

def targets_for_kcal(kcal_min: int | None, kcal_max: int | None) -> dict[str, tuple[float, float]]:
    if kcal_min is None or kcal_max is None:
        return {}
    return dict(_TARGETS.get((int(kcal_min), int(kcal_max)), {}))

def _norm(value: str) -> str:
    return re.sub(r"[\s　()（）\-_/.,，、]", "", value or "").lower()

_DAIRY = ("鮮奶", "牛奶", "乳酪", "起司", "優格", "優酪乳", "奶粉", "保久乳")
_FRUIT = ("香蕉", "蘋果", "芭樂", "番石榴", "鳳梨", "西瓜", "木瓜", "芒果", "柳丁", "橘子", "柑橘", "葡萄", "奇異果", "火龍果", "百香果", "柚", "水梨", "梨子", "草莓", "哈密瓜", "蓮霧")
_PROTEIN = ("豆腐", "豆干", "豆乾", "干絲", "豆皮", "豆包", "毛豆", "黃豆", "黑豆", "豆漿", "麵筋", "麵腸", "麵輪", "雞蛋", "鴨蛋", "蛋", "雞", "豬", "牛", "羊", "鴨", "鵝", "魚", "蝦", "蛤", "蚵", "牡蠣", "魷", "花枝", "小卷", "干貝", "海參", "肉", "排骨", "肉羹", "貢丸", "魚丸", "花枝丸", "雞塊", "火腿", "香腸", "熱狗", "培根", "豬排", "雞排", "魚排")
_GRAINS = ("白米", "糙米", "紫米", "米飯", "米", "燕麥", "大麥", "小麥", "薏仁", "小米", "藜麥", "紅藜", "麵", "冬粉", "粉絲", "米粉", "粿", "年糕", "玉米", "馬鈴薯", "洋芋", "地瓜", "甘藷", "芋頭", "山藥", "南瓜", "蓮藕", "紅豆", "綠豆", "花豆", "蠶豆", "皇帝豆", "栗子", "蓮子", "菱角", "西谷米", "西米", "粉圓", "湯圓", "米血", "甜不辣", "饅頭", "包子", "吐司", "麵包")
_OIL = ("沙拉油", "芥花油", "大豆油", "橄欖油", "麻油", "香油", "豬油", "雞油", "牛油", "芝麻", "花生", "腰果", "杏仁", "核桃", "堅果", "瓜子")
_VEGETABLE = ("青菜", "蔬菜", "高麗菜", "白菜", "大白菜", "小白菜", "青江", "菠菜", "油菜", "萵苣", "地瓜葉", "空心菜", "芥藍", "花椰", "青花菜", "紅鳳菜", "莧菜", "a菜", "皇宮菜", "洋蔥", "紅蘿蔔", "胡蘿蔔", "白蘿蔔", "蘿蔔", "番茄", "蕃茄", "小黃瓜", "大黃瓜", "黃瓜", "冬瓜", "蒲瓜", "扁蒲", "瓠瓜", "絲瓜", "苦瓜", "竹筍", "筍", "四季豆", "長豆", "豆芽", "銀芽", "彩椒", "甜椒", "青椒", "茄子", "芹菜", "韭菜", "蔥", "蒜", "薑", "九層塔", "香菜", "香菇", "鮮菇", "秀珍菇", "杏鮑菇", "金針菇", "洋菇", "木耳", "海帶", "海結", "海茸", "紫菜", "酸菜", "榨菜", "雪菜", "筍乾")

def food_group(name: str) -> str | None:
    n = _norm(name)
    if not n:return None
    if "椰奶" not in n and any(k in n for k in _DAIRY):return "dairy"
    if any(k in n for k in _FRUIT):return "fruit"
    if any(k in n for k in _PROTEIN):return "protein"
    if any(k in n for k in _GRAINS):return "grains"
    if any(k in n for k in _OIL):return "oil"
    if any(k in n for k in _VEGETABLE):return "vegetable"
    return None

def _kcal_per_100g(ingredient) -> Decimal | None:
    raw=getattr(ingredient,"kcal_per_100g",None)
    if raw is not None:return Decimal(str(raw))
    return match_ingredient_name(ingredient.name).kcal_per_100g

def _protein_grams_per_serving(name: str) -> Decimal:
    n=_norm(name)
    if "嫩豆腐" in n:return Decimal("140")
    if "豆腐" in n:return Decimal("80")
    if "豆干" in n or "豆乾" in n or "干絲" in n:return Decimal("40")
    if "毛豆" in n:return Decimal("50")
    if "黃豆" in n:return Decimal("20")
    if "黑豆" in n:return Decimal("25")
    if "豆漿" in n:return Decimal("240")
    if "雞蛋" in n or "鴨蛋" in n or n=="蛋":return Decimal("50")
    if "蝦仁" in n:return Decimal("50")
    if "牡蠣" in n or "蚵" in n:return Decimal("65")
    if "文蛤" in n or "蛤蜊" in n:return Decimal("160")
    if "海參" in n:return Decimal("100")
    return Decimal("35")

def _dairy_grams_per_serving(name: str) -> Decimal:
    n=_norm(name)
    if "起司" in n or "乳酪" in n:return Decimal("45")
    if "優格" in n:return Decimal("210")
    if "奶粉" in n:return Decimal("30")
    return Decimal("240")

@dataclass
class RecipeServings:
    values: dict[str, Decimal] = field(default_factory=lambda: {key: Decimal("0") for key in GROUP_ORDER})
    unclassified: list[str] = field(default_factory=list)
    missing_amount: list[str] = field(default_factory=list)
    def rounded(self) -> dict[str,float]:
        q=Decimal("0.1");return {key:float(value.quantize(q,rounding=ROUND_HALF_UP)) for key,value in self.values.items()}

def recipe_servings(recipe) -> RecipeServings:
    result=RecipeServings()
    for component in list(recipe.ingredients or ()):
        ingredient=component.ingredient
        if ingredient is None:continue
        group=food_group(ingredient.name)
        if not group:result.unclassified.append(ingredient.name);continue
        grams=ingredient_grams(ingredient,component.grams_per_person)
        if grams is None:
            if group=="protein" and (ingredient.base_unit or "g")=="個" and "蛋" in (ingredient.name or ""):
                try:result.values["protein"]+=Decimal(str(component.grams_per_person or 0))
                except Exception:result.missing_amount.append(ingredient.name)
            else:result.missing_amount.append(ingredient.name)
            continue
        if grams<=0:continue
        if group=="grains":
            kcal=_kcal_per_100g(ingredient)
            result.values[group]+=(grams/Decimal("100")*kcal)/KCAL_PER_SERVING[group] if kcal is not None and kcal>0 else grams/Decimal("20")
        elif group=="protein":result.values[group]+=grams/_protein_grams_per_serving(ingredient.name)
        elif group in {"vegetable","fruit"}:result.values[group]+=grams/Decimal("100")
        elif group=="dairy":result.values[group]+=grams/_dairy_grams_per_serving(ingredient.name)
        elif group=="oil":
            kcal=_kcal_per_100g(ingredient);result.values[group]+=(grams/Decimal("100")*kcal)/KCAL_PER_SERVING[group] if kcal is not None and kcal>0 else grams/Decimal("5")
    return result

def finalize_day_servings(base_values,total_kcal) -> dict[str,float]:
    values={key:Decimal(str(base_values.get(key,0) or 0)) for key in GROUP_ORDER}
    if total_kcal is not None:
        kcal=Decimal(str(total_kcal));non_oil=sum(values[key]*KCAL_PER_SERVING[key] for key in GROUP_ORDER if key!="oil");residual=max(Decimal("0"),(kcal-non_oil)/KCAL_PER_SERVING["oil"]);values["oil"]=max(values["oil"],residual)
    q=Decimal("0.1");return {key:float(values[key].quantize(q,rounding=ROUND_HALF_UP)) for key in GROUP_ORDER}

def merge_recipe_servings(recipe_values):
    total={key:Decimal("0") for key in GROUP_ORDER}
    for values in recipe_values:
        for key in GROUP_ORDER:total[key]+=Decimal(str(values.get(key,0) or 0))
    return total

def serving_deviation(values,targets):
    deviation=0.0;warnings=[];ok=[]
    for key,(low,high) in targets.items():
        value=float(values.get(key,0) or 0);label=GROUP_LABELS[key];advisory_only=key in {"fruit","dairy"}
        if value+0.05<low:
            if not advisory_only:deviation+=low-value
            warnings.append(f"⚠ {label} {value:g} 份，低於建議 {low:g}～{high:g} 份")
        elif value-0.05>high:
            if not advisory_only:deviation+=value-high
            warnings.append(f"⚠ {label} {value:g} 份，高於建議 {low:g}～{high:g} 份")
        else:ok.append(f"✅ {label} {value:g} 份")
    return deviation,warnings,ok
