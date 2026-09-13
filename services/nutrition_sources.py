"""官方營養資料與團膳食材對應。

優先使用衛福部食藥署資料；找不到精確 mapping 時，不再留下空白，
而是依食材名稱做保守類別估值。所有估值都標示 estimated/low confidence，
避免把推估值冒充官方數字。
"""
from __future__ import annotations
import csv, re, unicodedata
from dataclasses import dataclass
from decimal import Decimal
from functools import lru_cache
from pathlib import Path
from typing import Iterable

DATA_DIR=Path(__file__).resolve().parents[1]/"data"/"nutrition"
ENERGY_CSV=DATA_DIR/"tfda_food_energy.csv";INGREDIENT_MAP_CSV=DATA_DIR/"ingredient_tfda_map.csv"
SOURCE_LABEL="衛福部食藥署 臺灣食品營養成分資料庫"
FALLBACK_SOURCE_LABEL="團膳熱量估算值（名稱類別／TFDA 近似樣品）"
STATUS_MATCHED="matched";STATUS_ESTIMATED="estimated";STATUS_PENDING="pending";STATUS_VERIFIED="verified"

@dataclass(frozen=True)
class OfficialFood:
    code:str;name:str;alias:str;category:str;kcal_per_100g:Decimal;waste_percent:str;unit_weight:str

def normalize_name(value:str)->str:
    text=unicodedata.normalize("NFKC",value or "")
    return "".join(text.split()).lower()

def _decimal(value):
    try:return Decimal(str(value or "").strip()) if str(value or "").strip() else None
    except Exception:return None

@lru_cache(maxsize=1)
def official_foods():
    foods={}
    if not ENERGY_CSV.exists():return foods
    with ENERGY_CSV.open(encoding="utf-8",newline="") as f:
        for row in csv.DictReader(f):
            kcal=_decimal(row.get("kcal_per_100g"));code=(row.get("tfda_code") or "").strip()
            if kcal is None or not code:continue
            foods[code]=OfficialFood(code,(row.get("sample_name") or "").strip(),(row.get("alias") or "").strip(),(row.get("category") or "").strip(),kcal,(row.get("waste_percent") or "").strip(),(row.get("unit_weight") or "").strip())
    return foods

@dataclass(frozen=True)
class MapEntry:
    code:str;confidence:str;note:str;fallback_kcal:Decimal|None=None

@lru_cache(maxsize=1)
def ingredient_code_map():
    mapping={}
    if not INGREDIENT_MAP_CSV.exists():return mapping
    with INGREDIENT_MAP_CSV.open(encoding="utf-8",newline="") as f:
        for row in csv.DictReader(f):
            name=normalize_name(row.get("ingredient_name",""))
            if name:mapping[name]=MapEntry((row.get("tfda_code") or "").strip(),(row.get("confidence") or "").strip(),(row.get("note") or "").strip(),_decimal(row.get("fallback_kcal")))
    return mapping

@dataclass
class MatchResult:
    status:str;food:OfficialFood|None=None;fallback_kcal:Decimal|None=None;confidence:str="";note:str=""
    @property
    def kcal_per_100g(self):return self.food.kcal_per_100g if self.food else self.fallback_kcal
    @property
    def source_label(self):
        return SOURCE_LABEL if self.status==STATUS_MATCHED else (FALLBACK_SOURCE_LABEL if self.status==STATUS_ESTIMATED else "")

# 名稱類別保守估值；只在 mapping 完全找不到時使用。
_ESTIMATE_RULES=(
    (("沙拉油","芥花油","大豆油","橄欖油","麻油","香油","豬油","雞油","牛油"),884,"食用油脂通用估值"),
    (("芝麻","花生","腰果","杏仁","核桃","堅果","瓜子"),580,"堅果種子通用估值"),
    (("糖","砂糖","黑糖","冰糖","蜂蜜","糖漿","果醬"),360,"糖類調味品通用估值"),
    (("麵粉","太白粉","地瓜粉","玉米粉","樹薯粉","澱粉","冬粉","粉絲","米粉","乾麵","意麵"),350,"乾燥澱粉／麵製品通用估值"),
    (("白米","糙米","紫米","小米","燕麥","薏仁","藜麥"),355,"乾穀物通用估值"),
    (("香腸","火腿","熱狗","培根","肉鬆","貢丸","魚丸","花枝丸","雞塊","甜不辣"),220,"加工肉／魚漿製品通用估值"),
    (("豬","排骨","肉絲","肉片","肉丁","絞肉","里肌","梅花","五花"),210,"豬肉類通用估值"),
    (("牛","羊"),220,"牛羊肉類通用估值"),
    (("雞","鴨","鵝","火雞"),175,"禽肉類通用估值"),
    (("魚","蝦","蛤","蚵","牡蠣","魷","花枝","小卷","干貝","海參"),130,"魚貝海鮮通用估值"),
    (("蛋",),145,"蛋類通用估值"),
    (("豆腐","豆干","豆乾","豆皮","豆包","毛豆","黃豆","黑豆","豆漿","麵筋","麵輪","麵腸"),120,"豆製品通用估值"),
    (("鮮奶","牛奶","優格","優酪乳","乳品"),65,"乳品通用估值"),
    (("起司","乳酪"),300,"起司乳酪通用估值"),
    (("香蕉","蘋果","芭樂","鳳梨","西瓜","木瓜","芒果","柳丁","橘子","葡萄","奇異果","火龍果","梨","草莓","哈密瓜","蓮霧"),60,"水果類通用估值"),
    (("馬鈴薯","洋芋","地瓜","甘藷","芋頭","山藥","南瓜","玉米","蓮藕"),110,"根莖／澱粉蔬菜通用估值"),
    (("高麗菜","白菜","青江","菠菜","油菜","萵苣","地瓜葉","空心菜","芥藍","花椰","青花菜","洋蔥","蘿蔔","番茄","蕃茄","黃瓜","冬瓜","絲瓜","苦瓜","竹筍","四季豆","豆芽","甜椒","青椒","茄子","芹菜","韭菜","蔥","蒜","薑","香菜","香菇","菇","木耳","海帶","紫菜","青菜","蔬菜"),35,"蔬菜／菇藻類通用估值"),
    (("醬油","醬","味噌","沙茶","咖哩","番茄醬","蠔油","烏醋","白醋","米酒","料理酒"),100,"醬料調味品通用估值"),
    (("大骨","骨頭","高湯","湯底"),25,"湯底／骨類採購品通用估值"),
)

def _heuristic_estimate(name:str):
    n=normalize_name(name)
    for keys,kcal,note in _ESTIMATE_RULES:
        if any(normalize_name(k) in n for k in keys):return Decimal(str(kcal)),note
    return Decimal("100"),"無法分類，採團膳通用保守估值 100 kcal/100g"

def estimated_unit_grams(name:str)->Decimal:
    """個/人食材未設定可食克數時的保守估計。"""
    n=normalize_name(name)
    rules=(("蛋",50),("雞腿",90),("棒棒腿",80),("翅腿",70),("雞翅",45),("熱狗",35),("香腸",45),("魚丸",20),("貢丸",20),("包子",70),("饅頭",80),("水果",100),("蘋果",180),("香蕉",120))
    for key,grams in rules:
        if normalize_name(key) in n:return Decimal(str(grams))
    return Decimal("50")

def match_ingredient_name(name:str)->MatchResult:
    key=normalize_name(name);entry=ingredient_code_map().get(key)
    if entry:
        if entry.code:
            food=official_foods().get(entry.code)
            if food:return MatchResult(STATUS_MATCHED,food=food,confidence=entry.confidence or "medium",note=entry.note)
        if entry.fallback_kcal is not None:return MatchResult(STATUS_ESTIMATED,fallback_kcal=entry.fallback_kcal,confidence=entry.confidence or "low",note=entry.note or "既有團膳 fallback")
    kcal,note=_heuristic_estimate(name)
    return MatchResult(STATUS_ESTIMATED,fallback_kcal=kcal,confidence="low",note=note)

def apply_match(ingredient,result):
    if getattr(ingredient,"nutrition_verified",False):return False
    before=(ingredient.kcal_per_100g,ingredient.nutrition_source,ingredient.nutrition_food_name,ingredient.nutrition_code,ingredient.nutrition_confidence,ingredient.nutrition_status,ingredient.nutrition_note)
    ingredient.kcal_per_100g=result.kcal_per_100g;ingredient.nutrition_source=result.source_label;ingredient.nutrition_food_name=result.food.name if result.food else "團膳估算代表值";ingredient.nutrition_code=result.food.code if result.food else None;ingredient.nutrition_confidence=result.confidence or "low";ingredient.nutrition_status=result.status;ingredient.nutrition_note=result.note or None
    return before!=(ingredient.kcal_per_100g,ingredient.nutrition_source,ingredient.nutrition_food_name,ingredient.nutrition_code,ingredient.nutrition_confidence,ingredient.nutrition_status,ingredient.nutrition_note)

REPORT_FIELDS=("ingredient_id","current_name","base_unit","matched_official_name","official_code","kcal_per_100g","source","confidence","status","note")
def build_report_rows(ingredients:Iterable):
    rows=[]
    for ing in ingredients:
        m=match_ingredient_name(ing.name);k=getattr(ing,"kcal_per_100g",None) or m.kcal_per_100g
        rows.append({"ingredient_id":ing.id,"current_name":ing.name,"base_unit":ing.base_unit,"matched_official_name":getattr(ing,"nutrition_food_name",None) or (m.food.name if m.food else "團膳估算代表值"),"official_code":getattr(ing,"nutrition_code",None) or (m.food.code if m.food else ""),"kcal_per_100g":str(k),"source":getattr(ing,"nutrition_source",None) or m.source_label,"confidence":getattr(ing,"nutrition_confidence",None) or m.confidence,"status":getattr(ing,"nutrition_status",None) or m.status,"note":getattr(ing,"nutrition_note",None) or m.note})
    return rows
def write_report(path:Path,rows:list[dict]):
    path.parent.mkdir(parents=True,exist_ok=True)
    with path.open("w",encoding="utf-8",newline="") as f:
        w=csv.DictWriter(f,fieldnames=list(REPORT_FIELDS));w.writeheader();w.writerows(rows)
