"""不依賴語言模型的團膳自動排菜規則引擎。"""
from __future__ import annotations
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import date, timedelta
from decimal import Decimal, ROUND_HALF_UP
import random

from sqlalchemy.orm import joinedload

from models import KitchenRecipe, KitchenRecipeIngredient
from services.nutrition_sources import match_ingredient_name

CATEGORY_ORDER = ("主食", "主菜", "副菜", "青菜", "湯品", "點心")
DEFAULT_STRUCTURE = {"主食": 1, "主菜": 1, "副菜": 2, "青菜": 1, "湯品": 1, "點心": 0}
STRUCTURE_LIMITS = {"主食": 2, "主菜": 2, "副菜": 4, "青菜": 2, "湯品": 2, "點心": 2}
DIET_ALLOWED = {
    "regular": {"meat", "shared"},
    "vegetarian": {"vegetarian", "shared"},
    "all": {"meat", "vegetarian", "shared"},
}

@dataclass
class Rules:
    repeat_days: int = 5
    fish_per_week_min: int = 1
    fried_per_week_max: int = 1
    sweet_soup_per_week_max: int = 1
    exclude_incomplete_nutrition: bool = True

@dataclass
class Candidate:
    id: int
    name: str
    category: str
    kcal: int | None
    diet_type: str
    tags: set[str] = field(default_factory=set)

@dataclass
class DayResult:
    service_date: date
    items: list[Candidate]
    kcal: int | None
    warnings: list[str]


def service_dates(start: date, end: date, include_weekends=False):
    rows=[]; current=start
    while current <= end:
        if include_weekends or current.weekday() < 5:
            rows.append(current)
        current += timedelta(days=1)
    return rows


def build_structure(values):
    result={}
    for category in CATEGORY_ORDER:
        try: count=int(values.get(category, DEFAULT_STRUCTURE[category]))
        except (TypeError, ValueError): count=DEFAULT_STRUCTURE[category]
        result[category]=max(0,min(STRUCTURE_LIMITS[category],count))
    return result


def recipe_kcal(recipe):
    """依配方即時計算每人熱量；主檔未填時讀 TFDA／fallback mapping。"""
    rows=list(recipe.ingredients or ())
    if not rows:
        return None
    total=Decimal("0")
    for row in rows:
        ing=row.ingredient
        if ing is None or row.grams_per_person is None:
            return None
        amount=Decimal(str(row.grams_per_person))
        if (ing.base_unit or "g") == "g":
            grams=amount
        else:
            per_unit=getattr(ing,"edible_grams_per_unit",None)
            if per_unit is None or Decimal(str(per_unit)) <= 0:
                return None
            grams=amount*Decimal(str(per_unit))
        kcal=getattr(ing,"kcal_per_100g",None)
        if kcal is None:
            kcal=match_ingredient_name(ing.name).kcal_per_100g
        if kcal is None:
            return None
        total += grams/Decimal("100")*Decimal(str(kcal))
    return int(total.quantize(Decimal("1"),rounding=ROUND_HALF_UP))


def load_candidates():
    rows=(KitchenRecipe.query
        .options(joinedload(KitchenRecipe.ingredients).joinedload(KitchenRecipeIngredient.ingredient), joinedload(KitchenRecipe.tags))
        .filter(KitchenRecipe.active.is_(True)).order_by(KitchenRecipe.name).all())
    candidates=[]
    for recipe in rows:
        category=(recipe.category or "").strip()
        diet_type=(getattr(recipe,"diet_type",None) or "").strip()
        if category not in CATEGORY_ORDER or diet_type not in DIET_ALLOWED["all"]:
            continue
        tags={tag.tag for tag in (recipe.tags or []) if getattr(tag,"source","manual") == "manual"}
        candidates.append(Candidate(recipe.id, recipe.name, category, recipe_kcal(recipe), diet_type, tags))
    return candidates


def generate(start, end, structure, rules, kcal_min=None, kcal_max=None, include_weekends=False,
             seed=20260913, diet_mode="regular"):
    """依已確認 diet_type 排菜。

    regular：只可用 meat + shared。
    vegetarian：只可用 vegetarian + shared。
    all：僅供中央候選池／管理用途，不應直接當單一葷或素材單。
    """
    if diet_mode not in DIET_ALLOWED:
        raise ValueError("diet_mode must be regular, vegetarian or all")

    rng=random.Random(seed)
    candidates=load_candidates()
    allowed=DIET_ALLOWED[diet_mode]
    pools=defaultdict(list)
    for c in candidates:
        if c.diet_type not in allowed:
            continue
        if rules.exclude_incomplete_nutrition and c.kcal is None:
            continue
        pools[c.category].append(c)
    for bucket in pools.values(): rng.shuffle(bucket)
    dates=service_dates(start,end,include_weekends)
    last_used={}; usage=Counter(); week_counts=defaultdict(Counter); results=[]
    for day_index,d in enumerate(dates):
        chosen=[]; warnings=[]; week=d.isocalendar()[:2]
        for category in CATEGORY_ORDER:
            for _slot in range(structure.get(category,0)):
                pool=pools.get(category,[])
                if not pool:
                    warnings.append(f"{category} 沒有符合葷素分類與營養條件的可用菜色"); continue
                same={x.id for x in chosen}
                scored=[]
                for c in pool:
                    if c.id in same: continue
                    gap=day_index-last_used.get(c.id,-999)
                    penalty=usage[c.id]*5
                    if gap < rules.repeat_days: penalty += (rules.repeat_days-gap)*100
                    if "fried" in c.tags and week_counts[week]["fried"] >= rules.fried_per_week_max: penalty += 50
                    if "sweet_soup" in c.tags and week_counts[week]["sweet_soup"] >= rules.sweet_soup_per_week_max: penalty += 50
                    scored.append((penalty,rng.random(),c))
                if not scored: continue
                c=min(scored,key=lambda x:(x[0],x[1]))[2]
                chosen.append(c); usage[c.id]+=1; last_used[c.id]=day_index
                for tag in c.tags: week_counts[week][tag]+=1
        kcal=None if any(c.kcal is None for c in chosen) else sum(c.kcal or 0 for c in chosen)
        if kcal is None: warnings.append("營養資料不完整，無法計算當日總熱量")
        elif kcal_min is not None and kcal < kcal_min: warnings.append(f"熱量 {kcal} kcal 低於基準 {kcal_min}～{kcal_max} kcal")
        elif kcal_max is not None and kcal > kcal_max: warnings.append(f"熱量 {kcal} kcal 高於基準 {kcal_min}～{kcal_max} kcal")
        results.append(DayResult(d,chosen,kcal,warnings))
    # 素食菜單不要求魚類；葷食／all 才檢查每週魚類下限。
    if diet_mode != "vegetarian":
        for week,cnt in week_counts.items():
            if cnt["fish"] < rules.fish_per_week_min:
                for row in results:
                    if row.service_date.isocalendar()[:2] == week:
                        row.warnings.append(f"本週魚類 {cnt['fish']} 次，未達 {rules.fish_per_week_min} 次")
                        break
    return results
