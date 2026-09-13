"""菜色熱量計算。

熱量一律即時由 Recipe BOM 算出，不另外存一份快取欄位，
所以食材熱量或每人用量一改，畫面上的數字就跟著變。

計算方式（依需求規格）：
    每人熱量 = Σ（每人用量換算成公克 ÷ 100 × 食材 kcal/100g）

缺少任何一項食材的官方熱量時，整道菜視為「營養資料不完整」，
不會用 0 代替，也不會只算得到的部分當成正確值。
"""

from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal, ROUND_HALF_UP

HUNDRED = Decimal("100")
KCAL_QUANTUM = Decimal("1")


@dataclass
class RecipeNutrition:
    """一道菜的每人熱量結果。"""

    recipe_id: int
    kcal_per_person: Decimal | None
    missing_ingredients: list[str] = field(default_factory=list)
    has_ingredients: bool = True

    @property
    def complete(self) -> bool:
        return self.kcal_per_person is not None

    @property
    def kcal_int(self) -> int | None:
        if self.kcal_per_person is None:
            return None
        return int(self.kcal_per_person.quantize(KCAL_QUANTUM, rounding=ROUND_HALF_UP))

    @property
    def reason(self) -> str:
        if self.complete:
            return ""
        if not self.has_ingredients:
            return "尚未建立配方原料"
        if self.missing_ingredients:
            return "缺少：" + "、".join(self.missing_ingredients)
        return "營養資料不完整"

    @property
    def label(self) -> str:
        if self.complete:
            return f"預估熱量：{self.kcal_int} kcal / 人"
        return "營養資料不完整"


def ingredient_grams(ingredient, amount_per_person) -> Decimal | None:
    if amount_per_person is None:
        return None
    amount = Decimal(str(amount_per_person))
    base_unit = (ingredient.base_unit or "g").strip() or "g"
    if base_unit == "g":
        return amount
    per_unit = getattr(ingredient, "edible_grams_per_unit", None)
    if per_unit is None:
        return None
    per_unit = Decimal(str(per_unit))
    if per_unit <= 0:
        return None
    return amount * per_unit


def _ingredient_kcal(ingredient) -> Decimal | None:
    kcal = getattr(ingredient, "kcal_per_100g", None)
    if kcal is not None:
        return Decimal(str(kcal))
    from services.nutrition_sources import STATUS_MATCHED, match_ingredient_name
    matched = match_ingredient_name(ingredient.name)
    if matched.status == STATUS_MATCHED and matched.food is not None:
        return Decimal(str(matched.food.kcal_per_100g))
    return None


def recipe_nutrition(recipe) -> RecipeNutrition:
    rows = list(recipe.ingredients or ())
    if not rows:
        return RecipeNutrition(recipe_id=recipe.id, kcal_per_person=None, has_ingredients=False)

    total = Decimal("0")
    missing: list[str] = []
    for row in rows:
        ingredient = row.ingredient
        if ingredient is None:
            missing.append("未知食材")
            continue
        grams = ingredient_grams(ingredient, row.grams_per_person)
        kcal_per_100g = _ingredient_kcal(ingredient)
        if grams is None or kcal_per_100g is None:
            missing.append(ingredient.name)
            continue
        total += grams / HUNDRED * kcal_per_100g

    if missing:
        return RecipeNutrition(recipe_id=recipe.id, kcal_per_person=None, missing_ingredients=missing)
    return RecipeNutrition(recipe_id=recipe.id, kcal_per_person=total)


def bulk_recipe_nutrition(recipes) -> dict[int, RecipeNutrition]:
    return {recipe.id: recipe_nutrition(recipe) for recipe in recipes}


def day_kcal(nutritions) -> tuple[int | None, list[str]]:
    total = Decimal("0")
    incomplete: list[str] = []
    for name, nutrition in nutritions:
        if nutrition is None or not nutrition.complete:
            incomplete.append(name)
            continue
        total += nutrition.kcal_per_person
    if incomplete:
        return None, incomplete
    return int(total.quantize(KCAL_QUANTUM, rounding=ROUND_HALF_UP)), []
