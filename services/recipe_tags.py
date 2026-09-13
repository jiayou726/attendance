"""菜色標記（排菜規則用）。

人工標記仍是菜色分類的主要來源；另外提供純關鍵字的建議標記工具，
只看菜名與配方原料名稱，不呼叫任何語言模型。
為避免「甜湯最多一次」等規則因漏勾標記而失效，魚類、炸物、甜湯
三個排菜規則會自動採用關鍵字建議；其餘分類仍只採人工確認標記。
"""

from __future__ import annotations
from dataclasses import dataclass

SOURCE_MANUAL = "manual"
SOURCE_SUGGESTED = "suggested"
AUTO_RULE_TAGS = frozenset({"fish", "fried", "sweet_soup"})
NON_FISH_SEAFOOD = ("魷魚", "章魚", "墨魚", "花枝", "蝦", "小卷", "海鮮")

@dataclass(frozen=True)
class TagDef:
    key: str
    label: str
    keywords: tuple[str, ...]
    hint: str = ""

TAG_DEFS: tuple[TagDef, ...] = (
    TagDef("fish", "魚類", ("魚", "鮭", "鯖", "鱈", "旗魚", "虱目", "鯛", "柳葉", "鯊", "鮪"), "實際魚類，用於每週魚類規則；蝦、魷魚、花枝、小卷等海鮮不計入"),
    TagDef("chicken", "雞肉", ("雞", "翅", "棒棒腿", "骨腿", "清腿", "翅腿"), "主要蛋白質為雞肉"),
    TagDef("pork", "豬肉", ("豬", "肉絲", "肉片", "肉丁", "絞肉", "排骨", "梅花", "五花", "里肌", "貢丸", "肉羹", "熱狗", "火腿", "培根"), "主要蛋白質為豬肉"),
    TagDef("beef", "牛肉", ("牛",), "主要蛋白質為牛肉"),
    TagDef("egg", "蛋類", ("蛋",), "含蛋"),
    TagDef("tofu", "豆製品", ("豆腐", "豆干", "豆乾", "干絲", "豆皮", "豆包", "百頁", "油豆腐", "麵腸", "素肉", "毛豆", "豆腸"), "含豆製品"),
    TagDef("fried", "炸物", ("炸", "酥", "天婦羅", "甜不辣", "薯條", "薯餅", "雞塊", "麵筋泡", "可樂餅", "春捲"), "油炸或半成品炸物"),
    TagDef("sweet_soup", "甜湯", ("甜湯", "紅豆湯", "綠豆湯", "薏仁湯", "花生湯", "西米露", "芋圓", "粉圓", "山粉圓", "銀耳", "蓮子", "湯圓", "麥片"), "甜湯或甜點心"),
    TagDef("curry", "咖哩", ("咖哩",), "咖哩口味"),
    TagDef("vegetarian", "素食", ("素", "全素", "蔬"), "全素可用"),
    TagDef("dark_vegetable", "深色蔬菜", ("青江", "菠菜", "地瓜葉", "空心菜", "小白菜", "芥藍", "青花", "花椰", "紅鳳菜", "莧菜", "油菜", "A菜", "皇宮菜"), "深色蔬菜"),
)
TAG_BY_KEY = {definition.key: definition for definition in TAG_DEFS}
TAG_KEYS = tuple(definition.key for definition in TAG_DEFS)

def tag_label(key: str) -> str:
    definition = TAG_BY_KEY.get(key)
    return definition.label if definition else key

def suggest_tags(recipe) -> set[str]:
    haystack = recipe.name or ""
    for row in recipe.ingredients or ():
        if row.ingredient is not None:
            haystack += " " + (row.ingredient.name or "")
    suggested = set()
    for definition in TAG_DEFS:
        target = haystack
        if definition.key == "fish":
            # 「魷魚／章魚／墨魚」字面含「魚」，但不應拿來充每週魚類次數。
            for token in NON_FISH_SEAFOOD:
                target = target.replace(token, " ")
        if any(keyword in target for keyword in definition.keywords):
            suggested.add(definition.key)
    if "vegetarian" in suggested and suggested & {"chicken", "pork", "beef", "fish"}:
        suggested.discard("vegetarian")
    return suggested

def effective_rule_tags(recipe) -> set[str]:
    """Tags used by the planner.

    Manual tags are always honored. Fish/fried/sweet-soup tags additionally
    use deterministic keyword detection so hard weekly rules cannot silently
    fail only because an operator forgot to tick a checkbox.
    """
    manual = {row.tag for row in recipe.tags or () if row.source == SOURCE_MANUAL}
    return manual | (suggest_tags(recipe) & AUTO_RULE_TAGS)

def set_manual_tags(session, recipe, keys) -> bool:
    from models import KitchenRecipeTag
    wanted = {key for key in keys if key in TAG_BY_KEY}
    existing = {row.tag: row for row in recipe.tags}
    changed = False
    for key in wanted - set(existing):
        session.add(KitchenRecipeTag(recipe_id=recipe.id, tag=key, source=SOURCE_MANUAL))
        changed = True
    for key in set(existing) - wanted:
        session.delete(existing[key]); changed = True
    for key in wanted & set(existing):
        row = existing[key]
        if row.source != SOURCE_MANUAL:
            row.source = SOURCE_MANUAL; changed = True
    return changed

def manual_tag_map(recipes) -> dict[int, set[str]]:
    return {recipe.id: {row.tag for row in recipe.tags if row.source == SOURCE_MANUAL} for recipe in recipes}
