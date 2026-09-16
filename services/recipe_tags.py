"""菜色標記與葷素主分類（排菜規則用）。

葷／素／葷素共用使用 KitchenRecipe.diet_type，必須人工／資料清理後確認，
不再因為菜名含「素」或以「(素)」結尾就直接判定。
魚類、炸物、甜湯等次要標記仍保留人工勾選；suggest_tags 只提供提示。
"""

from __future__ import annotations
from dataclasses import dataclass

SOURCE_MANUAL = "manual"
SOURCE_SUGGESTED = "suggested"

DIET_TYPES = ("meat", "vegetarian", "shared")
DIET_LABELS = {
    "meat": "葷",
    "vegetarian": "素",
    "shared": "葷素共用",
}
DIET_ALLOWED = {
    "regular": {"meat", "shared"},
    "vegetarian": {"vegetarian", "shared"},
}
DIET_TAG_KEYS = set(DIET_TYPES)


@dataclass(frozen=True)
class TagDef:
    key: str
    label: str
    keywords: tuple[str, ...]
    hint: str = ""


# diet_type 已獨立成互斥主分類，因此這裡不再提供「vegetarian」關鍵字標籤。
TAG_DEFS: tuple[TagDef, ...] = (
    TagDef("fish", "魚類", ("魚", "鮭", "鯖", "鱈", "旗", "虱目", "鯛", "柳葉", "鯊", "鮪", "花枝", "魷", "蝦", "小卷", "海鮮"), "含魚類或海鮮，用於每週魚類規則"),
    TagDef("chicken", "雞肉", ("雞", "翅", "棒棒腿", "骨腿", "清腿", "翅腿"), "主要蛋白質為雞肉"),
    TagDef("pork", "豬肉", ("豬", "肉絲", "肉片", "肉丁", "絞肉", "排骨", "梅花", "五花", "里肌", "貢丸", "肉羹", "熱狗", "火腿", "培根"), "主要蛋白質為豬肉"),
    TagDef("beef", "牛肉", ("牛",), "主要蛋白質為牛肉"),
    TagDef("egg", "蛋類", ("蛋",), "含蛋"),
    TagDef("tofu", "豆製品", ("豆腐", "豆干", "豆乾", "干絲", "豆皮", "豆包", "百頁", "油豆腐", "麵腸", "毛豆", "豆腸"), "含豆製品"),
    TagDef("fried", "炸物", ("炸", "酥", "天婦羅", "甜不辣", "薯條", "薯餅", "雞塊", "麵筋泡", "可樂餅", "春捲"), "油炸或半成品炸物"),
    TagDef("sweet_soup", "甜湯", ("甜湯", "紅豆湯", "綠豆湯", "薏仁湯", "花生湯", "西米露", "芋圓", "粉圓", "山粉圓", "銀耳", "蓮子", "湯圓"), "甜湯或甜點心"),
    TagDef("curry", "咖哩", ("咖哩",), "咖哩口味"),
    TagDef("dark_vegetable", "深色蔬菜", ("青江", "菠菜", "地瓜葉", "空心菜", "小白菜", "芥藍", "青花", "花椰", "紅鳳菜", "莧菜", "油菜", "A菜", "皇宮菜"), "深色蔬菜"),
)
TAG_BY_KEY = {definition.key: definition for definition in TAG_DEFS}
TAG_KEYS = tuple(definition.key for definition in TAG_DEFS)


def tag_label(key: str) -> str:
    definition = TAG_BY_KEY.get(key)
    return definition.label if definition else key


def diet_label(value: str | None) -> str:
    return DIET_LABELS.get(value or "", "未分類")


def _ingredient_names(recipe) -> list[str]:
    return [
        (row.ingredient.name or "").strip()
        for row in (recipe.ingredients or ())
        if row.ingredient is not None and (row.ingredient.name or "").strip()
    ]


def suggest_tags(recipe) -> set[str]:
    """只提供次要標籤建議；葷素主分類絕不在這裡猜。"""
    names = _ingredient_names(recipe)
    ingredient_text = " ".join(names)
    name_text = recipe.name or ""
    suggested: set[str] = set()

    for definition in TAG_DEFS:
        # 肉類／魚類只看實際配方，不因為「素魚排」「素雞」等菜名誤判。
        source = ingredient_text if definition.key in {"fish", "chicken", "pork", "beef", "egg", "tofu", "dark_vegetable"} else f"{name_text} {ingredient_text}"
        if any(keyword in source for keyword in definition.keywords):
            suggested.add(definition.key)

    # 已確認非葷食時，不應出現魚／雞／豬／牛建議；若真的出現代表配方需再檢查。
    if getattr(recipe, "diet_type", None) in {"vegetarian", "shared"}:
        suggested -= {"fish", "chicken", "pork", "beef"}
    return suggested


def set_diet_type(session, recipe, value: str) -> bool:
    """設定唯一葷素主分類，並同步舊 kitchen_recipe_tag 供相容舊程式。"""
    from models import KitchenRecipeTag

    if value not in DIET_TYPES:
        raise ValueError("diet_type must be meat, vegetarian or shared")

    changed = recipe.diet_type != value
    recipe.diet_type = value

    existing = {row.tag: row for row in recipe.tags if row.tag in DIET_TAG_KEYS}
    for key, row in list(existing.items()):
        if key != value:
            session.delete(row)
            changed = True
    row = existing.get(value)
    if row is None:
        session.add(KitchenRecipeTag(recipe_id=recipe.id, tag=value, source=SOURCE_MANUAL))
        changed = True
    elif row.source != SOURCE_MANUAL:
        row.source = SOURCE_MANUAL
        changed = True
    return changed


def set_manual_tags(session, recipe, keys) -> bool:
    from models import KitchenRecipeTag

    # 葷素分類由 diet_type 控制，不允許從一般 checkbox 寫入。
    wanted = {key for key in keys if key in TAG_BY_KEY}
    existing = {row.tag: row for row in recipe.tags if row.tag not in DIET_TAG_KEYS}
    changed = False
    for key in wanted - set(existing):
        session.add(KitchenRecipeTag(recipe_id=recipe.id, tag=key, source=SOURCE_MANUAL))
        changed = True
    for key in set(existing) - wanted:
        session.delete(existing[key])
        changed = True
    for key in wanted & set(existing):
        row = existing[key]
        if row.source != SOURCE_MANUAL:
            row.source = SOURCE_MANUAL
            changed = True
    return changed


def manual_tag_map(recipes) -> dict[int, set[str]]:
    return {
        recipe.id: {
            row.tag
            for row in recipe.tags
            if row.source == SOURCE_MANUAL and row.tag not in DIET_TAG_KEYS
        }
        for recipe in recipes
    }
