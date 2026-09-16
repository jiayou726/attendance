"""青菜菜色與有機蔬菜食材整理規則。

青菜只保留「產銷履歷蔬菜」「有機蔬菜」兩道。
葉菜炒青菜類副菜不進菜單，改成有機蔬菜食材。
"""
from __future__ import annotations
import re

VEG_FIXED = frozenset({"有機蔬菜", "產銷履歷蔬菜"})
VEG_RETIRE = frozenset({"季節蔬菜", "生產追溯蔬菜"})

LEAF_VEGETABLES = (
    "青江菜", "菠菜", "地瓜葉", "空心菜", "小白菜", "芥藍菜", "芥藍",
    "花椰菜", "綠花椰菜", "青花菜", "紅鳳菜", "莧菜", "油菜", "小油菜",
    "A菜", "皇宮菜", "大陸妹", "大陸A", "黑葉白菜", "青松菜", "小松菜",
    "小松葉", "青松葉", "芥菜", "茼蒿", "福山萵苣", "福山", "蚵白菜",
    "鵝白菜", "油麥菜", "千寶菜", "荷葉白菜", "皺葉白菜", "高麗菜",
)
LEAF_BY_LENGTH = tuple(sorted(LEAF_VEGETABLES, key=len, reverse=True))
COOK_PREFIXES = ("清炒", "蒜炒", "脆炒", "鮮炒", "蒜香", "薑絲", "翠綠", "清燙", "炒", "燙")
CERT_PREFIXES = ("吉園圃", "有機", "產銷履歷", "產銷")
GENERIC_LEAF_DISHES = frozenset({
    "時令蔬菜", "鮮炒時蔬", "炒青菜", "燙青菜", "青菜",
})
MIXED_KEEP = re.compile(
    r"冬粉|寬粉|干片|年糕|滷|杏鮑|南瓜|茄子|玉米|毛豆|豆腐|培根|豆包|豆芽|銀芽|木耳|香菇|四季豆|鮮菇|彩椒"
)
VEG_SIDE_SUJI = re.compile(r"^(彩椒|杏鮑|青椒|芹菜|木耳|金針|海帶|茄子|四季豆|白菜|高麗|豆芽|玉米)")


def _bare(name: str) -> str:
    return re.sub(r"\(素\)$", "", (name or "").strip())


def leafy_core(name: str) -> str | None:
    """若菜名是葉菜炒青菜，回傳蔬菜本名。"""
    n = _bare(name)
    if n in GENERIC_LEAF_DISHES:
        return "青菜"
    if n in VEG_FIXED or n in VEG_RETIRE:
        return None
    if MIXED_KEEP.search(n) and n not in GENERIC_LEAF_DISHES:
        return None
    rest = n
    for prefix in CERT_PREFIXES:
        if rest.startswith(prefix):
            rest = rest[len(prefix):]
            break
    for cook in COOK_PREFIXES:
        if rest.startswith(cook):
            rest = rest[len(cook):]
            break
    if rest in ("時蔬", "青菜"):
        return "青菜"
    if rest.endswith("花椰") or rest.endswith("花椰菜") or rest.endswith("青花菜"):
        return "花椰菜"
    for leaf in LEAF_BY_LENGTH:
        if rest == leaf or rest == leaf.rstrip("菜"):
            return "A菜" if leaf in {"A菜", "大陸A"} and rest in {"A菜", "大陸A", "A"} else leaf
    return None


def is_leafy_side(name: str, category: str | None = None) -> bool:
    if (category or "").strip() not in {"", "副菜", "青菜"}:
        return False
    if _bare(name) in VEG_FIXED:
        return False
    if _bare(name) in VEG_RETIRE:
        return True
    return leafy_core(name) is not None


def organic_ingredient_name(dish_or_veg: str) -> str | None:
    core = leafy_core(dish_or_veg)
    if core is None:
        n = _bare(dish_or_veg)
        if n.startswith("有機") and n not in VEG_FIXED:
            return n
        return None
    if core == "青菜":
        return None
    if core == "大陸A":
        core = "A菜"
    if core == "福山":
        core = "福山萵苣"
    if core == "芥藍":
        core = "芥藍菜"
    if not core.startswith("有機"):
        return f"有機{core}"
    return core


def is_vegetable_suji_side(name: str) -> bool:
    """彩椒素雞這類『蔬菜名 + 素雞』是副菜，不是主菜。"""
    n = _bare(name)
    if "素雞排" in n or n.startswith(("照燒素雞", "三杯素雞", "宮保素雞", "糖醋素雞", "香煎素雞", "紅燒素雞")):
        return False
    return "素雞" in n and bool(VEG_SIDE_SUJI.search(n))


def classify_category(name: str, current: str | None = None) -> str:
    n = (name or "").strip()
    if _bare(n) in VEG_FIXED:
        return "青菜"
    if is_leafy_side(n, current or "副菜"):
        return "青菜" if _bare(n) in VEG_FIXED else current or "副菜"
    if is_vegetable_suji_side(n):
        return "副菜"
    return current or "副菜"
