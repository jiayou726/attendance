"""Recategorize kitchen recipes and expand the catalog from school-lunch open data.

Usage:
    DATABASE_URL=postgresql://... python scripts/expand_kitchen_catalog.py
    DATABASE_URL=postgresql://... python scripts/expand_kitchen_catalog.py --dry-run
"""
from __future__ import annotations

import argparse
import os
import re
import sys
from collections import defaultdict
from decimal import Decimal
from pathlib import Path

import psycopg2
from psycopg2.extras import execute_values

ROOT = Path(__file__).resolve().parents[1]
DATA = Path(__file__).resolve().parent / "data"
DISH_CANDIDATES = DATA / "gov_lunch_dish_candidates.txt"
ING_CANDIDATES = DATA / "gov_lunch_ingredient_candidates.txt"

VEG_FIXED = {"季節蔬菜", "有機蔬菜", "生產追溯蔬菜", "產銷履歷蔬菜"}
TARGET_TOTAL = 1200
CATEGORIES = ("主食", "主菜", "副菜", "青菜", "湯品", "點心")

OVERRIDE = {
    "芝麻烤腿排": "主菜",
    "蘑菇肉醬麵": "主食",
    "沙茶麵腸": "副菜",
    "海結麵輪": "副菜",
    "蘿蔔麵輪": "副菜",
    "麵輪海結": "副菜",
    "香菇麵筋": "副菜",
    "冬瓜麵筋": "副菜",
    "洋芋燒麵腸": "副菜",
    "洋芋燒麵腸(素)": "副菜",
    "蜜燒麵腸": "副菜",
    "蜜燒麵腸(素)": "副菜",
    "滷麵輪": "副菜",
    "滷麵輪(素)": "副菜",
    "酸菜麵腸": "副菜",
    "酸菜麵腸(素)": "副菜",
    "醬燒麵腸(素)": "副菜",
    "醬爆麵腸": "副菜",
    "原味蒸蛋": "副菜",
    "古早味蒸蛋": "副菜",
    "日式蒸蛋": "副菜",
    "柴魚蒸蛋": "副菜",
    "海結滷蛋": "副菜",
    "紅仁炒蛋": "副菜",
    "紅蔘炒蛋": "副菜",
    "茶葉蛋": "副菜",
    "蕃茄炒蛋": "副菜",
    "蕃茄豆腐蛋": "副菜",
    "起司玉米蒸蛋": "副菜",
    "法式炒蛋": "副菜",
    "肉燥蒸蛋": "副菜",
    "麥克雞塊": "主菜",
    "麥克雞塊*2": "主菜",
    "麥香雞堡": "主菜",
    "和風里肌": "主菜",
    "沙茶豬肉": "主菜",
    "泰式打拋豬": "主菜",
    "香酥蝦捲": "主菜",
    "獅子頭": "主菜",
    "白菜獅子頭": "主菜",
    "紅燒獅子頭": "主菜",
    "紅燒肉丸": "主菜",
    "茄汁獅子頭": "主菜",
    "蔬菜素排": "主菜",
    "牛蒡排": "主菜",
    "牛蒡排(素)": "主菜",
    "海帶串": "副菜",
    "海帶串(素)": "副菜",
    "玉米濃湯(素)": "湯品",
    "鮮竹筍湯(素)": "湯品",
    "肉羹炒竹筍": "副菜",
    "奶皇包": "點心",
    "芝麻包": "點心",
    "紅豆包": "點心",
    "摩摩喳喳": "點心",
    "奶香南瓜": "副菜",
    "紅豆西谷米": "點心",
    "三杯豆包": "主菜",
    "什錦豆包絲": "主菜",
    "炸豆包": "主菜",
    "糖醋豆包丁": "主菜",
    "紅燒豆包": "主菜",
    "銀芽豆包": "主菜",
    "什錦年糕": "副菜",
    "白菜年糕": "副菜",
    "白菜年糕條": "副菜",
    "滷味米血糕": "副菜",
    "滷味血糕": "副菜",
    "滷米血糕": "副菜",
    "肉燥冬瓜": "副菜",
    "肉燥四季豆": "副菜",
    "玉米肉茸": "副菜",
    "什錦豆腐燒": "主菜",
    "什錦豆腐燒(素)": "主菜",
    "海苔豆腐燒": "主菜",
    "海苔豆腐燒(素)": "主菜",
    "炸雙拚": "主菜",
    "炸雙拚(素)": "主菜",
    "宜蘭西魯肉": "主菜",
    "瓜仔肉": "主菜",
    "白玉燒肉": "主菜",
    "義式燉肉": "主菜",
    "馬鈴薯燉肉": "主菜",
    "黃金燉肉": "主菜",
    "鐵板肉柳": "主菜",
}

MEAT_ING_RE = re.compile(r"雞|豬|牛|魚|鴨|骨腿|排骨|里肌|絞肉|肉絲|肉片|肉丁|肉末|肉燥|蝦|海鮮|鯛")
VEG_SWAP = {
    "骨腿丁": "素雞",
    "素雞丁": "素雞丁",
    "絞肉": "素肉",
    "肉絲": "素肉",
    "肉片": "素肉",
    "排骨丁": "素排",
    "排骨": "素排",
    "魚丁": "素排",
    "魚排": "素排",
}


def classify(name: str) -> str:
    if name in VEG_FIXED:
        return "青菜"
    if name in OVERRIDE:
        return OVERRIDE[name]
    n = name
    if any(w in n for w in ("湯", "羹", "西米露")) or "濃湯" in n:
        return "湯品"
    if any(w in n for w in ("飯", "粥")):
        return "主食"
    if any(w in n for w in ("義大利麵", "炒麵", "乾麵", "酢醬麵", "炒米粉", "肉醬麵")):
        return "主食"
    if n.endswith("包") and "豆包" not in n and "鍋貼" not in n:
        return "點心"
    if any(w in n for w in ("豆包", "素雞", "素排", "素魚", "素肉", "素腰", "烤麩", "百頁", "豆腸")):
        return "主菜"
    if any(
        w in n
        for w in (
            "雞", "豬", "牛", "魚", "鴨", "排骨", "大排", "豬排", "雞排", "翅", "腿",
            "里肌", "肉絲", "肉片", "肉丁", "肉末", "肉燥", "打拋", "扣肉", "東坡",
            "控肉", "咕咾", "蝦", "海鮮", "鯛", "虱目", "柳葉魚", "燉肉", "燒肉", "西魯肉",
        )
    ):
        if any(w in n for w in ("蒸蛋", "炒蛋", "滷蛋")):
            return "副菜"
        return "主菜"
    if "肉" in n and not any(w in n for w in ("冬瓜", "四季豆", "肉茸")):
        return "主菜"
    if any(w in n for w in ("蒸蛋", "炒蛋", "滑蛋", "滷蛋", "茶葉蛋")):
        return "副菜"
    return "副菜"


def is_meat_dish(name: str) -> bool:
    if "(素)" in name or name.startswith("素") or "素雞" in name or "素排" in name or "素肉" in name:
        return False
    return bool(re.search(r"雞|豬|牛|魚|鴨|肉|排骨|蝦|海鮮|里肌|翅|腿", name))


def veg_name(name: str) -> str:
    base = re.sub(r"\(素\)$", "", name)
    return f"{base}(素)"


def norm_name(name: str) -> str:
    name = re.sub(r"[\s　]+", "", name)
    name = name.replace("（", "(").replace("）", ")")
    name = re.sub(r"[*＊]\d+$", "", name)
    return name


def connect():
    url = os.environ.get("DATABASE_URL") or os.environ.get("PROD_DB")
    if not url:
        raise SystemExit("Set DATABASE_URL (or PROD_DB) before running.")
    if "sslmode=" not in url:
        url += ("&" if "?" in url else "?") + "sslmode=require"
    return psycopg2.connect(url)


def fetch_recipes(cur):
    cur.execute("SELECT id, name, category FROM kitchen_recipe")
    return list(cur.fetchall())


def fetch_ingredients(cur):
    cur.execute("SELECT id, name FROM kitchen_ingredient")
    return {name: i for i, name in cur.fetchall()}


def recategorize(cur, dry: bool) -> int:
    recipes = fetch_recipes(cur)
    changed = []
    for recipe_id, name, category in recipes:
        new = classify(name)
        if new != category:
            changed.append((new, recipe_id, name, category))
    if dry:
        return len(changed)
    for new, recipe_id, _name, _old in changed:
        cur.execute(
            "UPDATE kitchen_recipe SET category=%s, updated_at=now() WHERE id=%s",
            (new, recipe_id),
        )
    return len(changed)


def ensure_ingredient(cur, cache: dict, name: str, dry: bool) -> int | None:
    if name in cache:
        return cache[name]
    if dry:
        cache[name] = -len(cache) - 1
        return cache[name]
    cur.execute(
        """
        INSERT INTO kitchen_ingredient (
            name, base_unit, purchase_unit, grams_per_purchase_unit, unit_price, order_increment,
            note, active, created_at, updated_at
        )
        VALUES (%s, 'g', 'kg', 1000, 0, 0.001, '政府午餐名單／估計新增', true, now(), now())
        ON CONFLICT (name) DO UPDATE SET name = EXCLUDED.name
        RETURNING id
        """,
        (name,),
    )
    row = cur.fetchone()
    cache[name] = row[0]
    return row[0]


def recipe_has_bom(cur, recipe_id: int) -> bool:
    cur.execute("SELECT 1 FROM kitchen_recipe_ingredient WHERE recipe_id=%s LIMIT 1", (recipe_id,))
    return cur.fetchone() is not None


def guessed_bom(name: str, vegetarian: bool = False) -> list[tuple[str, Decimal]]:
    n = name
    items: list[tuple[str, Decimal]] = []
    if vegetarian or "(素)" in n:
        if "湯" in n or "羹" in n or "濃湯" in n:
            items = [("豆腐", Decimal("25")), ("香菇", Decimal("8"))]
        elif any(w in n for w in ("飯", "麵", "米粉")):
            items = [("白米", Decimal("50"))] if "飯" in n else [("麵腸", Decimal("50"))]
            if "義大利" in n or "麵" in n and "飯" not in n:
                items = [("麵腸", Decimal("55"))]
            if "米粉" in n:
                items = [("冬粉", Decimal("40"))]
        else:
            items = [("素雞", Decimal("70")), ("豆包", Decimal("20"))]
        return items
    if any(w in n for w in ("湯", "羹", "濃湯", "西米露")):
        if "蛋" in n:
            items.append(("雞蛋", Decimal("8")))
        if "玉米" in n:
            items.append(("玉米粒", Decimal("24")))
        if "豆腐" in n:
            items.append(("豆腐", Decimal("20")))
        if "排骨" in n or "大骨" in n:
            items.append(("排骨丁", Decimal("20")))
        if "雞" in n:
            items.append(("骨腿丁", Decimal("20")))
        if "蘿蔔" in n:
            items.append(("冬瓜", Decimal("20")))
        if not items:
            items.append(("冬瓜", Decimal("25")))
        return items
    if any(w in n for w in ("飯", "粥")):
        return [("白米", Decimal("50"))]
    if "米粉" in n:
        return [("冬粉", Decimal("40"))]
    if any(w in n for w in ("義大利麵", "炒麵", "乾麵")):
        return [("麵腸", Decimal("55"))]
    if any(w in n for w in ("雞", "翅", "腿")):
        items.append(("骨腿丁", Decimal("80")))
    elif any(w in n for w in ("豬", "排骨", "里肌", "大排", "豬排")):
        items.append(("排骨丁", Decimal("80")))
    elif "魚" in n:
        items.append(("魚丁", Decimal("70")))
    elif "蝦" in n:
        items.append(("豆包", Decimal("50")))
    elif "蛋" in n:
        items.append(("雞蛋", Decimal("35")))
    elif any(w in n for w in ("豆腐", "豆干", "油腐", "百頁")):
        items.append(("豆腐", Decimal("45")))
    elif "豆包" in n:
        items.append(("豆包", Decimal("50")))
    elif "高麗" in n or "花椰" in n or "青菜" in n or "時蔬" in n:
        items.append(("青菜", Decimal("65")))
    else:
        items.append(("豆腐", Decimal("40")))
    return items


def add_bom(cur, recipe_id: int, cache: dict, items: list[tuple[str, Decimal]], dry: bool, note: str):
    if dry:
        for name, _qty in items:
            ensure_ingredient(cur, cache, name, dry)
        return
    for name, qty in items:
        ingredient_id = ensure_ingredient(cur, cache, name, dry)
        cur.execute(
            """
            INSERT INTO kitchen_recipe_ingredient
                (recipe_id, ingredient_id, grams_per_person, quantity_status, source_note)
            VALUES (%s, %s, %s, 'estimated', %s)
            ON CONFLICT (recipe_id, ingredient_id) DO NOTHING
            """,
            (recipe_id, ingredient_id, qty, note[:255]),
        )


def clone_bom_vegetarian(cur, source_id: int, dest_id: int, cache: dict, dry: bool):
    cur.execute(
        """
        SELECT i.name, ri.grams_per_person
        FROM kitchen_recipe_ingredient ri
        JOIN kitchen_ingredient i ON i.id = ri.ingredient_id
        WHERE ri.recipe_id=%s
        """,
        (source_id,),
    )
    rows = cur.fetchall()
    if not rows:
        add_bom(cur, dest_id, cache, guessed_bom("", vegetarian=True), dry, "缺素對菜／估計")
        return
    mapped: list[tuple[str, Decimal]] = []
    for name, qty in rows:
        mapped.append((VEG_SWAP.get(name, name if not MEAT_ING_RE.search(name) else "素肉"), qty))
    # de-dup names
    merged: dict[str, Decimal] = defaultdict(lambda: Decimal("0"))
    for name, qty in mapped:
        merged[name] += qty
    add_bom(cur, dest_id, cache, list(merged.items()), dry, "由葷菜對應估計")


def insert_recipe(cur, name: str, category: str, dry: bool) -> int | None:
    if dry:
        return None
    cur.execute(
        """
        INSERT INTO kitchen_recipe (name, category, note, active, created_at, updated_at)
        VALUES (%s, %s, %s, true, now(), now())
        ON CONFLICT (name) DO NOTHING
        RETURNING id
        """,
        (name[:120], category, "政府午餐菜色／素食對菜，配方為估計"),
    )
    row = cur.fetchone()
    if row:
        return row[0]
    cur.execute("SELECT id FROM kitchen_recipe WHERE name=%s", (name[:120],))
    found = cur.fetchone()
    return found[0] if found else None


def fill_empty_boms(cur, cache: dict, dry: bool) -> int:
    cur.execute(
        """
        SELECT r.id, r.name FROM kitchen_recipe r
        WHERE NOT EXISTS (
            SELECT 1 FROM kitchen_recipe_ingredient i WHERE i.recipe_id = r.id
        )
        """
    )
    rows = cur.fetchall()
    for recipe_id, name in rows:
        add_bom(cur, recipe_id, cache, guessed_bom(name, vegetarian="(素)" in name), dry, "空白配方補估計")
    return len(rows)


def add_vegetarian_pairs(cur, cache: dict, dry: bool) -> int:
    recipes = fetch_recipes(cur)
    names = {norm_name(name): (recipe_id, name, category) for recipe_id, name, category in recipes}
    created = 0
    for recipe_id, name, category in recipes:
        if category in ("青菜", "點心"):
            continue
        if not is_meat_dish(name):
            continue
        pair = veg_name(name)
        if norm_name(pair) in names:
            continue
        new_id = insert_recipe(cur, pair, category if category != "主食" else "主食", dry)
        names[norm_name(pair)] = (new_id or -1, pair, category)
        if new_id:
            clone_bom_vegetarian(cur, recipe_id, new_id, cache, dry)
        created += 1
    return created


def add_government_dishes(cur, cache: dict, dry: bool, target: int) -> tuple[int, int]:
    recipes = fetch_recipes(cur)
    have = {norm_name(name) for _i, name, _c in recipes}
    current = len(recipes)
    added = 0
    veg_added = 0
    if not DISH_CANDIDATES.exists():
        return 0, 0
    for raw in DISH_CANDIDATES.read_text(encoding="utf-8").splitlines():
        if current + added + veg_added >= target:
            break
        name = norm_name(raw)
        if not name or name in have or name in VEG_FIXED:
            continue
        if name.endswith("(素)") and norm_name(name[:-3]) in have:
            # pair handled separately
            continue
        category = classify(name)
        if category == "青菜":
            continue
        new_id = insert_recipe(cur, name, category, dry)
        have.add(name)
        added += 1
        if new_id:
            add_bom(cur, new_id, cache, guessed_bom(name), dry, "政府午餐菜色／估計")
        if is_meat_dish(name) and current + added + veg_added < target:
            pair = veg_name(name)
            if norm_name(pair) not in have:
                pair_id = insert_recipe(cur, pair, category, dry)
                have.add(norm_name(pair))
                veg_added += 1
                if pair_id:
                    add_bom(cur, pair_id, cache, guessed_bom(pair, vegetarian=True), dry, "政府午餐菜色素食對菜／估計")
    return added, veg_added


def add_missing_ingredients(cur, cache: dict, dry: bool) -> int:
    if not ING_CANDIDATES.exists():
        return 0
    skip = re.compile(
        r"公司|商行|農場|企業|股份|有限|物流|工廠|糧行|農產行|公糧|食品$|碾米|市場|農糧署|電話"
    )
    added = 0
    for raw in ING_CANDIDATES.read_text(encoding="utf-8").splitlines():
        name = raw.strip()
        if not name or skip.search(name) or name.endswith("行") or name.endswith("廠"):
            continue
        if re.match(r"^[\dP.E+\-]+$", name):
            continue
        if name in cache:
            continue
        ensure_ingredient(cur, cache, name, dry)
        added += 1
        if added >= 120:
            break
    return added


def counts(cur):
    cur.execute("SELECT category, count(*) FROM kitchen_recipe GROUP BY 1 ORDER BY 2 DESC")
    cats = cur.fetchall()
    cur.execute("SELECT count(*) FROM kitchen_recipe")
    n_r = cur.fetchone()[0]
    cur.execute("SELECT count(*) FROM kitchen_ingredient")
    n_i = cur.fetchone()[0]
    cur.execute(
        """
        SELECT count(*) FROM kitchen_recipe r
        WHERE NOT EXISTS (SELECT 1 FROM kitchen_recipe_ingredient i WHERE i.recipe_id=r.id)
        """
    )
    empty = cur.fetchone()[0]
    return n_r, n_i, empty, cats


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--target", type=int, default=TARGET_TOTAL)
    args = parser.parse_args()
    conn = connect()
    cur = conn.cursor()
    before = counts(cur)
    cache = fetch_ingredients(cur)
    cat_n = recategorize(cur, args.dry_run)
    empty_n = fill_empty_boms(cur, cache, args.dry_run)
    veg_n = add_vegetarian_pairs(cur, cache, args.dry_run)
    gov_n, gov_veg_n = add_government_dishes(cur, cache, args.dry_run, args.target)
    ing_n = add_missing_ingredients(cur, cache, args.dry_run)
    if args.dry_run:
        conn.rollback()
    else:
        conn.commit()
    after = counts(cur)
    print("before recipes/ings/empty", before[0], before[1], before[2], dict(before[3]))
    print("recategorized", cat_n)
    print("filled empty boms", empty_n)
    print("vegetarian pairs for existing", veg_n)
    print("gov dishes", gov_n, "gov vegetarian pairs", gov_veg_n)
    print("new ingredients", ing_n)
    print("after recipes/ings/empty", after[0], after[1], after[2], dict(after[3]))
    cur.execute("SELECT name, category FROM kitchen_recipe WHERE category='青菜' ORDER BY name")
    print("青菜", cur.fetchall())
    conn.close()


if __name__ == "__main__":
    main()
