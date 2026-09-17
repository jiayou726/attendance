"""整理青菜菜色與有機蔬菜食材。

- 青菜只留 產銷履歷蔬菜、有機蔬菜
- 葉菜類副菜停用／刪除，改成有機蔬菜食材
- 彩椒素雞這類蔬菜+素雞改回副菜

Usage:
    DATABASE_URL=postgresql://... python scripts/reorg_qingcai_catalog.py
    DATABASE_URL=postgresql://... python scripts/reorg_qingcai_catalog.py --dry-run
"""
from __future__ import annotations

import argparse
import os
import sys
from decimal import Decimal
from pathlib import Path

import psycopg2

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from services.qingcai_catalog import (  # noqa: E402
    VEG_FIXED,
    VEG_RETIRE,
    is_leafy_side,
    is_vegetable_suji_side,
    organic_ingredient_name,
)

ORGANIC_VEGETABLES = (
    "有機青江菜", "有機菠菜", "有機地瓜葉", "有機空心菜", "有機小白菜",
    "有機油菜", "有機小油菜", "有機A菜", "有機芥藍菜", "有機花椰菜",
    "有機青花菜", "有機莧菜", "有機高麗菜", "有機大陸妹", "有機蚵白菜",
    "有機鵝白菜", "有機油麥菜", "有機千寶菜", "有機小松菜", "有機青松菜",
    "有機黑葉白菜", "有機荷葉白菜", "有機皺葉白菜", "有機芥菜",
    "有機紅鳳菜", "有機皇宮菜", "有機茼蒿", "有機福山萵苣",
)


def connect():
    url = os.environ.get("DATABASE_URL") or os.environ.get("PROD_DB")
    if not url:
        raise SystemExit("Set DATABASE_URL (or PROD_DB) before running.")
    if "sslmode=" not in url:
        url += ("&" if "?" in url else "?") + "sslmode=require"
    return psycopg2.connect(url)


def fetch_recipes(cur):
    cur.execute("SELECT id, name, category, active FROM kitchen_recipe ORDER BY name")
    return list(cur.fetchall())


def fetch_ingredients(cur):
    cur.execute("SELECT id, name FROM kitchen_ingredient")
    return {name: recipe_id for recipe_id, name in cur.fetchall()}


def recipe_in_use(cur, recipe_id: int) -> bool:
    cur.execute(
        """
        SELECT
            (SELECT count(*) FROM kitchen_menu_plan_item WHERE recipe_id=%s)
          + (SELECT count(*) FROM kitchen_menu_draft_item WHERE recipe_id=%s)
          + (SELECT count(*) FROM kitchen_daily_dish_note WHERE recipe_id=%s)
        """,
        (recipe_id, recipe_id, recipe_id),
    )
    return (cur.fetchone()[0] or 0) > 0


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
        VALUES (%s, 'g', 'kg', 1000, 0, 0.001, '有機蔬菜', true, now(), now())
        ON CONFLICT (name) DO UPDATE
            SET note='有機蔬菜', active=true, updated_at=now()
        RETURNING id
        """,
        (name,),
    )
    cache[name] = cur.fetchone()[0]
    return cache[name]


def mark_organic_note(cur, name: str, dry: bool):
    if dry:
        return
    cur.execute(
        """
        UPDATE kitchen_ingredient
        SET note='有機蔬菜', updated_at=now()
        WHERE name=%s AND coalesce(note,'') <> '有機蔬菜'
        """,
        (name,),
    )


def replace_recipe_refs(cur, old_id: int, new_id: int, dry: bool):
    if dry or old_id == new_id:
        return
    cur.execute(
        """
        DELETE FROM kitchen_menu_plan_item a
        WHERE a.recipe_id=%s
          AND EXISTS (
              SELECT 1 FROM kitchen_menu_plan_item b
              WHERE b.plan_id=a.plan_id AND b.recipe_id=%s
          )
        """,
        (old_id, new_id),
    )
    cur.execute(
        "UPDATE kitchen_menu_plan_item SET recipe_id=%s WHERE recipe_id=%s",
        (new_id, old_id),
    )
    cur.execute(
        "UPDATE kitchen_menu_draft_item SET recipe_id=%s WHERE recipe_id=%s",
        (new_id, old_id),
    )
    cur.execute(
        """
        DELETE FROM kitchen_daily_dish_note a
        WHERE a.recipe_id=%s
          AND EXISTS (
              SELECT 1 FROM kitchen_daily_dish_note b
              WHERE b.service_date=a.service_date AND b.variant=a.variant AND b.recipe_id=%s
          )
        """,
        (old_id, new_id),
    )
    cur.execute(
        "UPDATE kitchen_daily_dish_note SET recipe_id=%s WHERE recipe_id=%s",
        (new_id, old_id),
    )


def deactivate_or_delete(cur, recipe_id: int, dry: bool):
    if dry:
        return
    if recipe_in_use(cur, recipe_id):
        cur.execute(
            "UPDATE kitchen_recipe SET active=false, updated_at=now() WHERE id=%s",
            (recipe_id,),
        )
        return
    cur.execute("DELETE FROM kitchen_recipe_tag WHERE recipe_id=%s", (recipe_id,))
    cur.execute("DELETE FROM kitchen_recipe_ingredient WHERE recipe_id=%s", (recipe_id,))
    cur.execute("DELETE FROM kitchen_recipe WHERE id=%s", (recipe_id,))


def ensure_organic_recipe_bom(cur, cache: dict, dry: bool):
    cur.execute("SELECT id FROM kitchen_recipe WHERE name='有機蔬菜'")
    row = cur.fetchone()
    if not row:
        return
    recipe_id = row[0]
    ingredient_id = ensure_ingredient(cur, cache, "有機蔬菜", dry)
    if dry:
        return
    cur.execute(
        "SELECT count(*) FROM kitchen_recipe_ingredient WHERE recipe_id=%s",
        (recipe_id,),
    )
    if cur.fetchone()[0]:
        return
    cur.execute(
        """
        INSERT INTO kitchen_recipe_ingredient
            (recipe_id, ingredient_id, grams_per_person, quantity_status, source_note)
        VALUES (%s, %s, %s, 'estimated', '有機蔬菜代表用量')
        """,
        (recipe_id, ingredient_id, Decimal("60")),
    )


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()
    conn = connect()
    cur = conn.cursor()
    recipes = fetch_recipes(cur)
    cache = fetch_ingredients(cur)
    organic_added = []
    retired = []
    suji_fixed = []

    for name in ORGANIC_VEGETABLES:
        if name not in cache:
            ensure_ingredient(cur, cache, name, args.dry_run)
            organic_added.append(name)
        else:
            mark_organic_note(cur, name, args.dry_run)

    cur.execute("SELECT id FROM kitchen_recipe WHERE name='有機蔬菜'")
    organic_recipe = cur.fetchone()
    organic_id = organic_recipe[0] if organic_recipe else None

    for recipe_id, name, category, active in recipes:
        if is_vegetable_suji_side(name) and category == "主菜":
            suji_fixed.append(name)
            if not args.dry_run:
                cur.execute(
                    "UPDATE kitchen_recipe SET category='副菜', updated_at=now() WHERE id=%s",
                    (recipe_id,),
                )
        if name in VEG_FIXED:
            continue
        if not is_leafy_side(name, category) and name not in VEG_RETIRE:
            continue
        ing_name = organic_ingredient_name(name)
        if ing_name and ing_name not in cache:
            ensure_ingredient(cur, cache, ing_name, args.dry_run)
            organic_added.append(ing_name)
        elif ing_name:
            mark_organic_note(cur, ing_name, args.dry_run)
        if organic_id and recipe_in_use(cur, recipe_id):
            replace_recipe_refs(cur, recipe_id, organic_id, args.dry_run)
        deactivate_or_delete(cur, recipe_id, args.dry_run)
        retired.append(name)

    ensure_organic_recipe_bom(cur, cache, args.dry_run)

    if args.dry_run:
        conn.rollback()
    else:
        conn.commit()

    cur.execute("SELECT name FROM kitchen_recipe WHERE category='青菜' AND active ORDER BY name")
    qingcai = [row[0] for row in cur.fetchall()]
    cur.execute("SELECT count(*) FROM kitchen_recipe WHERE active")
    active_n = cur.fetchone()[0]
    print("organic ingredients added", len(organic_added), organic_added)
    print("leafy dishes retired", len(retired))
    for name in retired:
        print(" -", name)
    print("suji recategorized", suji_fixed)
    print("active 青菜", qingcai)
    print("active recipes", active_n)
    conn.close()


if __name__ == "__main__":
    main()
