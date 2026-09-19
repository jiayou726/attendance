"""把魚類／炸物／甜湯標記寫進資料庫，並把素雞改成副菜。

Usage:
    DATABASE_URL=postgresql://... python scripts/sync_recipe_rule_tags.py
    DATABASE_URL=postgresql://... python scripts/sync_recipe_rule_tags.py --dry-run
"""
from __future__ import annotations

import argparse
import os
import sys
from collections import defaultdict
from pathlib import Path
from types import SimpleNamespace

import psycopg2

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from services.qingcai_catalog import is_vegetable_suji_side  # noqa: E402
from services.recipe_tags import AUTO_RULE_TAGS, suggest_tags  # noqa: E402


def connect():
    url = os.environ.get("DATABASE_URL") or os.environ.get("PROD_DB")
    if not url:
        raise SystemExit("Set DATABASE_URL (or PROD_DB) before running.")
    if "sslmode=" not in url:
        url += ("&" if "?" in url else "?") + "sslmode=require"
    return psycopg2.connect(url)


def load_recipes(cur):
    cur.execute("SELECT id, name, category FROM kitchen_recipe WHERE active")
    recipes = {
        recipe_id: SimpleNamespace(id=recipe_id, name=name, category=category, ingredients=[], tags=[])
        for recipe_id, name, category in cur.fetchall()
    }
    cur.execute(
        """
        SELECT ri.recipe_id, i.name
        FROM kitchen_recipe_ingredient ri
        JOIN kitchen_ingredient i ON i.id=ri.ingredient_id
        """
    )
    for recipe_id, ingredient_name in cur.fetchall():
        recipe = recipes.get(recipe_id)
        if recipe is None:
            continue
        recipe.ingredients.append(SimpleNamespace(ingredient=SimpleNamespace(name=ingredient_name)))
    return recipes


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()
    conn = connect()
    cur = conn.cursor()
    recipes = load_recipes(cur)

    suji_fixed = []
    for recipe in recipes.values():
        if is_vegetable_suji_side(recipe.name) and recipe.category == "主菜":
            suji_fixed.append(recipe.name)
            if not args.dry_run:
                cur.execute(
                    "UPDATE kitchen_recipe SET category='副菜', updated_at=now() WHERE id=%s",
                    (recipe.id,),
                )
            recipe.category = "副菜"

    tag_counts = defaultdict(int)
    inserted = 0
    examples = defaultdict(list)
    for recipe in recipes.values():
        wanted = suggest_tags(recipe) & AUTO_RULE_TAGS
        for tag in sorted(wanted):
            tag_counts[tag] += 1
            if len(examples[tag]) < 8:
                examples[tag].append(recipe.name)
            if args.dry_run:
                continue
            cur.execute(
                """
                INSERT INTO kitchen_recipe_tag (recipe_id, tag, source, created_at)
                VALUES (%s, %s, 'manual', now())
                ON CONFLICT (recipe_id, tag) DO UPDATE
                    SET source='manual'
                """,
                (recipe.id, tag),
            )
            if cur.rowcount:
                inserted += 1

    if args.dry_run:
        conn.rollback()
    else:
        conn.commit()

    print("suji recategorized", suji_fixed)
    print("rule tag dishes", dict(tag_counts))
    print("tags written", inserted)
    for tag, names in examples.items():
        print(tag, names)

    if not args.dry_run:
        cur.execute("SELECT tag, count(*) FROM kitchen_recipe_tag GROUP BY 1 ORDER BY 1")
        print("db tags", cur.fetchall())
        cur.execute(
            """
            SELECT r.name, r.category, t.tag
            FROM kitchen_recipe r
            JOIN kitchen_recipe_tag t ON t.recipe_id=r.id
            WHERE r.name LIKE '%鍋貼%' OR r.name LIKE '%照燒素雞%'
            ORDER BY r.name, t.tag
            """
        )
        print("spot check", cur.fetchall())
    conn.close()


if __name__ == "__main__":
    main()
