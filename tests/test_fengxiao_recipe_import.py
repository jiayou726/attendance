from pathlib import Path

import pytest
from openpyxl import Workbook

from app import create_app
from extensions import db
from models import KitchenIngredient, KitchenRecipe, KitchenRecipeIngredient
from scripts.import_fengxiao_recipe_amounts import import_amounts


@pytest.fixture()
def database(tmp_path):
    path = tmp_path / "recipes.db"
    app = create_app(
        {
            "TESTING": True,
            "SQLALCHEMY_DATABASE_URI": f"sqlite:///{path}",
            "AUTO_CREATE_DB": True,
            "PRODUCTION": False,
        }
    )
    with app.app_context():
        pending = KitchenRecipe(name="待補菜")
        existing = KitchenRecipe(name="原有菜")
        conflict = KitchenRecipe(name="衝突菜")
        new_component = KitchenRecipe(name="可加材料菜")
        missing_component = KitchenIngredient(name="待補材料", base_unit="g")
        fixed_component = KitchenIngredient(name="原有材料", base_unit="g")
        conflict_component = KitchenIngredient(name="衝突材料", base_unit="g")
        extra_component = KitchenIngredient(name="新增材料", base_unit="g")
        inferred_component = KitchenIngredient(name="特殊蔬菜", base_unit="g")
        inferred_recipe = KitchenRecipe(name="待估菜")
        db.session.add_all(
            [
                pending,
                existing,
                conflict,
                new_component,
                missing_component,
                fixed_component,
                conflict_component,
                extra_component,
                inferred_component,
                inferred_recipe,
            ]
        )
        db.session.flush()
        db.session.add_all(
            [
                KitchenRecipeIngredient(
                    recipe_id=pending.id,
                    ingredient_id=missing_component.id,
                    grams_per_person=0,
                    quantity_status="pending",
                ),
                KitchenRecipeIngredient(
                    recipe_id=existing.id,
                    ingredient_id=fixed_component.id,
                    grams_per_person=9,
                    quantity_status="estimated",
                ),
                KitchenRecipeIngredient(
                    recipe_id=inferred_recipe.id,
                    ingredient_id=inferred_component.id,
                    grams_per_person=0,
                    quantity_status="pending",
                ),
            ]
        )
        db.session.commit()
    return path


def _write_source(root: Path):
    workbook = Workbook()
    sheet = workbook.active
    rows = [[None] * 6 for _ in range(4)]
    rows.extend(
        [
            ["待補菜", None, "待補材料", 12, None, None],
            ["原有菜", None, "原有材料", 99, None, None],
            ["可加材料菜", None, "新增材料", "7g", None, None],
            ["衝突菜", None, "衝突材料", 4, None, None],
            ["衝突菜", None, "衝突材料", 5, None, None],
            ["資料庫沒有的菜", None, "新增材料", 8, None, None],
            ["熱量", None, "新增材料", 10, None, None],
        ]
    )
    for row in rows:
        sheet.append(row)
    workbook.save(root / "來源.xlsx")


def test_conservative_import_and_idempotency(database, tmp_path):
    source_root = tmp_path / "source"
    source_root.mkdir()
    _write_source(source_root)

    preview = import_amounts(database, source_root)
    assert preview["updated_pending"] == 1
    assert preview["added_components"] == 1
    assert preview["preserved_existing"] == 1
    assert preview["conflicting_pairs_skipped"] == 1
    assert preview["missing_dishes"] == 1

    applied = import_amounts(database, source_root, apply=True)
    assert applied["total_changes"] == 2

    app = create_app(
        {
            "TESTING": True,
            "SQLALCHEMY_DATABASE_URI": f"sqlite:///{database}",
            "AUTO_CREATE_DB": False,
            "PRODUCTION": False,
        }
    )
    with app.app_context():
        pending = KitchenRecipe.query.filter_by(name="待補菜").one()
        assert float(pending.ingredients[0].grams_per_person) == 12
        assert pending.ingredients[0].quantity_status == "manual"
        existing = KitchenRecipe.query.filter_by(name="原有菜").one()
        assert float(existing.ingredients[0].grams_per_person) == 9
        added = KitchenRecipe.query.filter_by(name="可加材料菜").one()
        assert float(added.ingredients[0].grams_per_person) == 7
        assert KitchenRecipe.query.filter_by(name="衝突菜").one().ingredients == []

    second_run = import_amounts(database, source_root, apply=True)
    assert second_run["total_changes"] == 0
    assert second_run["preserved_existing"] == 3


def test_fill_estimates_marks_remaining_pending_components(database, tmp_path):
    source_root = tmp_path / "source"
    source_root.mkdir()
    _write_source(source_root)

    applied = import_amounts(database, source_root, apply=True, fill_estimates=True)
    assert applied["updated_pending"] == 1
    assert applied["added_components"] == 1
    assert applied["estimated_components"] == 1

    connection = __import__("sqlite3").connect(database)
    try:
        amount, status, note = connection.execute(
            """select c.grams_per_person, c.quantity_status, c.source_note
                 from kitchen_recipe_ingredient c
                 join kitchen_recipe r on r.id = c.recipe_id
                where r.name = '待估菜'"""
        ).fetchone()
        assert amount == 25
        assert status == "estimated"
        assert "估算" in note
        assert connection.execute(
            "select count(*) from kitchen_recipe_ingredient where grams_per_person <= 0"
        ).fetchone()[0] == 0
    finally:
        connection.close()
