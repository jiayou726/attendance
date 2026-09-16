from datetime import date

from app import create_app
from extensions import db
from models import KitchenIngredient, KitchenRecipe, KitchenRecipeIngredient
from services.ai_menu_engine import Rules, generate


def _app(tmp_path):
    return create_app({
        "TESTING": True,
        "SECRET_KEY": "diet-rule-test",
        "SQLALCHEMY_DATABASE_URI": f"sqlite:///{tmp_path / 'diet.db'}",
        "AUTO_CREATE_DB": True,
        "KITCHEN_CSRF_ENABLED": False,
        "PRODUCTION": False,
    })


def test_ai_menu_uses_explicit_diet_type_instead_of_recipe_name(tmp_path):
    app = _app(tmp_path)
    with app.app_context():
        ingredient = KitchenIngredient(
            name="測試豆腐",
            base_unit="g",
            purchase_unit="kg",
            grams_per_purchase_unit=1000,
            unit_price=1,
            order_increment=1,
            active=True,
            kcal_per_100g=100,
            nutrition_status="verified",
            nutrition_verified=True,
        )
        db.session.add(ingredient)
        db.session.flush()

        recipes = [
            KitchenRecipe(name="沒有肉字但其實是葷菜", category="主菜", active=True, diet_type="meat"),
            KitchenRecipe(name="名稱沒有加素字的素菜", category="主菜", active=True, diet_type="vegetarian"),
            KitchenRecipe(name="大家都能吃的豆腐", category="主菜", active=True, diet_type="shared"),
        ]
        db.session.add_all(recipes)
        db.session.flush()
        for recipe in recipes:
            db.session.add(KitchenRecipeIngredient(
                recipe_id=recipe.id,
                ingredient_id=ingredient.id,
                grams_per_person=50,
            ))
        db.session.commit()

        structure = {"主食": 0, "主菜": 2, "副菜": 0, "青菜": 0, "湯品": 0, "點心": 0}
        rules = Rules(
            repeat_days=0,
            fish_per_week_min=0,
            fried_per_week_max=99,
            sweet_soup_per_week_max=99,
            exclude_incomplete_nutrition=True,
        )

        regular = generate(
            date(2026, 9, 17), date(2026, 9, 17), structure, rules,
            kcal_min=None, kcal_max=None, diet_mode="regular",
        )[0]
        vegetarian = generate(
            date(2026, 9, 17), date(2026, 9, 17), structure, rules,
            kcal_min=None, kcal_max=None, diet_mode="vegetarian",
        )[0]

        regular_types = {item.diet_type for item in regular.items}
        vegetarian_types = {item.diet_type for item in vegetarian.items}
        regular_names = {item.name for item in regular.items}
        vegetarian_names = {item.name for item in vegetarian.items}

        assert regular_types == {"meat", "shared"}
        assert vegetarian_types == {"vegetarian", "shared"}
        assert "名稱沒有加素字的素菜" not in regular_names
        assert "沒有肉字但其實是葷菜" not in vegetarian_names
