"""Targeted query optimizations for recipe list/detail pages.

The recipe list renders a per-recipe AP total, so letting SQLAlchemy lazily load
Recipe.ingredients and RecipeIngredient.ingredient creates an N+1 query pattern.
Keep the model defaults unchanged for other kitchen pages and only eager-load the
relationships on the two recipe GET views that need them.
"""

from functools import wraps

from flask import abort, render_template, request
from sqlalchemy.orm import selectinload

from extensions import db
from models import KitchenIngredient, KitchenRecipe, KitchenRecipeIngredient
from blueprints.order_tool import CATEGORIES, _int, _recipe_cost, _recipe_total_g


def _recipe_bom_load():
    return (
        selectinload(KitchenRecipe.ingredients)
        .selectinload(KitchenRecipeIngredient.ingredient)
    )


def install_recipe_performance_views(app):
    """Replace only recipe GET views with eager-loading equivalents."""

    original_recipes = app.view_functions["order_tool.recipes"]

    @wraps(original_recipes)
    def recipes_view(*args, **kwargs):
        # Preserve the existing POST/create behavior exactly as-is.
        if request.method != "GET":
            return original_recipes(*args, **kwargs)

        q = request.args.get("q", "").strip()
        query = KitchenRecipe.query.options(_recipe_bom_load())
        if q:
            query = query.filter(KitchenRecipe.name.ilike(f"%{q}%"))
        rows = query.order_by(
            KitchenRecipe.active.desc(),
            KitchenRecipe.category,
            KitchenRecipe.name,
        ).all()
        edit_row = (
            db.session.get(KitchenRecipe, _int(request.args.get("edit"), default=0))
            if request.args.get("edit")
            else None
        )
        return render_template(
            "kitchen/recipes.html",
            rows=rows,
            edit_row=edit_row,
            categories=CATEGORIES,
            q=q,
        )

    original_recipe_detail = app.view_functions["order_tool.recipe_detail"]

    @wraps(original_recipe_detail)
    def recipe_detail_view(recipe_id: int, *args, **kwargs):
        recipe = (
            KitchenRecipe.query.options(_recipe_bom_load())
            .filter(KitchenRecipe.id == recipe_id)
            .one_or_none()
        )
        if recipe is None:
            abort(404)

        ingredients_all = (
            KitchenIngredient.query.filter_by(active=True)
            .order_by(KitchenIngredient.name)
            .all()
        )
        return render_template(
            "kitchen/recipe_detail.html",
            recipe=recipe,
            ingredients=ingredients_all,
            ingredient_options=[
                {
                    "id": ingredient.id,
                    "name": ingredient.name,
                    "base_unit": ingredient.base_unit or "g",
                    "purchase_unit": ingredient.purchase_unit,
                }
                for ingredient in ingredients_all
            ],
            categories=CATEGORIES,
            total_g=_recipe_total_g(recipe),
            total_cost=_recipe_cost(recipe),
        )

    app.view_functions["order_tool.recipes"] = recipes_view
    app.view_functions["order_tool.recipe_detail"] = recipe_detail_view
