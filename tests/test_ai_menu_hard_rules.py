from datetime import date
from types import SimpleNamespace

import pytest

import services.ai_menu_engine as engine
from services.ai_menu_engine import CATEGORY_ORDER, Candidate, RuleFeasibilityError, Rules
from services.recipe_tags import suggest_tags


def _structure(**overrides):
    result = {category: 0 for category in CATEGORY_ORDER}
    result.update(overrides)
    return result


def test_keyword_rules_detect_sweet_soups_and_exclude_non_fish_seafood():
    def recipe(name):
        return SimpleNamespace(name=name, ingredients=[])

    assert "sweet_soup" in suggest_tags(recipe("紅豆薏仁湯"))
    assert "sweet_soup" in suggest_tags(recipe("地瓜芋圓湯"))
    assert "fish" in suggest_tags(recipe("香煎鯖魚"))
    assert "fish" not in suggest_tags(recipe("魷魚羹"))
    assert "fish" not in suggest_tags(recipe("花枝丸"))


def test_generate_never_exceeds_weekly_sweet_soup_max(monkeypatch):
    candidates = [
        Candidate(1, "紅豆薏仁湯", "湯品", 100, frozenset({"sweet_soup"})),
        Candidate(2, "地瓜芋圓湯", "湯品", 100, frozenset({"sweet_soup"})),
        Candidate(3, "海帶芽湯", "湯品", 80, frozenset()),
    ]
    monkeypatch.setattr(engine, "load_candidates", lambda: candidates)
    rules = Rules(
        recipe_repeat_days=0,
        main_repeat_days=0,
        fish_per_week_min=0,
        fried_per_week_max=99,
        sweet_soup_per_week_max=1,
        prefer_kcal_in_range=False,
        exclude_incomplete_nutrition=False,
    )

    result = engine.generate(
        date(2026, 9, 14),
        date(2026, 9, 18),
        _structure(湯品=1),
        rules,
    )

    sweet_count = sum(
        "sweet_soup" in result.candidates[recipe_id].tags
        for recipe_id in result.assignment.values()
        if recipe_id
    )
    assert sweet_count <= 1


def test_generate_fails_instead_of_breaking_sweet_soup_hard_cap(monkeypatch):
    candidates = [
        Candidate(1, "紅豆薏仁湯", "湯品", 100, frozenset({"sweet_soup"})),
        Candidate(2, "地瓜芋圓湯", "湯品", 100, frozenset({"sweet_soup"})),
    ]
    monkeypatch.setattr(engine, "load_candidates", lambda: candidates)
    rules = Rules(
        recipe_repeat_days=0,
        main_repeat_days=0,
        fish_per_week_min=0,
        fried_per_week_max=99,
        sweet_soup_per_week_max=1,
        prefer_kcal_in_range=False,
        exclude_incomplete_nutrition=False,
    )

    with pytest.raises(RuleFeasibilityError):
        engine.generate(
            date(2026, 9, 14),
            date(2026, 9, 15),
            _structure(湯品=1),
            rules,
        )


def test_generate_guarantees_weekly_fish_minimum_when_feasible(monkeypatch):
    candidates = [
        Candidate(1, "清蒸鯖魚", "主菜", 180, frozenset({"fish"})),
        Candidate(2, "滷豬肉", "主菜", 220, frozenset()),
    ]
    monkeypatch.setattr(engine, "load_candidates", lambda: candidates)
    rules = Rules(
        recipe_repeat_days=0,
        main_repeat_days=0,
        fish_per_week_min=1,
        fried_per_week_max=99,
        sweet_soup_per_week_max=99,
        prefer_kcal_in_range=False,
        exclude_incomplete_nutrition=False,
    )

    result = engine.generate(
        date(2026, 9, 14),
        date(2026, 9, 18),
        _structure(主菜=1),
        rules,
    )

    fish_count = sum(
        "fish" in result.candidates[recipe_id].tags
        for recipe_id in result.assignment.values()
        if recipe_id
    )
    assert fish_count >= 1
