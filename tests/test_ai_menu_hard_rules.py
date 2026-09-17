from datetime import date
from types import SimpleNamespace

import pytest

import services.ai_menu_engine as engine
from services.ai_menu_engine import CATEGORY_ORDER, Candidate, RuleFeasibilityError, Rules, _vegetarian_compatible
from services.recipe_tags import suggest_tags


def _structure(**overrides):
    result = {category: 0 for category in CATEGORY_ORDER}
    result.update(overrides)
    return result


def test_each_generation_requests_a_fresh_random_seed(monkeypatch):
    candidates = [Candidate(1, "白米飯", "主食", 200)]
    seeds = iter((101, 202))
    seen = []
    original_random = engine.random.Random

    monkeypatch.setattr(engine, "load_candidates", lambda: candidates)
    monkeypatch.setattr(engine.secrets, "randbits", lambda _bits: next(seeds))
    monkeypatch.setattr(engine.random, "Random", lambda seed: (seen.append(seed), original_random(seed))[1])
    rules = Rules(fish_per_week_min=0, fried_per_week_max=9, sweet_soup_per_week_max=9,
                  fish_per_week_max=9,
                  prefer_kcal_in_range=False, exclude_incomplete_nutrition=False)

    for _ in range(2):
        engine.generate(date(2026, 9, 14), date(2026, 9, 14), _structure(主食=1), rules)

    assert seen == [101, 202]


def test_keyword_rules_detect_sweet_soups_and_exclude_non_fish_seafood():
    def recipe(name, category=""):
        return SimpleNamespace(name=name, category=category, ingredients=[])

    assert "sweet_soup" in suggest_tags(recipe("紅豆薏仁湯"))
    assert "sweet_soup" in suggest_tags(recipe("地瓜芋圓湯"))
    assert "sweet_soup" in suggest_tags(recipe("紅豆紫米湯", "湯品"))
    assert "sweet_soup" not in suggest_tags(recipe("*麥片飯", "主食"))
    assert "fish" in suggest_tags(recipe("香煎鯖魚"))
    assert "fish" not in suggest_tags(recipe("魚香紫茄"))
    assert "fish" not in suggest_tags(recipe("素魚排"))
    assert "fish" not in suggest_tags(recipe("柴魚蒸蛋"))
    assert "fish" not in suggest_tags(recipe("蘿蔔燉煮"))
    assert "fish" not in suggest_tags(recipe("魷魚羹"))
    assert "fish" not in suggest_tags(recipe("花枝丸"))
    assert "fried" in suggest_tags(recipe("芝麻球"))
    assert "fried" in suggest_tags(recipe("鍋貼"))
    assert "fried" in suggest_tags(recipe("遊龍鍋貼"))
    assert "fried" in suggest_tags(recipe("香煎鍋貼"))
    assert "fried" not in suggest_tags(recipe("香煎鯖魚"))
    assert "fried" not in suggest_tags(recipe("炸醬麵"))
    assert "fried" not in suggest_tags(recipe("蛋酥白菜"))
    assert "sweet_soup" not in suggest_tags(recipe("冬瓜排骨湯", "湯品"))
    assert "sweet_soup" not in suggest_tags(recipe("銀芽豆包", "點心"))
    assert "sweet_soup" not in suggest_tags(recipe("紅豆包", "點心"))
    assert "fried" not in suggest_tags(recipe("綜合滷味"))
    assert "fried" not in suggest_tags(recipe("敏豆甜不辣"))


def test_vegetarian_compatibility_checks_recipe_ingredients_not_only_name():
    def recipe(name, ingredients):
        return SimpleNamespace(
            name=name,
            ingredients=[SimpleNamespace(ingredient=SimpleNamespace(name=item)) for item in ingredients],
        )

    assert _vegetarian_compatible(recipe("麻婆豆腐", ["豆腐", "素肉"]), set())
    assert _vegetarian_compatible(recipe("麻婆豆腐(素)", ["豆腐", "素肉"]), set())
    assert _vegetarian_compatible(recipe("素魚排", ["素魚排"]), set())
    assert not _vegetarian_compatible(recipe("麻婆豆腐", ["豆腐", "豬絞肉"]), {"vegetarian"})
    assert not _vegetarian_compatible(recipe("白菜羹", ["白菜", "肉羹"]), set())
    assert not _vegetarian_compatible(recipe("蒲瓜黑輪", ["蒲瓜", "黑輪"]), set())
    assert not _vegetarian_compatible(recipe("遊龍鍋貼", ["鍋貼"]), set())
    assert _vegetarian_compatible(recipe("素鍋貼", ["素鍋貼"]), set())
    assert not _vegetarian_compatible(recipe("冬瓜大骨湯", ["冬瓜", "大骨"]), set())
    assert _vegetarian_compatible(recipe("番茄炒蛋", ["番茄", "雞蛋"]), set())
    assert not _vegetarian_compatible(recipe("福州丸", ["福州丸"]), set())
    assert not _vegetarian_compatible(recipe("羅宋湯", ["番茄", "洋芋"]), set())


def test_regular_and_vegetarian_generation_use_separate_main_dish_pools(monkeypatch):
    candidates = [
        Candidate(1, "白飯", "主食", 300, vegetarian_compatible=True),
        Candidate(2, "滷雞腿", "主菜", 200, vegetarian_compatible=False),
        Candidate(3, "素肉燥(素)", "主菜", 180, frozenset({"vegetarian"}), vegetarian_compatible=True),
    ]
    monkeypatch.setattr(engine, "load_candidates", lambda: candidates)
    rules = Rules(recipe_repeat_days=0, main_repeat_days=0, fish_per_week_min=0,
                  fried_per_week_max=9, sweet_soup_per_week_max=9, fish_per_week_max=9,
                  prefer_kcal_in_range=False, exclude_incomplete_nutrition=False)

    regular = engine.generate(date(2026, 9, 14), date(2026, 9, 14),
                              _structure(主食=1, 主菜=1), rules, meal_variant="regular")
    vegetarian = engine.generate(date(2026, 9, 14), date(2026, 9, 14),
                                 _structure(主食=1, 主菜=1), rules, meal_variant="vegetarian")

    assert regular.recipe_at(0, "主菜", 1).name == "滷雞腿"
    assert vegetarian.recipe_at(0, "主菜", 1).name == "素肉燥(素)"


def test_regular_menu_accepts_shared_main_but_rejects_explicit_vegetarian_main():
    shared = Candidate(1, "番茄炒蛋", "主菜", 180, vegetarian_compatible=True)
    labelled = Candidate(2, "素魚排(素)", "主菜", 180, vegetarian_compatible=True)
    regular_copy = Candidate(3, "素魚排", "主菜", 180, vegetarian_compatible=True)

    assert engine.candidate_allowed(shared, "regular")
    assert engine.candidate_allowed(shared, "vegetarian")
    assert not engine.candidate_allowed(labelled, "regular")
    assert engine.candidate_allowed(labelled, "vegetarian")
    assert engine.candidate_allowed(regular_copy, "regular")


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
        fish_per_week_max=99,
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
        fish_per_week_max=99,
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
        fish_per_week_max=1,
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
    assert fish_count <= 1


def test_fried_respects_weekly_max_and_selected_weekday(monkeypatch):
    candidates = [
        Candidate(1, "白飯", "主食", 200),
        Candidate(2, "炸雞", "主菜", 300, frozenset({"fried"})),
        Candidate(3, "滷肉", "主菜", 250, frozenset()),
    ]
    monkeypatch.setattr(engine, "load_candidates", lambda: candidates)
    rules = Rules(
        recipe_repeat_days=0,
        main_repeat_days=0,
        fish_per_week_min=0,
        fish_per_week_max=9,
        fried_per_week_max=1,
        fried_weekdays=(2,),
        sweet_soup_per_week_max=9,
        prefer_kcal_in_range=False,
        exclude_incomplete_nutrition=False,
    )

    result = engine.generate(
        date(2026, 9, 14),
        date(2026, 9, 18),
        _structure(主食=1, 主菜=1),
        rules,
    )

    fried_days = [
        result.dates[key[0]]
        for key, recipe_id in result.assignment.items()
        if recipe_id and "fried" in result.candidates[recipe_id].tags
    ]
    assert fried_days == [date(2026, 9, 16)]


def test_fish_max_and_weekday_from_dict():
    rules = Rules.from_dict({
        "fish_per_week_max": "1",
        "fried_per_week_max": "1",
        "sweet_soup_per_week_max": "1",
        "fish_weekdays": ["2"],
        "fried_weekdays": ["4"],
        "sweet_soup_weekdays": [],
    })
    assert rules.fish_per_week_max == 1
    assert rules.fish_weekdays == (2,)
    assert rules.fried_weekdays == (4,)
    assert rules.sweet_soup_weekdays == ()
    lines = engine.describe_rules(rules)
    assert any("魚類每週最多 1 次" in line and "週三" in line for line in lines)
    assert any("炸物每週最多 1 次" in line and "週五" in line for line in lines)
