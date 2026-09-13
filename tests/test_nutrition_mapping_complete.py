import csv

from services.nutrition_sources import (
    INGREDIENT_MAP_CSV,
    STATUS_ESTIMATED,
    STATUS_MATCHED,
    match_ingredient_name,
)


def test_every_mapping_row_resolves_to_kcal():
    """既有 mapping 不應再留下會讓配方熱量中斷的空洞。"""
    with INGREDIENT_MAP_CSV.open(encoding="utf-8", newline="") as handle:
        rows = list(csv.DictReader(handle))

    assert rows
    unresolved = []
    for row in rows:
        name = (row.get("ingredient_name") or "").strip()
        if not name:
            continue
        result = match_ingredient_name(name)
        if result.kcal_per_100g is None:
            unresolved.append((name, result.note))
        else:
            assert result.status in {STATUS_MATCHED, STATUS_ESTIMATED}

    assert unresolved == []


def test_known_ambiguous_names_use_explicit_fallbacks():
    expected = {
        "地瓜": 126,
        "玉米粒": 174,
        "香菇": 39,
        "排骨": 287,
        "絞肉": 212,
        "雞肉": 119,
        "雞塊": 229,
        "三色丁": 100,
    }
    for name, kcal in expected.items():
        result = match_ingredient_name(name)
        assert result.status == STATUS_ESTIMATED
        assert int(result.kcal_per_100g) == kcal


def test_exact_tfda_mapping_still_has_priority():
    result = match_ingredient_name("白米")
    assert result.status == STATUS_MATCHED
    assert result.food is not None
    assert result.food.code == "A05002"
    assert int(result.kcal_per_100g) == 354
