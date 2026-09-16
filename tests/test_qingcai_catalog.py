from services.qingcai_catalog import (
    classify_category,
    is_leafy_side,
    is_vegetable_suji_side,
    organic_ingredient_name,
)


def test_qingcai_keeps_only_two_fixed_dishes():
    assert classify_category("有機蔬菜") == "青菜"
    assert classify_category("產銷履歷蔬菜") == "青菜"
    assert is_leafy_side("季節蔬菜", "青菜")
    assert is_leafy_side("生產追溯蔬菜", "青菜")
    assert not is_leafy_side("有機蔬菜", "青菜")
    assert not is_leafy_side("產銷履歷蔬菜", "青菜")


def test_leafy_sides_become_organic_ingredients():
    assert is_leafy_side("炒青江菜", "副菜")
    assert is_leafy_side("蒜香菠菜", "副菜")
    assert is_leafy_side("有機A菜", "副菜")
    assert is_leafy_side("吉園圃空心菜", "副菜")
    assert is_leafy_side("炒綠花椰菜", "副菜")
    assert organic_ingredient_name("炒青江菜") == "有機青江菜"
    assert organic_ingredient_name("有機A菜") == "有機A菜"
    assert organic_ingredient_name("蒜香花椰") == "有機花椰菜"
    assert organic_ingredient_name("鮮炒時蔬") is None


def test_mixed_vegetable_sides_stay_dishes():
    assert not is_leafy_side("田園時蔬", "副菜")
    assert not is_leafy_side("彩繪時蔬", "副菜")
    assert not is_leafy_side("時蔬冬粉", "副菜")
    assert not is_leafy_side("培根高麗菜", "副菜")
    assert not is_leafy_side("彩椒花椰", "副菜")
    assert not is_leafy_side("木耳炒高麗菜", "副菜")


def test_pepper_suji_is_side_not_main():
    assert is_vegetable_suji_side("彩椒素雞")
    assert is_vegetable_suji_side("杏鮑素雞")
    assert is_vegetable_suji_side("杏鮑素雞丁")
    assert not is_vegetable_suji_side("照燒素雞")
    assert not is_vegetable_suji_side("蔬菜素排")
    assert classify_category("彩椒素雞", "主菜") == "副菜"
    assert classify_category("照燒素雞", "主菜") == "主菜"
