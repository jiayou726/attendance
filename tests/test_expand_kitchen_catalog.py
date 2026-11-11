from scripts.expand_kitchen_catalog import classify, veg_name, is_meat_dish


def test_fixed_vegetables_stay_qingcai():
    assert classify("有機蔬菜") == "青菜"
    assert classify("產銷履歷蔬菜") == "青菜"
    assert classify("季節蔬菜") == "青菜"
    assert classify("生產追溯蔬菜") == "青菜"


def test_vegetarian_steak_is_main():
    assert classify("蔬菜素排") == "主菜"


def test_steamed_egg_is_side():
    assert classify("原味蒸蛋") == "副菜"
    assert classify("番茄炒蛋") == "副菜"


def test_bean_curd_roll_is_main_not_snack():
    assert classify("炸豆包") == "主菜"
    assert classify("奶皇包") == "點心"


def test_meat_pair_name():
    assert is_meat_dish("三杯雞")
    assert not is_meat_dish("三杯雞(素)")
    assert veg_name("三杯雞") == "三杯雞(素)"
