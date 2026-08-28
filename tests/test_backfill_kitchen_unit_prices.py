from datetime import date
from decimal import Decimal

from scripts.backfill_kitchen_unit_prices import convert_price, merge_source_note


def test_converts_catti_price_to_price_per_kg():
    converted, note = convert_price(Decimal("26"), "斤", "kg", None)
    assert converted.quantize(Decimal("0.0001")) == Decimal("43.3333")
    assert note == "1斤＝0.6kg"


def test_converts_package_price_when_mass_is_explicit():
    converted, note = convert_price(Decimal("650"), "件", "kg", "1件＝50斤")
    assert converted.quantize(Decimal("0.0001")) == Decimal("21.6667")
    assert note == "1件＝50斤"


def test_rejects_count_only_or_mismatched_package_conversion():
    assert convert_price(Decimal("350"), "包", "kg", "1件＝4包") is None
    assert convert_price(Decimal("1200"), "件", "kg", "1件＝6包") is None


def test_same_unit_needs_no_conversion():
    assert convert_price(Decimal("155"), "kg", "kg", None) == (
        Decimal("155"),
        "kg→kg",
    )


def test_price_source_is_appended_without_discarding_existing_note():
    assert merge_source_note("原有備註", "價格來源") == "原有備註；價格來源"
    assert merge_source_note(None, "價格來源") == "價格來源"
