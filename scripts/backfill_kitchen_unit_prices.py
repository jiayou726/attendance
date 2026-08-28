"""Backfill ingredient prices from linked historical supplier catalog rows.

The operation is deliberately conservative.  A historical price is eligible
only when it belongs to the ingredient's selected supplier (or the ingredient
has no selected supplier), falls inside the requested date window, and can be
converted exactly into the ingredient's purchase unit.
"""

from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass
from datetime import date
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from app import create_app
from extensions import db
from models import KitchenIngredient, KitchenSupplierItem


TAIWAN_CATTI_KG = Decimal("0.6")
PRICE_QUANTUM = Decimal("0.0001")
PACKAGE_MASS_RE = re.compile(
    r"^1\s*([^=＝\s]+)\s*[=＝]\s*(\d+(?:\.\d+)?)\s*(kg|斤)$",
    re.IGNORECASE,
)


@dataclass(frozen=True)
class PriceCandidate:
    ingredient: KitchenIngredient
    supplier_item: KitchenSupplierItem
    converted_price: Decimal
    conversion_note: str


def convert_price(
    price: Decimal,
    source_unit: str,
    target_unit: str,
    package_conversion: str | None,
) -> tuple[Decimal, str] | None:
    """Convert one source-unit price to the target purchase unit."""
    if source_unit == target_unit:
        return price, f"{source_unit}→{target_unit}"
    if source_unit == "斤" and target_unit == "kg":
        return price / TAIWAN_CATTI_KG, "1斤＝0.6kg"
    if source_unit == "kg" and target_unit == "斤":
        return price * TAIWAN_CATTI_KG, "1斤＝0.6kg"

    match = PACKAGE_MASS_RE.fullmatch((package_conversion or "").strip())
    if not match or match.group(1) != source_unit:
        return None
    mass = Decimal(match.group(2))
    mass_unit = match.group(3).lower()
    mass_kg = mass if mass_unit == "kg" else mass * TAIWAN_CATTI_KG
    if mass_kg <= 0:
        return None
    if target_unit == "kg":
        return price / mass_kg, package_conversion.strip()
    if target_unit == "斤":
        return price / (mass_kg / TAIWAN_CATTI_KG), package_conversion.strip()
    return None


def select_candidate(
    ingredient: KitchenIngredient,
    start_date: date,
    end_date: date,
    minimum_price: Decimal,
    maximum_price: Decimal,
) -> PriceCandidate | None:
    query = KitchenSupplierItem.query.filter(
        KitchenSupplierItem.ingredient_id == ingredient.id,
        KitchenSupplierItem.active.is_(True),
        KitchenSupplierItem.last_unit_price.isnot(None),
        KitchenSupplierItem.last_unit_price > 0,
        KitchenSupplierItem.last_purchase_date >= start_date,
        KitchenSupplierItem.last_purchase_date <= end_date,
    )
    if ingredient.supplier_id is not None:
        query = query.filter(KitchenSupplierItem.supplier_id == ingredient.supplier_id)

    candidates: list[PriceCandidate] = []
    for item in query.all():
        converted = convert_price(
            Decimal(item.last_unit_price),
            item.unit,
            ingredient.purchase_unit,
            item.package_conversion,
        )
        if converted is None:
            continue
        converted_price, conversion_note = converted
        if not minimum_price <= converted_price <= maximum_price:
            continue
        candidates.append(PriceCandidate(
            ingredient=ingredient,
            supplier_item=item,
            converted_price=converted_price.quantize(PRICE_QUANTUM, rounding=ROUND_HALF_UP),
            conversion_note=conversion_note,
        ))
    if not candidates:
        return None
    return max(candidates, key=lambda row: (
        row.supplier_item.last_purchase_date or date.min,
        row.supplier_item.order_count or 0,
        row.supplier_item.id,
    ))


def price_source_note(candidate: PriceCandidate) -> str:
    item = candidate.supplier_item
    supplier_name = item.supplier.name if item.supplier else "未知廠商"
    return (
        f"單價由歷史帳單回填：{supplier_name}/{item.name}，"
        f"{item.last_purchase_date.isoformat()}，{item.last_unit_price}/{item.unit}，"
        f"換算 {candidate.conversion_note}"
    )[:255]


def merge_source_note(existing_note: str | None, source_note: str) -> str:
    existing = (existing_note or "").strip()
    if not existing:
        return source_note[:255]
    if source_note in existing:
        return existing[:255]
    return f"{existing}；{source_note}"[:255]


def backfill_prices(
    start_date: date,
    end_date: date,
    minimum_price: Decimal,
    maximum_price: Decimal,
    overwrite: bool = False,
) -> list[PriceCandidate]:
    selected = []
    for ingredient in KitchenIngredient.query.order_by(KitchenIngredient.name).all():
        if not overwrite and Decimal(ingredient.unit_price or 0) > 0:
            continue
        candidate = select_candidate(
            ingredient, start_date, end_date, minimum_price, maximum_price
        )
        if candidate is None:
            continue
        ingredient.unit_price = candidate.converted_price
        ingredient.note = merge_source_note(
            ingredient.note,
            price_source_note(candidate),
        )
        selected.append(candidate)
    return selected


def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument("--start-date", type=date.fromisoformat, default=date(2012, 1, 1))
    parser.add_argument("--end-date", type=date.fromisoformat, default=date(2025, 12, 31))
    parser.add_argument("--minimum-price", type=Decimal, default=Decimal("5"))
    parser.add_argument("--maximum-price", type=Decimal, default=Decimal("2000"))
    parser.add_argument("--overwrite", action="store_true")
    parser.add_argument("--apply", action="store_true")
    return parser.parse_args()


def main():
    args = parse_args()
    if args.end_date < args.start_date:
        raise SystemExit("end date must be on or after start date")
    app = create_app({"AUTO_CREATE_DB": True})
    with app.app_context():
        selected = backfill_prices(
            args.start_date,
            args.end_date,
            args.minimum_price,
            args.maximum_price,
            overwrite=args.overwrite,
        )
        for candidate in selected:
            item = candidate.supplier_item
            print(
                f"{candidate.ingredient.name}: {candidate.converted_price}/"
                f"{candidate.ingredient.purchase_unit} <- {item.last_unit_price}/"
                f"{item.unit} {item.supplier.name} {item.last_purchase_date}"
            )
        if args.apply:
            db.session.commit()
        else:
            db.session.rollback()
        print(f"eligible={len(selected)}")
        print("APPLIED" if args.apply else "DRY RUN ONLY")


if __name__ == "__main__":
    main()
