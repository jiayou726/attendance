"""官方營養資料的讀取與食材對應。

熱量數值一律來自衛福部食藥署「臺灣食品營養成分資料庫」開放資料，
存放在 data/nutrition/tfda_food_energy.csv（見 docs/nutrition_sources.md）。
本模組不會自行推估任何熱量；對不上官方樣品的食材一律標記為待人工確認。
"""

from __future__ import annotations

import csv
import unicodedata
from dataclasses import dataclass
from decimal import Decimal
from functools import lru_cache
from pathlib import Path
from typing import Iterable

DATA_DIR = Path(__file__).resolve().parents[1] / "data" / "nutrition"
ENERGY_CSV = DATA_DIR / "tfda_food_energy.csv"
INGREDIENT_MAP_CSV = DATA_DIR / "ingredient_tfda_map.csv"

SOURCE_LABEL = "衛福部食藥署 臺灣食品營養成分資料庫"

STATUS_MATCHED = "matched"
STATUS_PENDING = "pending"
STATUS_VERIFIED = "verified"

PENDING_NOTE = "名稱無法明確對應官方樣品，待人工確認"


@dataclass(frozen=True)
class OfficialFood:
    """官方資料庫的一筆樣品。"""
    code: str
    name: str
    alias: str
    category: str
    kcal_per_100g: Decimal
    waste_percent: str
    unit_weight: str


def normalize_name(value: str) -> str:
    if not value:
        return ""
    text = unicodedata.normalize("NFKC", value)
    return "".join(text.split())


def _decimal(value: str) -> Decimal | None:
    value = (value or "").strip()
    if not value:
        return None
    try:
        return Decimal(value)
    except (ArithmeticError, ValueError):
        return None


@lru_cache(maxsize=1)
def official_foods() -> dict[str, OfficialFood]:
    foods: dict[str, OfficialFood] = {}
    if not ENERGY_CSV.exists():
        return foods
    with ENERGY_CSV.open(encoding="utf-8", newline="") as handle:
        for row in csv.DictReader(handle):
            kcal = _decimal(row.get("kcal_per_100g", ""))
            if kcal is None:
                continue
            code = (row.get("tfda_code") or "").strip()
            if not code:
                continue
            foods[code] = OfficialFood(
                code=code,
                name=(row.get("sample_name") or "").strip(),
                alias=(row.get("alias") or "").strip(),
                category=(row.get("category") or "").strip(),
                kcal_per_100g=kcal,
                waste_percent=(row.get("waste_percent") or "").strip(),
                unit_weight=(row.get("unit_weight") or "").strip(),
            )
    return foods


@dataclass(frozen=True)
class MapEntry:
    code: str
    confidence: str
    note: str


@lru_cache(maxsize=1)
def ingredient_code_map() -> dict[str, MapEntry]:
    mapping: dict[str, MapEntry] = {}
    if not INGREDIENT_MAP_CSV.exists():
        return mapping
    with INGREDIENT_MAP_CSV.open(encoding="utf-8", newline="") as handle:
        for row in csv.DictReader(handle):
            name = normalize_name(row.get("ingredient_name", ""))
            if not name:
                continue
            mapping[name] = MapEntry(
                code=(row.get("tfda_code") or "").strip(),
                confidence=(row.get("confidence") or "").strip(),
                note=(row.get("note") or "").strip(),
            )
    return mapping


@dataclass
class MatchResult:
    status: str
    food: OfficialFood | None = None
    confidence: str = ""
    note: str = ""


def match_ingredient_name(name: str) -> MatchResult:
    key = normalize_name(name)
    entry = ingredient_code_map().get(key)
    if entry is None:
        return MatchResult(status=STATUS_PENDING, note=PENDING_NOTE)
    if not entry.code:
        return MatchResult(status=STATUS_PENDING, note=entry.note or PENDING_NOTE)
    food = official_foods().get(entry.code)
    if food is None:
        return MatchResult(status=STATUS_PENDING, note=f"對應編號 {entry.code} 不在官方資料檔中，待人工確認")
    return MatchResult(status=STATUS_MATCHED, food=food, confidence=entry.confidence or "medium", note=entry.note)


def apply_match(ingredient, result: MatchResult) -> bool:
    if ingredient.nutrition_verified:
        return False
    before = (
        ingredient.kcal_per_100g, ingredient.nutrition_source,
        ingredient.nutrition_food_name, ingredient.nutrition_code,
        ingredient.nutrition_confidence, ingredient.nutrition_status,
        ingredient.nutrition_note,
    )
    if result.status == STATUS_MATCHED and result.food is not None:
        ingredient.kcal_per_100g = result.food.kcal_per_100g
        ingredient.nutrition_source = SOURCE_LABEL
        ingredient.nutrition_food_name = result.food.name
        ingredient.nutrition_code = result.food.code
        ingredient.nutrition_confidence = result.confidence
        ingredient.nutrition_status = STATUS_MATCHED
        ingredient.nutrition_note = result.note or None
    else:
        ingredient.kcal_per_100g = None
        ingredient.nutrition_source = None
        ingredient.nutrition_food_name = None
        ingredient.nutrition_code = None
        ingredient.nutrition_confidence = None
        ingredient.nutrition_status = STATUS_PENDING
        ingredient.nutrition_note = result.note or PENDING_NOTE
    after = (
        ingredient.kcal_per_100g, ingredient.nutrition_source,
        ingredient.nutrition_food_name, ingredient.nutrition_code,
        ingredient.nutrition_confidence, ingredient.nutrition_status,
        ingredient.nutrition_note,
    )
    return before != after


REPORT_FIELDS = (
    "ingredient_id", "current_name", "base_unit", "matched_official_name",
    "official_code", "kcal_per_100g", "source", "confidence", "status", "note",
)


def build_report_rows(ingredients: Iterable) -> list[dict]:
    rows = []
    for ingredient in ingredients:
        rows.append({
            "ingredient_id": ingredient.id,
            "current_name": ingredient.name,
            "base_unit": ingredient.base_unit,
            "matched_official_name": ingredient.nutrition_food_name or "",
            "official_code": ingredient.nutrition_code or "",
            "kcal_per_100g": "" if ingredient.kcal_per_100g is None else str(ingredient.kcal_per_100g.normalize()),
            "source": ingredient.nutrition_source or "",
            "confidence": ingredient.nutrition_confidence or "",
            "status": ingredient.nutrition_status or STATUS_PENDING,
            "note": ingredient.nutrition_note or "",
        })
    return rows


def write_report(path: Path, rows: list[dict]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(REPORT_FIELDS))
        writer.writeheader()
        writer.writerows(rows)
