"""官方營養資料與團膳食材對應。

優先使用衛福部食藥署「臺灣食品營養成分資料庫」樣品。
若食材名稱無法唯一對應官方樣品，mapping 可提供明確標示的 low-confidence
fallback_kcal；fallback 必須在 note 說明採用的代表品、區間或估值依據，
不可冒充為 TFDA 精確樣品。
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
FALLBACK_SOURCE_LABEL = "團膳熱量預設值（TFDA 近似樣品／官方區間／產品估值）"

STATUS_MATCHED = "matched"
STATUS_ESTIMATED = "estimated"
STATUS_PENDING = "pending"
STATUS_VERIFIED = "verified"

PENDING_NOTE = "名稱無法明確對應官方樣品，且尚未提供 fallback_kcal"


@dataclass(frozen=True)
class OfficialFood:
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


def _decimal(value) -> Decimal | None:
    value = str(value or "").strip()
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
            code = (row.get("tfda_code") or "").strip()
            if kcal is None or not code:
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
    fallback_kcal: Decimal | None = None


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
                fallback_kcal=_decimal(row.get("fallback_kcal")),
            )
    return mapping


@dataclass
class MatchResult:
    status: str
    food: OfficialFood | None = None
    fallback_kcal: Decimal | None = None
    confidence: str = ""
    note: str = ""

    @property
    def kcal_per_100g(self) -> Decimal | None:
        if self.food is not None:
            return self.food.kcal_per_100g
        return self.fallback_kcal

    @property
    def source_label(self) -> str:
        if self.status == STATUS_MATCHED:
            return SOURCE_LABEL
        if self.status == STATUS_ESTIMATED:
            return FALLBACK_SOURCE_LABEL
        return ""


def match_ingredient_name(name: str) -> MatchResult:
    key = normalize_name(name)
    entry = ingredient_code_map().get(key)
    if entry is None:
        return MatchResult(status=STATUS_PENDING, note=PENDING_NOTE)

    if entry.code:
        food = official_foods().get(entry.code)
        if food is not None:
            return MatchResult(
                status=STATUS_MATCHED,
                food=food,
                confidence=entry.confidence or "medium",
                note=entry.note,
            )
        if entry.fallback_kcal is None:
            return MatchResult(
                status=STATUS_PENDING,
                note=f"對應編號 {entry.code} 不在本專案官方資料檔中，且無 fallback_kcal",
            )

    if entry.fallback_kcal is not None:
        return MatchResult(
            status=STATUS_ESTIMATED,
            fallback_kcal=entry.fallback_kcal,
            confidence=entry.confidence or "low",
            note=entry.note,
        )

    return MatchResult(status=STATUS_PENDING, note=entry.note or PENDING_NOTE)


def apply_match(ingredient, result: MatchResult) -> bool:
    """把 match 寫到食材主檔；人工 verified 的值永遠不覆蓋。"""
    if ingredient.nutrition_verified:
        return False

    before = (
        ingredient.kcal_per_100g, ingredient.nutrition_source,
        ingredient.nutrition_food_name, ingredient.nutrition_code,
        ingredient.nutrition_confidence, ingredient.nutrition_status,
        ingredient.nutrition_note,
    )

    kcal = result.kcal_per_100g
    if kcal is not None:
        ingredient.kcal_per_100g = kcal
        ingredient.nutrition_source = result.source_label
        ingredient.nutrition_food_name = result.food.name if result.food else "團膳預設代表值"
        ingredient.nutrition_code = result.food.code if result.food else None
        ingredient.nutrition_confidence = result.confidence or ("high" if result.food else "low")
        ingredient.nutrition_status = result.status
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
