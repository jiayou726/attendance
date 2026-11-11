"""Private platform references and unambiguous, editable export defaults."""
import base64
import binascii
import json
import os
import zlib

from flask import current_app


def _private_reference_data():
    empty = dict(materials=[], products=[], parties=[])
    raw = current_app.config.get("FTRACE_REFERENCES_JSON",
                                 os.environ.get("FTRACE_REFERENCES_JSON", ""))
    try:
        if not raw:
            packed = current_app.config.get("FTRACE_REFERENCES_ZLIB_BASE64",
                os.environ.get("FTRACE_REFERENCES_ZLIB_BASE64", ""))
            if not packed:
                return empty
            decoder = zlib.decompressobj()
            raw = decoder.decompress(base64.b64decode(packed, validate=True), 2_000_001)
            if len(raw) > 2_000_000 or not decoder.eof:
                raise ValueError
        data = json.loads(raw) if isinstance(raw, (str, bytes)) else raw
        result = {}
        for kind in empty:
            rows = data.get(kind, [])
            if not isinstance(rows, list) or len(rows) > 20_000:
                raise ValueError
            result[kind] = []
            for row in rows:
                if not all(isinstance(row.get(key), str) for key in ("name", "code")):
                    raise ValueError
                clean = dict(name=row["name"], code=row["code"])
                aliases = row.get("aliases", [])
                if not isinstance(aliases, list) or not all(isinstance(a, str) for a in aliases):
                    raise ValueError
                clean["aliases"] = aliases
                if kind == "materials":
                    specs = row.get("specs", [])
                    if not isinstance(specs, list):
                        raise ValueError
                    clean["specs"] = []
                    for spec in specs:
                        if not isinstance(spec.get("code"), str) or not isinstance(spec.get("supplier_name", ""), str):
                            raise ValueError
                        clean_spec = dict(code=spec["code"], supplier_name=spec.get("supplier_name", ""))
                        for key in ("product_name", "package_label", "manufacturer_name"):
                            if isinstance(spec.get(key), str):
                                clean_spec[key] = spec[key]
                        clean["specs"].append(clean_spec)
                if kind == "products" and isinstance(row.get("grams"), (int, float)):
                    clean["grams"] = row["grams"]
                result[kind].append(clean)
        return result
    except (ValueError, TypeError, KeyError, AttributeError, zlib.error, binascii.Error):
        current_app.logger.warning("Invalid private ftrace references; leave codes editable.")
        return empty


def _db_parties():
    """Read the restricted shared transaction-party registry; never expose contacts."""
    from extensions import db
    from sqlalchemy import inspect, text

    try:
        engine = db.engine
        schema = "public" if engine.dialect.name == "postgresql" else None
        if not inspect(engine).has_table("kitchen_ftrace_party", schema=schema):
            return []
        table = "public.kitchen_ftrace_party" if schema else "kitchen_ftrace_party"
        with engine.connect() as connection:
            rows = connection.execute(text(
                f"SELECT name, code, relation, aliases FROM {table} "
                "WHERE active = true ORDER BY code"
            )).mappings().all()
        return [
            dict(name=row["name"], code=row["code"], relation=row["relation"],
                 aliases=list(row["aliases"] or []))
            for row in rows
        ]
    except Exception as error:
        current_app.logger.warning(
            "Could not read ftrace transaction parties: %s", type(error).__name__)
        return []


def reference_data():
    """Combine private material/product mappings with database-stored party codes.

    Merge by code, not name: one legal name may have multiple platform codes.
    Preserve verified aliases from the existing private configuration.
    """
    result = _private_reference_data()
    by_code = {row["code"]: dict(row) for row in result["parties"]}
    for party in _db_parties():
        old = by_code.get(party["code"])
        if old:
            merged_aliases = [
                old["name"], *(old.get("aliases") or []), *(party.get("aliases") or [])
            ]
            party["aliases"] = [
                value for value in dict.fromkeys(merged_aliases)
                if value != party["name"]
            ]
        by_code[party["code"]] = party
    result["parties"] = list(by_code.values())
    return result


def _normalise(value):
    return "".join((value or "").split())


def _unique(rows, name):
    name = _normalise(name)
    if not name:
        return None
    matches = [r for r in rows if name in {_normalise(r["name"]),
                                          *(_normalise(a) for a in r.get("aliases", []))}]
    return matches[0] if len(matches) == 1 else None


def apply_reference_defaults(data, references):
    parties = references["parties"]
    supplier_parties = [
        party for party in parties
        if not party.get("relation") or party["relation"][1:2] == "1"
    ]
    customer_parties = [
        party for party in parties
        if party.get("relation", "")[:1] == "1"
    ]
    for row in data["receipts"]:
        material = _unique(references["materials"], row["name"])
        if material:
            row["material_code"] = material["code"]
            specs = material.get("specs", [])
            supplier_specs = [s for s in specs if _normalise(s["supplier_name"]) == _normalise(row["supplier_name"])]
            if len(supplier_specs) == 1:
                row["spec_code"] = supplier_specs[0]["code"]
            elif len(specs) == 1:
                row["spec_code"] = specs[0]["code"]
        supplier = _unique(supplier_parties, row["supplier_name"])
        if supplier:
            row["supplier_code"] = supplier["code"]
    for meal in data["meals"]:
        if not meal.get("customer_code"):
            customer = _unique(customer_parties, meal.get("school_key") or meal["name"])
            if customer:
                meal["customer_code"] = customer["code"]
        for material in meal["materials"]:
            match = _unique(references["materials"], material["name"])
            if match:
                material["code"] = match["code"]
    return data
