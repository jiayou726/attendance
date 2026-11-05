"""Private platform references and unambiguous, editable export defaults."""
import base64
import binascii
import json
import os
import zlib

from flask import current_app


def reference_data():
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
        supplier = _unique(references["parties"], row["supplier_name"])
        if supplier:
            row["supplier_code"] = supplier["code"]
    for meal in data["meals"]:
        for material in meal["materials"]:
            match = _unique(references["materials"], material["name"])
            if match:
                material["code"] = match["code"]
    return data
