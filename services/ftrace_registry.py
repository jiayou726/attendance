"""Read-only, provenance-preserving view of Ftrace's existing records.

This module never writes, reassigns batches, or infers official platform submission.
"""
import json
from decimal import Decimal, InvalidOperation


def _text(value):
    return "" if value is None else str(value).strip()


def _day(value):
    return value.isoformat() if hasattr(value, "isoformat") else _text(value)


def _json(value):
    try:
        parsed = json.loads(value)
        return parsed if isinstance(parsed, dict) else {}
    except (ValueError, TypeError):
        return {}


def _weight_key(value):
    try:
        return ("number", Decimal(_text(value)))
    except (InvalidOperation, ValueError):
        return ("raw", _text(value))


def receipt_registry(imported, documents):
    """One row per *complete* receipt identity; retain every contributing source.

    Incomplete identities are deliberately not merged: two blank batches do not
    establish that two records refer to the same received lot.
    """
    found = {}

    def add(lot, source, fallback):
        report = _text(lot.get("report_date"))
        seq = _text(lot.get("seq"))
        material = _text(lot.get("material_code"))
        spec = _text(lot.get("spec_code"))
        batch = _text(lot.get("batch"))
        received = _text(lot.get("received_date"))
        identity = (report, seq, material, spec, batch, received)
        if not report or not material or not spec or not (batch or received):
            identity = ("incomplete", fallback)
        item = found.setdefault(identity, dict(
            day=report, seq=seq, code=material, spec=spec,
            batch=batch, received_date=received,
            name=_text(lot.get("name")), sources=[],
        ))
        if not item["name"] and lot.get("name"):
            item["name"] = _text(lot["name"])
        item["sources"].append(source)

    for record in imported:
        lot = _json(record.payload_json)
        lot.setdefault("report_date", _text(record.report_date))
        lot.setdefault("material_code", _text(record.material_code))
        lot.setdefault("name", _text(record.name))
        issues = lot.get("issues", [])
        if not isinstance(issues, list):
            issues = ["歷史資料待核對"]
        add(lot, dict(origin="import", id=record.id,
                      weight=_text(lot.get("receipt_kg")), issues=issues,
                      disabled=bool(record.disabled), status="imported"),
            "import:" + _text(record.id))

    for document in documents:
        if document.kind != "receiving":
            continue
        payload = _json(document.payload_json)
        for index, row in enumerate(payload.get("receipts", [])):
            if not isinstance(row, dict) or row.get("include") is not True:
                continue
            batch_mode = payload.get("batch_mode") is True
            add(dict(report_date=_day(document.day),seq=row.get("seq"),
                     material_code=row.get("material_code"),spec_code=row.get("spec_code"),
                     batch=row.get("batch") if batch_mode else "",
                     received_date="" if batch_mode else row.get("received_date"),
                     name=row.get("name")),
                dict(origin="shared", id=document.id, weight=_text(row.get("kg")),
                     issues=[], disabled=False, status=document.status),
                "shared:" + _text(document.id) + ":" + str(index))

    results = []
    for item in found.values():
        sources = item["sources"]
        unique_weights = {}
        for source in sources:
            key = _weight_key(source["weight"])
            unique_weights.setdefault(key, source["weight"])
        conflict = len(unique_weights) > 1
        has_pending = any(source["issues"] for source in sources)
        all_disabled = all(source["disabled"] for source in sources)
        shared_statuses = {s["status"] for s in sources if s["origin"] == "shared"}
        if conflict:
            state = "重量不一致，需核對"
            level = "warn"
        elif has_pending:
            state = "歷史資料待核對"
            level = "warn"
        elif all_disabled:
            state = "已停用"
            level = "muted"
        elif "submitted" in shared_statuses:
            state = "平台已申報（人工確認）"
            level = "ok"
        elif "exported" in shared_statuses:
            state = "XLS 已下載，平台未確認"
            level = "neutral"
        elif any(s["origin"] == "import" for s in sources):
            state = "歷史 XLS 匯入，平台未查證"
            level = "neutral"
        else:
            state = "共用草稿，尚未下載"
            level = "neutral"
        imported_ids = sorted({s["id"] for s in sources if s["origin"] == "import"})
        document_ids = sorted({s["id"] for s in sources if s["origin"] == "shared"})
        item.update(
            weight="／".join(unique_weights.values()),
            source_label="、".join(filter(None, [
                "歷史 XLS" + (f" {sum(s['origin']=='import' for s in sources)} 筆" if sum(s["origin"]=="import" for s in sources)>1 else "") if imported_ids else "",
                "共用收貨" + (f" {sum(s['origin']=='shared' for s in sources)} 筆" if sum(s["origin"]=="shared" for s in sources)>1 else "") if document_ids else "",
            ])),
            imported_ids=imported_ids,
            document_ids=document_ids,
            duplicate=len(sources)>1,
            conflict=conflict,
            state=state,
            level=level,
            issues=list(dict.fromkeys(str(issue) for s in sources for issue in s["issues"])),
        )
        results.append(item)
    return sorted(results, key=lambda x: (x["day"], x["code"], x["seq"]), reverse=True)


def receipt_dashboard(rows, candidates, as_of):
    """Mark the actual *ready* candidates for a manufacturing day.

    "Ready" matches the existing batch selector's source catalog, excluding
    ambiguous, incomplete and conflicting unified records. It is not official
    platform-submission evidence, and it does not imply available stock.
    """
    day = _day(as_of)
    def identity(record):
        return (record["day"], record["seq"], record["code"], record["spec"],
                record["batch"], record["received_date"])

    def candidate_identity(lot):
        return tuple(_text(lot.get(key)) for key in (
            "report_date", "seq", "material_code", "spec_code", "batch", "received_date"
        ))

    usable = {candidate_identity(lot) for lot in candidates if lot.get("ready") is True}
    totals = {"total": len(rows), "ready": 0, "review": 0, "other": 0, "as_of": day}
    for row in rows:
        date_value = row["day"]
        in_range = date_value and (date_value[:7] <= day[:7] if len(date_value) == 7
                                   else date_value <= day)
        try:
            positive_weight = all(Decimal(_text(s["weight"])) > 0 for s in row["sources"]
                                  if not s["disabled"])
        except (InvalidOperation, ValueError, KeyError):
            positive_weight = False
        fully_identified = bool(row["day"] and row["code"] and row["spec"]
                                and (row["batch"] or row["received_date"]))
        if all(s["disabled"] for s in row["sources"]):
            state, label = "other", "已停用"
        elif not in_range:
            state, label = "other", "晚於指定製造日期"
        elif row["conflict"] or row["issues"] or not fully_identified or not positive_weight:
            state, label = "review", "需核對後才能使用"
        elif identity(row) not in usable:
            state, label = "review", "尚未符合製造批次選用條件"
        else:
            state, label = "ready", "製造頁可選用"
        row["use_state"] = state
        row["use_label"] = label
        totals[state] += 1
    return totals


def _too_much(delivered, produced):
    try:
        return Decimal(_text(delivered)) > Decimal(_text(produced))
    except (InvalidOperation, ValueError):
        return False


def meal_registry(documents, kind):
    """Flatten latest manufacture/delivery docs; compare deliveries with saved make."""
    made = {}
    for doc in documents:
        if doc.kind == "manufacture":
            made[_day(doc.day)] = _json(doc.payload_json)

    results = []
    for doc in documents:
        if doc.kind != kind:
            continue
        day = _day(doc.day)
        payload = _json(doc.payload_json)
        batch_mode = payload.get("batch_mode") is True
        for item in payload.get("meals", []):
            if not isinstance(item, dict) or item.get("include") is not True:
                continue
            notes = []
            if kind == "manufacture":
                materials = item.get("materials", [])
                if batch_mode:
                    for material in materials if isinstance(materials, list) else []:
                        if not isinstance(material, dict) or not material.get("code") or not material.get("lots"):
                            notes.append("部分原料尚未對應收貨批次")
                            break
                        if any(not isinstance(lot, dict) or not lot.get("report_date") or not lot.get("spec_code") or not (lot.get("batch") or lot.get("received_date")) for lot in material["lots"]):
                            notes.append("部分原料批次欄位待補")
                            break
                else:
                    notes.append("未使用批號串接")
            else:
                manufacture = made.get(day)
                if not manufacture:
                    notes.append("本系統無同日製造紀錄，無法自動核對")
                else:
                    if (manufacture.get("batch_mode") is True) != batch_mode:
                        notes.append("製造與交貨的批號模式不同")
                    source = next((row for row in manufacture.get("meals", [])
                                   if isinstance(row, dict) and row.get("include") is True and row.get("key") == item.get("key")), None)
                    if not source:
                        notes.append("缺少對應的製造餐食")
                    else:
                        checked = ["product_code"]
                        if batch_mode:
                            checked += ["batch", "manufactured_date", "expiry_date"]
                        if any(_text(source.get(key)) != _text(item.get(key)) for key in checked):
                            notes.append("產品編碼／批號／日期與製造不同")
                        if _too_much(item.get("quantity"), source.get("quantity")):
                            notes.append("交貨份數大於製造份數")
                        if _too_much(item.get("kg"), source.get("kg")):
                            notes.append("交貨重量大於製造重量")
            results.append(dict(
                day=day, name=_text(item.get("name")), key=_text(item.get("key")),
                code=_text(item.get("product_code")), seq=_text(item.get("seq")),
                batch=_text(item.get("batch")) if batch_mode else "",
                quantity=_text(item.get("quantity")), kg=_text(item.get("kg")),
                customer=_text(item.get("customer_code")),
                document_id=doc.id, document_status=doc.status, batch_mode=batch_mode,
                state="；".join(dict.fromkeys(notes)) if notes else "本系統基本資料比對通過",
                level="warn" if notes else "ok",
            ))
    return sorted(results, key=lambda x: (x["day"], x["code"], x["seq"]), reverse=True)
