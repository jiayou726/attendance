"""Official lot-based XLS schemas and validation, isolated from legacy exports."""
from datetime import date
from decimal import Decimal
import json
import re

from services.ftrace import ROOT, ExportError, _number, _value

BATCH_SCHEMAS = json.loads((ROOT / 'data/ftrace/batch_schemas.json').read_text())


def suggested_batch(code, day):
    """Existing kitchen convention: material/product code plus ROC receipt/made day."""
    value = f'{code}{day.year - 1911:03}{day.month:02}{day.day:02}' if code else ''
    return value if len(value) <= 40 else ''


def identifier(row, field, label, required=True):
    value = _value(row, field, label, required)
    if len(value) > 40 or any(ord(c) < 32 for c in value):
        raise ExportError(f'{label}最多 40 字，不可含控制字元。')
    return value


def iso_date(row, field, label, required=True):
    value = _value(row, field, label, required)
    if not value:
        return None
    try:
        return date.fromisoformat(value)
    except ValueError:
        raise ExportError(f'{label}格式錯誤，請使用 YYYY-MM-DD。') from None


def fmt(day):
    return day.strftime('%Y/%m/%d') if day else ''


def build_batch_sheets(kind, day, payload):
    sheets = {name: [] for name in BATCH_SCHEMAS[kind]}
    source = payload.get('receipts' if kind == 'receiving' else 'meals')
    if not isinstance(source, list) or len(source) > 99 or any(not isinstance(r, dict) for r in source):
        raise ExportError('批號資料格式錯誤，每日最多 99 筆。')
    rows = [r for r in source if r.get('include') is True]
    if not rows:
        raise ExportError('請勾選要匯出的資料。')
    seen = set()
    total_links = 0
    consumption = {}
    for i, row in enumerate(rows, 1):
        code = identifier(row, 'material_code' if kind == 'receiving' else 'product_code', '平台自編碼')
        seq = identifier(row, 'seq', '申報日期流水號')
        batch = identifier(row, 'batch', '原料批號' if kind == 'receiving' else '製造批號')
        identity = (code, seq)
        if identity in seen:
            raise ExportError('同產品／原料的申報流水號重複，請分開編號。')
        seen.add(identity)
        key = ['2', fmt(day), seq, code]
        if kind == 'receiving':
            spec = identifier(row, 'spec_code', '原料規格碼')
            supplier = identifier(row, 'supplier_code', '供應商自編碼')
            received = iso_date(row, 'received_date', '實際收貨日期')
            if received != day:
                raise ExportError('日申報請選實際收貨當日；收貨日期須與申報日期一致。')
            made = iso_date(row, 'manufactured_date', '原料製造日期', False)
            expiry = iso_date(row, 'expiry_date', '原料有效日期', False)
            if made and made > received or expiry and expiry < received or made and expiry and expiry < made:
                raise ExportError('原料製造、收貨或有效日期先後不正確。')
            kg = _number(row, 'kg', '實收淨重（公斤）', 7)
            key += [spec]
            sheets['收貨資料'].append(key + ['', '', ''])
            sheets['收貨資料批號資訊'].append(key + [batch, fmt(made), fmt(expiry), fmt(received), kg, '', ''])
            sheets['收貨資料供應商'].append(key + [supplier])
            continue
        made = iso_date(row, 'manufactured_date', '餐食製造日期')
        expiry = iso_date(row, 'expiry_date', '餐食有效日期')
        if made > day or expiry < made or (kind == 'delivery' and expiry < day):
            raise ExportError('餐食製造、交貨或有效日期先後不正確。')
        kg = _number(row, 'kg', '成品總淨重（公斤）', 3)
        qty = _number(row, 'quantity', '餐食份數', 1 if kind == 'manufacture' else 2)
        if kind == 'delivery':
            customer = identifier(row, 'customer_code', '下游業者自編碼')
            key += [batch, fmt(made), fmt(expiry)]
            sheets['交貨資料'].append(key + [kg, qty, '0', '1', ''])
            sheets['交貨資料下游業者'].append(key + [customer, kg, qty, ''])
            sheets['交貨資料其他'].append(key + ['0', '', '0', '', '', '', '0', '0', '', ''])
            continue
        sheets['製造資料'].append(key + [batch, fmt(made), fmt(expiry), kg, qty])
        materials = row.get('materials')
        if not isinstance(materials, list) or len(materials) > 1000:
            raise ExportError('請核對製造使用的原材料批次。')
        materials = [m for m in materials if not isinstance(m, dict) or m.get('include') is not False]
        if not materials:
            raise ExportError('已排除全部製造原料，請取消這筆製造或還原原料。')
        for material in materials:
            if not isinstance(material, dict):
                raise ExportError('原材料批次格式錯誤。')
            material_code = identifier(material, 'code', '原材料自編碼')
            lots = material.get('lots')
            if not isinstance(lots, list) or not lots or len(lots) > 100:
                raise ExportError(f'{material.get("name", "原材料")}：請選擇或填寫收貨批次。')
            seen_lots = set()
            for lot in lots:
                total_links += 1
                if total_links > 10000 or not isinstance(lot, dict):
                    raise ExportError('使用批次資料過多或格式錯誤。')
                report = _value(lot, 'report_date', '原料收貨申報年月／日期')
                if not re.fullmatch(r'\d{4}-\d{2}(-\d{2})?', report):
                    raise ExportError('原料收貨申報日期須為 YYYY-MM 或 YYYY-MM-DD。')
                try:
                    report_day = date.fromisoformat(report if len(report) == 10 else report + '-01')
                except ValueError:
                    raise ExportError('原料收貨申報日期無效。') from None
                if report_day > made:
                    raise ExportError('不能使用製造日期之後申報收貨的原料。')
                lot_seq = identifier(lot, 'seq', '收貨流水號', False)
                spec = identifier(lot, 'spec_code', '收貨原料規格碼')
                lot_batch = identifier(lot, 'batch', '收貨批號', False)
                received = iso_date(lot, 'received_date', '舊制實際收貨日期', False)
                if bool(lot_batch) == bool(received):
                    raise ExportError('收貨批號與舊制收貨日期請擇一填寫，須對應平台原紀錄。')
                if received and received > made:
                    raise ExportError('實際收貨日期不可晚於餐食製造日期。')
                used = _number(lot, 'kg', '原料使用淨重（公斤）', 7)
                capacity = _value(lot, 'receipt_kg', '該收貨批次重量', False)
                if capacity:
                    cap = _number(lot, 'receipt_kg', '該收貨批次重量', 7)
                    if Decimal(used) > Decimal(cap):
                        raise ExportError('單筆原料使用量大於該收貨批次重量。')
                lot_key = (report, lot_seq, spec, lot_batch, received)
                if lot_key in seen_lots:
                    raise ExportError('同一原材料重複選取相同收貨批次，請合併使用量。')
                seen_lots.add(lot_key)
                shared_key = (material_code,) + lot_key
                used_total, known_cap = consumption.get(shared_key, (Decimal('0'), None))
                known_cap = min(known_cap, Decimal(cap)) if capacity and known_cap is not None else Decimal(cap) if capacity else known_cap
                consumption[shared_key] = (used_total + Decimal(used), known_cap)
                sheets['製造餐食資料原材料供應商'].append(key + [batch,
                    report[:4], report[5:7], report[8:10] if len(report) == 10 else '',
                    lot_seq, material_code, spec, lot_batch, fmt(received), used])
    if any(cap is not None and used > cap for used, cap in consumption.values()):
        raise ExportError('所選餐食合計原料使用量大於該收貨批次重量，請核對拆批或庫存。')
    return sheets
