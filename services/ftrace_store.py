"""Persistent ftrace snapshots and imported receipt provenance."""
from datetime import date
import hashlib
import json
from sqlalchemy import update
from sqlalchemy.exc import IntegrityError
from extensions import db
from ftrace_models import FtraceDocument, FtraceRevision, FtraceImport, FtraceReceipt, now
from services.ftrace import ExportError, KINDS
from services.ftrace_import import parse_receipts

class Conflict(ExportError):
    pass

def encoded(value):
    return json.dumps(value, ensure_ascii=False, allow_nan=False, separators=(',',':'))

def document(day, kind):
    return FtraceDocument.query.filter_by(day=day,kind=kind).first()

def snapshot(row):
    return dict(version=row.version,status=row.status,payload=json.loads(row.payload_json),
                updated_at=row.updated_at.isoformat()) if row else dict(version=0,status='draft',payload=None)

def validate_payload(day, kind, payload):
    if kind not in dict(KINDS) or not isinstance(payload,dict) or payload.get('date')!=day.isoformat():
        raise ExportError('儲存日期或資料類型不一致，請重新整理。')
    if any(not isinstance(payload.get(group),list) for group in ('receipts','meals')):
        raise ExportError('收貨與餐食資料格式錯誤。')
    if sum(len(payload[g]) for g in ('receipts','meals'))>3000 or len(encoded(payload))>2_000_000:
        raise ExportError('資料過多，請分日期儲存。')
    if any(not isinstance(r,dict) for g in ('receipts','meals') for r in payload[g]):
        raise ExportError('紀錄格式錯誤。')

def save_document(day, kind, payload, expected, status='draft'):
    validate_payload(day,kind,payload)
    if status not in ('draft','exported'):
        raise ExportError('儲存狀態無效。')
    try:
        expected=int(expected)
    except (ValueError,TypeError):
        raise ExportError('缺少紀錄版本。') from None
    text=encoded(payload)
    row=document(day,kind)
    if row:
        if row.version!=expected:
            raise Conflict('共用紀錄已由另一個分頁或同事更新。你的修改仍在瀏覽器暫存，請重新整理並核對，避免覆蓋。')
        if row.status=='submitted':
            raise Conflict('這份紀錄已標記平台申報完成，請在資料區建立修訂草稿後再編輯。')
        result=db.session.execute(update(FtraceDocument).where(FtraceDocument.id==row.id,FtraceDocument.version==expected)
            .values(payload_json=text,version=expected+1,status=status,updated_at=now()))
        if result.rowcount!=1:
            db.session.rollback();raise Conflict('共用紀錄版本已變更，請重新整理。')
        version=expected+1
    else:
        if expected!=0:
            raise Conflict('共用紀錄不存在或版本不一致，請重新整理。')
        row=FtraceDocument(day=day,kind=kind,payload_json=text,version=1,status=status)
        db.session.add(row);version=1
    try:
        db.session.flush()
        db.session.add(FtraceRevision(document_id=row.id,version=version,status=status,payload_json=text))
        db.session.commit()
    except IntegrityError:
        db.session.rollback();raise Conflict('同日期資料已由另一位使用者建立，請重新整理後核對。') from None
    db.session.expire_all()
    return snapshot(db.session.get(FtraceDocument,row.id))

def set_status(row, expected, status):
    if row.version!=expected or status not in ('draft','submitted'):
        raise Conflict('紀錄版本已更新，請重新整理後再操作。')
    # Revision is append-only; returning to draft never erases the submitted snapshot.
    result=db.session.execute(update(FtraceDocument).where(FtraceDocument.id==row.id,FtraceDocument.version==expected)
        .values(version=expected+1,status=status,updated_at=now()))
    if result.rowcount!=1:
        db.session.rollback();raise Conflict('紀錄版本已更新，請重新整理。')
    db.session.add(FtraceRevision(document_id=row.id,version=expected+1,status=status,payload_json=row.payload_json))
    db.session.commit()

def import_receipts(content,filename,report_is_received=False):
    digest=hashlib.sha256(content).hexdigest()
    existing=FtraceImport.query.filter_by(digest=digest).first()
    if existing:
        return existing,False
    lots,summary=parse_receipts(content)
    if report_is_received:
        for lot in lots:
            if not lot['batch'] and not lot['received_date'] and len(lot['report_date'])==10:
                lot['received_date']=lot['report_date']
                lot['actual_received_date']=lot['report_date']
                lot['date_basis']='使用者確認本檔申報日期即實際收貨日期'
                lot['issues']=[s for s in lot['issues'] if s!='缺實際收貨日期／批號']
        summary['pending']=sum(bool(lot['issues']) for lot in lots)
        summary['date_basis']='使用者確認本檔日申報日期即實際收貨日期'
    row=FtraceImport(digest=digest,filename=filename[:255],original_file=content,summary_json=encoded(summary))
    db.session.add(row)
    try:
        db.session.flush()
        for lot in lots:
            db.session.add(FtraceReceipt(import_id=row.id,source_row=lot['source_row'],report_date=lot['report_date'],
                material_code=lot['material_code'],name=lot['name'][:255],payload_json=encoded(lot)))
        db.session.commit()
    except IntegrityError:
        db.session.rollback()
        existing=FtraceImport.query.filter_by(digest=digest).first()
        if existing:
            return existing,False
        raise
    return row,True

def receipt_catalog(day):
    choices=[]
    records=FtraceReceipt.query.filter(FtraceReceipt.report_date<=day.isoformat(),FtraceReceipt.disabled.is_(False)).order_by(FtraceReceipt.report_date.desc(),FtraceReceipt.id).all()
    seen={}
    for record in records:
        lot=json.loads(record.payload_json)
        lot.update(storage_id=record.id,origin='平台匯入',review_note=record.review_note)
        lot['ready']=not lot.get('issues') and bool(lot.get('batch') or lot.get('received_date'))
        lot['source_ids']=[record.id]
        identity=tuple(lot.get(k,'') for k in ('report_date','seq','material_code','spec_code','batch','received_date'))
        # Repeated exports never double the apparent available weight. Conflicting values stay pending.
        if identity in seen:
            previous=seen[identity]
            previous['source_ids'].append(record.id)
            if previous['receipt_kg']!=lot['receipt_kg'] or previous.get('issues') or lot.get('issues'):
                previous['ready']=False
                previous['issues']=list(dict.fromkeys(previous.get('issues',[])+lot.get('issues',[])+['同識別紀錄重複或重量不同，需核對']))
            continue
        seen[identity]=lot;choices.append(lot)
    for doc in FtraceDocument.query.filter(FtraceDocument.kind=='receiving',FtraceDocument.day<=day).order_by(FtraceDocument.day.desc()).all():
        payload=json.loads(doc.payload_json)
        for r in payload.get('receipts',[]):
            if not r.get('include') or not r.get('material_code'):
                continue
            lot=dict(report_date=doc.day.isoformat(),seq=str(r.get('seq','')),material_code=r['material_code'],
                spec_code=r.get('spec_code',''),batch=r.get('batch','') if payload.get('batch_mode') else '',
                received_date='' if payload.get('batch_mode') else r.get('received_date',''),
                receipt_kg=str(r.get('kg','')),kg='',name=r.get('name',''),origin='共用收貨紀錄',
                material_key=str(r.get('material_key') or ''),document_id=doc.id,document_status=doc.status)
            lot['ready']=bool(lot['seq'] and lot['spec_code'] and (lot['batch'] or lot['received_date']))
            try:
                lot['ready']=lot['ready'] and float(lot['receipt_kg'])>0
            except (ValueError,TypeError):
                lot['ready']=False
            identity=tuple(lot.get(k,'') for k in ('report_date','seq','material_code','spec_code','batch','received_date'))
            if identity not in seen:
                seen[identity]=lot;choices.append(lot)
            else:
                previous=seen[identity]
                # Two sources with an identical batch identity but different
                # weights cannot be treated as one verified available lot.
                from decimal import Decimal, InvalidOperation
                try:
                    disagreement=Decimal(str(previous['receipt_kg']))!=Decimal(str(lot['receipt_kg']))
                except (InvalidOperation, ValueError, TypeError):
                    disagreement=True
                if disagreement:
                    previous['ready']=False
                    previous['issues']=list(dict.fromkeys(previous.get('issues',[])+['同一收貨識別的來源重量不同，請先核對']))
    return choices


def validate_daily_manufacture(day, payload):
    """Allow known receipts on/before manufacturing day, never a future lot."""
    if not (payload.get('daily_receiving_only') and payload.get('batch_mode')):
        return
    from decimal import Decimal, InvalidOperation
    catalog={}
    for lot in receipt_catalog(day):
        if not lot.get('report_date') or lot.get('report_date')>day.isoformat() or lot.get('ready') is not True:
            continue
        key=(str(lot.get('material_code') or ''),) + tuple(str(lot.get(field) or '') for field in
              ('report_date','seq','spec_code','batch','received_date'))
        catalog[key]=lot
    if not catalog:
        raise ExportError('製造日前尚無可使用的收貨批次，請先完成收貨或匯入歷史紀錄。')
    for meal in payload.get('meals',[]):
        if not meal.get('include'):
            continue
        for material in meal.get('materials',[]):
            if material.get('include') is False:
                continue
            code=str(material.get('code') or '')
            if not material.get('lots'):
                raise ExportError(f'{material.get("name","原料")}：尚未對應已儲存的收貨批次。')
            for lot in material['lots']:
                if not lot.get('report_date') or lot.get('report_date')>day.isoformat():
                    raise ExportError(f'{material.get("name","原料")}：收貨申報日期不可晚於 {day.isoformat()} 製造日期。')
                key=(code,) + tuple(str(lot.get(field) or '') for field in
                    ('report_date','seq','spec_code','batch','received_date'))
                source=catalog.get(key)
                if not source:
                    raise ExportError(f'{material.get("name","原料")}：批號不在製造日前的可用收貨紀錄中，請重新對應。')
                try:
                    actual=Decimal(str(source.get('receipt_kg') or '0'))
                    used=Decimal(str(lot.get('kg') or '0'))
                except (InvalidOperation, ValueError):
                    raise ExportError('原料收貨／製造重量無效。') from None
                if actual<=0 or used<=0 or used>actual:
                    raise ExportError(f'{material.get("name","原料")}：使用重量超過該批次收貨量或數量無效。')
                # Always use the saved source amount to protect the cross-school
                # capacity check in build_batch_sheets from altered browser data.
                lot['receipt_kg']=str(source['receipt_kg'])

def validate_stored_links(payload,kind):
    if kind!='manufacture' or not payload.get('batch_mode'):
        return
    if not any(lot.get('storage_id') for meal in payload.get('meals',[]) if meal.get('include')
               for material in meal.get('materials',[]) if material.get('include') is not False
               for lot in material.get('lots',[])):
        return
    try:
        day=date.fromisoformat(payload.get('date',''))
    except ValueError:
        raise ExportError('歷史收貨連結缺少製造申報日期。') from None
    available={}
    for candidate in receipt_catalog(day):
        for source_id in candidate.get('source_ids',[]):
            available[source_id]=candidate
    for meal in payload.get('meals',[]):
        if not meal.get('include'):
            continue
        for material in meal.get('materials',[]):
            if material.get('include') is False:
                continue
            for lot in material.get('lots',[]):
                if not lot.get('storage_id'):
                    continue
                record=db.session.get(FtraceReceipt,lot['storage_id'])
                if not record or record.disabled:
                    raise ExportError('使用的歷史收貨紀錄不存在或已停用，請重新選取。')
                source=json.loads(record.payload_json)
                if source.get('issues') or not available.get(record.id,{}).get('ready'):
                    raise ExportError(f'{material.get("name", "原料")}：歷史收貨尚待核對，請先到共用資料區補齊。')
                if source['material_code']!=material.get('code') or any(str(source.get(k,'') or '')!=str(lot.get(k,'') or '') for k in ('report_date','seq','spec_code','batch','received_date')):
                    raise ExportError('使用的歷史收貨欄位已變更，請重新選取批次。')
                lot['receipt_kg']=source['receipt_kg']
