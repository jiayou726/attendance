"""Additive routes on the existing, CSRF-protected kitchen blueprint."""
from datetime import date, datetime
import json
from zoneinfo import ZoneInfo

from flask import jsonify, render_template, request, send_file, redirect, url_for, flash, abort

from blueprints.order_tool import order_bp
from services.ftrace import KINDS, SCHEMAS, ExportError, build_sheets, draft, export_workbook
from services.ftrace_references import reference_data
from services import ftrace_store as store
from ftrace_models import FtraceDocument, FtraceRevision, FtraceImport, FtraceReceipt
from extensions import db
from io import BytesIO
from werkzeug.utils import secure_filename


@order_bp.get("/ftrace")
def ftrace():
    raw_date = request.args.get("date")
    try:
        day = date.fromisoformat(raw_date) if raw_date else datetime.now(ZoneInfo("Asia/Taipei")).date()
    except ValueError:
        return "日期格式錯誤，請使用 YYYY-MM-DD。", 400
    kind = request.args.get("kind", "receiving")
    if kind not in SCHEMAS:
        return "資料類型錯誤。", 400
    cloud=store.snapshot(store.document(day,kind))
    made=store.snapshot(store.document(day,'manufacture'))
    response = render_template("kitchen/ftrace.html", data=draft(day), kind=kind,
                               kinds=KINDS, references=reference_data(), cloud=cloud,
                               manufacture=made, receipts=store.receipt_catalog(day))
    return response, 200, {"Cache-Control": "no-store"}


@order_bp.post("/ftrace/export")
def ftrace_export():
    try:
        if len(request.form.get("payload", "")) > 2_000_000:
            raise ExportError("匯出資料過大。")
        payload = json.loads(request.form.get("payload", ""))
        if not isinstance(payload, dict):
            raise ExportError("匯出資料格式錯誤。")
        kind = request.form.get("kind")
        day = date.fromisoformat(request.form.get("date", ""))
        store.validate_stored_links(payload,kind)
        if kind=='delivery' and payload.get('batch_mode'):
            made=store.document(day,'manufacture')
            if made:
                sources=json.loads(made.payload_json).get('meals',[])
                from decimal import Decimal, InvalidOperation
                for row in payload.get('meals',[]):
                    if not row.get('include'):
                        continue
                    source=next((r for r in sources if r.get('key')==row.get('key') and r.get('include')),None)
                    if source:
                        if any(row.get(k)!=source.get(k) for k in ('product_code','batch','manufactured_date','expiry_date')):
                            raise ExportError('交貨與共用製造批號／日期不同，請重新帶入共用製造紀錄。')
                        try:
                            if any(Decimal(str(row.get(k) or 0))>Decimal(str(source.get(k) or 0)) for k in ('quantity','kg')):
                                raise ExportError('交貨數量／重量大於共用製造紀錄。')
                        except InvalidOperation:
                            raise ExportError('製造或交貨數量無效。') from None
        sheets = build_sheets(kind, day, payload)
        stream = export_workbook(kind, sheets, batch_mode=payload.get("batch_mode") is True)
    except (ValueError, TypeError, KeyError) as error:
        message = str(error) if isinstance(error, ExportError) else "日期或資料格式錯誤，請重新檢查。"
        return jsonify(error=message), 400
    label = dict(KINDS)[kind]
    response = send_file(stream, as_attachment=True,
        download_name=f"非追不可_{label}_{day.isoformat()}.xls",
        mimetype="application/vnd.ms-excel")
    response.headers["Cache-Control"] = "no-store"
    return response


@order_bp.get('/ftrace/shared')
def ftrace_shared():
    try:
        day=date.fromisoformat(request.args.get('date',''))
        kind=request.args.get('kind','')
        if kind not in dict(KINDS):
            raise ValueError
    except ValueError:
        return jsonify(error='日期或資料類型無效。'),400
    return jsonify(**store.snapshot(store.document(day,kind))),200,{'Cache-Control':'no-store'}


@order_bp.post('/ftrace/shared')
def ftrace_shared_save():
    try:
        day=date.fromisoformat(request.form.get('date',''))
        result=store.save_document(day,request.form.get('kind'),json.loads(request.form.get('payload','')),
            request.form.get('version'),request.form.get('status','draft'))
        return jsonify(**result)
    except (ValueError,TypeError,KeyError) as error:
        db.session.rollback()
        return jsonify(error=str(error) if isinstance(error,ExportError) else '儲存格式錯誤。'),409 if isinstance(error,store.Conflict) else 400


@order_bp.get('/ftrace/receipts')
def ftrace_receipts():
    try:
        day=date.fromisoformat(request.args.get('date',''))
    except ValueError:
        return jsonify(error='日期格式錯誤。'),400
    return jsonify(receipts=store.receipt_catalog(day)),200,{'Cache-Control':'no-store'}


@order_bp.route('/ftrace/records',methods=['GET','POST'])
def ftrace_records():
    if request.method=='POST':
        upload=request.files.get('file')
        try:
            if not upload or not (upload.filename or '').lower().endswith('.xls'):
                raise ExportError('請選擇非追不可匯出的原始 .xls 收貨檔。')
            from services.ftrace_import import MAX_BYTES
            row,created=store.import_receipts(upload.read(MAX_BYTES+1),upload.filename.rsplit('/',1)[-1],report_is_received=request.form.get('report_is_received')=='1')
            summary=json.loads(row.summary_json)
            flash(('匯入完成' if created else '相同檔案已存在，未重複匯入')+f'：{summary["records"]} 筆，{summary["pending"]} 筆待核對。','success')
        except (ValueError,TypeError) as error:
            db.session.rollback();flash(str(error) if isinstance(error,ExportError) else '檔案格式錯誤。','error')
        return redirect(url_for('order_tool.ftrace_records'))
    imports=FtraceImport.query.order_by(FtraceImport.id.desc()).all()
    query=FtraceReceipt.query
    search=request.args.get('q','').strip()[:100]
    month=request.args.get('month','').strip()
    if search:
        from sqlalchemy import or_
        query=query.filter(or_(FtraceReceipt.name.contains(search,autoescape=True),FtraceReceipt.material_code.contains(search,autoescape=True)))
    if month:
        try:
            date.fromisoformat(month+'-01')
            query=query.filter(FtraceReceipt.report_date.startswith(month))
        except ValueError:
            flash('月份格式請使用 YYYY-MM。','error')
    page=max(1,request.args.get('page',1,type=int) or 1)
    pagination=query.order_by(FtraceReceipt.report_date.desc(),FtraceReceipt.id).paginate(page=page,per_page=50,error_out=False)
    docs_query=FtraceDocument.query
    if month and len(month)==7:
        try:
            start=date.fromisoformat(month+'-01')
            end=date(start.year+1,1,1) if start.month==12 else date(start.year,start.month+1,1)
            docs_query=docs_query.filter(FtraceDocument.day>=start,FtraceDocument.day<end)
        except ValueError:
            pass
    docs_page=max(1,request.args.get('docs_page',1,type=int) or 1)
    docs_pagination=docs_query.order_by(FtraceDocument.day.desc(),FtraceDocument.kind).paginate(page=docs_page,per_page=50,error_out=False)
    docs=docs_pagination.items
    return render_template('kitchen/ftrace_records.html',imports=[(r,json.loads(r.summary_json)) for r in imports],
        receipts=[(r,json.loads(r.payload_json)) for r in pagination.items],pagination=pagination,
        documents=docs,docs_pagination=docs_pagination,kinds=dict(KINDS),q=search,month=month),200,{'Cache-Control':'no-store'}


@order_bp.get('/ftrace/imports/<int:import_id>/original')
def ftrace_import_original(import_id):
    row=db.session.get(FtraceImport,import_id)
    if not row:
        abort(404)
    return send_file(BytesIO(row.original_file),as_attachment=True,download_name=row.filename,mimetype='application/vnd.ms-excel')


@order_bp.get('/ftrace/documents/<int:document_id>')
def ftrace_document_history(document_id):
    row=db.session.get(FtraceDocument,document_id)
    if not row:
        abort(404)
    revisions=FtraceRevision.query.filter_by(document_id=row.id).order_by(FtraceRevision.version.desc()).all()
    return render_template('kitchen/ftrace_history.html',record=row,kinds=dict(KINDS),
        revisions=[(r,json.loads(r.payload_json)) for r in revisions])


@order_bp.post('/ftrace/documents/<int:document_id>/status')
def ftrace_document_status(document_id):
    row=db.session.get(FtraceDocument,document_id)
    if not row:
        abort(404)
    try:
        status=request.form.get('status')
        if status=='submitted':
            if request.form.get('confirmed')!='1':
                raise ExportError('請確認已在非追不可看到成功申報紀錄。')
            payload=json.loads(row.payload_json)
            store.validate_stored_links(payload,row.kind)
            build_sheets(row.kind,row.day,payload)
        store.set_status(row,int(request.form.get('version','')),status)
        flash('已更新狀態；過去各版本仍完整保留。','success')
    except (ValueError,TypeError) as error:
        db.session.rollback();flash(str(error) if isinstance(error,ExportError) else '版本格式錯誤。','error')
    return redirect(url_for('order_tool.ftrace_document_history',document_id=document_id))


@order_bp.route('/ftrace/receipts/<int:receipt_id>',methods=['GET','POST'])
def ftrace_receipt_edit(receipt_id):
    row=db.session.get(FtraceReceipt,receipt_id)
    if not row:
        abort(404)
    lot=json.loads(row.payload_json)
    if request.method=='POST':
        try:
            if int(request.form.get('version',''))!=row.version:
                raise store.Conflict('資料已更新，請重新整理再核對。')
            note=request.form.get('note','').strip()[:1000]
            if not note:
                raise ExportError('請填寫核對依據或停用原因。')
            batch=request.form.get('batch','').strip()
            received=request.form.get('received_date','').strip()
            if len(batch)>40 or batch and received:
                raise ExportError('批號與舊制實際收貨日期擇一填寫，批號最多 40 字。')
            if received:
                date.fromisoformat(received)
            original=dict(batch=lot.get('batch',''),received_date=lot.get('received_date',''),issues=lot.get('issues',[]),disabled=row.disabled)
            lot['batch']=batch;lot['received_date']=received
            issues=[s for s in lot.get('issues',[]) if s!='缺實際收貨日期／批號']
            if request.form.get('reviewed')=='1':
                issues=[s for s in issues if '同收貨識別' not in s and '多個收貨日期' not in s]
            if not batch and not received:
                issues.append('缺實際收貨日期／批號')
            lot['issues']=issues
            lot.setdefault('corrections',[]).append(dict(before=original,note=note,at=store.now().isoformat()))
            from sqlalchemy import update
            result=db.session.execute(update(FtraceReceipt).where(FtraceReceipt.id==row.id,FtraceReceipt.version==row.version)
                .values(payload_json=store.encoded(lot),review_note=note,disabled=request.form.get('disabled')=='1',version=row.version+1))
            if result.rowcount!=1:
                raise store.Conflict('資料版本已更新，請重新整理。')
            db.session.commit();flash('已儲存核對結果，原始檔與修改紀錄保留。','success')
            return redirect(url_for('order_tool.ftrace_records',q=row.material_code))
        except (ValueError,TypeError) as error:
            db.session.rollback();flash(str(error) if isinstance(error,ExportError) else '日期或版本格式錯誤。','error')
    return render_template('kitchen/ftrace_receipt_edit.html',record=row,lot=lot)
