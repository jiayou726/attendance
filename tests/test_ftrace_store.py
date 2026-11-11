from datetime import date
from io import BytesIO
import json
import pytest
import xlwt
from tests.test_ftrace_export import app, DAY
from ftrace_models import FtraceDocument, FtraceRevision, FtraceReceipt, FtraceImport
from services.ftrace_import import parse_receipts
from services.ftrace_store import receipt_catalog


def workbook(*,batch=False,duplicates=False,with_date=False):
    book=xlwt.Workbook()
    headers=['申報方式(必填)','申報日期(必填)','申報日期流水號','原材料自編碼(必填)',
        '原材料名稱','原材料廠商自編碼','收貨總淨重']
    row=['2','2026/10/05','0','TEST-RAW','測試原料','TEST-SPEC','20']
    sheets={'收貨資料':[headers,row]+([row.copy()] if duplicates else []),
        '收貨資料供應商':[['申報方式','申報日期','申報日期流水號','原材料自編碼','原材料廠商自編碼','供應商自編碼','供應商名稱'],
            ['2','2026/10/05','0','TEST-RAW','TEST-SPEC','TEST-V','測試供應商']]}
    if with_date:
        sheets['收貨資料收貨日期']=[['申報方式','申報日期','申報日期流水號','原材料自編碼','原材料廠商自編碼','收貨日期'],
            ['2','2026/10/05','0','TEST-RAW','TEST-SPEC','2026/10/04']]
    if batch:
        sheets['收貨資料批號資訊']=[['申報方式','申報日期','申報日期流水號','原材料自編碼','原材料廠商自編碼','批號','收貨日期','收貨淨重(必填)'],
            ['2','2026/10/05','0','TEST-RAW','TEST-SPEC','TEST-LOT','2026/10/04','20']]
    for name,rows in sheets.items():
        sheet=book.add_sheet(name)
        for i,row in enumerate(rows):
            for j,v in enumerate(row):sheet.write(i,j,v)
    output=BytesIO();book.save(output);return output.getvalue()


def empty_payload(day=DAY):
    return dict(date=day.isoformat(),batch_mode=True,receipts=[],meals=[])


def save(client,payload,version=0,kind='receiving',status='draft'):
    return client.post('/admin/order-tool/ftrace/shared',data=dict(date=payload['date'],kind=kind,
        payload=json.dumps(payload),version=version,status=status))


def test_parser_joins_by_headers_preserves_zero_and_flags_missing_actual_date():
    lots,summary=parse_receipts(workbook(duplicates=True))
    assert len(lots)==2 and summary['duplicate_groups']==1
    assert lots[0]['seq']=='0' and lots[0]['report_date']=='2026-10-05'
    assert lots[0]['received_date']==lots[0]['batch']==''
    assert lots[0]['suppliers']==[dict(code='TEST-V',name='測試供應商')]
    assert '缺實際收貨日期／批號' in lots[0]['issues']
    assert lots[0]['source_row']==2 and lots[1]['source_row']==3


def test_parser_supports_legacy_dates_and_batch_exclusive_link():
    legacy,_=parse_receipts(workbook(with_date=True))
    assert legacy[0]['received_date']=='2026-10-04' and not legacy[0]['issues']
    batch,_=parse_receipts(workbook(batch=True))
    assert batch[0]['batch']=='TEST-LOT' and batch[0]['received_date']==''
    assert batch[0]['actual_received_date']=='2026-10-04'
    assert not batch[0]['issues']


def test_import_is_durable_idempotent_and_original_file_downloadable(app):
    client=app.test_client();data=workbook()
    for _ in range(2):
        response=client.post('/admin/order-tool/ftrace/records',data={'file':(BytesIO(data),'receipt.xls')})
        assert response.status_code==302
    with app.app_context():
        assert FtraceReceipt.query.count()==1 and FtraceImport.query.count()==1
        assert FtraceDocument.query.count()==0
        row=FtraceImport.query.one();import_id=row.id
    assert client.get(f'/admin/order-tool/ftrace/imports/{import_id}/original').data==data
    independent=app.test_client()
    response=independent.get('/admin/order-tool/ftrace/records')
    assert response.status_code==200 and '測試原料'.encode() in response.data
    lots=independent.get('/admin/order-tool/ftrace/receipts?date=2026-10-06').json['receipts']
    assert len(lots)==1 and lots[0]['ready'] is False
    assert independent.get('/admin/order-tool/ftrace/receipts?date=2026-10-04').json['receipts']==[]


def test_shared_documents_cross_browser_cas_and_revision_history(app):
    a,b=app.test_client(),app.test_client();payload=empty_payload()
    assert save(a,payload).json['version']==1
    assert b.get(f'/admin/order-tool/ftrace/shared?date={DAY}&kind=receiving').json['payload']==payload
    assert save(b,payload,version=0).status_code==409
    payload['receipts']=[dict(name='手動草稿',key='r1')]
    assert save(a,payload,version=1).json['version']==2
    assert save(b,payload,version=1).status_code==409
    with app.app_context():
        assert FtraceDocument.query.count()==1
        assert FtraceRevision.query.count()==2
        revisions=FtraceRevision.query.order_by(FtraceRevision.version).all()
        assert json.loads(revisions[0].payload_json)['receipts']==[]
        assert json.loads(revisions[1].payload_json)['receipts'][0]['key']=='r1'


def test_wrong_date_or_kind_rejected_and_csrf_enforced(app):
    client=app.test_client()
    assert client.post('/admin/order-tool/ftrace/shared',data=dict(date='2026-10-04',kind='receiving',payload=json.dumps(empty_payload()),version=0)).status_code==400
    assert save(client,empty_payload(),kind='INVALID').status_code==400
    app.config['KITCHEN_CSRF_ENABLED']=True
    response=save(client,empty_payload())
    assert response.status_code==302
    with app.app_context():assert FtraceDocument.query.count()==0


def test_pending_receipt_correction_preserves_source_and_versions(app):
    client=app.test_client();client.post('/admin/order-tool/ftrace/records',data={'file':(BytesIO(workbook()),'receipt.xls')})
    with app.app_context():rid=FtraceReceipt.query.one().id
    response=client.post(f'/admin/order-tool/ftrace/receipts/{rid}',data=dict(version='1',received_date='2026-10-04',note='驗收單核對'))
    assert response.status_code==302
    lot=client.get('/admin/order-tool/ftrace/receipts?date=2026-10-06').json['receipts'][0]
    assert lot['ready'] and lot['received_date']=='2026-10-04' and lot['seq']=='0'
    with app.app_context():
        row=FtraceReceipt.query.one();payload=json.loads(row.payload_json)
        assert row.version==2 and payload['source_raw']['申報日期']=='2026/10/05'
        assert payload['corrections'][0]['before']['received_date']==''
    assert client.post(f'/admin/order-tool/ftrace/receipts/{rid}',data=dict(version='1',note='stale')).status_code==200


def test_saved_receipt_used_cross_browser_and_manufacturing_shared_to_delivery(app):
    client=app.test_client();payload=empty_payload()
    payload['receipts']=[dict(include=True,material_code='TEST-RAW',spec_code='TEST-SPEC',seq='01',batch='TEST-LOT',kg='20',name='測試原料')]
    assert save(client,payload).status_code==200
    lots=app.test_client().get(f'/admin/order-tool/ftrace/receipts?date={DAY}').json['receipts']
    assert lots[0]['ready'] and lots[0]['receipt_kg']=='20'
    made=empty_payload();made['meals']=[dict(key='school',include=True,product_code='TEST-P',batch='TEST-P-LOT',kg='5',quantity='10',manufactured_date=str(DAY),expiry_date=str(DAY))]
    assert save(client,made,kind='manufacture').status_code==200
    response=app.test_client().get(f'/admin/order-tool/ftrace?date={DAY}&kind=delivery')
    assert response.status_code==200 and b'TEST-P-LOT' in response.data


def test_repeated_exports_do_not_double_receipt_capacity(app):
    from services.ftrace_store import import_receipts
    from extensions import db
    with app.app_context():
        import_receipts(workbook(batch=True),'one.xls')
        original=FtraceReceipt.query.one()
        duplicate=FtraceReceipt(import_id=original.import_id,source_row=99,report_date=original.report_date,material_code=original.material_code,name=original.name,payload_json=original.payload_json)
        db.session.add(duplicate);db.session.commit()
        lots=receipt_catalog(date(2026,10,6))
        assert len(lots)==1 and lots[0]['receipt_kg']=='20' and len(lots[0]['source_ids'])==2


def test_confirmed_report_date_becomes_actual_date_without_invented_batch(app):
    client=app.test_client()
    client.post('/admin/order-tool/ftrace/records',data={'file':(BytesIO(workbook()),'receipt.xls'),'report_is_received':'1'})
    lots=client.get('/admin/order-tool/ftrace/receipts?date=2026-10-06').json['receipts']
    assert lots[0]['received_date']=='2026-10-05' and lots[0]['batch']==''
    assert lots[0]['ready'] and '使用者確認' in lots[0]['date_basis']
    with app.app_context():
        summary=json.loads(FtraceImport.query.one().summary_json)
        assert summary['pending']==0


def test_confirmed_date_keeps_conflicting_duplicates_pending(app):
    client=app.test_client()
    client.post('/admin/order-tool/ftrace/records',data={'file':(BytesIO(workbook(duplicates=True)),'receipt.xls'),'report_is_received':'1'})
    lots=client.get('/admin/order-tool/ftrace/receipts?date=2026-10-06').json['receipts']
    assert len(lots)==1 and not lots[0]['ready']
    assert any('重複' in i or '多列' in i for i in lots[0]['issues'])


def test_ftrace_practice_records_and_files_never_leak_to_formal_or_another_account(tmp_path):
    from tests.test_practice_database import _app, _login
    application=_app(tmp_path)
    learner=application.test_client();_login(learner,'資料測試甲')
    prefix='/admin/order-tool/practice-workspace'
    response=learner.post(prefix+'/ftrace/shared',data=dict(date=str(DAY),kind='receiving',payload=json.dumps(empty_payload()),version=0))
    assert response.status_code==200
    learner.post(prefix+'/ftrace/records',data={'file':(BytesIO(workbook()),'private.xls'),'report_is_received':'1'})
    other=application.test_client();_login(other,'資料測試乙')
    assert other.get(prefix+f'/ftrace/shared?date={DAY}&kind=receiving').json['version']==0
    assert other.get(prefix+'/ftrace/receipts?date=2026-10-06').json['receipts']==[]
    assert learner.get(f'/admin/order-tool/ftrace/shared?date={DAY}&kind=receiving').json['version']==0
    assert learner.get('/admin/order-tool/ftrace/receipts?date=2026-10-06').json['receipts']==[]
    with application.app_context():
        assert FtraceDocument.query.count()==0 and FtraceImport.query.count()==0


def test_pending_source_cannot_be_forged_by_editing_the_manufacturing_payload(app):
    from services.ftrace_store import validate_stored_links
    from services.ftrace import ExportError
    client=app.test_client()
    client.post('/admin/order-tool/ftrace/records',data={'file':(BytesIO(workbook()),'receipt.xls')})
    with app.app_context():
        source=FtraceReceipt.query.one()
        payload=dict(date='2026-10-06',batch_mode=True,meals=[dict(include=True,materials=[dict(name='測試原料',code='TEST-RAW',lots=[dict(storage_id=source.id,received_date='2026-10-05')])])])
        with pytest.raises(ExportError,match='尚待核對'):
            validate_stored_links(payload,'manufacture')


def test_conflicting_cross_file_capacity_is_blocked_even_if_first_source_is_complete(app):
    from services.ftrace_store import import_receipts, validate_stored_links
    from services.ftrace import ExportError
    from extensions import db
    with app.app_context():
        import_receipts(workbook(batch=True),'one.xls')
        first=FtraceReceipt.query.one();lot=json.loads(first.payload_json)
        changed=dict(lot,receipt_kg='999')
        db.session.add(FtraceReceipt(import_id=first.import_id,source_row=99,report_date=first.report_date,material_code=first.material_code,name=first.name,payload_json=json.dumps(changed)))
        db.session.commit()
        payload=dict(date='2026-10-06',batch_mode=True,meals=[dict(include=True,materials=[dict(name='測試原料',code='TEST-RAW',lots=[dict(lot,storage_id=first.id)])])])
        with pytest.raises(ExportError,match='尚待核對'):
            validate_stored_links(payload,'manufacture')


def test_daily_manufacturing_uses_matching_menu_material_and_saved_same_day_receipt(app):
    from tests.test_ftrace_export import batch_payload
    client=app.test_client()
    receive=empty_payload(day=DAY)
    receive['receipts']=[dict(
        key='menu:1',include=True,material_code='B001',spec_code='B001A025',
        supplier_code='A025',seq='91',batch='TEST-RAW-LOT',
        received_date=DAY.isoformat(),kg='1.235',name='高麗菜',
    )]
    assert save(client,receive,kind='receiving').status_code==200
    made=batch_payload()
    made['date']=DAY.isoformat()
    made['daily_receiving_only']=True
    lot=made['meals'][0]['materials'][0]['lots'][0]
    lot.update(report_date=DAY.isoformat(),seq='91',
               spec_code='B001A025',batch='TEST-RAW-LOT',received_date='',
               kg='0.96')
    def export():
        return client.post('/admin/order-tool/ftrace/export',data=dict(
            date=DAY.isoformat(),kind='manufacture',payload=json.dumps(made)))
    response=export()
    assert response.status_code==200
    assert 'ms-excel' in response.content_type

    lot['report_date']='2026-11-04'
    response=export()
    assert response.status_code==400
    assert '當日收貨批次' in response.get_json()['error']

    lot['report_date']=DAY.isoformat()
    lot['batch']='NONEXISTENT'
    response=export()
    assert response.status_code==400
    assert '不在當日可選用' in response.get_json()['error']


def test_duplicate_shared_and_imported_identity_with_different_weight_not_auto_usable(app):
    from services.ftrace_store import import_receipts
    with app.app_context():
        import_receipts(workbook(batch=True),'original.xls')
    client=app.test_client()
    doc=empty_payload(day=date(2026,10,5))
    doc['receipts']=[dict(include=True,material_code='TEST-RAW',
       name='測試原料',seq='0',spec_code='TEST-SPEC',batch='TEST-LOT',kg='30')]
    assert save(client,doc).status_code==200
    choices=client.get('/admin/order-tool/ftrace/receipts?date=2026-10-05').json['receipts']
    matched=[item for item in choices if item['material_code']=='TEST-RAW']
    assert len(matched)==1
    assert matched[0]['ready'] is False
    assert any('重量不同' in issue for issue in matched[0]['issues'])
