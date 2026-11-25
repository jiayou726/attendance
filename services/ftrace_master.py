"""Parse and retain private FDA master exports; never modify procurement or historic declarations."""
import hashlib
import json
import xlrd
from sqlalchemy.exc import IntegrityError
from extensions import db
from ftrace_models import FtraceMasterImport
from services.ftrace import ExportError
from services.ftrace_import import MAX_BYTES, read_rows

def parse_master(content):
    if not content or len(content)>MAX_BYTES:
        raise ExportError('請提供 10 MB 以內的原始 XLS 主檔。')
    try:
        book=xlrd.open_workbook(file_contents=content,on_demand=True)
    except Exception:
        raise ExportError('無法讀取原始 XLS 主檔。') from None
    try:
        sheets=set(book.sheet_names())
        if '原材料' not in sheets:
            raise ExportError('缺少原材料工作表。')
        food='原材料廠商資料' in sheets
        packaging='原材料材質' in sheets
        if food==packaging:
            raise ExportError('請上傳原料或物料材料主檔，不是收貨或其他資料。')
        kind='food' if food else 'packaging'
        materials={}
        for n,row in read_rows(book,'原材料'):
            code=row.get('原材料自編碼','')
            name=row.get('原材料名稱','')
            if not code or not name or len(code)>40 or code in materials:
                raise ExportError(f'原材料第 {n} 列：編碼缺漏、過長或重複。')
            alias=row.get('原材料別名','')
            materials[code]=dict(code=code,name=name,aliases=[alias] if alias else [],
                specs=[],raw=row)
        if not materials:
            raise ExportError('主檔沒有原材料資料。')
        specs={}
        origins=[]
        if food:
            units={row.get('代碼',''):row.get('名稱','') for _,row in read_rows(book,'★包裝單位代碼')}
            for n,row in read_rows(book,'原材料廠商資料'):
                code=row.get('原材料自編碼','')
                spec=row.get('原材料廠商自編碼','')
                if code not in materials or not spec or len(spec)>40 or (code,spec) in specs:
                    raise ExportError(f'原材料廠商資料第 {n} 列：編碼缺漏、重複或無對應原料。')
                weight=row.get('包裝規格重量','')
                unit={'1':'公克','2':'公斤'}.get(row.get('包裝規格重量單位代碼',''),'')
                package_unit=units.get(row.get('包裝規格單位代碼',''),'')
                if row.get('包裝規格包裝形式')=='2':
                    package='散裝'+('／'+row['包裝規格散裝備註'] if row.get('包裝規格散裝備註') else '')
                else:
                    package=(weight+unit+('／'+package_unit if package_unit else '')) if weight and unit else ''
                # Domestic responsible business is a candidate, not proof of actual supplier.
                entry=dict(code=spec,supplier_name=row.get('國內負責廠商',''),
                    manufacturer_name=row.get('製造廠商',''),package_label=package,
                    master_imported=True,raw=row)
                specs[code,spec]=entry
                materials[code]['specs'].append(entry)
            for n,row in read_rows(book,'原材料原產地'):
                key=(row.get('原材料自編碼',''),row.get('原材料廠商自編碼',''))
                if key not in specs or not row.get('原產地代碼'):
                    raise ExportError(f'原材料原產地第 {n} 列：無對應規格或原產地代碼。')
                origins.append(row)
        else:
            for n,row in read_rows(book,'原材料材質'):
                code=row.get('原材料自編碼','')
                if code not in materials or not row.get('包裝容器材質代碼'):
                    raise ExportError(f'原材料材質第 {n} 列：無對應材料或材質代碼。')
                materials[code].setdefault('composition',[]).append(row)
        payload=dict(materials=list(materials.values()),origins=origins)
        summary=dict(materials=len(materials),specs=len(specs),origins=len(origins))
        return kind,payload,summary
    finally:
        book.release_resources()

def latest_master(kind):
    return FtraceMasterImport.query.filter_by(kind=kind).order_by(FtraceMasterImport.id.desc()).first()

def import_master(content,filename):
    kind,payload,summary=parse_master(content)
    digest=hashlib.sha256(content).hexdigest()
    existing=FtraceMasterImport.query.filter_by(kind=kind,digest=digest).first()
    if existing:
        return existing,False
    row=FtraceMasterImport(kind=kind,digest=digest,filename=filename[:255],
        original_file=content,payload_json=json.dumps(payload,ensure_ascii=False),
        summary_json=json.dumps(summary,ensure_ascii=False))
    db.session.add(row)
    try:
        db.session.commit()
    except IntegrityError:
        db.session.rollback()
        existing=FtraceMasterImport.query.filter_by(kind=kind,digest=digest).first()
        if existing:
            return existing,False
        raise
    return row,True

def merge_materials(old, imported):
    """Keep verified private aliases and old codes; enrich exact code matches only."""
    result={r['code']:dict(r) for r in old}
    for row in imported:
        code=row['code']
        prior=result.get(code,{})
        aliases=list(dict.fromkeys([*(prior.get('aliases') or []),*(row.get('aliases') or []),
                                    *([prior['name']] if prior.get('name') else [])]))
        specs={s['code']:dict(s) for s in prior.get('specs',[])}
        for source in row.get('specs',[]):
            new={k:source[k] for k in ('code','supplier_name','manufacturer_name','package_label','master_imported')}
            existing=specs.get(new['code'])
            if existing:
                # Keep verified legacy supplier names. The official responsible name is not the supplier relationship.
                new={**new,**existing}
                new['master_imported']=bool(existing.get('master_imported',False))
            specs[new['code']]=new
        result[code]=dict(code=code,name=row['name'],aliases=[a for a in aliases if a!=row['name']],
                          specs=list(specs.values()),master_imported=True)
    return list(result.values())
