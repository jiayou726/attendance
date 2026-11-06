"""Read official receipt exports by column headers; preserve every source row."""
from collections import Counter, defaultdict
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
import re
import xlrd
from services.ftrace import ExportError

MAX_BYTES = 10 * 1024 * 1024

def value(cell, datemode):
    if cell.ctype == xlrd.XL_CELL_DATE:
        return xlrd.xldate_as_datetime(cell.value, datemode).date().isoformat()
    if cell.ctype == xlrd.XL_CELL_NUMBER:
        return format(Decimal(str(cell.value)), 'f').rstrip('0').rstrip('.') if cell.value % 1 else str(int(cell.value))
    if cell.ctype == xlrd.XL_CELL_ERROR:
        raise ExportError('檔案含有 Excel 錯誤儲存格，請確認原始匯出檔。')
    return str(cell.value).strip()

def label(s):
    return re.sub(r'[\s（(].*$', '', s)

def iso(s, optional=False):
    if not s and optional:
        return ''
    s = s.replace('/', '-')
    if re.fullmatch(r'\d{4}-\d{2}', s):
        date.fromisoformat(s + '-01'); return s
    try:
        return date.fromisoformat(s).isoformat()
    except ValueError:
        raise ExportError(f'日期格式無效：{s[:30]}') from None

def read_rows(book, name):
    if name not in book.sheet_names():
        return []
    sheet = book.sheet_by_name(name)
    if sheet.nrows > 20001 or sheet.ncols > 80:
        raise ExportError('檔案資料過多，請分月份匯出。')
    if not sheet.nrows:
        return []
    headers = [label(value(c, book.datemode)) for c in sheet.row(0)]
    result = []
    for i in range(1, sheet.nrows):
        vals = [value(c, book.datemode) for c in sheet.row(i)]
        if any(vals):
            result.append((i+1, dict(zip(headers, vals))))
    return result

def identity(row):
    return (row.get('申報方式',''), iso(row.get('申報日期','')), row.get('申報日期流水號',''),
            row.get('原材料自編碼',''), row.get('原材料廠商自編碼',''))

def parse_receipts(content):
    if not content or len(content) > MAX_BYTES:
        raise ExportError('請提供 10 MB 以內的原始 XLS 收貨匯出檔。')
    try:
        book = xlrd.open_workbook(file_contents=content, on_demand=True)
    except Exception:
        raise ExportError('無法讀取 XLS，請上傳非追不可資料匯出的原始 .xls 檔。') from None
    try:
        base = read_rows(book, '收貨資料')
        if not base:
            raise ExportError('找不到「收貨資料」或沒有收貨紀錄。')
        suppliers, dates, batches = defaultdict(list), defaultdict(list), defaultdict(list)
        for _, r in read_rows(book, '收貨資料供應商'):
            suppliers[identity(r)].append({'code':r.get('供應商自編碼',''), 'name':r.get('供應商名稱','')})
        for _, r in read_rows(book, '收貨資料收貨日期'):
            dates[identity(r)].append(iso(r.get('收貨日期','')))
        for _, r in read_rows(book, '收貨資料批號資訊'):
            batches[identity(r)].append(r)
        counts = Counter(identity(r) for _, r in base)
        result = []
        for source_row, raw in base:
            mode, report, seq, code, spec = key = identity(raw)
            if mode not in ('1','2') or not code or not spec or len(code)>40 or len(spec)>40:
                raise ExportError(f'收貨資料第 {source_row} 列：缺少有效申報方式、原材料碼或廠商自編碼。')
            if (mode=='1') != (len(report)==7):
                raise ExportError(f'收貨資料第 {source_row} 列：申報方式與年月／日期不一致。')
            children = batches[key] or [None]
            for child in children:
                kg = (child.get('收貨淨重') or child.get('收貨總淨重')) if child else raw.get('收貨總淨重','')
                try:
                    weight=Decimal(kg or '0')
                    if not weight.is_finite() or weight<0:
                        raise InvalidOperation
                except InvalidOperation:
                    raise ExportError(f'收貨資料第 {source_row} 列：重量無效。') from None
                received = iso(child.get('收貨日期',''), True) if child else dates[key][0] if len(set(dates[key]))==1 else ''
                lot = dict(report_date=report, seq=seq, material_code=code, spec_code=spec,
                    name=raw.get('原材料名稱',''), receipt_kg=str(weight), kg='',
                    batch=child.get('批號','') if child else '', received_date='' if child and child.get('批號') else received,
                    actual_received_date=received,
                    suppliers=suppliers[key], source_row=source_row, source_raw=raw,
                    date_candidates=dates[key], source_batch=child)
                issues=[]
                if not lot['batch'] and not received:
                    issues.append('缺實際收貨日期／批號')
                if counts[key]>1:
                    issues.append('同收貨識別有多列，需核對重複或分批')
                if len(set(dates[key]))>1:
                    issues.append('有多個收貨日期，需核對分批重量')
                if not weight:
                    issues.append('收貨重量為零')
                lot['issues']=issues
                result.append(lot)
        summary=dict(source_rows=len(base), records=len(result),
            start=min(x['report_date'] for x in result), end=max(x['report_date'] for x in result),
            pending=sum(bool(x['issues']) for x in result), duplicate_groups=sum(n>1 for n in counts.values()),
            total_kg=str(sum(Decimal(x['receipt_kg']) for x in result)))
        return result, summary
    finally:
        book.release_resources()
