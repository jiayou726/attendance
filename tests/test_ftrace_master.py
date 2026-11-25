from io import BytesIO
import pytest
import xlwt
from services.ftrace import ExportError
from services.ftrace_master import parse_master, merge_materials
from services.ftrace_references import apply_reference_defaults

def workbook(sheets):
    book=xlwt.Workbook()
    for name,rows in sheets.items():
        sheet=book.add_sheet(name)
        for i,row in enumerate(rows):
            for j,value in enumerate(row):
                sheet.write(i,j,value)
    buf=BytesIO()
    book.save(buf)
    return buf.getvalue()

def food(duplicate=False):
    return workbook({
        '原材料':[['原材料自編碼(必填)','原材料名稱(必填)','原材料別名'],
                    ['X01','肉丁',''],['X01' if duplicate else 'X02','肉丁','']],
        '原材料廠商資料':[['原材料自編碼(必填)','原材料廠商自編碼(必填)','國內負責廠商'],
                          ['X01','X01A','甲廠商'],['X02','X02A','乙廠商']],
        '原材料原產地':[['原材料自編碼(必填)','原材料廠商自編碼(必填)','原產地代碼(必填)'],
                         ['X01','X01A','TW']]
    })

def test_food_master_and_supplier_ambiguity():
    kind,payload,summary=parse_master(food())
    assert kind=='food' and summary==dict(materials=2,specs=2,origins=1)
    refs=dict(materials=merge_materials([],payload['materials']),products=[],parties=[])
    for supplier,expected in [('甲廠商',('X01','X01A')),('乙廠商',('X02','X02A')),
                              ('不明廠商',('','')),('',('',''))]:
        row=dict(name='肉丁',supplier_name=supplier,material_code='',spec_code='')
        apply_reference_defaults(dict(receipts=[row],meals=[]),refs)
        assert (row['material_code'],row['spec_code'])==expected

def test_unique_new_spec_not_guessed_for_different_supplier():
    _,payload,_=parse_master(food())
    refs=dict(materials=merge_materials([],payload['materials'][:1]),products=[],parties=[])
    row=dict(name='肉丁',supplier_name='不明廠商',material_code='',spec_code='')
    apply_reference_defaults(dict(receipts=[row],meals=[]),refs)
    assert row['material_code']=='X01' and row['spec_code']==''

def test_keep_legacy_alias_and_spec_mapping():
    _,payload,_=parse_master(food())
    old=[dict(code='X01',name='肉',aliases=['舊名'],specs=[dict(code='X01A',supplier_name='核對過廠商')]),
         dict(code='OLD',name='歷史原料',aliases=[],specs=[])]
    refs=merge_materials(old,payload['materials'])
    by_code={r['code']:r for r in refs}
    assert 'OLD' in by_code
    assert set(by_code['X01']['aliases'])=={'肉','舊名'}
    assert by_code['X01']['specs'][0]['supplier_name']=='核對過廠商'

def test_packaging_separate_and_duplicate_codes_rejected():
    kind,payload,summary=parse_master(workbook({
        '原材料':[['原材料自編碼(必填)','原材料名稱(必填)'],['P01','餐盒']],
        '原材料材質':[['原材料自編碼(必填)','包裝容器材質代碼(必填)'],['P01','24']]
    }))
    assert kind=='packaging' and summary==dict(materials=1,specs=0,origins=0)
    assert payload['materials'][0]['composition'][0]['包裝容器材質代碼']=='24'
    with pytest.raises(ExportError):
        parse_master(food(duplicate=True))
