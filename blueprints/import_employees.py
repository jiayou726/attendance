# blueprints/import_employees.py
# -*- coding: utf-8 -*-
from flask import Blueprint, request, render_template_string, url_for
import pandas as pd
from extensions import db
from models import Employee
from . import CSS

import_bp = Blueprint('imp', __name__, url_prefix='/admin')

@import_bp.route('/import_employees', methods=['GET', 'POST'])
def import_employees():
    msg = ''
    results = []
    if request.method == 'POST':
        file = request.files.get('file')
        if not file or file.filename == '':
            msg = '請選擇檔案'
        else:
            try:
                if file.filename.lower().endswith('.csv'):
                    df = pd.read_csv(file)
                else:
                    df = pd.read_excel(file)
            except Exception as e:
                msg = f'讀取檔案失敗: {e}'
            else:
                required = {'id', 'name', 'area', 'default_break'}
                if not required.issubset(df.columns):
                    msg = '檔案欄位必須包含 id, name, area, default_break'
                else:
                    inserted = 0
                    for _, row in df.iterrows():
                        try:
                            eid = int(row['id'])
                            name = str(row['name']).strip()
                            area = str(row['area']).strip()
                            brk = float(row['default_break']) if row['default_break'] != '' else 0.0
                        except (ValueError, TypeError):
                            results.append((row.get('id', ''), '格式錯誤'))
                            continue

                        if Employee.query.get(eid):
                            results.append((eid, 'ID 已存在'))
                            continue

                        emp = Employee(id=eid, name=name, area=area, default_break=brk)
                        db.session.add(emp)
                        inserted += 1
                        results.append((eid, 'ok'))
                    db.session.commit()
                    msg = f'成功匯入 {inserted} 筆'

    rows = ''.join(f"<tr><td>{eid}</td><td>{res}</td></tr>" for eid, res in results)
    table = f"<table><tr><th>ID</th><th>結果</th></tr>{rows}</table>" if results else ''
    return render_template_string(f"""<!doctype html><html><head>{CSS}</head><body>
<h2>匯入員工</h2>
<p class='success'>{msg}</p>""" +
        "<form method='post' enctype='multipart/form-data'>"+
        "檔案：<input type='file' name='file' accept='.csv,.xls,.xlsx' required>"+
        "<button type='submit'>匯入</button>"+
        f"<a href='{url_for('emp.list_employees')}'>返回</a></form>"+
        table+
        "</body></html>"
    )
