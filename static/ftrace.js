/* Export-only editor. Never calls menu/procurement write endpoints. */
(() => {
  'use strict';
  const initial = JSON.parse(document.getElementById('ftrace-data').textContent);
  const settings = JSON.parse(document.getElementById('ftrace-settings').textContent);
  const scope = location.pathname.includes('practice-workspace') ? 'practice' : 'formal';
  const key = `ftrace:v1:${scope}:${initial.date}:${settings.kind}`;
  const referenceKey = `ftrace:references:v1:${scope}`;
  const editor = document.getElementById('ftrace-editor');
  const errorBox = document.getElementById('ftrace-error');
  let data = structuredClone(initial);
  let refs = {materials: [], products: [], parties: [], mappings: {}};
  const status = message => { document.getElementById('draft-status').textContent = message; };
  function fail(message) { errorBox.textContent = message; errorBox.hidden = false; }
  function readSaved(name) { try { return JSON.parse(localStorage.getItem(name)); } catch (_) { return null; } }
  const savedRefs = readSaved(referenceKey);
  if (savedRefs && Array.isArray(savedRefs.materials) && Array.isArray(savedRefs.products) && Array.isArray(savedRefs.parties)) refs = savedRefs;
  refs.mappings ||= {};
  const saved = readSaved(key);
  if (!saved && settings.kind === 'delivery') {
    const manufacture = readSaved(`ftrace:v1:${scope}:${initial.date}:manufacture`);
    if (manufacture?.date === initial.date && Array.isArray(manufacture.meals)) {
      for (const row of data.meals) {
        const made = manufacture.meals.find(m => m.key === row.key);
        if (made) for (const field of ['product_code','quantity','kg']) row[field] = made[field] ?? row[field];
      }
    }
  }
  if (saved && saved.date === initial.date && Array.isArray(saved.receipts) && Array.isArray(saved.meals)) {
    // Source membership comes from today's database; removed/cancelled rows stay removed.
    for (const group of ['receipts', 'meals']) {
      const old = new Map(saved[group].map(row => [row.key, row]));
      for (const row of data[group]) {
        const prior = old.get(row.key);
        if (!prior) continue;
        for (const field of ['include','material_code','spec_code','supplier_code','kg','received_date','product_code','customer_code','quantity']) {
          if (Object.hasOwn(prior, field)) row[field] = prior[field];
        }
        for (const material of row.materials || []) material.code = prior.materials?.find(m => m.key === material.key)?.code || '';
      }
    }
    status('已還原本次暫存；來源新增或移除的項目以目前系統為準，請重新核對數量。');
  }
  const normalized = value => String(value || '').replace(/[\s　]/g, '');
  function match(list, name) {
    const matches = list.filter(r => normalized(r.name) === normalized(name));
    return matches.length === 1 ? matches[0] : null;
  }
  function applyReferences() {
    for (const row of data.receipts) {
      const saved = refs.mappings[`material:${row.material_key}`];
      const material = saved || match(refs.materials, row.name);
      row.material_code ||= material?.code || '';
      const specs = refs.materials.find(r => r.code === row.material_code)?.specs || [];
      const supplierSpecs = specs.filter(s => normalized(s.supplier_name) === normalized(row.supplier_name));
      row.spec_code ||= refs.mappings[`spec:${row.material_key}:${row.supplier_key}`]?.code || (supplierSpecs.length === 1 ? supplierSpecs[0].code : specs.length === 1 ? specs[0].code : '');
      row.supplier_code ||= refs.mappings[`supplier:${row.supplier_key}`]?.code || match(refs.parties, row.supplier_name)?.code || '';
    }
    for (const row of data.meals) {
      const product = refs.mappings[`school:${row.school_key}`] || match(refs.products, row.name);
      row.product_code ||= product?.code || '';
      row.customer_code ||= product?.customer_code || match(refs.parties, row.name)?.code || '';
      const grams = product?.grams || refs.products.find(p => p.code === row.product_code)?.grams;
      if (!row.kg && grams) row.kg = String(Math.round(Number(row.quantity) * Number(grams)) / 1000);
      for (const material of row.materials) material.code ||= refs.mappings[`material:${material.key}`]?.code || match(refs.materials, material.name)?.code || '';
    }
  }
  function lists() {
    for (const [id, values] of [['material-codes',refs.materials],['product-codes',refs.products],['party-codes',refs.parties],['spec-codes',refs.materials.flatMap(r => r.specs || [])]]) {
      const target = document.getElementById(id); target.replaceChildren();
      for (const value of values) { const option = document.createElement('option'); option.value = value.code; option.label = value.name || value.supplier_name || ''; target.append(option); }
    }
  }
  function field(parent, row, name, label, type = 'text', list = null) {
    const wrapper = document.createElement('label'); wrapper.textContent = label;
    const input = document.createElement('input'); input.type = type; input.value = row[name] ?? '';
    input.setAttribute('aria-label', `${row.name || ''} ${label}`);
    if (type === 'number') { input.min = '0'; input.step = name === 'quantity' ? '.01' : '.001'; }
    if (list) input.setAttribute('list', list);
    input.addEventListener('input', () => { row[name] = input.value; invalidate(); });
    wrapper.append(input); parent.append(wrapper);
  }
  function invalidate() { document.getElementById('ftrace-confirm').checked = false; errorBox.hidden = true; status('有尚未暫存的修改。'); }
  function render() {
    editor.replaceChildren();
    const rows = settings.kind === 'receiving' ? data.receipts : data.meals;
    if (!rows.length) { const empty = document.createElement('p'); empty.textContent = '這一天沒有可帶入的資料，請選擇其他日期。'; editor.append(empty); }
    for (const row of rows) {
      const card = document.createElement('section'); card.className = 'ftrace-row';
      const title = document.createElement('h3'); const include = document.createElement('input'); include.type = 'checkbox'; include.checked = row.include;
      include.setAttribute('aria-label', `匯出 ${row.name}`); include.addEventListener('change', () => {row.include = include.checked; invalidate();});
      title.append(include, document.createTextNode(` ${row.name}${row.meal_type ? ' · '+row.meal_type : ''}`)); card.append(title);
      const source = document.createElement('div'); source.className = 'muted'; source.textContent = row.source; card.append(source);
      const fields = document.createElement('div'); fields.className = 'ftrace-fields'; card.append(fields);
      if (settings.kind === 'receiving') {
        source.textContent += `；供應商：${row.supplier_name || '待確認'}`;
        field(fields,row,'material_code','原材料自編碼','text','material-codes');
        field(fields,row,'spec_code','原材料廠商自編碼（規格碼）','text','spec-codes');
        field(fields,row,'supplier_code','供應商自編碼','text','party-codes');
        field(fields,row,'received_date','實際收貨日期','date');
        field(fields,row,'kg','實收淨重（公斤）','number');
      } else {
        field(fields,row,'product_code','產品自編碼','text','product-codes');
        field(fields,row,'quantity',settings.kind === 'manufacture' ? '實際製造份數' : '實際交貨份數','number');
        field(fields,row,'kg',settings.kind === 'manufacture' ? '成品總淨重（公斤）' : '交貨總淨重（公斤）','number');
        if (settings.kind === 'delivery') field(fields,row,'customer_code','下游業者自編碼','text','party-codes');
        else {
          const details = document.createElement('details'); const summary = document.createElement('summary'); summary.textContent = `使用原材料（${row.materials.length} 項，請核對編碼）`; details.append(summary);
          const materials = document.createElement('div'); materials.className = 'ftrace-materials';
          for (const material of row.materials) field(materials,material,'code',material.name,'text','material-codes');
          details.append(materials); card.append(details);
        }
      }
      editor.append(card);
    }
    document.getElementById('ftrace-download').disabled = rows.length === 0;
    lists();
  }
  function remember() {
    for (const row of data.receipts) {
      if (row.material_code) refs.mappings[`material:${row.material_key}`] = {code:row.material_code};
      if (row.spec_code) refs.mappings[`spec:${row.material_key}:${row.supplier_key}`] = {code:row.spec_code};
      if (row.supplier_key && row.supplier_code) refs.mappings[`supplier:${row.supplier_key}`] = {code:row.supplier_code};
    }
    for (const row of data.meals) {
      if (row.product_code) refs.mappings[`school:${row.school_key}`] = {...refs.mappings[`school:${row.school_key}`], code:row.product_code, ...(row.customer_code ? {customer_code:row.customer_code} : {})};
      for (const material of row.materials) if (material.code) refs.mappings[`material:${material.key}`] = {code:material.code};
    }
    try { localStorage.setItem(key, JSON.stringify(data)); localStorage.setItem(referenceKey, JSON.stringify(refs)); status('已暫存到這個瀏覽器。'); return true; }
    catch (_) { fail('瀏覽器無法暫存；仍可直接下載，離開前請先完成匯出。'); return false; }
  }
  document.getElementById('save-draft').addEventListener('click', remember);
  document.getElementById('reset-draft').addEventListener('click', () => {
    data = structuredClone(initial); applyReferences(); render(); invalidate();
    try {localStorage.removeItem(key);} catch (_) {}
    status('已重新帶入目前系統資料。');
  });
  document.getElementById('ftrace-reference').addEventListener('change', async event => {
    try {
      const file = event.target.files[0]; if (!file) return;
      if (file.size > 2_000_000) throw new Error('編碼對照檔過大。');
      const value = JSON.parse(await file.text());
      for (const type of ['materials','products','parties']) {
        if (!Array.isArray(value[type]) || value[type].some(r => !r || typeof r.name !== 'string' || typeof r.code !== 'string')) throw new Error('對照檔格式錯誤。');
      }
      refs = {...value, mappings:refs.mappings}; applyReferences(); render(); invalidate(); remember();
      document.getElementById('reference-status').textContent = `已載入 ${refs.materials.length} 項原料、${refs.products.length} 項餐食產品、${refs.parties.length} 個交易對象。`;
    } catch (error) {fail(error.message || '無法載入對照檔。');}
  });
  function download(blob, name) {const url = URL.createObjectURL(blob); const a = document.createElement('a'); a.href = url; a.download = name; a.click(); setTimeout(() => URL.revokeObjectURL(url), 1000);}
  document.getElementById('ftrace-save-reference').addEventListener('click', () => { remember(); download(new Blob([JSON.stringify(refs,null,2)],{type:'application/json'}),'非追不可編碼對照.json'); });
  document.getElementById('ftrace-download').addEventListener('click', async () => {
    if (!document.getElementById('ftrace-confirm').checked) {fail('請先核對實際數量並勾選確認。'); return;}
    const button = document.getElementById('ftrace-download'); button.disabled = true; errorBox.hidden = true;
    try {
      const body = new URLSearchParams({date:data.date, kind:settings.kind, payload:JSON.stringify(data), confirmed:'1', _csrf_token:settings.csrf});
      const response = await fetch(settings.url,{method:'POST',body});
      if (!response.ok || !response.headers.get('Content-Type')?.includes('application/vnd.ms-excel')) {
        let message = '下載失敗，請重新整理後再試。';
        try {message = (await response.json()).error || message;} catch (_) {}
        throw new Error(message);
      }
      const labels = {receiving:'收貨',manufacture:'製造',delivery:'交貨'};
      download(await response.blob(),`非追不可_${labels[settings.kind]}_${data.date}.xls`); remember(); status('匯入檔已下載，請到非追不可平台匯入。');
    } catch (error) {fail(error.message);} finally {button.disabled = false;}
  });
  applyReferences(); render();
})();
