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
  const systemRefs = JSON.parse(document.getElementById('ftrace-references').textContent);
  let refs = {...systemRefs, mappings: {}};
  const status = message => { document.getElementById('draft-status').textContent = message; };
  function fail(message) { errorBox.textContent = message; errorBox.hidden = false; }
  function readSaved(name) { try { return JSON.parse(localStorage.getItem(name)); } catch (_) { return null; } }
  const savedRefs = readSaved(referenceKey);
  if (savedRefs && Array.isArray(savedRefs.materials) && Array.isArray(savedRefs.products) && Array.isArray(savedRefs.parties)) {
    refs = {...savedRefs, ...Object.fromEntries(['materials','products','parties'].map(k => [k, systemRefs[k].length ? systemRefs[k] : savedRefs[k]]))};
  }
  refs.mappings ||= {};
  let cloudVersion = settings.cloud?.version || 0;
  let sharedReceipts = settings.receipts || [];
  let sharedManufacture = settings.manufacture?.payload || null;
  const saved = settings.cloud?.payload || readSaved(key);
  function manufactureDraft() { return sharedManufacture || readSaved(`ftrace:v1:${scope}:${initial.date}:manufacture`); }
  async function refreshReceipts() {
    const response=await fetch(`${settings.receipts_url}?date=${encodeURIComponent(data.date)}`);
    if(!response.ok) throw new Error('無法讀取共用收貨資料，請稍後再試。');
    sharedReceipts=(await response.json()).receipts;
  }
  async function saveShared(state='draft') {
    remember();
    const body=new URLSearchParams({date:data.date,kind:settings.kind,payload:JSON.stringify(data),version:String(cloudVersion),status:state,_csrf_token:settings.csrf});
    const response=await fetch(settings.shared_url,{method:'POST',body});
    const result=await response.json();
    if(!response.ok) throw new Error(result.error || '共用儲存失敗；修改仍保留在此瀏覽器，請稍後重試。');
    cloudVersion=result.version;
    status(`已儲存到共用資料區，版本 ${cloudVersion}；尚不代表平台已申報。`);
    return result;
  }
  if (!saved && settings.kind === 'delivery') {
    const manufacture = manufactureDraft();
    if (manufacture?.date === initial.date && Array.isArray(manufacture.meals)) {
      for (const row of data.meals) {
        const made = manufacture.meals.find(m => m.key === row.key);
        if (made) for (const field of ['product_code','quantity','kg','batch','manufactured_date','expiry_date']) row[field] = made[field] ?? row[field];
      }
    }
  }
  if (saved && saved.date === initial.date && Array.isArray(saved.receipts) && Array.isArray(saved.meals)) {
    if (settings.kind === 'receiving') {
      // Keep only ingredients present in today's served menus. Historical
      // out-of-menu rows remain in the shared document's version history.
      data.receipts = FtraceLots.reconcileMenuReceipts(data.receipts, saved.receipts);
    }
    // For manufacture/delivery, only currently served menu entries are shown.
    for (const group of ['receipts', 'meals']) {
      const old = new Map(saved[group].map(row => [row.key, row]));
      for (const row of data[group]) {
        const prior = old.get(row.key);
        if (!prior) continue;
        for (const field of ['include','material_code','spec_code','supplier_code','kg','received_date','product_code','customer_code','quantity','seq','batch','manufactured_date','expiry_date']) {
          if (Object.hasOwn(prior, field)) row[field] = prior[field];
        }
        if (!Object.hasOwn(prior,'batch')) row.batch = '';
        for (const material of row.materials || []) {
          const previous = prior.materials?.find(m => m.key === material.key);
          material.code = previous?.code || '';
          if (Array.isArray(previous?.lots)) material.lots = previous.lots;
        }
      }
    }
    status(settings.cloud?.payload ? `已載入共用資料版本 ${cloudVersion}；修改後請儲存。` : '已還原瀏覽器暫存；請儲存到共用資料區，才能跨電腦使用。');
  }
  data.batch_mode = saved?.batch_mode !== false;
  const batchToggle = document.getElementById('batch-mode');
  batchToggle.checked = data.batch_mode;
  data.receipts.forEach((row,i) => { row.seq ||= String(i+1).padStart(2,'0'); });
  data.meals.forEach((row,i) => {
    row.seq ||= String(i+1).padStart(2,'0');
    row.batch ||= historicalBatch(row.product_code,initial.date);
    row.manufactured_date ||= initial.date; row.expiry_date ||= initial.date;
  });
  const normalized = value => String(value || '').replace(/[\s　]/g, '');
  function match(list, name) {
    const matches = list.filter(r => [r.name, ...(r.aliases || [])].some(n => normalized(n) === normalized(name)));
    return matches.length === 1 ? matches[0] : null;
  }
  function applyReferences() {
    for (const row of data.receipts) {
      const saved = refs.mappings[`material:${row.material_key}`];
      const material = saved || match(refs.materials, row.name);
      if (saved) row.material_code = saved.code;
      else row.material_code ||= material?.code || '';
      const specs = refs.materials.find(r => r.code === row.material_code)?.specs || [];
      const supplierSpecs = specs.filter(s => normalized(s.supplier_name) === normalized(row.supplier_name));
      const savedSpec = refs.mappings[`spec:${row.material_key}:${row.supplier_key}`];
      const savedSupplier = refs.mappings[`supplier:${row.supplier_key}`];
      if (savedSpec) row.spec_code = savedSpec.code;
      else row.spec_code ||= supplierSpecs.length === 1 ? supplierSpecs[0].code : specs.length === 1 ? specs[0].code : '';
      if (savedSupplier) row.supplier_code = savedSupplier.code;
      else row.supplier_code ||= match(refs.parties.filter(p => !p.relation || p.relation[1] === '1'), row.supplier_name)?.code || '';
    }
    for (const row of data.meals) {
      const product = refs.mappings[`school:${row.school_key}`] || match(refs.products, row.name);
      row.product_code ||= product?.code || '';
      const customers = refs.parties.filter(p => !p.relation || p.relation[0] === '1');
      row.customer_code ||= product?.customer_code || match(customers, row.school_key)?.code || match(customers, row.name)?.code || '';
      const grams = product?.grams || row.grams_per_portion || refs.products.find(p => p.code === row.product_code)?.grams;
      if (!row.kg && grams) row.kg = String(Math.round(Number(row.quantity) * Number(grams)) / 1000);
      for (const material of row.materials) material.code ||= refs.mappings[`material:${material.key}`]?.code || match(refs.materials, material.name)?.code || '';
    }
  }
  function lists() {
    for (const [id, values] of [['material-codes',refs.materials],['product-codes',refs.products],['party-codes',refs.parties]]) {
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
    return input;
  }
  function specField(parent, row) {
    const input = field(parent,row,'spec_code','原材料廠商自編碼（規格碼）');
    const picker = document.createElement('details'); picker.className = 'ftrace-spec-picker';
    const summary = document.createElement('summary'); summary.textContent = '選擇產品／包裝規格';
    const options = document.createElement('div'); options.className = 'ftrace-spec-options';
    picker.append(summary, options);
    const description = document.createElement('small'); description.className = 'muted';
    input.parentElement.append(picker, description);
    const details = (spec, material) => [...new Set([
      spec.product_name || material.name,
      spec.package_label,
      spec.manufacturer_name && `製造廠：${spec.manufacturer_name}`,
      spec.supplier_name && `廠商：${spec.supplier_name}`
    ].filter(Boolean))].join('；');
    function update(reset = false) {
      if (reset) { row.spec_code = ''; input.value = ''; }
      const material = refs.materials.find(m => m.code === row.material_code);
      const specs = material?.specs || [];
      options.replaceChildren();
      for (const spec of specs) {
        const option = document.createElement('button'); option.type = 'button';
        option.textContent = `${spec.code}（${details(spec,material)}）`;
        option.addEventListener('click', () => {
          row.spec_code = spec.code; input.value = spec.code; picker.open = false;
          invalidate(); update();
        });
        options.append(option);
      }
      if (!specs.length) {
        const empty = document.createElement('p');
        empty.textContent = material ? '平台對照沒有此原料的規格，請核對後填寫。' : '請先選原材料，再選產品規格。';
        options.append(empty);
      }
      const chosen = specs.find(s => s.code === row.spec_code);
      description.textContent = chosen ? `（${details(chosen,material)}）` :
        !material ? '先選原材料，才會顯示這個食材的規格。' :
        specs.length ? '選項括號會顯示產品、包裝與廠商；也可直接輸入規格碼。' :
        '平台對照沒有此原料的規格，請核對後填寫。';
    }
    input.addEventListener('input', () => update());
    update(); return update;
  }
  function invalidate() { document.getElementById('ftrace-confirm').checked = false; errorBox.hidden = true; status('有尚未暫存的修改。'); }
  function historicalBatch(code, day) {
    const match=String(day || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if(!code || !match) return '';
    const value=`${code}${String(Number(match[1])-1911).padStart(3,'0')}${match[2]}${match[3]}`;
    return value.length<=40 ? value : '';
  }
  function action(parent, text, callback) {
    const button = document.createElement('button'); button.type='button'; button.className='btn secondary'; button.textContent=text;
    button.addEventListener('click', callback); parent.append(button); return button;
  }
  function receiptBatchFields(card, fields, row) {
    field(fields,row,'seq','收貨申報流水號');
    row.batch ||= historicalBatch(row.material_code,row.received_date);
    const batchInput=field(fields,row,'batch','原料批號（依包裝或收貨紀錄）');
    const hint=document.createElement('small');hint.className='muted';hint.textContent='內部批號先填原材料碼＋民國收貨日期；有原廠批號或既有批號時請改成原紀錄。';batchInput.parentElement.append(hint);
    field(fields,row,'manufactured_date','原料製造日期（選填）','date');
    field(fields,row,'expiry_date','原料有效日期（選填）','date');
    action(card,'依既有格式重帶：原材料碼＋民國收貨日期', () => {
      row.batch=historicalBatch(row.material_code,row.received_date); render(); invalidate();
    });
  }
  function mealBatchFields(fields,row) {
    row.batch ||= historicalBatch(row.product_code,row.manufactured_date);
    field(fields,row,'seq','申報流水號'); field(fields,row,'batch','餐食製造批號');
    field(fields,row,'manufactured_date','餐食製造日期','date');
    field(fields,row,'expiry_date','餐食有效日期','date');
  }
  function receiptCatalog(materialCode) {
    let local=[];
    try { local=FtraceLots.catalog(localStorage,scope,data.date,materialCode); } catch (_) {}
    const found=new Map();
    for(const lot of [...sharedReceipts.filter(r=>r.material_code===materialCode && r.report_date<=data.date),...local]) {
      const id=JSON.stringify(FtraceLots.identity.map(k=>lot[k] || ''));
      if(!found.has(id)) found.set(id,lot);
    }
    return [...found.values()];
  }
  function materialLots(parent,meal,material) {
    const note=document.createElement('p'); note.className='muted';
    material.lots ||= [{report_date:'',seq:'',spec_code:'',batch:'',received_date:'',kg:material.estimated_kg || ''}];
    const catalog=receiptCatalog(material.code);
    FtraceLots.prefill(material.lots,catalog);
    const availability = !material.code ? '請先選擇原材料自編碼。' : !catalog.length ?
      '找不到此原料收貨紀錄；請先儲存收貨，或到共用資料區匯入歷史 XLS。' :
      catalog.length === 1 ? (catalog[0].ready===false ? '這筆歷史收貨尚待核對，請先到共用資料區補齊，不能猜填日期或批號。' : '只有一筆完整收貨紀錄，空白批次已自動帶入；請確認實際使用這批。') :
      `找到 ${catalog.length} 筆收貨紀錄，請選實際使用的批次，不會自動猜最新一筆。`;
    note.textContent=`${availability} 配方估算使用 ${material.estimated_kg || '待補'} 公斤，請依實際領料修改。共用草稿不代表平台已申報，也未扣庫存。`;
    parent.append(note);
    action(parent,'重新讀取共用收貨紀錄',async()=>{try{await refreshReceipts();render();invalidate();}catch(error){fail(error.message);}});
    for (const [index,lot] of material.lots.entries()) {
      const box=document.createElement('div'); box.className='ftrace-lot'; parent.append(box);
      const heading=document.createElement('strong'); heading.textContent=`${material.name}：原料批次 ${index+1}`; box.append(heading);
      const select=document.createElement('select'); select.setAttribute('aria-label',`${meal.name} ${material.name} 選擇收貨批次 ${index+1}`);
      const empty=document.createElement('option'); empty.value=''; empty.textContent='選共用收貨紀錄，或在下方手填平台既有資料'; select.append(empty);
      catalog.forEach((r,i) => {const o=document.createElement('option');o.value=String(i);o.textContent=`${r.report_date}／流水號 ${r.seq}／${r.spec_code}／批號 ${r.batch}／收貨 ${r.receipt_kg} 公斤／${r.origin || '瀏覽器暫存'}${r.ready===false ? '（待核對）' : ''}`;select.append(o);});
      const selected=catalog.findIndex(r=>FtraceLots.identity.every(k=>(r[k] || '')===(lot[k] || '')));
      if(selected>=0) select.value=String(selected);
      select.addEventListener('change', () => {if(select.value!==''){Object.assign(lot,catalog[Number(select.value)],{kg:lot.kg});render();invalidate();}});
      box.append(select);
      const form=document.createElement('div'); form.className='ftrace-fields'; box.append(form);
      const label=`${meal.name} ${material.name} 批次${index+1}`;
      const edit={...lot,name:label};
      for (const [key,title,type] of [
        ['report_date','收貨申報年月／日期（YYYY-MM 或 YYYY-MM-DD）','text'],
        ['seq','收貨申報流水號（同平台）','text'],['spec_code','收貨原材料廠商自編碼','text'],
        ['batch','收貨批號（新制）','text'],['received_date','收貨日期（僅舊制無批號時填）','date'],
        ['kg','實際使用淨重（公斤）','number']]) {
        const input=field(form,edit,key,title,type); input.addEventListener('input',()=>{lot[key]=input.value;});
      }
      action(box,'移除此原料批次',()=>{material.lots.splice(index,1);render();invalidate();});
    }
    action(parent,'同原料再加一個使用批次',()=>{material.lots.push({report_date:'',seq:'',spec_code:'',batch:'',received_date:'',kg:''});render();invalidate();});
  }
  function render() {
    const expanded = new Set([...editor.querySelectorAll('.ftrace-row')].filter(el=>el.querySelector('details')?.open).map(el=>el.dataset.rowKey));
    editor.replaceChildren();
    const rows = settings.kind === 'receiving' ? data.receipts : data.meals;
    if (!rows.length) { const empty = document.createElement('p'); empty.textContent = '這一天沒有可帶入的資料，請選擇其他日期。'; editor.append(empty); }
    for (const row of rows) {
      const card = document.createElement('section'); card.className = 'ftrace-row'; card.dataset.rowKey=row.key;
      const title = document.createElement('h3'); const include = document.createElement('input'); include.type = 'checkbox'; include.checked = row.include;
      include.setAttribute('aria-label', `匯出 ${row.name}`); include.addEventListener('change', () => {row.include = include.checked; invalidate();});
      title.append(include, document.createTextNode(` ${row.name}${row.meal_type ? ' · '+row.meal_type : ''}`)); card.append(title);
      const source = document.createElement('div'); source.className = 'muted'; source.textContent = row.source; card.append(source);
      const fields = document.createElement('div'); fields.className = 'ftrace-fields'; card.append(fields);
      if (settings.kind === 'receiving') {
        source.textContent += `；供應商：${row.supplier_name || '待確認'}`;
        const materialInput = field(fields,row,'material_code','原材料自編碼','text','material-codes');
        const updateSpecs = specField(fields,row);
        materialInput.addEventListener('input', () => updateSpecs(true));
        field(fields,row,'supplier_code','供應商自編碼','text','party-codes');
        field(fields,row,'received_date','實際收貨日期','date');
        field(fields,row,'kg','實收淨重（公斤）','number');
        if (data.batch_mode) receiptBatchFields(card,fields,row);
      } else {
        field(fields,row,'product_code','產品自編碼','text','product-codes');
        field(fields,row,'quantity',settings.kind === 'manufacture' ? '實際製造份數' : '實際交貨份數','number');
        const weightInput = field(fields,row,'kg',settings.kind === 'manufacture' ? '成品總淨重（公斤）' : '交貨總淨重（公斤）','number');
        if (!row.kg) { const reason = document.createElement('small'); reason.className='muted'; reason.textContent='尚未設定每份成品重量，請填實際總公斤數。'; weightInput.parentElement.append(reason); }
        if (data.batch_mode) {
          mealBatchFields(fields,row);
          if (settings.kind==='manufacture') action(card,'依既有格式重帶：產品碼＋民國製造日期',()=>{row.batch=historicalBatch(row.product_code,row.manufactured_date);render();invalidate();});
        }
        if (settings.kind === 'delivery') field(fields,row,'customer_code','下游業者自編碼','text','party-codes');
        else {
          const details = document.createElement('details'); details.open=expanded.has(row.key); const summary = document.createElement('summary'); summary.textContent = `使用原材料（${row.materials.length} 項，請核對編碼）`; details.append(summary);
          const materials = document.createElement('div'); materials.className = 'ftrace-materials';
          for (const material of row.materials) {
            const materialCard = document.createElement('section'); materials.append(materialCard);
            const originalCode=material.code;
            const codeInput=field(materialCard,material,'code',material.name,'text','material-codes');
            codeInput.addEventListener('change',()=>{
              if(material.code!==originalCode) {delete material.lots;render();invalidate();}
            });
            if (data.batch_mode) materialLots(materialCard,row,material);
          }
          details.append(materials); card.append(details);
        }
      }
      editor.append(card);
    }
    document.getElementById('ftrace-download').disabled = rows.length === 0;
    lists();
    const uploadHelp = document.getElementById('ftrace-upload-help');
    const job = {receiving:'收貨資料(餐食製造業)',manufacture:'餐食業製造資料',delivery:data.batch_mode ? '交貨資料' : '交貨資料(餐食製造業)'}[settings.kind];
    uploadHelp.textContent = `下載後，到非追不可對應頁的「資料匯入」，作業名稱選「${job}」，使用下載的 XLS 檔。${data.batch_mode ? '請先收貨，再製造，最後交貨，確認批次已存在。' : ''}平台只新增，請先查是否已申報，避免重複。`;
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
  if (settings.kind === 'delivery') action(document.querySelector('.ftrace-actions'),'帶入共用製造的批號與數量', async () => {
    try {
      const response=await fetch(`${settings.shared_url}?date=${encodeURIComponent(data.date)}&kind=manufacture`);
      if(!response.ok) throw new Error('無法讀取共用製造紀錄。');
      sharedManufacture=(await response.json()).payload;
    } catch(error) {fail(error.message);return;}
    const made=manufactureDraft();
    if (!made?.batch_mode) {fail('請先在製造頁核對批號與數量並儲存到共用資料區。');return;}
    for (const row of data.meals) {
      const source=made.meals?.find(m=>m.key===row.key);
      if(source) for(const field of ['product_code','batch','manufactured_date','expiry_date','kg','quantity']) row[field]=source[field] ?? row[field];
    }
    render();invalidate();
  });
  document.getElementById('save-draft').addEventListener('click', async()=>{
    try{await saveShared();}catch(error){fail(error.message);}
  });
  document.getElementById('restore-browser-draft').addEventListener('click',()=>{
    const backup=readSaved(key);
    if(!backup || backup.date!==initial.date || !Array.isArray(backup.receipts) || !Array.isArray(backup.meals)) {
      fail('這個瀏覽器沒有本日、本類型的備份。');return;
    }
    data=structuredClone(backup);
    if (settings.kind === 'receiving') {
      data.receipts = FtraceLots.reconcileMenuReceipts(initial.receipts, backup.receipts);
    }
    batchToggle.checked=data.batch_mode!==false;render();invalidate();
    status('已載入本機備份，尚未覆蓋共用紀錄；請核對後再儲存。');
  });
  document.getElementById('migrate-drafts').addEventListener('click',async()=>{
    let imported=0,skipped=0;
    try {
      for(let i=0;i<localStorage.length;i++) {
        const name=localStorage.key(i);
        if(!name?.startsWith(`ftrace:v1:${scope}:`)) continue;
        const parts=name.split(':');const kind=parts.at(-1);const draft=readSaved(name);
        if(!['receiving','manufacture','delivery'].includes(kind) || !draft?.date) continue;
        const body=new URLSearchParams({date:draft.date,kind,payload:JSON.stringify(draft),version:'0',_csrf_token:settings.csrf});
        const response=await fetch(settings.shared_url,{method:'POST',body});
        if(response.status===409){skipped++;continue;}
        const result=await response.json();
        if(!response.ok) throw new Error(result.error || '移入失敗。');
        imported++;
        if(draft.date===data.date && kind===settings.kind) cloudVersion=result.version;
      }
      await refreshReceipts();render();status(`已移入 ${imported} 份舊暫存；${skipped} 份已有共用紀錄，保留共用版本未覆蓋。`);
    }catch(error){fail(`已移入 ${imported} 份。${error.message}`);}
  });
  document.getElementById('reset-draft').addEventListener('click', () => {
    const mode = data.batch_mode; data = structuredClone(initial); data.batch_mode = mode;
    data.receipts.forEach((r,i) => r.seq=String(i+1).padStart(2,'0'));
    data.meals.forEach((r,i) => {r.seq=String(i+1).padStart(2,'0');r.batch=historicalBatch(r.product_code,initial.date);r.manufactured_date=initial.date;r.expiry_date=initial.date;});
    batchToggle.addEventListener('change', () => {data.batch_mode=batchToggle.checked; render(); invalidate();});
  applyReferences(); render(); invalidate();
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
      if (data.batch_mode && settings.kind === 'delivery') {
        const made=manufactureDraft();
        for (const row of data.meals.filter(r=>r.include)) {
          const source=made?.batch_mode && made.meals?.find(m=>m.key===row.key && m.include);
          if(source && ['product_code','batch','manufactured_date','expiry_date'].some(k=>row[k]!==source[k])) throw new Error(`${row.name}：交貨批號／日期與已暫存製造不同，請先帶入製造批號。`);
          if(source && (Number(row.kg)>Number(source.kg) || Number(row.quantity)>Number(source.quantity))) throw new Error(`${row.name}：交貨數量或重量大於已暫存製造，請核對。`);
        }
      }
      const body = new URLSearchParams({date:data.date, kind:settings.kind, payload:JSON.stringify(data), confirmed:'1', _csrf_token:settings.csrf});
      const response = await fetch(settings.url,{method:'POST',body});
      if (!response.ok || !response.headers.get('Content-Type')?.includes('application/vnd.ms-excel')) {
        let message = '下載失敗，請重新整理後再試。';
        try {message = (await response.json()).error || message;} catch (_) {}
        throw new Error(message);
      }
      const labels = {receiving:'收貨',manufacture:'製造',delivery:'交貨'};
      const file=await response.blob();await saveShared('exported');
      download(file,`非追不可_${labels[settings.kind]}_${data.batch_mode ? '依批號_' : ''}${data.date}.xls`);status('已儲存共用紀錄並下載匯入檔；請到平台匯入成功後，在共用資料區標記申報完成。');
    } catch (error) {fail(error.message);} finally {button.disabled = false;}
  });
  batchToggle.addEventListener('change', () => {data.batch_mode=batchToggle.checked; render(); invalidate();});
  applyReferences(); render();
  if (systemRefs.materials.length || systemRefs.parties.length) document.getElementById('reference-status').textContent = `系統已載入 ${systemRefs.materials.length} 項原料、${systemRefs.parties.length} 個交易對象。`;
})();
