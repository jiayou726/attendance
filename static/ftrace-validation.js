/* Fast, side-effect-free preflight for the Ftrace editor.
 * The Python exporter remains authoritative; this mirrors its field rules
 * so all ordinary corrections can be shown at once.
 */
(function (root) {
  'use strict';

  function validate(data, kind, manufacture) {
    const issues = [];
    const batch = data.batch_mode === true;
    const source = kind === 'receiving' ? data.receipts : data.meals;
    const text = (row, key) => String(row && row[key] != null ? row[key] : '').trim();
    const add = (row, field, message, materialIndex = null, lotIndex = null) => {
      const rowKey = row ? String(row.key) : '';
      issues.push({
        id: [rowKey, materialIndex, lotIndex, field, message].join('|'),
        rowKey, rowName: row ? String(row.name || '資料') : '整份資料',
        field, materialIndex, lotIndex, message
      });
    };
    const validDate = value => {
      if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) return false;
      const d = new Date(value + 'T00:00:00Z');
      return !Number.isNaN(d.getTime()) && d.toISOString().slice(0, 10) === value;
    };
    const checkCode = (row, object, field, label, materialIndex = null, lotIndex = null, required = true) => {
      const value = text(object, field);
      if (!value && required) add(row, field, '請填寫' + label, materialIndex, lotIndex);
      else if (value && (value.length > (batch ? 40 : 160) || /^[=+@]/.test(value) || /[\x00-\x1f]/.test(value))) {
        add(row, field, label + '格式不正確（編碼最多 ' + (batch ? 40 : 160) + ' 字）', materialIndex, lotIndex);
      }
      return value;
    };
    const checkDate = (row, object, field, label, required = true, materialIndex = null, lotIndex = null) => {
      const value = text(object, field);
      if (!value && required) add(row, field, '請填寫' + label, materialIndex, lotIndex);
      else if (value && !validDate(value)) add(row, field, label + '日期無效，請使用 YYYY-MM-DD', materialIndex, lotIndex);
      return validDate(value) ? value : '';
    };
    const checkNumber = (row, object, field, label, places, materialIndex = null, lotIndex = null) => {
      const value = text(object, field);
      if (!value) {
        add(row, field, '請填寫' + label, materialIndex, lotIndex);
        return null;
      }
      const number = Number(value);
      const decimals = (value.split('.')[1] || '').length;
      if (!Number.isFinite(number) || number <= 0 || number > 999999999 || !/^\d+(\.\d+)?$/.test(value)) {
        add(row, field, label + '必須是有效的正數', materialIndex, lotIndex);
        return null;
      }
      if (decimals > places) add(row, field, label + '最多 ' + places + ' 位小數', materialIndex, lotIndex);
      return number;
    };
    if (!Array.isArray(source)) {
      add(null, '', '資料格式無效，請重新帶入當日資料');
      return issues;
    }
    if (source.length > 99) add(null, '', '每日最多 99 筆資料，請分批處理');
    const rows = source.filter(row => row && row.include === true);
    if (!rows.length) add(null, '', '請先勾選至少一筆要下載的資料');
    const seen = new Set();
    const consumption = new Map();

    for (const row of rows) {
      const codeKey = kind === 'receiving' ? 'material_code' : 'product_code';
      const code = checkCode(row, row, codeKey, kind === 'receiving' ? '原材料自編碼' : '產品自編碼');
      if (batch) {
        const seq = checkCode(row, row, 'seq', '申報流水號');
        const identity = code + '|' + seq;
        if (code && seq && seen.has(identity)) add(row, 'seq', '同一編碼的申報流水號重複');
        seen.add(identity);
        checkCode(row, row, 'batch', kind === 'receiving' ? '原料批號' : '製造批號');
      }
      if (kind === 'receiving') {
        checkCode(row, row, 'spec_code', '原材料廠商自編碼（規格碼）');
        checkCode(row, row, 'supplier_code', '供應商自編碼');
        const received = checkDate(row, row, 'received_date', '實際收貨日期');
        if (received && data.date) {
          if (batch && received !== data.date) add(row, 'received_date', '依批號日申報：收貨日期須等於申報日期 ' + data.date);
          else if (!batch && received > data.date) add(row, 'received_date', '收貨日期不能晚於申報日期');
        }
        const made = batch ? checkDate(row, row, 'manufactured_date', '原料製造日期', false) : '';
        const expiry = batch ? checkDate(row, row, 'expiry_date', '原料有效日期', false) : '';
        if (made && received && made > received) add(row, 'manufactured_date', '原料製造日期晚於收貨日期');
        if (expiry && received && expiry < received) add(row, 'expiry_date', '原料有效日期早於收貨日期');
        if (made && expiry && expiry < made) add(row, 'expiry_date', '有效日期早於製造日期');
        checkNumber(row, row, 'kg', '實收淨重（公斤）', batch ? 7 : 3);
        continue;
      }

      const label = kind === 'manufacture' ? '實際製造份數' : '實際交貨份數';
      checkNumber(row, row, 'kg', '成品總淨重（公斤）', 3);
      checkNumber(row, row, 'quantity', label, batch && kind === 'manufacture' ? 1 : 2);
      let made = '', expiry = '';
      if (batch) {
        made = checkDate(row, row, 'manufactured_date', '餐食製造日期');
        expiry = checkDate(row, row, 'expiry_date', '餐食有效日期');
        if (made && data.date && made > data.date) add(row, 'manufactured_date', '製造日期不可晚於申報日期');
        if (made && expiry && expiry < made) add(row, 'expiry_date', '有效日期早於製造日期');
        if (kind === 'delivery' && expiry && data.date && expiry < data.date) add(row, 'expiry_date', '有效日期早於交貨日期');
      }
      if (kind === 'delivery') {
        checkCode(row, row, 'customer_code', '下游業者自編碼');
        const sourceMeal = batch && manufacture && manufacture.batch_mode && Array.isArray(manufacture.meals)
          ? manufacture.meals.find(item => item.key === row.key && item.include) : null;
        if (sourceMeal) {
          if (['product_code', 'batch', 'manufactured_date', 'expiry_date'].some(key => text(row, key) !== text(sourceMeal, key))) {
            add(row, 'batch', '交貨編碼、批號或日期與共用製造紀錄不同，請重新帶入');
          }
          for (const key of ['kg', 'quantity']) {
            if (Number(row[key]) > Number(sourceMeal[key])) add(row, key, '交貨數量／重量大於共用製造紀錄');
          }
        }
        continue;
      }

      if (!Array.isArray(row.materials) || !row.materials.length || row.materials.length > 1000) {
        add(row, 'materials', '請核對製造使用的原材料項目');
        continue;
      }
      row.materials.forEach((material, materialIndex) => {
        if (!material || typeof material !== 'object') {
          add(row, 'materials', '原材料資料格式錯誤');
          return;
        }
        const materialCode = checkCode(row, material, 'code', material.name + '：原材料自編碼', materialIndex);
        if (!batch) return;
        if (!Array.isArray(material.lots) || !material.lots.length || material.lots.length > 100) {
          add(row, 'code', material.name + '：請選擇或填寫收貨批次', materialIndex);
          return;
        }
        const seenLots = new Set();
        material.lots.forEach((lot, lotIndex) => {
          if (!lot || typeof lot !== 'object') {
            add(row, 'code', '收貨批次格式錯誤', materialIndex);
            return;
          }
          const report = text(lot, 'report_date');
          if (!report || !/^\d{4}-\d{2}(-\d{2})?$/.test(report) ||
              !validDate(report.length === 7 ? report + '-01' : report)) {
            add(row, 'report_date', '請填有效的收貨申報年月／日期（YYYY-MM 或 YYYY-MM-DD）', materialIndex, lotIndex);
          } else if (made && (report.length === 7 ? report + '-01' : report) > made) {
            add(row, 'report_date', '不能使用製造日期之後申報收貨的原料', materialIndex, lotIndex);
          }
          checkCode(row, lot, 'seq', '收貨流水號', materialIndex, lotIndex, false);
          const spec = checkCode(row, lot, 'spec_code', '收貨原材料廠商自編碼', materialIndex, lotIndex);
          const lotBatch = checkCode(row, lot, 'batch', '收貨批號', materialIndex, lotIndex, false);
          const received = checkDate(row, lot, 'received_date', '舊制實際收貨日期', false, materialIndex, lotIndex);
          if (Boolean(lotBatch) === Boolean(received)) {
            add(row, lotBatch ? 'batch' : 'received_date', '收貨批號與舊制收貨日期須擇一填寫', materialIndex, lotIndex);
          }
          if (received && made && received > made) add(row, 'received_date', '收貨日期不可晚於製造日期', materialIndex, lotIndex);
          const used = checkNumber(row, lot, 'kg', '原料實際使用淨重（公斤）', 7, materialIndex, lotIndex);
          const capacity = text(lot, 'receipt_kg');
          if (capacity && used !== null) {
            const available = Number(capacity);
            if (Number.isFinite(available) && used > available) add(row, 'kg', '使用重量大於單筆收貨重量', materialIndex, lotIndex);
          }
          const key = [report, text(lot, 'seq'), spec, lotBatch, received].join('|');
          if (seenLots.has(key)) add(row, 'report_date', '同一原材料重複選取相同批次', materialIndex, lotIndex);
          seenLots.add(key);
          const sharedKey = materialCode + '|' + key;
          if (used !== null && capacity && Number(capacity) > 0) {
            const previous = consumption.get(sharedKey) || {used: 0, capacity: Number(capacity)};
            previous.used += used;
            previous.capacity = Math.min(previous.capacity, Number(capacity));
            consumption.set(sharedKey, previous);
            if (previous.used > previous.capacity) add(row, 'kg', '多筆餐食使用此收貨批次的總量超過收貨重量', materialIndex, lotIndex);
          }
        });
      });
    }
    return issues;
  }

  root.FtraceValidation = { validate };
  if (typeof module !== 'undefined' && module.exports) module.exports = { validate };
})(typeof window !== 'undefined' ? window : globalThis);
