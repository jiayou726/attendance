/* Receipt choices come only from saved drafts, never invented platform records. */
(function (root) {
  'use strict';
  const identity = ['report_date', 'seq', 'spec_code', 'batch', 'received_date'];
  function catalog(storage, scope, day, code) {
    if (!code) return [];
    const found = new Map();
    for (let i = 0; i < storage.length; i++) {
      const key = storage.key(i);
      if (!key?.startsWith(`ftrace:v1:${scope}:`) || !key.endsWith(':receiving')) continue;
      let saved;
      try { saved = JSON.parse(storage.getItem(key)); } catch (_) { continue; }
      if (!saved?.batch_mode || !/^\d{4}-\d{2}-\d{2}$/.test(saved.date || '') || saved.date > day) continue;
      for (const row of Array.isArray(saved.receipts) ? saved.receipts : []) {
        if (!row?.include || row.material_code !== code || !row.batch || !row.seq || !row.spec_code || !(Number(row.kg) > 0)) continue;
        const lot = {report_date: saved.date, seq: row.seq, spec_code: row.spec_code,
          batch: row.batch, received_date: '', kg: '', receipt_kg: row.kg, name: row.name || ''};
        found.set(JSON.stringify(identity.map(k => lot[k])), lot);
      }
    }
    return [...found.values()].sort((a, b) => b.report_date.localeCompare(a.report_date));
  }
  function prefill(lots, choices) {
    // Never replace a manual link, a split across batches, or guess between choices.
    if (lots.length !== 1 || choices.length !== 1 || identity.some(k => lots[0][k])) return false;
    Object.assign(lots[0], choices[0], {kg: lots[0].kg});
    return true;
  }
  const api = {catalog, prefill, identity};
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.FtraceLots = api;
})(typeof window === 'undefined' ? {} : window);
