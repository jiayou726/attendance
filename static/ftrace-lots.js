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
    if (lots.length !== 1 || choices.length !== 1 || choices[0].ready === false || identity.some(k => lots[0][k])) return false;
    Object.assign(lots[0], choices[0], {kg: lots[0].kg});
    return true;
  }
  // Only current day's menu materials belong in a new receiving export.
  // Preserve an unambiguous existing edit for the same ingredient, but never
  // reinsert archived purchases or choose arbitrarily between two old lots.
  function reconcileMenuReceipts(menuReceipts, savedReceipts) {
    if (!Array.isArray(menuReceipts)) return [];
    const saved = Array.isArray(savedReceipts) ? savedReceipts : [];
    const counts = new Map();
    for (const row of saved) {
      const key = String(row?.material_key || '');
      if (!key) continue;
      if (!counts.has(key)) counts.set(key, []);
      counts.get(key).push(row);
    }
    const fields = ['include','material_code','spec_code','supplier_code',
      'kg','received_date','seq','batch','manufactured_date','expiry_date'];
    return menuReceipts.map(row => {
      const matches = counts.get(String(row.material_key || '')) || [];
      const exact = matches.find(old => old.key === row.key);
      const prior = exact || (matches.length === 1 ? matches[0] : null);
      if (!prior) return {...row};
      const restored = {...row};
      for (const field of fields) {
        if (Object.hasOwn(prior, field)) restored[field] = prior[field];
      }
      return restored;
    });
  }
  const api = {catalog, prefill, identity, reconcileMenuReceipts};
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.FtraceLots = api;
})(typeof window === 'undefined' ? {} : window);
