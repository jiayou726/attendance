/* Same-day manufacturing assistant: menu requirements intersect same-day receipts.
 * No stock inference, no historical receipt defaults and no modifications to receipt records.
 */
(function (root) {
  'use strict';
  const identity = ['report_date', 'seq', 'spec_code', 'batch', 'received_date'];
  const clean = v => String(v == null ? '' : v).trim();
  const grams = v => Math.round(Number(v || 0) * 1e7);
  const kg = v => (v / 1e7).toFixed(7).replace(/0+$/, '').replace(/\.$/, '');
  const keyOf = row => identity.map(k => clean(row[k])).join('|');

  function candidatesFor(receipts, day, code) {
    const matches = new Map();
    for (const receipt of receipts || []) {
      if (clean(receipt.material_code) !== code || clean(receipt.report_date) !== day) continue;
      const id = keyOf(receipt);
      const amount = Number(receipt.receipt_kg);
      if (receipt.ready !== true || !clean(receipt.spec_code) ||
          !(clean(receipt.batch) || clean(receipt.received_date)) ||
          !(amount > 0)) continue;
      if (!matches.has(id)) matches.set(id, receipt);
      else if (Number(matches.get(id).receipt_kg) !== amount) {
        // Do not silently choose one of conflicting duplicate sources.
        matches.set(id, {...receipt, ready:false});
      }
    }
    return [...matches.values()].filter(r => r.ready === true);
  }
  function groupMeals(meals) {
    const groups = new Map();
    for (const meal of meals || []) {
      if (!meal || !meal.include) continue;
      for (const material of meal.materials || []) {
        const code = clean(material.code);
        const key = code || 'missing-code:' + clean(material.key || material.name);
        if (!groups.has(key)) {
          groups.set(key, {key, code, name:clean(material.name), items:[], need:0});
        }
        const group = groups.get(key);
        group.items.push({meal, material});
        const amount=Number(material.estimated_kg);
        if (Number.isFinite(amount) && amount > 0) group.need += grams(amount);
      }
    }
    return [...groups.values()];
  }
  function linked(item, candidate, day) {
    const lots = item.material.lots;
    return Array.isArray(lots) && lots.length === 1 &&
      clean(lots[0].report_date) === day && keyOf(lots[0]) === keyOf(candidate) &&
      Number(lots[0].kg) > 0;
  }
  function inspect(meals, receipts, day) {
    const groups = groupMeals(meals);
    for (const group of groups) {
      group.needKg = kg(group.need);
      group.candidates = group.code ? candidatesFor(receipts, day, group.code) : [];
      group.capacity = group.candidates.reduce((sum, candidate) => sum + grams(candidate.receipt_kg), 0);
      if (!group.code) group.state='no-code';
      else if (!group.candidates.length) group.state='missing';
      else if (group.need <= 0) group.state='no-usage';
      else if (group.capacity < group.need) group.state='insufficient';
      else if (group.candidates.some(candidate => group.items.every(item => linked(item,candidate,day))))
        group.state='linked';
      else if (group.candidates.length === 1 && grams(group.candidates[0].receipt_kg) >= group.need)
        group.state='auto';
      else group.state='multiple';
    }
    return groups;
  }
  function assign(group, candidate, day, overwrite = false) {
    if (clean(candidate.report_date) !== day || candidate.ready !== true ||
        clean(candidate.material_code) !== group.code ||
        grams(candidate.receipt_kg) < group.need || !(group.need > 0)) return false;
    if (!overwrite && group.items.some(({material}) =>
      Array.isArray(material.lots) && (material.lots.length !== 1 ||
        identity.some(key => clean(material.lots[0][key]))))) return false;
    for (const {material} of group.items) {
      const amount = Number(material.estimated_kg);
      if (!(amount > 0)) return false;
    }
    for (const {material} of group.items) {
      const selected={...candidate, kg:kg(grams(material.estimated_kg))};
      delete selected.issues;
      material.lots=[selected];
    }
    return true;
  }
  function autoApply(meals, receipts, day) {
    let count=0;
    for (const group of inspect(meals,receipts,day)) {
      if (group.state==='auto' && assign(group,group.candidates[0],day)) count++;
    }
    return count;
  }
  const api={inspect,assign,autoApply,candidatesFor,groupMeals,keyOf};
  if (typeof module !== 'undefined' && module.exports) module.exports=api;
  else root.FtraceDaily=api;
})(typeof window !== 'undefined' ? window : {});
