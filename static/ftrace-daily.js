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
  function receiptChoices(receipts, day) {
    // Search by the receipt's actual declared name and code, not the menu name.
    const byCode = new Map();
    for (const receipt of receipts || []) {
      if (clean(receipt.report_date)!==day || !clean(receipt.material_code)) continue;
      const code=clean(receipt.material_code);
      if (!byCode.has(code)) byCode.set(code,[]);
      byCode.get(code).push(receipt);
    }
    const choices=[];
    for (const [code, rows] of byCode) {
      const ready=candidatesFor(rows,day,code);
      for (const receipt of ready) {
        choices.push(receipt);
      }
    }
    return choices.sort((a,b)=>(clean(a.name)+clean(a.material_code)+keyOf(a)).localeCompare(
      clean(b.name)+clean(b.material_code)+keyOf(b), 'zh-Hant'));
  }
  function inheritReceiptCodes(meals, receipts, day) {
    // Both menu and saved receiving rows retain the same internal ingredient
    // key. This is stronger evidence than the external name or alias.
    const keyed = new Map();
    for (const receipt of receipts || []) {
      const materialKey=clean(receipt.material_key), code=clean(receipt.material_code);
      if (clean(receipt.report_date)!==day || !materialKey || !code) continue;
      if (!keyed.has(materialKey)) keyed.set(materialKey,new Set());
      keyed.get(materialKey).add(code);
    }
    let changed=0;
    for (const meal of meals || []) {
      for (const material of meal.materials || []) {
        const codes=keyed.get(clean(material.key));
        if (!codes || codes.size!==1) continue;
        const chosen=[...codes][0];
        if (clean(material.code)===chosen) continue;
        // Never overwrite an earlier user-selected lot/identity.
        if (Array.isArray(material.lots) && material.lots.some(lot =>
          identity.some(field=>clean(lot?.[field])))) continue;
        material.code=chosen;
        changed++;
      }
    }
    return changed;
  }
  function matchReceiving(group, receipt, receipts, day) {
    // Explicit user action: rename/code mismatch can be resolved once for all
    // meals, but only against a real valid same-day receipt.
    if (!group.items.length || !(group.need>0) ||
        !receiptChoices(receipts,day).includes(receipt)) return false;
    const code=clean(receipt.material_code);
    const candidates=candidatesFor(receipts,day,code);
    const capacity=candidates.reduce((sum,item)=>sum+grams(item.receipt_kg),0);
    if (capacity<group.need || group.items.some(({material})=>
      !(grams(material.estimated_kg)>0))) return false;
    const proposed={...group,code,candidates,capacity};
    const assigned=grams(receipt.receipt_kg)>=group.need
      ? assign(proposed,receipt,day,true)
      : distribute(proposed,day);
    if (!assigned) return false;
    for (const {material} of group.items) material.code=code;
    return true;
  }
  function groupMeals(meals) {
    const groups = new Map();
    for (const meal of meals || []) {
      if (!meal || !meal.include) continue;
      for (const material of meal.materials || []) {
        const code = clean(material.code);
        const key = clean(material.key) ? 'ingredient:' + clean(material.key) :
          (code ? 'code:' + code : 'missing-code:' + clean(material.name));
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
  function fullyLinked(group,day) {
    const remaining=new Map(group.candidates.map(candidate=>[keyOf(candidate),grams(candidate.receipt_kg)]));
    for (const {material} of group.items) {
      if (!Array.isArray(material.lots) || !material.lots.length) return false;
      for (const lot of material.lots) {
        if (clean(lot.report_date)!==day) return false;
        const key=keyOf(lot);
        const used=grams(lot.kg);
        if (!remaining.has(key) || !(used>0) || used>remaining.get(key)) return false;
        remaining.set(key,remaining.get(key)-used);
      }
    }
    return true;
  }
  // Explicitly requested allocation across several same-day lots. It is only
  // a menu-based proposal; real stock picking must be confirmed before export.
  function distribute(group, day) {
    if (!group.code || !(group.need>0) || group.capacity < group.need ||
        !group.candidates.length || group.items.some(item=>!(grams(item.material.estimated_kg)>0))) return false;
    const sorted=[...group.candidates].sort((a,b)=>String(a.seq).localeCompare(String(b.seq)));
    const left=sorted.map(lot=>grams(lot.receipt_kg));
    const proposals=[];
    for (const {material} of group.items) {
      let needed=grams(material.estimated_kg);
      const selected=[];
      for (let i=0;i<sorted.length && needed>0;i++) {
        const allocated=Math.min(needed,left[i]);
        if (allocated<=0) continue;
        selected.push({...sorted[i],kg:kg(allocated)});
        delete selected[selected.length-1].issues;
        left[i]-=allocated;needed-=allocated;
      }
      if (needed>0 || !selected.length) return false;
      proposals.push({material,selected});
    }
    for (const proposal of proposals) proposal.material.lots=proposal.selected;
    return true;
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
      else if (fullyLinked(group,day))
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
  const api={inspect,assign,distribute,autoApply,candidatesFor,receiptChoices,inheritReceiptCodes,matchReceiving,groupMeals,keyOf};
  if (typeof module !== 'undefined' && module.exports) module.exports=api;
  else root.FtraceDaily=api;
})(typeof window !== 'undefined' ? window : {});
