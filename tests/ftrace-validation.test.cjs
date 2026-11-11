const { test } = require('node:test');
const assert = require('node:assert/strict');
const { validate } = require('../static/ftrace-validation.js');

const receivingRow = () => ({
  key: 'rice', name: '白米', include: true, material_code: 'K011',
  spec_code: 'K011A005', supplier_code: 'A006', received_date: '2026-10-14',
  kg: '2.55', seq: '01', batch: 'K0111151014',
  manufactured_date: '', expiry_date: ''
});
const receiving = rows => ({
  date: '2026-10-14', batch_mode: true, receipts: rows, meals: []
});
const mealRow = () => ({
  key: 'school:1', name: '學校午餐', include: true, product_code: 'P001',
  seq: '01', batch: 'P0011151014', manufactured_date: '2026-10-14',
  expiry_date: '2026-10-14', kg: '50', quantity: '100',
  customer_code: 'C001', materials: [{
    key: 'rice', name: '白米', code: 'K011',
    lots: [{
      report_date: '2026-10-14', seq: '01', spec_code: 'K011A005',
      batch: 'K0111151014', received_date: '', kg: '20', receipt_kg: '30'
    }]
  }]
});
const meals = (kind, rows) => ({date: '2026-10-14',batch_mode:true,meals:rows,receipts:[]});

test('valid batch receiving has no problems', () => {
  assert.deepEqual(validate(receiving([receivingRow()]), 'receiving'), []);
});

test('shows every missing receiving field, not just the first', () => {
  const row = {...receivingRow(), spec_code:'', supplier_code:'', batch:'', kg:'', received_date:''};
  const issues=validate(receiving([row]),'receiving');
  assert.deepEqual(new Set(issues.map(i=>i.field)), new Set(['spec_code','supplier_code','batch','kg','received_date']));
  assert.ok(issues.every(i=>i.rowKey==='rice'));
});

test('issues disappear when a row is fixed or unchecked', () => {
  const row=receivingRow(); row.spec_code='';
  assert.equal(validate(receiving([row]),'receiving').length,1);
  row.spec_code='K011A005';
  assert.equal(validate(receiving([row]),'receiving').length,0);
  row.spec_code=''; row.include=false;
  assert.equal(validate(receiving([row]),'receiving').length,1);
  assert.equal(validate(receiving([row]),'receiving')[0].rowKey,''); // only no selected rows
});

test('duplicate report sequence, invalid date and invalid weight all appear', () => {
  const a=receivingRow(),b={...receivingRow(),key:'rice2',name:'第二批白米',kg:'0',received_date:'2026-10-15'};
  const issues=validate(receiving([a,b]),'receiving');
  assert.ok(issues.some(i=>i.rowKey==='rice2' && i.field==='seq'));
  assert.ok(issues.some(i=>i.rowKey==='rice2' && i.field==='received_date'));
  assert.ok(issues.some(i=>i.rowKey==='rice2' && i.field==='kg'));
});

test('nested manufacturing lot problems identify material and lot indices', () => {
  const row=mealRow();
  row.materials[0].lots[0].spec_code='';
  row.materials[0].lots[0].kg='0';
  const issues=validate(meals('manufacture',[row]),'manufacture');
  assert.ok(issues.some(i=>i.field==='spec_code' && i.materialIndex===0 && i.lotIndex===0));
  assert.ok(issues.some(i=>i.field==='kg' && i.materialIndex===0 && i.lotIndex===0));
});

test('valid manufacturing and delivery pass', () => {
  const row=mealRow();
  assert.equal(validate(meals('manufacture',[row]),'manufacture').length,0);
  assert.equal(validate(meals('delivery',[row]),'delivery',meals('manufacture',[row])).length,0);
});

test('delivery mismatches with manufacture point at affected fields', () => {
  const made=mealRow(),delivered={...mealRow(),batch:'OTHER',kg:'99',quantity:'101'};
  const issues=validate(meals('delivery',[delivered]),'delivery',meals('manufacture',[made]));
  assert.ok(issues.some(i=>i.field==='batch'));
  assert.ok(issues.some(i=>i.field==='kg'));
  assert.ok(issues.some(i=>i.field==='quantity'));
});

test('batch total over available receipt points to the later used lot', () => {
  const a=mealRow(),b=mealRow();
  a.materials[0].lots[0].kg='20';
  b.key='school:2';b.product_code='P002';b.materials[0].lots[0].kg='20';
  const issues=validate(meals('manufacture',[a,b]),'manufacture');
  assert.ok(issues.some(i=>i.rowKey==='school:2' && i.field==='kg' && i.lotIndex===0));
});

test('legacy manufacture does not require batches but still checks material codes', () => {
  const row=mealRow();
  row.materials[0].code='';
  const data={...meals('manufacture',[row]),batch_mode:false};
  const issues=validate(data,'manufacture');
  assert.equal(issues.length,1);
  assert.equal(issues[0].field,'code');
});
