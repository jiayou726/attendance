const {test} = require('node:test');
const assert = require('node:assert/strict');
const {catalog, prefill} = require('../static/ftrace-lots.js');
const row = {include:true,material_code:'TEST-RAW',seq:'01',spec_code:'TEST-SPEC',batch:'TEST-BATCH',kg:'20'};
function storage(entries) {
  const map = new Map(entries);
  return {length:map.size,key:i=>[...map.keys()][i],getItem:k=>map.get(k)};
}
const draft = (date, receipts=[row]) => JSON.stringify({date,batch_mode:true,receipts});
test('only complete, included receipts of the exact material, scope and past date are offered',()=>{
  const data = storage([
    ['ftrace:v1:formal:2026-10-05:receiving',draft('2026-10-05',[
      row,{...row,include:false},{...row,material_code:'OTHER'},
      {...row,spec_code:''},{...row,seq:''},{...row,batch:''},{...row,kg:'0'}])],
    ['ftrace:v1:practice:2026-10-05:receiving',draft('2026-10-05')],
    ['ftrace:v1:formal:2026-10-07:receiving',draft('2026-10-07')],
    ['ftrace:v1:formal:2026-10-04:receiving','broken JSON'],
    ['ftrace:v1:formal:2026-10-03:manufacture',draft('2026-10-03')],
  ]);
  const choices=catalog(data,'formal','2026-10-06','TEST-RAW');
  assert.equal(choices.length,1);
  assert.equal(choices[0].report_date,'2026-10-05');
  assert.deepEqual(catalog(data,'formal','2026-10-06',''),[]);
});
test('single saved batch fills the link and preserves recipe/actual usage',()=>{
  const choices=catalog(storage([['ftrace:v1:formal:2026-10-05:receiving',draft('2026-10-05')]]),'formal','2026-10-06','TEST-RAW');
  const lots=[{kg:'2.5'}];
  assert.equal(prefill(lots,choices),true);
  assert.equal(lots[0].seq,'01');
  assert.equal(lots[0].spec_code,'TEST-SPEC');
  assert.equal(lots[0].batch,'TEST-BATCH');
  assert.equal(lots[0].kg,'2.5');
});
test('never chooses between batches, overwrites manual identity, or fills split batches',()=>{
  const choice={report_date:'2026-10-05',seq:'01',spec_code:'TEST-SPEC',batch:'TEST-BATCH'};
  assert.equal(prefill([{kg:'2'}],[]),false);
  assert.equal(prefill([{kg:'2'}],[choice,{...choice,batch:'OTHER'}]),false);
  for(const key of ['report_date','seq','spec_code','batch','received_date']) {
    const lots=[{kg:'2',[key]:'MANUAL'}];
    assert.equal(prefill(lots,[choice]),false);
    assert.equal(lots[0][key],'MANUAL');
  }
  assert.equal(prefill([{kg:'1'},{kg:'1'}],[choice]),false);
});
test('pending imported receipts are never automatically linked',()=>{
  assert.equal(prefill([{kg:'2'}],[{report_date:'2026-10-05',seq:'0',spec_code:'TEST-SPEC',received_date:'2026-10-05',ready:false}]),false);
});
