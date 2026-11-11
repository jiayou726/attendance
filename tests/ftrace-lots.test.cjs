const {test} = require('node:test');
const assert = require('node:assert/strict');
const {catalog, prefill, reconcileMenuReceipts} = require('../static/ftrace-lots.js');
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

test('current receiving shows menu ingredients only, without appending archived purchases',()=>{
  const current=[
    {key:'menu:1', material_key:'1', name:'白米',kg:'100',batch:'NEW'},
    {key:'menu:2', material_key:'2', name:'高麗菜',kg:'40',batch:'NEW2'}
  ];
  const prior=[
    {key:'101',material_key:'1',name:'白米',kg:'105',batch:'REVIEWED',seq:'07'},
    {key:'102',material_key:'3',name:'前一天採購的食材',kg:'9',batch:'OLD'}
  ];
  const rows=reconcileMenuReceipts(current,prior);
  assert.deepEqual(rows.map(x=>x.key),['menu:1','menu:2']);
  assert.equal(rows[0].kg,'105');
  assert.equal(rows[0].batch,'REVIEWED');
  assert.equal(rows[0].seq,'07');
  assert.equal(rows[1].kg,'40');
  assert.equal(current[0].kg,'100'); // Does not change fresh estimates.
  assert.equal(prior[1].kg,'9'); // Old shared revisions remain unchanged.
});

test('never guess a batch when archived records repeat the same ingredient',()=>{
  const current=[{key:'menu:4',material_key:'4',name:'木耳',kg:'5',batch:'DEFAULT'}];
  const old=[
    {key:'10',material_key:'4',kg:'2',batch:'LOT-A'},
    {key:'11',material_key:'4',kg:'3',batch:'LOT-B'}
  ];
  assert.deepEqual(reconcileMenuReceipts(current,old),current);
  assert.equal(reconcileMenuReceipts([],old).length,0);
});

test('current menu-key revision wins over historic duplicate receipt IDs',()=>{
  const current=[{key:'menu:7',material_key:'7',name:'紅蘿蔔',kg:'8'}];
  const old=[
    {key:'21',material_key:'7',kg:'5'},
    {key:'menu:7',material_key:'7',kg:'7',include:false}
  ];
  const rows=reconcileMenuReceipts(current,old);
  assert.equal(rows.length,1);
  assert.equal(rows[0].kg,'7');
  assert.equal(rows[0].include,false);
});
