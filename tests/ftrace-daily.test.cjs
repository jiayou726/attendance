const {test} = require('node:test');
const assert = require('node:assert/strict');
const {inspect,autoApply,assign,distribute,candidatesFor} = require('../static/ftrace-daily.js');

const DAY='2026-10-14';
const makeMeal=(key,qty,code='K011',name='白米')=>({
 key,include:true,materials:[{key:'ingredient:1',name,code,estimated_kg:String(qty)}]
});
const lot=(code='K011',kg='50',batch='RICE-1014',date=DAY)=>({
 report_date:date,material_code:code,seq:'01',spec_code:code+'A005',
 batch,received_date:'',receipt_kg:kg,origin:'共用收貨紀錄',
 document_status:'exported',ready:true
});
test('menus choose ingredients, same-day receipts provide batch, quantities use menu',()=>{
 const meals=[makeMeal('schoolA',12),makeMeal('schoolB',18)];
 const receipts=[lot('K011','35'),lot('EGG','10','EGG-01'),lot('K011','100','OLD','2026-10-13')];
 assert.equal(autoApply(meals,receipts,DAY),1);
 const groups=inspect(meals,receipts,DAY);
 assert.equal(groups[0].state,'linked');
 assert.equal(groups[0].needKg,'30');
 assert.equal(meals[0].materials[0].lots[0].kg,'12');
 assert.equal(meals[1].materials[0].lots[0].kg,'18');
 assert.equal(meals[0].materials[0].lots[0].batch,'RICE-1014');
 assert.equal(receipts[0].receipt_kg,'35'); // Do not alter real receipt
});
test('do not include receipts absent from menu',()=>{
 const meals=[makeMeal('school',3)];
 const groups=inspect(meals,[lot(),lot('PORK','9','PORK-1014')],DAY);
 assert.equal(groups.length,1);
 assert.equal(groups[0].code,'K011');
});
test('missing current-day receipts stay unresolved; older matching batch not auto selected',()=>{
 const meals=[makeMeal('schoolA',2)];
 assert.equal(autoApply(meals,[lot('K011','9','YESTERDAY','2026-10-13')],DAY),0);
 assert.equal(inspect(meals,[lot('K011','9','YESTERDAY','2026-10-13')],DAY)[0].state,'missing');
 assert.equal(meals[0].materials[0].lots,undefined);
});
test('receipt less than combined school menu need cannot be silently linked',()=>{
 const meals=[makeMeal('schoolA',12),makeMeal('schoolB',18)];
 assert.equal(autoApply(meals,[lot('K011','25')],DAY),0);
 assert.equal(inspect(meals,[lot('K011','25')],DAY)[0].state,'insufficient');
});
test('multiple batches require explicit choice, preserve manual source link',()=>{
 const meals=[makeMeal('schoolA',12)];
 const receipts=[lot('K011','20','LOT-A'),lot('K011','25','LOT-B')];
 assert.equal(autoApply(meals,receipts,DAY),0);
 assert.equal(inspect(meals,receipts,DAY)[0].state,'multiple');
 const group=inspect(meals,receipts,DAY)[0];
 assert.equal(assign(group,receipts[1],DAY,true),true);
 assert.equal(meals[0].materials[0].lots[0].batch,'LOT-B');
 const prior=[makeMeal('schoolA',12)];
 prior[0].materials[0].lots=[{report_date:'2026-10-13',batch:'HISTORIC'}];
 assert.equal(autoApply(prior,[receipts[0]],DAY),0);
 assert.equal(prior[0].materials[0].lots[0].batch,'HISTORIC');
});
test('one-click explicit allocation splits estimated menu usage across same-day lots',()=>{
 const meals=[makeMeal('schoolA',9),makeMeal('schoolB',6)];
 const receipts=[lot('K011','10','LOT-A'),lot('K011','10','LOT-B')];
 const group=inspect(meals,receipts,DAY)[0];
 assert.equal(group.state,'multiple');
 assert.equal(distribute(group,DAY),true);
 assert.deepEqual(meals[0].materials[0].lots.map(x=>x.kg),['9']);
 assert.deepEqual(meals[1].materials[0].lots.map(x=>x.kg),['1','5']);
 assert.equal(inspect(meals,receipts,DAY)[0].state,'linked');
 assert.equal(receipts[0].receipt_kg,'10'); // Do not modify receipts
});
test('conflicting equal-identity weights must not enter automatic choices',()=>{
 const a=lot('K011','10'),b=lot('K011','12');
 assert.equal(candidatesFor([a,b],DAY,'K011').length,0);
 assert.equal(autoApply([makeMeal('school',5)],[a,b],DAY),0);
});
test('manual old-day and missing material code never become auto valid',()=>{
 const meals=[makeMeal('school',3,'')];
 assert.equal(autoApply(meals,[lot()],DAY),0);
 assert.equal(inspect(meals,[lot()],DAY)[0].state,'no-code');
});
test('do not fill a meal that has no recipe weight',()=>{
 const meals=[makeMeal('school',0)];
 assert.equal(autoApply(meals,[lot()],DAY),0);
 assert.equal(inspect(meals,[lot()],DAY)[0].state,'no-usage');
});
