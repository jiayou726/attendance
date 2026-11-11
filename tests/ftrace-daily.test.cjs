const {test} = require('node:test');
const assert = require('node:assert/strict');
const {inspect,autoApply,assign,distribute,candidatesFor,receiptChoices,inheritReceiptCodes,matchReceiving,setGroupIncluded} = require('../static/ftrace-daily.js');

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
 assert.equal(inspect(meals,[lot('K011','9','YESTERDAY','2026-10-13')],DAY)[0].state,'manual');
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

test('inherit platform code from same-day saved receipt internal ingredient ID without relying on name',()=>{
 const meals=[makeMeal('schoolA',5,''),makeMeal('schoolB',6,'WRONG-CODE')];
 const receipts=[{...lot('CORRECT','25','REAL-LOT'),material_key:'ingredient:1',name:'實際申報：CAS去骨雞胸'}];
 assert.equal(inheritReceiptCodes(meals,receipts,DAY),2);
 assert.deepEqual(meals.map(x=>x.materials[0].code),['CORRECT','CORRECT']);
 assert.equal(autoApply(meals,receipts,DAY),1);
 assert.deepEqual(meals.map(x=>x.materials[0].lots[0].kg),['5','6']);
});
test('two different codes for the same internal menu ingredient never auto guess',()=>{
 const meals=[makeMeal('school',2,'')];
 const receipts=[
  {...lot('CODE-A','10','LOT-A'),material_key:'ingredient:1'},
  {...lot('CODE-B','10','LOT-B'),material_key:'ingredient:1'}
 ];
 assert.equal(inheritReceiptCodes(meals,receipts,DAY),0);
 assert.equal(meals[0].materials[0].code,'');
});
test('do not override a manually saved manufacture lot when inheriting same-day code',()=>{
 const meals=[makeMeal('school',2,'OLD')];
 meals[0].materials[0].lots=[{report_date:DAY,seq:'01',spec_code:'OLD-A',batch:'MANUAL',kg:'2'}];
 assert.equal(inheritReceiptCodes(meals,[{...lot('NEW'),material_key:'ingredient:1'}],DAY),0);
 assert.equal(meals[0].materials[0].code,'OLD');
});
test('search can match a differently named official receiving material once for multiple schools',()=>{
 const meals=[makeMeal('schoolA',8,'MENU-CODE','菜單雞肉'),makeMeal('schoolB',12,'MENU-CODE','菜單雞肉')];
 const source={...lot('REAL-CODE','25','ACTUAL-LOT'),name:'CAS去骨雞胸肉'};
 const receipts=[source,{...lot('OTHER','10','LOT-OTHER'),name:'另一個食材'}];
 const choices=receiptChoices(receipts,DAY);
 assert.ok(choices.some(x=>x.name==='CAS去骨雞胸肉'));
 const group=inspect(meals,receipts,DAY)[0];
 assert.equal(group.state,'missing');
 assert.equal(matchReceiving(group,source,receipts,DAY),true);
 assert.deepEqual(meals.map(x=>x.materials[0].code),['REAL-CODE','REAL-CODE']);
 assert.deepEqual(meals.map(x=>x.materials[0].lots[0].batch),['ACTUAL-LOT','ACTUAL-LOT']);
 assert.deepEqual(meals.map(x=>x.materials[0].lots[0].kg),['8','12']);
 assert.equal(inspect(meals,receipts,DAY)[0].state,'linked');
});
test('manual remap uses multiple same-day batches when aggregate stock covers menu',()=>{
 const meals=[makeMeal('schoolA',9,'WRONG'),makeMeal('schoolB',6,'WRONG')];
 const first={...lot('RIGHT','10','FIRST'),name:'實際收貨品名'};
 const second={...lot('RIGHT','10','SECOND'),name:'實際收貨品名',seq:'02'};
 const group=inspect(meals,[first,second],DAY)[0];
 assert.equal(matchReceiving(group,first,[first,second],DAY),true);
 assert.deepEqual(meals[0].materials[0].lots.map(l=>l.kg),['9']);
 assert.deepEqual(meals[1].materials[0].lots.map(l=>l.kg),['1','5']);
 assert.equal(inspect(meals,[first,second],DAY)[0].state,'linked');
});
test('cannot map to another date, pending receipt or insufficient receipt',()=>{
 const source={...lot('REAL','10'),name:'真正收貨'};
 const pending={...source,ready:false};
 const meals=[makeMeal('school',11,'WRONG')];
 const original=JSON.stringify(meals);
 const group=inspect(meals,[source],DAY)[0];
 assert.equal(matchReceiving(group,source,[source],DAY),false);
 assert.equal(matchReceiving(group,pending,[pending],DAY),false);
 assert.equal(matchReceiving(group,{...source,report_date:'2026-10-15'},[source],DAY),false);
 assert.equal(JSON.stringify(meals),original);
});
test('name search excludes unresolved or conflicting duplicated receipt batches',()=>{
 const a={...lot('ACTUAL','10'),name:'實際品名'},b={...a,receipt_kg:'11'};
 const choices=receiptChoices([a,b],DAY);
 assert.equal(choices.length,0);
});


test('older receipts can be searched by real name and manually applied to all schools',()=>{
 const meals=[makeMeal('s1',5,'WRONG','菜單絞肉'),makeMeal('s2',7,'WRONG','菜單絞肉')];
 const historic={...lot('ACTUAL','20','OLDER-LOT','2026-10-13'),name:'CAS 豬絞肉'};
 const receipts=[historic,lot('OTHER','10')];
 assert.equal(autoApply(meals,receipts,DAY),0);
 const choices=receiptChoices(receipts,DAY);
 assert.ok(choices.includes(historic));
 assert.equal(matchReceiving(inspect(meals,receipts,DAY)[0],historic,receipts,DAY),true);
 assert.deepEqual(meals.map(m=>m.materials[0].lots[0].report_date),['2026-10-13','2026-10-13']);
 assert.equal(inspect(meals,receipts,DAY)[0].state,'linked-past');
});
test('same-day always wins automatic mapping even if old receipt is listed first',()=>{
 const meals=[makeMeal('s1',5)];
 const historic=lot('K011','50','OLD','2026-10-13');
 const current=lot('K011','50','TODAY',DAY);
 assert.equal(autoApply(meals,[historic,current],DAY),1);
 assert.equal(meals[0].materials[0].lots[0].batch,'TODAY');
});
test('future receipts cannot appear in manually searchable choices',()=>{
 const past=lot('K011','10','YESTERDAY','2026-10-13');
 const future=lot('K011','10','TOMORROW','2026-10-15');
 assert.deepEqual(receiptChoices([past,future],DAY),[past]);
 const meals=[makeMeal('s1',3)];
 const group=inspect(meals,[past],DAY)[0];
 assert.equal(matchReceiving(group,future,[past,future],DAY),false);
});
test('x excludes one ingredient in all schools but does not change receiving or menu',()=>{
 const schoolA={key:'A',include:true,materials:[
  {key:'rice',name:'白米',code:'R',estimated_kg:'5',lots:[{batch:'R1'}]},
  {key:'carrot',name:'紅蘿蔔',code:'C',estimated_kg:'2',lots:[{batch:'C1'}]}
 ]};
 const schoolB={key:'B',include:true,materials:[
  {key:'rice',name:'白米',code:'R',estimated_kg:'7',lots:[{batch:'R1'}]},
  {key:'carrot',name:'紅蘿蔔',code:'C',estimated_kg:'3',lots:[{batch:'C1'}]}
 ]};
 const meals=structuredClone([schoolA,schoolB]);
 const receipts=[{name:'白米',kg:'20'},{name:'紅蘿蔔',kg:'10'}];
 const before=JSON.stringify([schoolA,schoolB]);
 const beforeReceipt=JSON.stringify(receipts);
 const group=inspect(meals,[],DAY).find(g=>g.name==='白米');
 assert.equal(setGroupIncluded(meals,group,false),2);
 assert.deepEqual(meals.map(m=>m.materials[0].include),[false,false]);
 assert.deepEqual(meals.map(m=>m.materials[1].include),[undefined,undefined]);
 assert.equal(meals[0].include,true);
 assert.equal(inspect(meals,[],DAY).find(g=>g.name==='白米').state,'excluded');
 assert.equal(JSON.stringify([schoolA,schoolB]),before);
 assert.equal(JSON.stringify(receipts),beforeReceipt);
 const restore=inspect(meals,[],DAY).find(g=>g.name==='白米');
 assert.equal(setGroupIncluded(meals,restore,true),2);
 assert.deepEqual(meals.map(m=>m.materials[0].include),[true,true]);
 assert.equal(meals[0].materials[0].lots[0].batch,'R1');
});
test('x on last material disables meal, restore makes it visible again even after JSON save',()=>{
 const meals=[makeMeal('A',8),makeMeal('B',3)];
 const group=inspect(meals,[],DAY)[0];
 assert.equal(setGroupIncluded(meals,group,false),2);
 assert.ok(meals.every(m=>m.include===false && m._material_excluded_all));
 const saved=JSON.parse(JSON.stringify(meals));
 const excluded=inspect(saved,[],DAY)[0];
 assert.equal(excluded.state,'excluded');
 assert.equal(setGroupIncluded(saved,excluded,true),2);
 assert.ok(saved.every(m=>m.include===true && !m._material_excluded_all));
 assert.equal(inspect(saved,[],DAY)[0].needKg,'11');
});
test('excluded ingredient cannot be automatically filled or used for matching',()=>{
 const meals=[makeMeal('schoolA',3)];
 const receipts=[lot('K011','9')];
 setGroupIncluded(meals,inspect(meals,receipts,DAY)[0],false);
 assert.equal(autoApply(meals,receipts,DAY),0);
 assert.equal(inspect(meals,receipts,DAY)[0].state,'excluded');
});
