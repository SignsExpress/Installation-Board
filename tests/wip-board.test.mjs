import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import fs from 'node:fs';
const source=fs.readFileSync('src/ModernApp.jsx','utf8');
const extract=name=>{const start=source.indexOf('function '+name+'(');const end=source.indexOf('\n}',start);return source.slice(start,end+2);};
test('WIP starts today with ten working days, skips weekends and crosses year boundaries',()=>{
 const ctx=vm.createContext({Intl,Date,parseIsoDate:s=>new Date(s+'T00:00:00Z'),toIsoDate:d=>d.toISOString().slice(0,10),addDays:(d,n)=>new Date(d.getTime()+n*86400000)});
 vm.runInContext(extract('getWipBoardDays'),ctx);
 for(const [start,first,last] of [['2026-10-08','2026-10-08','2026-10-21'],['2026-10-10','2026-10-12','2026-10-23'],['2026-12-28','2026-12-28','2027-01-08']]){
  const days=ctx.getWipBoardDays(start);assert.equal(days.length,10);assert.equal(days[0].id,first);assert.equal(days[9].id,last);assert.ok(days.every(d=>![0,6].includes(new Date(d.id).getUTCDay())));
 }
});
test('dropping a multi-day job into any status preserves its details and leaves other jobs alone',()=>{
 let cards=[{id:'one',lane:'day:2026-10-08',extraLanes:['day:2026-10-09'],tab:'wip',company:'Customer',productionAssignees:['MC'],preTaxTotal:700},{id:'two',lane:'backlog',tab:'pre-wip'}];
 const start=source.indexOf('  function moveCard(cardId, lane, tab = activeTab)');const end=source.indexOf('  function toggleWipAssignee',start);
 const ctx=vm.createContext({activeTab:'wip',setCards:update=>{cards=update(cards);}});vm.runInContext(source.slice(start,end),ctx);
 for(const lane of ['completed','ordered','awaiting-installation','salesperson','hold','day:2026-10-21','backlog']){
  ctx.moveCard('one',lane);assert.equal(cards[0].lane,lane);assert.equal(cards[0].company,'Customer');assert.equal(cards[0].productionAssignees[0],'MC');assert.equal(cards[0].preTaxTotal,700);assert.equal(cards[0].extraLanes.length,0);assert.equal(cards[1].lane,'backlog');assert.equal(cards[1].tab,'pre-wip');
 }
});
