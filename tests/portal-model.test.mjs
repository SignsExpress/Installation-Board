import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import {portalWeekDays,filterPortalWip,isPortalWrite,PORTAL_GROUPS} from '../src/portal-model.mjs';

test('production weeks cover Monday to Friday across year and Sunday boundaries',()=>{
  assert.deepEqual(portalWeekDays('2026-01-01').map(d=>d.id),['2025-12-29','2025-12-30','2025-12-31','2026-01-01','2026-01-02']);
  assert.equal(portalWeekDays('2026-10-11')[0].id,'2026-10-05');
  assert.equal(portalWeekDays('2026-10-08',1)[0].id,'2026-10-12');
  assert.equal(portalWeekDays('2026-10-08',-1)[0].id,'2026-09-28');
});
test('search finds references, customers and salespeople without mixing tabs or changing saved cards',()=>{
  const cards=[{id:'a',tab:'wip',orderNumber:'ORD-123',company:'Example Ltd',salesperson:'Amber'},{id:'b',tab:'pre-wip',orderNumber:'ORD-124',company:'Example Ltd'}];
  const original=structuredClone(cards);
  assert.deepEqual(filterPortalWip(cards,'wip',' EXAMPLE ').map(c=>c.id),['a']);
  assert.deepEqual(filterPortalWip(cards,'wip','amber').map(c=>c.id),['a']);
  assert.equal(filterPortalWip(cards,'wip','no-match').length,0);
  assert.deepEqual(cards,original);
});
test('feedback recognises saved changes, not sign-ins, source pulls or read requests',()=>{
  assert.equal(isPortalWrite('/api/wip-board',{method:'PUT'}),true);
  assert.equal(isPortalWrite('/api/jobs/abc',{method:'DELETE'}),true);
  assert.equal(isPortalWrite('/api/attendance',{method:'POST'}),true);
  assert.equal(isPortalWrite('/api/jobs'),false);
  assert.equal(isPortalWrite('/api/auth/login',{method:'POST'}),false);
  assert.equal(isPortalWrite('/api/pro-forma/pull',{method:'POST'}),false);
  assert.equal(isPortalWrite('/api/wip-board/enrich',{method:'POST'}),false);
});
test('navigation groups have no duplicate destinations or removed modules',()=>{
  const keys=PORTAL_GROUPS.flatMap(g=>g.keys);
  assert.equal(keys.length,new Set(keys).size);
  for(const removed of ['filtering','materials','mustang','morning-meeting'])assert.equal(keys.includes(removed),false);
});
test('moving between visible weeks preserves all saved dated placements, including multiple days',()=>{
  const source=fs.readFileSync(new URL('../src/App.jsx',import.meta.url),'utf8').replace(/\r/g,'');
  const getFunction=name=>{const start=source.indexOf('function '+name+'(');return source.slice(start,source.indexOf('\nfunction ',start+1));};
  const context=vm.createContext({getWipVisibleLaneIds:days=>new Set(['backlog',...days.map(d=>'day:'+d.id)])});
  vm.runInContext(getFunction('keepWipCardInVisibleLane')+'\n'+getFunction('keepWipCardsInVisibleLanes'),context);
  const cards=[{id:'future',lane:'day:2026-12-15',extraLanes:['day:2026-12-16']},{id:'past',lane:'day:2026-09-30',extraLanes:[]}];
  const result=context.keepWipCardsInVisibleLanes(cards,portalWeekDays('2026-10-08'));
  assert.equal(result[0].lane,'day:2026-12-15');assert.equal(result[0].extraLanes[0],'day:2026-12-16');assert.equal(result[1].lane,'day:2026-09-30');
});
