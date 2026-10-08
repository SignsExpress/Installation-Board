import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import {resolveSkin,skinLink} from '../src/portal-skin.mjs';
test('classic is default; explicit links override a remembered selection',()=>{
 assert.equal(resolveSkin('',null),'classic');
 assert.equal(resolveSkin('','new'),'new');
 assert.equal(resolveSkin('?skin=classic','new'),'classic');
 assert.equal(resolveSkin('?skin=new','classic'),'new');
 assert.equal(resolveSkin('?skin=unknown','classic'),'classic');
});
test('switch links keep the current module and other parameters',()=>{
 assert.equal(skinLink('https://www.sxpreston.com/wip?tab=pre#job','new'),'https://www.sxpreston.com/wip?tab=pre&skin=new#job');
 assert.equal(skinLink('https://www.sxpreston.com/?skin=new','classic'),'https://www.sxpreston.com/?skin=classic');
});
test('the classic application and new skin use identical PDF output functions',()=>{
 const old=fs.readFileSync('src/App.jsx','utf8').replaceAll('\r\n','\n');
 const modern=fs.readFileSync('src/ModernApp.jsx','utf8').replaceAll('\r\n','\n');
 for(const [start,end] of [['function buildProFormaPreviewHtml(','\nfunction '],['  async function openPrintPreview()','\n  function ']]){
  const part=s=>s.slice(s.indexOf(start),s.indexOf(end,s.indexOf(start)));
  assert.ok(part(old).length>500);assert.equal(part(modern),part(old));
 }
});
