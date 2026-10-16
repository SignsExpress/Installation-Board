import test from 'node:test';
import assert from 'node:assert/strict';
import {calculateIrelandDelivery,IRELAND_DEFAULTS} from '../src/ireland/calculator.mjs';
const base={...IRELAND_DEFAULTS,goods:280,delivery:102.37,insurance:0};
test('Ireland estimate exactly matches supplied workbook cached EUR and GBP outputs',()=>{
 const r=calculateIrelandDelivery(base);
 assert.deepEqual(r.eur,{customs:444.87,duty:12.01,vat:105.08,processing:14.5,additional:131.59,deliveryTotal:250.69,landed:576.46});
 assert.deepEqual(r.gbp,{customs:382.37,duty:10.32,vat:90.32,processing:12.46,additional:113.1,deliveryTotal:215.47,landed:495.47});
});
test('DHL thresholds are applied to rounded customs value with inclusive small-fee upper bound',()=>{
 for(const [goods,expected] of [[0,0],[4.99,0],[4.999,5],[5,5],[20,5],[20.01,14.5],[580,14.5],[1000,25]]){
  const r=calculateIrelandDelivery({...base,goods,delivery:0,insurance:0,exchange:1});
  assert.equal(r.eur.processing,expected,'Customs value '+goods);
 }
});
test('insurance, zero duty and edited rates flow through the same model',()=>{
 const r=calculateIrelandDelivery({...base,goods:100,delivery:20,insurance:5,exchange:2,duty:0,vat:10,processing:5,minimum:1});
 assert.deepEqual(r.eur,{customs:250,duty:0,vat:25,processing:12.5,additional:37.5,deliveryTotal:87.5,landed:287.5});
 assert.equal(r.gbp.additional,18.75);
});
test('missing, negative and invalid inputs do not display misleading totals; explicit zero works',()=>{
 for(const override of [{goods:''},{delivery:null},{insurance:-1},{duty:'bad'},{exchange:0},{exchange:Infinity},{lower:21,upper:20}]){
  const r=calculateIrelandDelivery({...base,...override});assert.ok(r.error);assert.equal(r.gbp,undefined);
 }
 assert.equal(calculateIrelandDelivery({...base,goods:0,delivery:0,insurance:0}).gbp.additional,0);
});
