import test from 'node:test';
import assert from 'node:assert/strict';
import {isRetiredModulePath} from '../src/retired-modules.mjs';
test('retired module links including client and nested routes return home',()=>{
 for(const path of ['/morning-meeting','/morning-meeting/materials','/filtering','/client/filtering','/materials','/mustang','/mustang/photos']) assert.equal(isRetiredModulePath(path),true,path);
});
test('all remaining module routes and similarly named paths stay available',()=>{
 for(const path of ['/','/client','/board','/client/board','/wip','/design-board','/client/design-board','/attendance','/holidays','/mileage','/order-panels','/van-estimator','/rams','/client/rams','/rams-logic','/social-post','/description-pull','/pro-forma','/client/pro-forma','/pro-forma-template','/installer','/igloo-admin','/notifications','/corebridge-explorer','/tv-installs','/credit-application','/materials-report']) assert.equal(isRetiredModulePath(path),false,path);
});
