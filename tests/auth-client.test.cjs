'use strict';
const {test} = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const fs = require('node:fs');
const path = require('node:path');
const config = require('../netlify/lib/config.cjs');
function harness({popupError,silentError,missingLibrary=false} = {}) {
  const calls=[]; let settings;
  class Client {
    constructor(c) {settings=c;}
    async initialize() {calls.push('initialize');}
    async loginPopup(request) {calls.push({popup:request});if(popupError)throw new Error('popup blocked');return {account:{homeAccountId:'member-id'}};}
    setActiveAccount(account) {calls.push({active:account});}
    async acquireTokenSilent(request) {calls.push({silent:request});if(silentError)throw new Error('interaction required');return {accessToken:'custom-api-token'};}
    async clearCache() {calls.push('clearCache');}
  }
  const msal={PublicClientApplication:Client};
  const context=vm.createContext({window:missingLibrary?{}:{msal},msal,URL,location:{origin:config.origin}});
  vm.runInContext(fs.readFileSync(path.join(__dirname,'../js/auth.js'),'utf8'),context);
  return {api:context.window.trainingAuth,calls,settings};
}
test('initialization uses tenant authority, memory cache, callback and never redirects iframe',async () => {
  const h=harness(); await new Promise(resolve=>setImmediate(resolve));
  assert.equal(h.settings.auth.authority,`https://login.microsoftonline.com/${config.tenant}`);
  assert.equal(h.settings.auth.redirectUri,config.origin+'/auth.html');
  assert.equal(h.settings.cache.cacheLocation,'memoryStorage');
  assert.equal(h.settings.system.allowRedirectInIframe,false);
  assert.equal(h.calls.length,1);
});
test('only explicit connect opens popup; token uses custom API scope, never Graph/ID token',async () => {
  const h=harness();
  assert.equal((await h.api.connect()).id,'member-id');
  const request=h.calls.find(x=>x.popup).popup;
  assert.equal(request.scopes[0],`api://${config.client}/Training.Progress`);
  assert.equal(request.prompt,'select_account');
  assert.equal(await h.api.token(),'custom-api-token');
  assert.equal(h.calls.filter(x=>x.popup).length,1);
});
test('silent interaction-required failure does not trigger unsolicited popup',async () => {
  const h=harness({silentError:true});await h.api.connect();
  await assert.rejects(h.api.token());
  assert.equal(h.calls.filter(x=>x.popup).length,1);
});
test('disconnect clears training account/cache; no token without reconnect',async () => {
  const h=harness();await h.api.connect();await h.api.disconnect();
  assert.ok(h.calls.includes('clearCache'));
  await assert.rejects(h.api.token());
});
test('blocked popup and missing pinned library remain retryable failures',async () => {
  const h=harness({popupError:true});
  await assert.rejects(h.api.connect());await assert.rejects(h.api.connect());
  assert.equal(h.calls.filter(x=>x.popup).length,2);
  await assert.rejects(harness({missingLibrary:true}).api.connect());
});
