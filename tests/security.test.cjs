'use strict';
const {test, before} = require('node:test');
const assert = require('node:assert/strict');
const {authenticate} = require('../netlify/lib/auth.cjs');
const {handler,ownerKey,completionKey,known} = require('../netlify/lib/progress.cjs');
const {createGraph,base,pages} = require('../netlify/lib/graph.cjs');
const {SafeError} = require('../netlify/lib/errors.cjs');
const config = require('../netlify/lib/config.cjs');
const user = {tenant:config.tenant,oid:'11111111-1111-4111-8111-111111111111',email:'staff@example.test'};
const other = {...user,oid:'22222222-2222-4222-8222-222222222222'};
const course = 'zoho-crm-coach', lesson = 'lesson-1';
let jose, keys, attacker;
before(async () => { jose = await import('jose'); keys = await jose.generateKeyPair('RS256'); attacker = await jose.generateKeyPair('RS256'); });
function claims() { const now = Math.floor(Date.now()/1000); return {ver:'2.0',iss:`https://login.microsoftonline.com/${config.tenant}/v2.0`,aud:config.client,tid:config.tenant,oid:user.oid,azp:config.client,scp:'Training.Progress',acct:0,iat:now,nbf:now-5,exp:now+3600,preferred_username:user.email}; }
async function sign(changes = {}, key) {
  const p = {...claims(),...changes};
  for (const k of Object.keys(p)) if (p[k] === undefined) delete p[k];
  return new jose.SignJWT(p).setProtectedHeader({alg:'RS256',typ:'JWT'}).sign(key || keys.privateKey);
}
const authenticateTest = async token => authenticate(token,{key:keys.publicKey});
test('valid signed member API access token produces immutable learner identity', async () => {
  assert.deepEqual(await authenticateTest('Bearer '+await sign()), user);
});
test('explicit string-valued member claim is accepted; no truthy/falsy coercion', async () => {
  assert.equal((await authenticateTest('Bearer '+await sign({acct:'0'}))).oid,user.oid);
});
for (const [name,changes] of Object.entries({
  'wrong issuer':{iss:'https://attacker.test/'},
  'common issuer':{iss:'https://login.microsoftonline.com/common/v2.0'},
  'wrong audience':{aud:'another-api'},
  'Graph audience':{aud:'00000003-0000-0000-c000-000000000000'},
  'multiple audiences':{aud:[config.client,'another-api']},
  'wrong client':{azp:'00000000-0000-0000-0000-000000000000'},
  'missing client':{azp:undefined},
  'wrong scope':{scp:'User.Read Training.Progress.Admin'},
  'ID token (no API scope)':{scp:undefined,nonce:'id-token-nonce'},
  'app token':{scp:undefined,idtyp:'app',roles:['Training.Progress']},
  'wrong tenant':{tid:other.oid},
  'missing oid':{oid:undefined},
  'invalid oid':{oid:'somebody@example.test'},
  'guest':{acct:1}, 'guest string':{acct:'1'},
  'missing membership':{acct:undefined}, 'boolean membership':{acct:false},
  'v1 token':{ver:'1.0',iss:`https://sts.windows.net/${config.tenant}/`,aud:`api://${config.client}`,appid:config.client},
  'expired':{exp:Math.floor(Date.now()/1000)-60},
  'not yet valid':{nbf:Math.floor(Date.now()/1000)+300},
  'future issued':{iat:Math.floor(Date.now()/1000)+300},
  'missing expiry':{exp:undefined},
})) test(`rejects ${name}`,async () => {
  await assert.rejects(authenticateTest('Bearer '+await sign(changes)), error => error.status===401 && error.message==='sign_in_required');
});
test('rejects forged signature and unsigned/malformed bearer tokens',async () => {
  await assert.rejects(authenticateTest('Bearer '+await sign({},attacker.privateKey)), {status:401});
  for (const token of [undefined,'Basic abc','Bearer abc','Bearer '+Buffer.from(JSON.stringify({alg:'none'})).toString('base64url')+'.e30.','Bearer '+ 'a'.repeat(17000)]) {
    await assert.rejects(authenticateTest(token), {status:401});
  }
});
function schema() {
  return {value:[
    {name:'Title',text:{maxLength:255},indexed:true},
    {name:'CompletionKey',text:{maxLength:64},indexed:true,enforceUniqueValues:true},
    ...['TrainingId','LessonId','PersonEmail'].map(name => ({name,text:{maxLength:255}})),
    {name:'CompletedOn',dateTime:{format:'dateTime'}},
  ]};
}
function row(who = user, trainingId = course, lessonId = lesson) {
  return {id:'17',fields:{Title:ownerKey(who),CompletionKey:completionKey(who,trainingId,lessonId),TrainingId:trainingId,LessonId:lessonId,CompletedOn:'2026-10-16T12:00:00Z'}};
}
function event(method = 'POST', body = {trainingId:course,lessonId:lesson}) {
  return {httpMethod:method,headers:{Authorization:'Bearer test','Content-Type':'application/json',Origin:config.origin},body:JSON.stringify(body)};
}
const run = (request, auth = async () => user) => handler({request,auth,enabled:()=>true});
const json = response => JSON.parse(response.body);
test('missing authentication rejected before any Graph request',async () => {
  const e = event(); delete e.headers.Authorization;
  const h = run(() => assert.fail('Graph called'), authenticateTest);
  assert.equal((await h(e)).statusCode,401);
});
test('forbidden methods and cross-origin requests cannot reach Graph',async () => {
  const h = run(() => assert.fail('Graph called'));
  for (const method of ['PUT','DELETE','PATCH','OPTIONS']) assert.equal((await h(event(method))).statusCode,405);
  const e = event(); e.headers.Origin = 'https://attacker.test';
  assert.equal((await h(e)).statusCode,403);
});
test('tracking remains disabled until explicitly configured',async () => {
  const h = handler({request:()=>assert.fail('Graph called'),enabled:()=>false});
  assert.equal(json(await h(event())).error,'tracking_configuration_required');
});
for (const [name,body] of Object.entries({
  'unknown course':{trainingId:'unknown',lessonId:lesson},
  'unknown lesson':{trainingId:course,lessonId:'unknown'},
  'path traversal':{trainingId:'../zoho-crm-coach',lessonId:lesson},
  'forged email':{trainingId:course,lessonId:lesson,personEmail:other.email},
  'forged oid':{trainingId:course,lessonId:lesson,oid:other.oid},
  'forged identity':{trainingId:course,lessonId:lesson,Title:ownerKey(other)},
  'forged completion key':{trainingId:course,lessonId:lesson,CompletionKey:'0'.repeat(64)},
  'self-reported score':{trainingId:course,lessonId:lesson,score:100},
  'array':[], 'null':null,
})) test(`invalid input rejected: ${name}`, async () => {
  assert.equal((await run(()=>assert.fail('Graph called'))(event('POST',body))).statusCode,400);
});
test('bounded body, malformed JSON, base64, wrong content type, and legacy query rejected', async () => {
  const h = run(()=>assert.fail('Graph called'));
  const events = [];
  let e = event(); e.body = 'x'.repeat(2049); events.push(e);
  e = event(); e.body = '{'; events.push(e);
  e = event(); e.isBase64Encoded = true; events.push(e);
  e = event(); e.headers['Content-Type'] = 'text/plain'; events.push(e);
  e = event('GET'); e.queryStringParameters = {personEmail:'another@example.test'}; events.push(e);
  for (const input of events) assert.equal((await h(input)).statusCode,400);
});
test('pagination filters every result by tenant+oid, key and known lesson',async () => {
  const mine = row();
  const mixed = [mine,row(other),{...row(),fields:{...mine.fields,CompletionKey:'forged'}},row(user,'unknown','lesson'),{id:'legacy',fields:{PersonEmail:user.email,TrainingId:course,LessonId:lesson}}];
  let calls = 0;
  const h = run(async url => {
    calls++;
    if (calls===1) { assert.match(decodeURIComponent(url),new RegExp(ownerKey(user))); return {value:mixed,'@odata.nextLink':base+'/items?$skiptoken=next'}; }
    return {value:[row(other)]};
  });
  const r = await h(event('GET'));
  assert.equal(r.statusCode,200); assert.equal(calls,2);
  assert.deepEqual(json(r),{completions:[{trainingId:course,lessonId:lesson,completedOn:mine.fields.CompletedOn}]});
  assert.equal(r.headers['Cache-Control'],'no-store');
  assert.ok(!r.body.includes(user.email)); assert.ok(!r.body.includes(user.oid));
});
for (const name of ['missing-key','not-unique','not-indexed','wrong-length','title-not-indexed']) test(`schema fails closed: ${name}`,async () => {
  const s = schema();
  const key = s.value.find(c=>c.name==='CompletionKey');
  if (name==='missing-key') s.value=s.value.filter(c=>c!==key);
  if (name==='not-unique') key.enforceUniqueValues=false;
  if (name==='not-indexed') key.indexed=false;
  if (name==='wrong-length') key.text.maxLength=255;
  if (name==='title-not-indexed') s.value.find(c=>c.name==='Title').indexed=false;
  const r = await run(async (url,method) => { assert.equal(method,undefined); assert.equal(url,base+'/columns'); return s; })(event());
  assert.equal(r.statusCode,503); assert.equal(json(r).error,'tracking_schema_required');
});
test('write uses server identity/time, omits Person and Score, acknowledges Graph success only',async () => {
  let written;
  const h = run(async (url,method,body) => {
    if (url.endsWith('/columns')) return schema();
    if (method==='POST') {written=body.fields;return {id:'20'};}
    return {value:[]};
  });
  const r = await h(event());
  assert.equal(r.statusCode,200); assert.equal(json(r).ok,true);
  assert.equal(written.Title,ownerKey(user)); assert.equal(written.CompletionKey,completionKey(user,course,lesson));
  assert.equal(written.PersonEmail,user.email); assert.ok(Date.parse(written.CompletedOn));
  assert.ok(!('Score' in written)); assert.ok(!('Person' in written));
});
test('sequential retry does not create another item',async () => {
  const h = run(async (url,method) => {
    if (url.endsWith('/columns')) return schema();
    assert.notEqual(method,'POST'); return {value:[row()]};
  });
  assert.equal(json(await h(event())).duplicate,true);
});
test('unique-key conflict rereads and acknowledges only exact same-owner row',async () => {
  let reads=0;
  const h = run(async (url,method) => {
    if (url.endsWith('/columns')) return schema();
    if (method==='POST') throw new SafeError(503,'possible_duplicate');
    return {value:++reads===1 ? [] : [row(other),row()]};
  });
  assert.equal(json(await h(event())).duplicate,true);
});
test('Graph write conflict/failure cannot acknowledge another learner row',async () => {
  const h = run(async (url,method) => {
    if (url.endsWith('/columns')) return schema();
    if (method==='POST') throw new Error('sensitive upstream details bearer secret');
    return {value:[row(other)]};
  });
  const r = await h(event());
  assert.equal(r.statusCode,503); assert.deepEqual(json(r),{ok:false,error:'tracking_unavailable'});
  assert.ok(!r.body.includes('secret'));
});
test('Graph read errors are sanitized and cannot imply an empty successful history',async () => {
  const r = await run(async () => {throw new Error('secret access_token');})(event('GET'));
  assert.equal(r.statusCode,503); assert.equal(json(r).error,'tracking_unavailable');
});
test('Graph token never follows off-list or off-host pagination or redirect',async () => {
  let calls=0;
  const graph = createGraph({fetcher:async (_url,opts) => {calls++;assert.equal(opts.redirect,'error');return {ok:true,json:async()=>({value:[]})};},token:async()=> 'not-a-real-token'});
  for (const url of ['https://attacker.test/items',base.replace(config.list,user.oid)+'/items',base+'/../other/items']) await assert.rejects(graph(url));
  assert.equal(calls,0);
  await graph(base+'/items');assert.equal(calls,1);
});
test('Graph HTTP errors and network failures are bounded sanitized failures',async () => {
  for (const fetcher of [async()=>({ok:false,status:403}),async()=>({ok:false,status:429}),async()=>{throw new Error('bearer secret');}]) {
    const graph=createGraph({fetcher,token:async()=> 'test'});
    await assert.rejects(graph(base+'/items'),{status:503,message:'tracking_unavailable'});
  }
});
test('pagination rejects loops, malformed pages, and pagination failures',async () => {
  const url=base+'/items';
  await assert.rejects(pages(async()=>({value:[],'@odata.nextLink':url}),url),{status:503});
  await assert.rejects(pages(async()=>({value:{}}),url),{status:503});
});
test('catalog allowlist includes all 7 courses / 44 ready lessons unchanged',() => {
  const catalog=require('../data/trainings-index.json');
  assert.equal(catalog.trainings.length,7);
  let count=0;
  for (const t of catalog.trainings) for (const l of t.lessons) {assert.ok(known(t.id,l.id));count++;}
  assert.equal(count,44);
});
test('identity/key canonicalization is deterministic and separates learners/lessons',() => {
  assert.match(completionKey(user,course,lesson),/^[0-9a-f]{64}$/);
  assert.equal(completionKey(user,course,lesson),completionKey({...user,email:'changed@example.test'},course,lesson));
  assert.notEqual(completionKey(user,course,lesson),completionKey(other,course,lesson));
  assert.notEqual(completionKey(user,course,lesson),completionKey(user,course,'lesson-2'));
});
