'use strict';
const {createHash} = require('node:crypto');
const catalog = require('../../data/trainings-index.json');
const {base, pages, createGraph} = require('./graph.cjs');
const {authenticate} = require('./auth.cjs');
const {SafeError} = require('./errors.cjs');
const config = require('./config.cjs');
const allowed = new Set(catalog.trainings.flatMap(t => t.lessons.filter(l => l.status === 'ready').map(l => `${t.id}:${l.id}`)));
const ownerKey = user => `oa-lms:v1:${user.tenant}:${user.oid}`;
const completionKey = (user, trainingId, lessonId) => createHash('sha256').update(JSON.stringify([user.tenant,user.oid,trainingId,lessonId])).digest('hex');
const known = (trainingId, lessonId) => typeof trainingId === 'string' && typeof lessonId === 'string' && trainingId.length <= 100 && lessonId.length <= 100 && allowed.has(`${trainingId}:${lessonId}`);
async function assertSchema(request) {
  const columns = await pages(request, `${base}/columns`);
  const byName = Object.fromEntries(columns.map(c => [c.name,c]));
  const text = (name, max) => byName[name]?.text && byName[name].text.maxLength >= max;
  const ready = text('Title', 83) && byName.Title.indexed === true &&
    text('TrainingId',100) && text('LessonId',100) && text('PersonEmail',255) &&
    text('CompletionKey',64) && byName.CompletionKey.text.maxLength === 64 &&
    byName.CompletionKey.indexed === true && byName.CompletionKey.enforceUniqueValues === true &&
    byName.CompletedOn?.dateTime;
  if (!ready) throw new SafeError(503,'tracking_schema_required');
}
async function ownRows(request, user) {
  // Title has only fixed text and validated UUIDs: no browser identity or OData fragments.
  const url = `${base}/items?$expand=fields($select=Title,CompletionKey,TrainingId,LessonId,CompletedOn)&$filter=${encodeURIComponent(`fields/Title eq '${ownerKey(user)}'`)}&$top=200`;
  const rows = await pages(request,url);
  // Defense in depth even if Graph ignores a filter/malformed older rows exist.
  return rows.filter(row => row.fields?.Title === ownerKey(user) &&
    known(row.fields.TrainingId,row.fields.LessonId) &&
    row.fields.CompletionKey === completionKey(user,row.fields.TrainingId,row.fields.LessonId));
}
function handler({request = createGraph(), auth = authenticate, enabled = () => process.env.TRAINING_TRACKING_ENABLED === 'true'} = {}) {
  return async event => {
    const headers = {'Content-Type':'application/json', 'Cache-Control':'no-store', 'X-Content-Type-Options':'nosniff'};
    const respond = (statusCode,data) => ({statusCode,headers,body:JSON.stringify(data)});
    try {
      if (!['GET','POST'].includes(event.httpMethod)) { headers.Allow = 'GET, POST'; throw new SafeError(405,'method_not_allowed'); }
      if (!enabled()) throw new SafeError(503,'tracking_configuration_required');
      const h = Object.fromEntries(Object.entries(event.headers || {}).map(([k,v]) => [k.toLowerCase(),v]));
      if (h.origin && h.origin !== config.origin) throw new SafeError(403,'origin_not_allowed');
      const user = await auth(h.authorization);
      // Legacy email/identity query strings are rejected, never used as selectors.
      if (Object.keys(event.queryStringParameters || {}).length) throw new SafeError(400,'invalid_request');
      if (event.httpMethod === 'GET') {
        const rows = await ownRows(request,user);
        return respond(200,{completions:rows.map(({fields:f}) => ({trainingId:f.TrainingId,lessonId:f.LessonId,completedOn:f.CompletedOn}))});
      }
      if (!/^application\/json(?:\s*;.*)?$/i.test(h['content-type'] || '') || event.isBase64Encoded || typeof event.body !== 'string' || Buffer.byteLength(event.body) > 2048) throw new SafeError(400,'invalid_request');
      let input;
      try { input = JSON.parse(event.body); } catch (_) { throw new SafeError(400,'invalid_request'); }
      if (!input || Array.isArray(input) || Object.keys(input).some(k => !['trainingId','lessonId'].includes(k)) || !known(input.trainingId,input.lessonId)) throw new SafeError(400,'invalid_lesson_or_request');
      // Self-reported score is deliberately NOT accepted or written as assessment evidence.
      await assertSchema(request);
      const {trainingId,lessonId} = input;
      const key = completionKey(user,trainingId,lessonId);
      const existing = await ownRows(request,user);
      const found = existing.find(r => r.fields.CompletionKey === key);
      if (found) return respond(200,{ok:true,id:String(found.id),duplicate:true});
      const fields = {Title:ownerKey(user),CompletionKey:key,TrainingId:trainingId,LessonId:lessonId,CompletedOn:new Date().toISOString()};
      if (user.email) fields.PersonEmail = user.email;
      try {
        const item = await request(`${base}/items`,'POST',{fields});
        if (!item.id) throw new SafeError(503,'tracking_unavailable');
        return respond(200,{ok:true,id:String(item.id)});
      } catch (error) {
        // Also handles an uncertain network outcome after Graph committed the write.
        const saved = (await ownRows(request,user)).find(r => r.fields.CompletionKey === key);
        if (saved) return respond(200,{ok:true,id:String(saved.id),duplicate:true});
        throw new SafeError(503,'tracking_unavailable');
      }
    } catch (error) {
      return respond(error instanceof SafeError ? error.status : 503,{ok:false,error:error instanceof SafeError && error.code !== 'possible_duplicate' ? error.code : 'tracking_unavailable'});
    }
  };
}
module.exports = {handler,ownerKey,completionKey,assertSchema,ownRows,known};
