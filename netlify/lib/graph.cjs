'use strict';
const config = require('./config.cjs');
const { SafeError } = require('./errors.cjs');
const base = `https://graph.microsoft.com/v1.0/sites/${config.site}/lists/${config.list}`;
let cachedToken;
let pendingToken;
async function graphToken(fetcher = fetch) {
  if (cachedToken && cachedToken.until > Date.now() + 60000) return cachedToken.value;
  if (!process.env.TRAINING_GRAPH_CLIENT_SECRET) throw new SafeError(503, 'tracking_configuration_required');
  pendingToken ||= (async () => {
    try {
      const r = await fetcher(`https://login.microsoftonline.com/${config.tenant}/oauth2/v2.0/token`, {
        method:'POST', headers:{'Content-Type':'application/x-www-form-urlencoded'},
        body:new URLSearchParams({client_id:config.client, client_secret:process.env.TRAINING_GRAPH_CLIENT_SECRET, scope:'https://graph.microsoft.com/.default', grant_type:'client_credentials'}),
        signal:AbortSignal.timeout(8000), redirect:'error',
      });
      if (!r.ok) throw new Error('token');
      const data = await r.json();
      if (typeof data.access_token !== 'string' || !Number.isFinite(data.expires_in) || data.expires_in <= 60) throw new Error('token response');
      cachedToken = {value:data.access_token,until:Date.now()+Math.min(data.expires_in,3600)*1000};
      return cachedToken.value;
    } catch (_) { throw new SafeError(503, 'tracking_unavailable'); }
  })();
  try { return await pendingToken; } finally { pendingToken = undefined; }
}
function createGraph({fetcher = fetch, token = () => graphToken(fetcher)} = {}) {
  return async function request(url, method = 'GET', body) {
    // Graph pagination may never send the application token to another host/list.
    const parsed = new URL(url);
    if (!url.startsWith(base + '/') || parsed.origin !== 'https://graph.microsoft.com' || !parsed.pathname.startsWith(new URL(base).pathname + '/') || parsed.username || parsed.password) throw new SafeError(503,'tracking_unavailable');
    try {
      const response = await fetcher(url, {
        method, headers:{Authorization:`Bearer ${await token()}`, ...(body ? {'Content-Type':'application/json'} : {})},
        ...(body ? {body:JSON.stringify(body)} : {}), signal:AbortSignal.timeout(10000), redirect:'error',
      });
      if (!response.ok) {
        // Concurrent same-key creation can return 400 or 409; caller must reread
        // and confirm the exact row before acknowledging success.
        throw new SafeError(503, method === 'POST' && [400,409].includes(response.status) ? 'possible_duplicate' : 'tracking_unavailable');
      }
      return await response.json();
    } catch (error) {
      if (error instanceof SafeError) throw error;
      throw new SafeError(503,'tracking_unavailable');
    }
  };
}
async function pages(request, url) {
  const rows = [];
  const seen = new Set();
  while (url) {
    if (seen.has(url) || seen.size >= 100) throw new SafeError(503, 'tracking_unavailable');
    seen.add(url);
    const page = await request(url);
    if (!Array.isArray(page.value) || page.value.length > 1000) throw new SafeError(503,'tracking_unavailable');
    rows.push(...page.value);
    url = page['@odata.nextLink'];
    if (url !== undefined && typeof url !== 'string') throw new SafeError(503,'tracking_unavailable');
  }
  return rows;
}
module.exports = {base, createGraph, pages};
