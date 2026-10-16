'use strict';
const config = require('./config.cjs');
const { SafeError } = require('./errors.cjs');
const UUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
let remoteKeys;
async function authenticate(authorization, options = {}) {
  try {
    if (typeof authorization !== 'string' || authorization.length > 16384 || !/^Bearer [A-Za-z0-9_-]+\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+$/i.test(authorization)) throw new Error('bearer');
    const token = authorization.slice(7);
    const jose = await import('jose');
    // Entra requestedAccessTokenVersion=2 is an explicit setup prerequisite.
    // Reject v1 rather than guessing resource/audience aliases.
    const issuer = `https://login.microsoftonline.com/${config.tenant}/v2.0`;
    remoteKeys ||= jose.createRemoteJWKSet(new URL(`https://login.microsoftonline.com/${config.tenant}/discovery/v2.0/keys`), {timeoutDuration:5000, cooldownDuration:30000});
    const {payload: p} = await jose.jwtVerify(token, options.key || remoteKeys, {
      algorithms:['RS256'], issuer,
      audience:config.client,
      requiredClaims:['exp','nbf','iat','iss','aud','tid','oid','scp','ver','acct'],
      clockTolerance:30,
    });
    if (p.ver !== '2.0' || p.aud !== config.client || p.tid !== config.tenant || typeof p.oid !== 'string' || !UUID.test(p.oid) ||
        p.azp !== config.client ||
        typeof p.scp !== 'string' || !p.scp.split(' ').includes('Training.Progress') ||
        ![0,'0'].includes(p.acct) || p.idtyp === 'app' || typeof p.iat !== 'number' ||
        p.iat > Date.now()/1000 + 30 || typeof p.nbf !== 'number' || p.exp <= p.nbf) throw new Error('claims');
    // Email is optional display metadata; it NEVER determines ownership.
    const email = p.preferred_username || p.upn;
    return Object.freeze({tenant:config.tenant, oid:p.oid.toLowerCase(),
      email: typeof email === 'string' && email.length <= 255 && /^[^\s@]+@[^\s@]+$/.test(email) ? email : null});
  } catch (_) {
    // Never expose token, claims, library exceptions, or signing-key details.
    throw new SafeError(401, 'sign_in_required');
  }
}
module.exports = { authenticate };
