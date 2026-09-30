/**
 * Shared helper: authenticate to Microsoft Graph as the Azure AD app
 * registration (client-credentials flow, app-only — no signed-in user),
 * and resolve the SharePoint site ID once per invocation.
 *
 * Required environment variables (set in Netlify site settings):
 *   AZURE_TENANT_ID      — Azure AD tenant ID (or the .onmicrosoft.com domain)
 *   AZURE_CLIENT_ID      — App registration's Application (client) ID
 *   AZURE_CLIENT_SECRET  — A client secret generated for that app registration
 *   SP_HOSTNAME          — e.g. "oneaccord.sharepoint.com"
 *   SP_SITE_PATH         — e.g. "/sites/Ops"
 *   SP_TRAININGS_LIST_ID       — GUID of the Trainings list
 *   SP_COMPLETIONS_LIST_ID     — GUID of the Training Completions list
 *
 * The app registration needs the Microsoft Graph **Application** permission
 * Sites.Selected (preferred, scoped to just this site) or Sites.ReadWrite.All,
 * with admin consent granted. See README.md for the full setup walkthrough.
 */

let cachedToken = null; // { accessToken, expiresAt } — reused across warm invocations
let cachedSiteId = null;

async function getAccessToken() {
  if (cachedToken && cachedToken.expiresAt > Date.now() + 30_000) {
    return cachedToken.accessToken;
  }

  const tenant = requireEnv("AZURE_TENANT_ID");
  const clientId = requireEnv("AZURE_CLIENT_ID");
  const clientSecret = requireEnv("AZURE_CLIENT_SECRET");

  const res = await fetch(`https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: clientId,
      client_secret: clientSecret,
      scope: "https://graph.microsoft.com/.default",
      grant_type: "client_credentials",
    }),
  });

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Graph token request failed (${res.status}): ${text}`);
  }

  const data = await res.json();
  cachedToken = {
    accessToken: data.access_token,
    expiresAt: Date.now() + data.expires_in * 1000,
  };
  return cachedToken.accessToken;
}

async function getSiteId(accessToken) {
  if (cachedSiteId) return cachedSiteId;

  const hostname = requireEnv("SP_HOSTNAME");
  const sitePath = requireEnv("SP_SITE_PATH");

  const res = await graphFetch(accessToken, `/sites/${hostname}:${sitePath}`);
  cachedSiteId = res.id;
  return cachedSiteId;
}

async function graphFetch(accessToken, path, options = {}) {
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    ...options,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
      ...(options.headers || {}),
    },
  });

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Graph request to ${path} failed (${res.status}): ${text}`);
  }

  if (res.status === 204) return null;
  return res.json();
}

function requireEnv(name) {
  const value = process.env[name];
  if (!value) throw new Error(`Missing required environment variable: ${name}`);
  return value;
}

module.exports = { getAccessToken, getSiteId, graphFetch, requireEnv };
