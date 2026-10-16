# Tracking readiness and safe release plan

## Scope of this change

- Keep the complete GitHub JSON catalog authoritative. Deploying a partial SharePoint catalog must not hide published courses.
- Only mark a new completion saved when the server explicitly acknowledges it.
- Show a clear save failure and a retry action. Retry stays attached to the original lesson, even after navigation.
- Add dependency-free regression tests: `node --test tests/*.test.cjs`.
- No tracking functions have been activated or relocated in this change. It does **not** enable SharePoint tracking or Microsoft sign-in.

## Why the legacy functions must not be switched on as-is

The root-level `record-completion.js` and `get-completions.js` trust a browser-supplied `personEmail`. The public Netlify address can be accessed outside the SharePoint iframe. An email prompt, CORS, a permitted email domain, or frame restrictions do not authenticate a learner. Exposing these handlers unchanged would allow impersonation and completion-record disclosure.

The present completion checkmark was optimistic: it appeared before the SharePoint write succeeded. This patch removes that misleading success, but identity verification is still a prerequisite for activating the backend.

## Required secure tracking design

1. Use single-tenant Microsoft Entra sign-in for learners. Configure the exact redirect URI and tenant/client identifiers in an appropriate app registration. Keep client secrets server-side. Do not put secrets in course JSON, GitHub, or browser code.
2. Support SharePoint iframe constraints: use an explicit sign-in popup or top-level sign-in fallback; Microsoft login generally cannot run inside an iframe. Test supported browsers and blocked third-party storage.
3. Validate token signature, issuer, audience, expiry, and tenant on the backend. Reject unauthenticated and wrong-tenant requests. Derive learner identity from validated claims, never from the submitted email. Prefer stable tenant/object IDs for identity, retaining email for display/reporting. Decide explicitly whether tenant guests may participate.
4. Read/write only that learner's records. Keep any administrative reporting separate and permission-checked. Preserve the app's limited SharePoint site grant.
5. Validate training/lesson IDs against the published catalog. Validate scores server-side or label them as self-reported; browser submissions are not evidence of assessed proficiency.
6. Make repeated completion submissions safe, including retries. Define how legacy records map to verified identities. Support Graph pagination. Return generic errors to browsers and keep diagnostic details server-side without tokens or personal data.
7. Confirm Netlify runtime environment variables without exposing values. Test on a restricted preview/staging list, not the production list by default. Note that preview deployments may not have production-scoped variables.
8. Verify: actual Microsoft sign-in → lesson completion → one correctly attributed SharePoint row → refreshed browser restores the checkmark → second device restores it. Test expired sign-in, offline errors, duplicate retries, identity tampering, and wrong-tenant access.

## Other follow-up work

- Persist partial lesson progress if the business requires it. Current implementation records lesson completion, not durable block-by-block progress. A completion percentage can be derived from completed lessons once the backend is verified.
- The repository and training JSON are public. Internal videos require SharePoint access, but text, rubrics, embedded identifiers, and other repository content are not made private by embedding the site. Review content sensitivity and hosting/repository privacy with the owner.
- Review AI coach authentication, abuse protection, budget limits, and confidential-data handling before increasing use.
- Keep the JSON catalog as the publishing source until a deliberate catalog migration is approved and all entries are populated.

## Release gate

This frontend patch may be reviewed independently. It is not a claim that the portal is a secure LMS or that SharePoint tracking works. Do not describe tracking as operational until the authenticated end-to-end checks pass.
