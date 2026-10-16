/* Pinned MSAL Browser 4.5.1, authorization code + PKCE; no secret in browser.
 * Popup is deliberate: Microsoft login pages cannot run in a SharePoint iframe.
 * Tokens remain in MSAL memory, not our localStorage or URL query parameters.
 */
window.trainingAuth = (() => {
  const clientId = 'c46f3256-59ec-4a09-97ad-e2af9a87e42b';
  const scope = `api://${clientId}/Training.Progress`;
  let account = null;
  let busy = false;
  let instance;
  const ready = (async () => {
    if (!window.msal) throw new Error('sign_in_library_unavailable');
    instance = new msal.PublicClientApplication({
      auth: {clientId, authority:'https://login.microsoftonline.com/b794c469-a090-4070-a4a2-b7a40cde25b6',
        redirectUri: new URL('/auth.html', location.origin).href, navigateToLoginRequestUrl:false},
      cache:{cacheLocation:'memoryStorage'},
      system:{allowRedirectInIframe:false, asyncPopups:false, loggerOptions:{piiLoggingEnabled:false,loggerCallback:()=>{}}},
    });
    await instance.initialize();
  })();
  // Prevent an unhandled rejection before the learner clicks the sign-in button.
  ready.catch(() => {});
  return {
    async connect() {
      if (busy) throw new Error('sign_in_in_progress');
      busy = true;
      try {
        await ready;
        const result = await instance.loginPopup({scopes:[scope],prompt:'select_account'});
        account = result.account;
        if (!account) throw new Error('sign_in_required');
        instance.setActiveAccount(account);
        return {id:account.homeAccountId};
      } finally { busy = false; }
    },
    async token() {
      await ready;
      if (!account) throw new Error('sign_in_required');
      // Never launch an unsolicited popup during loading/saving. A failed silent
      // renewal surfaces a reconnect button that the user explicitly clicks.
      const result = await instance.acquireTokenSilent({scopes:[scope],account});
      if (!result.accessToken) throw new Error('sign_in_required');
      return result.accessToken;
    },
    async disconnect() {
      account = null;
      if (instance) {
        instance.setActiveAccount(null);
        await instance.clearCache();
      }
    },
  };
})();
