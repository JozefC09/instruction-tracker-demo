// Microsoft 365 sign-in with MSAL.js (SPA, auth code + PKCE). No passwords or secrets are stored by the app.
const SCOPES = ['User.Read', 'Files.ReadWrite.All'];

let pca = null;
let account = null;

const inIframe = () => {
  try { return window.self !== window.top; } catch { return true; }
};

export async function initAuth(cfg) {
  if (!window.msal) throw new Error('The Microsoft sign-in library did not load. Check that cdn.jsdelivr.net is allowed, or host msal-browser locally (see README).');
  const redirectUri = cfg.redirectUri || location.origin + location.pathname;
  pca = new window.msal.PublicClientApplication({
    auth: {
      clientId: cfg.clientId,
      authority: `https://login.microsoftonline.com/${cfg.tenantId || 'organizations'}`,
      redirectUri,
    },
    cache: { cacheLocation: 'localStorage' },
  });
  await pca.initialize();
  const result = await pca.handleRedirectPromise();
  account = result?.account ?? pca.getAllAccounts()[0] ?? null;
  if (account) pca.setActiveAccount(account);
  return account;
}

export async function signIn() {
  // Redirect sign-in doesn't work inside an iframe (e.g. a LibGuides embed), so use a popup there.
  if (inIframe()) {
    const r = await pca.loginPopup({ scopes: SCOPES });
    account = r.account;
    pca.setActiveAccount(account);
    return account;
  }
  await pca.loginRedirect({ scopes: SCOPES });
  return null;
}

export async function signOut() {
  const req = { account };
  if (inIframe()) await pca.logoutPopup(req);
  else await pca.logoutRedirect(req);
}

export function currentUser() {
  return account ? { name: account.name || account.username, upn: account.username } : null;
}

export async function getToken() {
  try {
    const r = await pca.acquireTokenSilent({ scopes: SCOPES, account });
    return r.accessToken;
  } catch (e) {
    if (e instanceof window.msal.InteractionRequiredAuthError) {
      if (inIframe()) return (await pca.acquireTokenPopup({ scopes: SCOPES, account })).accessToken;
      await pca.acquireTokenRedirect({ scopes: SCOPES, account });
    }
    throw e;
  }
}
