// ===================== auth.js =====================
// Entra ID (App registrations > Your app > Overview)
const tenantId = '09bcf3cb-a7ba-4d82-aabb-d6005731b877';   // Directory (tenant) ID
const clientId = '7e8e1393-eeac-4b56-ab11-10de17e3b219';   // Application (client) ID

// App authority and redirect locations (GitHub Pages)
const authority = `https://login.microsoftonline.com/${tenantId}`;
const redirectUri = 'https://leejameslop-maker.github.io/Staffing-Request-portal/';
const postLogoutRedirectUri = redirectUri;

// Graph scopes: user profile + SharePoint Sites scopes (read/write lists)
const graphScopes = [
  'openid',
  'profile',
  'offline_access',
  'User.Read',
  'Sites.Read.All',
  'Sites.ReadWrite.All'
];

// Ensure msal-browser is loaded in index.html BEFORE this file:
// <script src="https://alcdn.msauth.net/browser/2.38.0/js/msal-browser.min.js"></script>

const msalConfig = {
  auth: { clientId, authority, redirectUri, postLogoutRedirectUri },
  cache: { cacheLocation: 'localStorage', storeAuthStateInCookie: false },
  system: { allowNativeBroker: false }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

async function signInIfNeeded() {
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length === 0) {
    await msalInstance.loginPopup({ scopes: graphScopes });
  }
}

async function getAccessToken() {
  const acct = msalInstance.getAllAccounts()[0];
  const req = { scopes: graphScopes, account: acct };
  try {
    const res = await msalInstance.acquireTokenSilent(req);
    return res.accessToken;
  } catch {
    const res = await msalInstance.acquireTokenPopup({ scopes: graphScopes });
    return res.accessToken;
  }
}

window.signInIfNeeded = signInIfNeeded;
window.getAccessToken = getAccessToken;
// =================== end auth.js ===================
