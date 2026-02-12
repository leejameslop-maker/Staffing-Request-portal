// ===================== auth.js =====================
// Fill in these two values from Entra ID (App registrations > Your app > Overview)
// Replace the placeholders exactly once and commit this file to your repo root.
const tenantId = 'YOUR_TENANT_ID'; // e.g., 11111111-2222-3333-4444-555555555555
const clientId = 'YOUR_CLIENT_ID'; // e.g., 00000000-0000-0000-0000-000000000000

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
  auth: {
    clientId,
    authority,
    redirectUri,
    postLogoutRedirectUri
  },
  cache: {
    cacheLocation: 'localStorage', // persist across refresh
    storeAuthStateInCookie: false  // set true only for legacy IE/EdgeHTML
  },
  system: { allowNativeBroker: false }
};

// Create the MSAL instance
const msalInstance = new msal.PublicClientApplication(msalConfig);

// Sign-in helper used by index.html
async function signInIfNeeded() {
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length === 0) {
    await msalInstance.loginPopup({ scopes: graphScopes });
  }
}

// Acquire token for Microsoft Graph
async function getAccessToken() {
  const acct = msalInstance.getAllAccounts()[0];
  const req = { scopes: graphScopes, account: acct };
  try {
    const res = await msalInstance.acquireTokenSilent(req);
    return res.accessToken;
  } catch (e) {
    const res = await msalInstance.acquireTokenPopup({ scopes: graphScopes });
    return res.accessToken;
  }
}

// Expose to global (so index.html can call them)
window.signInIfNeeded = signInIfNeeded;
window.getAccessToken = getAccessToken;
// =================== end auth.js ===================
