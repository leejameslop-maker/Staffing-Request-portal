// ===================== graph.js =====================
// This file handles Microsoft Graph calls to your SharePoint List.
// Replace SITE_ID and LIST_ID with your actual SharePoint values.

// IMPORTANT:
// - You MUST replace the placeholders below after uploading this file.
// - I will help you get your site ID and list ID next.

const SITE_ID = 'YOUR_SITE_ID';
const LIST_ID = 'YOUR_LIST_ID';

// ----- Helper: GET auth token from auth.js -----
async function graphRequest(url, method = 'GET', body = null) {
  const token = await getAccessToken();
  const headers = {
    'Authorization': `Bearer ${token}`,
    'Content-Type': 'application/json'
  };

  const options = { method, headers };
  if (body) options.body = JSON.stringify(body);

  const res = await fetch(url, options);
  if (!res.ok) {
    throw new Error(`Graph API error: ${res.status} ${res.statusText}`);
  }
  return await res.json();
}

// ----- CREATE request (submit form) -----
async function createRequest(fields) {
  const url = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/items`;
  return await graphRequest(url, 'POST', {
    fields
  });
}

// ----- GET all requests (dashboard) -----
async function getRequests() {
  const url = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/items?expand=fields`;
  const data = await graphRequest(url);
  return data.value || [];
}

// expose globally
window.createRequest = createRequest;
window.getRequests = getRequests;

// =================== end graph.js ===================
