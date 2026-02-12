// ===================== graph.js =====================
// Microsoft Graph calls wired to your SharePoint List.

// ✔ Your real IDs (from earlier steps)
const SITE_ID = 'interimwm.sharepoint.com,b7376cd3-3e98-4800-a514-952b910353bd,6b26dce2-4242-4c0f-9fb8-45d6f4c128d1';
const LIST_ID = '34e0c73f-8309-4143-8de6-c1b1bdcf05d6';

// ----- auth helper from auth.js -----
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
    const text = await res.text().catch(() => '');
    throw new Error(`Graph API error: ${res.status} ${res.statusText} - ${text}`);
  }
  return await res.json();
}

// ----- CREATE (form submit) -----
async function createRequest(fields) {
  const url = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/items`;
  return await graphRequest(url, 'POST', { fields });
}

// ----- READ (dashboard) -----
async function getRequests() {
  const url = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/items?expand=fields`;
  const data = await graphRequest(url);
  return data.value || [];
}

// ----- UPDATE (NEW) → change fields on an item (e.g., Stage, Status) -----
async function updateRequest(itemId, fieldsPatch) {
  // PATCH the list item's *fields* entity
  const url = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/items/${itemId}/fields`;
  return await graphRequest(url, 'PATCH', fieldsPatch);
}

// Expose globally for index.html
window.createRequest = createRequest;
window.getRequests = getRequests;
window.updateRequest = updateRequest;

// =================== end graph.js ===================
