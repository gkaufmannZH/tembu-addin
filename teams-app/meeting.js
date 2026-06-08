/* globals microsoftTeams, msal */

const CLIENT_ID  = '6a0f7ccb-afe3-4045-9b45-721d2046fafb';
const AUTH_URL   = 'https://gkaufmannzh.github.io/tembu.app/teams-app/auth.html';
const SCOPES     = ['User.Read', 'Tasks.Read'];
const TEMBU_LIST = 'Tembu';

let _token = null;
let _allRumbles = [];
let _msalInstance = null;

// ── MSAL (lazy, so a CDN failure doesn't crash the whole script) ──────────
function getMsal() {
  if (_msalInstance) return _msalInstance;
  if (typeof msal === 'undefined') throw new Error('MSAL CDN nicht geladen');
  _msalInstance = new msal.PublicClientApplication({
    auth: {
      clientId: CLIENT_ID,
      authority: 'https://login.microsoftonline.com/common',
      redirectUri: AUTH_URL,
    },
    cache: { cacheLocation: 'localStorage' },
  });
  return _msalInstance;
}

// ── Init ──────────────────────────────────────────────────────────────────
wireEvents();

try {
  const inst = getMsal();
  const initFn = inst.initialize ? inst.initialize.bind(inst) : () => Promise.resolve();
  initFn().then(async () => {
    const accounts = inst.getAllAccounts();
    if (accounts.length > 0) {
      try {
        const result = await inst.acquireTokenSilent({ scopes: SCOPES, account: accounts[0] });
        _token = result.accessToken;
        await loadRumbles();
        showSignedIn();
      } catch (_) {}
    }
  }).catch(e => setStatus('Init-Fehler: ' + e.message));
} catch (e) {
  setStatus('Fehler: ' + e.message);
}

// ── Auth ──────────────────────────────────────────────────────────────────
function signIn() {
  const btn = document.getElementById('btnSignIn');
  btn.textContent = 'Verbinden…';

  if (typeof microsoftTeams === 'undefined') {
    btn.textContent = 'Fehler: Teams SDK nicht geladen';
    return;
  }

  microsoftTeams.authentication.authenticate({
    url: AUTH_URL,
    width: 600,
    height: 535,
  }).then(token => {
    _token = token;
    loadRumbles().then(showSignedIn);
  }).catch(err => {
    btn.textContent = 'Fehler: ' + (err?.message || String(err));
  });
}

// ── Graph ─────────────────────────────────────────────────────────────────
async function graphGet(path) {
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    headers: { Authorization: `Bearer ${_token}` },
  });
  if (!res.ok) throw new Error(`Graph ${res.status}`);
  return res.json();
}

function parseBody(text) {
  const r = {};
  for (const line of (text || '').split('\n')) {
    const i = line.indexOf(':');
    if (i > 0) r[line.slice(0, i).trim()] = line.slice(i + 1).trim();
  }
  return r;
}

async function loadRumbles() {
  const lists = await graphGet('/me/todo/lists');
  const list = lists.value?.find(l => l.displayName === TEMBU_LIST);
  if (!list) { _allRumbles = []; return; }

  const tasks = await graphGet(
    `/me/todo/lists/${list.id}/tasks?$filter=status ne 'completed'&$top=200`
  );

  _allRumbles = (tasks.value || []).map(task => {
    const f = parseBody(task.body?.content);
    return {
      contactName: f.CONTACT || task.title.replace(/^Tembu:\s*/i, ''),
      text: f.TEXT || task.title.replace(/^Tembu:\s*/i, ''),
      sourceUrl: f.SOURCE_URL || null,
      sourceType: f.SOURCE_TYPE || null,
      createdAt: f.CREATED || task.createdDateTime,
    };
  });
}

// ── Render ────────────────────────────────────────────────────────────────
function renderRumbles(filter) {
  const query = (filter || '').toLowerCase().trim();
  const list = document.getElementById('rumbleList');

  const grouped = {};
  for (const r of _allRumbles) {
    if (query && !r.contactName.toLowerCase().includes(query) && !r.text.toLowerCase().includes(query)) continue;
    if (!grouped[r.contactName]) grouped[r.contactName] = [];
    grouped[r.contactName].push(r);
  }

  const contacts = Object.keys(grouped).sort();
  if (contacts.length === 0) {
    list.innerHTML = `<div class="empty">${query ? 'Keine Treffer.' : 'Keine aktiven Rumbles.'}</div>`;
    return;
  }

  list.innerHTML = contacts.map(name => {
    const items = grouped[name];
    const initials = name.split(' ').map(w => w[0]).join('').slice(0, 2).toUpperCase();
    const rumbleRows = items.map(r => {
      const date = r.createdAt ? new Date(r.createdAt).toLocaleDateString('de-CH') : '';
      const sourceLink = r.sourceUrl
        ? `<a class="rumble-source" href="${r.sourceUrl}" target="_blank">↗ In Outlook öffnen</a>`
        : '';
      return `
        <div class="rumble-item">
          <div class="rumble-text">${escapeHtml(r.text)}</div>
          <div class="rumble-meta">${date}</div>
          ${sourceLink}
        </div>`;
    }).join('');

    return `
      <div class="section">
        <div class="section-header">
          <div class="avatar">${initials}</div>
          <div class="contact-name">${escapeHtml(name)}</div>
          <div class="rumble-count">${items.length} Rumble${items.length !== 1 ? 's' : ''}</div>
        </div>
        ${rumbleRows}
      </div>`;
  }).join('');
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// ── UI helpers ────────────────────────────────────────────────────────────
function setStatus(msg) {
  const el = document.querySelector('.status');
  if (el) el.textContent = msg;
}

function showSignedIn() {
  document.getElementById('notSignedIn').classList.add('hidden');
  document.getElementById('signedIn').classList.remove('hidden');
  renderRumbles('');
}

function showNotSignedIn() {
  document.getElementById('notSignedIn').classList.remove('hidden');
  document.getElementById('signedIn').classList.add('hidden');
}

function wireEvents() {
  document.getElementById('btnSignIn')?.addEventListener('click', signIn);
  document.getElementById('searchInput')?.addEventListener('input', e => {
    renderRumbles(e.target.value);
  });
}
