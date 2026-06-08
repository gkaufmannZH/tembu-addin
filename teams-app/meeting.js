/* globals microsoftTeams, msal */

const CLIENT_ID  = '6a0f7ccb-afe3-4045-9b45-721d2046fafb';
const AUTH_URL   = 'https://gkaufmannzh.github.io/tembu.app/teams-app/auth.html';
const SCOPES     = ['User.Read', 'Tasks.Read'];
const TEMBU_LIST = 'Tembu';

let _token = null;
let _allRumbles = [];

// ── Init: only wire UI, no CDN calls at load time ─────────────────────────
wireEvents();

// ── Auth ──────────────────────────────────────────────────────────────────
async function signIn() {
  const btn = document.getElementById('btnSignIn');
  btn.textContent = 'Verbinden…';

  let tokenReceived = false;

  // Poll localStorage for token written by auth.html popup
  const pollInterval = setInterval(() => {
    const pending = localStorage.getItem('tembu_pending_token');
    const ts = parseInt(localStorage.getItem('tembu_pending_ts') || '0');
    if (pending && (Date.now() - ts) < 60000) {
      clearInterval(pollInterval);
      localStorage.removeItem('tembu_pending_token');
      localStorage.removeItem('tembu_pending_ts');
      tokenReceived = true;
      _token = pending;
      loadRumbles().then(showSignedIn);
    }
  }, 500);

  try {
    if (typeof microsoftTeams === 'undefined') throw new Error('Teams SDK nicht geladen');

    btn.textContent = 'Teams init…';
    await Promise.race([
      microsoftTeams.app.initialize(),
      new Promise((_, reject) => setTimeout(() => reject(new Error('Teams init timeout (3s)')), 3000)),
    ]);

    btn.textContent = 'Anmelden…';
    const token = await microsoftTeams.authentication.authenticate({
      url: AUTH_URL,
      width: 600,
      height: 535,
    });

    clearInterval(pollInterval);
    _token = token;
    await loadRumbles();
    showSignedIn();
  } catch (err) {
    clearInterval(pollInterval);
    if (!tokenReceived) {
      btn.textContent = 'Fehler: ' + (err?.message || String(err));
    }
  }
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
