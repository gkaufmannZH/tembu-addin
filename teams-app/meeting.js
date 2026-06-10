/* globals msal, microsoftTeams */

const CLIENT_ID   = '6a0f7ccb-afe3-4045-9b45-721d2046fafb';
const AUTH_URL    = 'https://gkaufmannzh.github.io/tembu.app/teams-app/auth.html';
const SCOPES      = ['User.Read', 'Tasks.Read', 'Chat.ReadBasic'];
const TEMBU_LIST  = 'Tembu';

let _token = null;
let _allRumbles = [];
let _msal = null;
let _teamsReady = false;
let _chatMemberNames = []; // display names of participants in current chat/meeting

// ── Init ──────────────────────────────────────────────────────────────────
wireEvents();

(async () => {
  try {
    await microsoftTeams.app.initialize();
    _teamsReady = true;
  } catch {}

  // Try silent token on startup
  try {
    const inst = await getMsal();
    const accounts = inst.getAllAccounts();
    if (accounts.length > 0) {
      const result = await inst.acquireTokenSilent({ scopes: SCOPES, account: accounts[0] });
      _token = result.accessToken;
      await loadRumbles();
      showSignedIn();
    }
  } catch {}
})();

// ── MSAL (silent refresh only) ────────────────────────────────────────────
async function getMsal() {
  if (_msal) return _msal;
  _msal = new msal.PublicClientApplication({
    auth: {
      clientId: CLIENT_ID,
      authority: 'https://login.microsoftonline.com/common',
      redirectUri: AUTH_URL,
    },
    cache: { cacheLocation: 'localStorage' },
  });
  if (_msal.initialize) await _msal.initialize();
  return _msal;
}

// ── Auth ──────────────────────────────────────────────────────────────────
async function signIn() {
  const btn = document.getElementById('btnSignIn');
  const status = document.querySelector('.status');
  btn.textContent = 'Verbinden…';
  btn.disabled = true;
  if (status) status.textContent = '';

  try {
    if (_teamsReady) {
      // Teams-managed auth dialog — no window.open() needed
      _token = await microsoftTeams.authentication.authenticate({
        url: AUTH_URL,
        width: 600,
        height: 535,
      });
    } else {
      // Fallback for browser testing
      const inst = await getMsal();
      const result = await inst.loginPopup({ scopes: SCOPES, prompt: 'select_account' });
      _token = result.accessToken;
    }
    await loadRumbles();
    showSignedIn();
  } catch (err) {
    btn.textContent = 'Mit Microsoft anmelden';
    btn.disabled = false;
    if (status) status.textContent = 'Fehler: ' + (err?.message || String(err));
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
      createdAt: f.CREATED || task.createdDateTime,
    };
  });
  await loadChatMembers();
}

async function loadChatMembers() {
  _chatMemberNames = [];
  if (!_teamsReady) return;
  try {
    const context = await microsoftTeams.app.getContext();
    const chatId = context.chat?.id;
    if (!chatId) return;
    const result = await graphGet(`/chats/${chatId}/members`);
    _chatMemberNames = (result.value || [])
      .map(m => m.displayName)
      .filter(Boolean);
  } catch {}
}

// ── Render ────────────────────────────────────────────────────────────────
function isParticipant(contactName) {
  if (!_chatMemberNames.length) return false;
  const cn = contactName.toLowerCase();
  return _chatMemberNames.some(m => {
    const mn = m.toLowerCase();
    return mn === cn || mn.includes(cn) || cn.includes(mn);
  });
}

function renderRumbles(filter) {
  const query = (filter || '').toLowerCase().trim();
  const list = document.getElementById('rumbleList');

  const grouped = {};
  for (const r of _allRumbles) {
    if (query && !r.contactName.toLowerCase().includes(query) && !r.text.toLowerCase().includes(query)) continue;
    if (!grouped[r.contactName]) grouped[r.contactName] = [];
    grouped[r.contactName].push(r);
  }

  const allContacts = Object.keys(grouped).sort();
  if (allContacts.length === 0) {
    list.innerHTML = `<div class="empty">${query ? 'Keine Treffer.' : 'Keine aktiven Rumbles.'}</div>`;
    return;
  }

  const chatContacts = allContacts.filter(n => isParticipant(n));
  const otherContacts = allContacts.filter(n => !isParticipant(n));

  const renderContact = (name, highlight) => {
    const items = grouped[name];
    const initials = name.split(' ').map(w => w[0]).join('').slice(0, 2).toUpperCase();
    const rows = items.map(r => {
      const date = r.createdAt ? new Date(r.createdAt).toLocaleDateString('de-CH') : '';
      const link = r.sourceUrl
        ? `<a class="rumble-source" href="${r.sourceUrl}" target="_blank">↗ In Outlook öffnen</a>` : '';
      return `<div class="rumble-item"><div class="rumble-text">${escapeHtml(r.text)}</div><div class="rumble-meta">${date}</div>${link}</div>`;
    }).join('');
    const border = highlight ? 'border-left:3px solid #2d5a5a;' : '';
    return `<div class="section" style="${border}"><div class="section-header"><div class="avatar">${initials}</div><div class="contact-name">${escapeHtml(name)}</div><div class="rumble-count">${items.length} Rumble${items.length !== 1 ? 's' : ''}</div></div>${rows}</div>`;
  };

  let html = '';
  if (chatContacts.length > 0) {
    html += `<div class="chat-section-label">Teilnehmer dieses Chats</div>`;
    html += chatContacts.map(n => renderContact(n, true)).join('');
    if (otherContacts.length > 0) {
      html += `<div class="chat-section-label" style="opacity:0.5">Weitere Rumbles</div>`;
      html += otherContacts.map(n => renderContact(n, false)).join('');
    }
  } else {
    html += allContacts.map(n => renderContact(n, false)).join('');
  }
  list.innerHTML = html;
}

function escapeHtml(str) {
  return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// ── UI helpers ────────────────────────────────────────────────────────────
function showSignedIn() {
  document.getElementById('notSignedIn').classList.add('hidden');
  document.getElementById('signedIn').classList.remove('hidden');
  renderRumbles('');
}

function wireEvents() {
  document.getElementById('btnSignIn')?.addEventListener('click', signIn);
  document.getElementById('searchInput')?.addEventListener('input', e => renderRumbles(e.target.value));
}
