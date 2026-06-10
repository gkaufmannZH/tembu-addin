/* globals Office, msal */

const CLIENT_ID    = '6a0f7ccb-afe3-4045-9b45-721d2046fafb';
const AUTH_URL     = 'https://gkaufmannzh.github.io/tembu.app/outlook-addin/auth.html';
const SCOPES       = ['User.Read', 'Tasks.ReadWrite'];
const TEMBU_LIST   = 'Tembu';
const SESSION_KEY  = '@tembu_outlook_session';

let _token = null;
let _account = null;
let _sourceUrl = null;
let _itemType = null;

// ── MSAL instance (silent refresh only — auth happens via dialog) ──────────
const msalInstance = new msal.PublicClientApplication({
  auth: {
    clientId: CLIENT_ID,
    authority: 'https://login.microsoftonline.com/common',
    redirectUri: AUTH_URL,
  },
  cache: { cacheLocation: 'localStorage' },
});

// ── Office init ───────────────────────────────────────────────────────────
Office.initialize = async function () {
  await msalInstance.initialize();
  await msalInstance.handleRedirectPromise();

  let authed = false;

  // Try MSAL silent refresh (works when MSAL cache is in taskpane's localStorage)
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length > 0) {
    try {
      const result = await msalInstance.acquireTokenSilent({ scopes: SCOPES, account: accounts[0] });
      _token = result.accessToken;
      _account = accounts[0];
      authed = true;
    } catch {}
  }

  // Fallback: own session cache — handles Outlook Desktop where dialog/taskpane WebViews are isolated
  if (!authed) {
    try {
      const cached = JSON.parse(localStorage.getItem(SESSION_KEY) || 'null');
      if (cached && cached.expiry > Date.now()) {
        _token = cached.token;
        _account = cached.account;
        authed = true;
      }
    } catch {}
  }

  authed ? showForm() : showSignIn();
  wireEvents();        // always wire first so buttons work even if context load fails
  loadOutlookContext();
};

// ── Read context from current Outlook item ────────────────────────────────
function loadOutlookContext() {
  const item = Office.context.mailbox?.item;
  if (!item) return;

  _itemType = item.itemType;
  const badge = document.getElementById('sourceBadge');
  const typeEl = document.getElementById('sourceType');
  const subjectEl = document.getElementById('sourceSubject');
  const contactInput = document.getElementById('contactName');

  // In read mode item.subject is a string; in compose mode it's an Office.Subject object
  function setSubject(el) {
    if (typeof item.subject === 'string') {
      el.textContent = item.subject;
    } else if (item.subject?.getAsync) {
      item.subject.getAsync(r => { if (r.status === Office.AsyncResultStatus.Succeeded) el.textContent = r.value || ''; });
    }
  }

  if (item.itemType === Office.MailboxEnums.ItemType.Message) {
    typeEl.textContent = 'E-Mail';
    badge.classList.remove('hidden');
    setSubject(subjectEl);

    // Read mode: item.from.displayName is a string; compose mode: item.from.getAsync()
    if (typeof item.from?.displayName === 'string') {
      contactInput.value = item.from.displayName;
    } else if (item.from?.getAsync) {
      item.from.getAsync(r => { if (r.status === Office.AsyncResultStatus.Succeeded) contactInput.value = r.value?.displayName || ''; });
    }

    try {
      const restId = Office.context.mailbox.convertToRestId(item.itemId, Office.MailboxEnums.RestVersion.v2_0);
      _sourceUrl = `https://outlook.office.com/mail/deeplink/read/${encodeURIComponent(restId)}`;
    } catch {}

  } else if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
    typeEl.textContent = 'Termin';
    badge.classList.remove('hidden');
    setSubject(subjectEl);

    // Read mode: attendees are plain arrays; compose mode: Office.Recipients objects with getAsync()
    if (Array.isArray(item.requiredAttendees)) {
      const all = [...(item.requiredAttendees || []), ...(item.optionalAttendees || [])];
      const first = all.find(a => a.displayName);
      if (first) contactInput.value = first.displayName;
    } else if (item.requiredAttendees?.getAsync) {
      item.requiredAttendees.getAsync(r => {
        if (r.status === Office.AsyncResultStatus.Succeeded) {
          const first = (r.value || []).find(a => a.displayName);
          if (first) contactInput.value = first.displayName;
        }
      });
    }

    try {
      const restId = Office.context.mailbox.convertToRestId(item.itemId, Office.MailboxEnums.RestVersion.v2_0);
      _sourceUrl = `https://outlook.office.com/calendar/item/${encodeURIComponent(restId)}`;
    } catch {}
  }
}

// ── Auth ──────────────────────────────────────────────────────────────────
function startSignIn() {
  Office.context.ui.displayDialogAsync(
    AUTH_URL,
    { height: 60, width: 35, promptBeforeOpen: false },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        showStatus('Dialog-Fehler: ' + (asyncResult.error?.message || asyncResult.error?.code || 'unbekannt'), 'error');
        return;
      }
      const dialog = asyncResult.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (args) => {
        dialog.close();
        try {
          const msg = JSON.parse(args.message);
          if (msg.success && msg.token) {
            _token = msg.token;
            _account = msg.account;
            // Persist token for subsequent reloads (Outlook Desktop isolates dialog/taskpane WebViews)
            try {
              localStorage.setItem(SESSION_KEY, JSON.stringify({
                token: msg.token,
                account: msg.account,
                expiry: Date.now() + 55 * 60 * 1000,
              }));
            } catch {}
            showForm();
            clearStatus();
          } else {
            showStatus(msg.error || 'Anmeldung fehlgeschlagen.', 'error');
          }
        } catch {
          showStatus('Unbekannter Fehler.', 'error');
        }
      });
      dialog.addEventHandler(Office.EventType.DialogEventReceived, () => dialog.close());
    }
  );
}

function signOut() {
  msalInstance.getAllAccounts().forEach(a => msalInstance.logoutPopup({ account: a }).catch(() => {}));
  localStorage.clear();
  _token = null;
  _account = null;
  showSignIn();
}

// ── Graph helpers ─────────────────────────────────────────────────────────
async function graphFetch(method, path, body) {
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    method,
    headers: {
      Authorization: `Bearer ${_token}`,
      'Content-Type': 'application/json',
    },
    body: body ? JSON.stringify(body) : undefined,
  });
  if (!res.ok) throw new Error(`Graph ${res.status}`);
  if (res.status === 204) return null;
  return res.json();
}

async function getOrCreateTembuList() {
  const lists = await graphFetch('GET', '/me/todo/lists');
  const existing = lists.value.find(l => l.displayName === TEMBU_LIST);
  if (existing) return existing.id;
  const created = await graphFetch('POST', '/me/todo/lists', { displayName: TEMBU_LIST });
  return created.id;
}

async function saveRumble(contactName, rumbleText) {
  const listId = await getOrCreateTembuList();

  const bodyLines = [
    `TEXT:${rumbleText}`,
    `CONTACT:${contactName}`,
    `SOURCE_TYPE:${_itemType === Office.MailboxEnums.ItemType.Appointment ? 'appointment' : 'message'}`,
  ];
  if (_sourceUrl) bodyLines.push(`SOURCE_URL:${_sourceUrl}`);
  bodyLines.push(`CREATED:${new Date().toISOString()}`);

  await graphFetch('POST', `/me/todo/lists/${listId}/tasks`, {
    title: `Tembu: ${contactName}`,
    body: { content: bodyLines.join('\n'), contentType: 'text' },
    importance: 'normal',
  });
}

// ── Save handler ──────────────────────────────────────────────────────────
async function handleSave() {
  const contactName = document.getElementById('contactName').value.trim();
  const rumbleText  = document.getElementById('rumbleText').value.trim();

  if (!contactName) { showStatus('Bitte einen Kontaktnamen eingeben.', 'error'); return; }
  if (!rumbleText)  { showStatus('Bitte einen Rumble-Text eingeben.', 'error'); return; }

  const btn = document.getElementById('btnSave');
  btn.disabled = true;
  btn.textContent = 'Speichern…';
  clearStatus();

  try {
    await saveRumble(contactName, rumbleText);
    showStatus('Rumble gespeichert ✓ Wird beim nächsten App-Start synchronisiert.', 'success');
    document.getElementById('rumbleText').value = '';
  } catch (e) {
    showStatus(`Fehler: ${e.message}`, 'error');
  } finally {
    btn.disabled = false;
    btn.textContent = 'Rumble speichern';
  }
}

// ── UI helpers ────────────────────────────────────────────────────────────
function showForm() {
  document.getElementById('signInSection').classList.add('hidden');
  document.getElementById('formSection').classList.remove('hidden');
  if (_account) {
    document.getElementById('userLabel').textContent = _account.name || _account.username || '';
  }
}

function showSignIn() {
  document.getElementById('signInSection').classList.remove('hidden');
  document.getElementById('formSection').classList.add('hidden');
}

function showStatus(msg, type) {
  const el = document.getElementById('statusMsg');
  el.textContent = msg;
  el.className = `status ${type}`;
  el.classList.remove('hidden');
}

function clearStatus() {
  document.getElementById('statusMsg').classList.add('hidden');
}

function wireEvents() {
  document.getElementById('btnSignIn').addEventListener('click', startSignIn);
  document.getElementById('btnSignOut').addEventListener('click', signOut);
  document.getElementById('btnSave').addEventListener('click', handleSave);
  document.getElementById('btnDiag').addEventListener('click', runDiag);
}

function runDiag() {
  const out = document.getElementById('diagOut');
  out.style.display = 'block';

  const d = Office.context.diagnostics || {};
  const info = [
    'Host: ' + (d.host || '?'),
    'Platform: ' + (d.platform || '?'),
    'Version: ' + (d.version || '?'),
    'DisplayLanguage: ' + (Office.context.displayLanguage || '?'),
  ].join('\n');

  const results = [];
  out.textContent = info + '\n\nTeste URLs…';

  const tests = [
    { label: 'github.io/auth.html', url: AUTH_URL },
    { label: 'login.microsoft.com', url: 'https://login.microsoftonline.com/common' },
    { label: 'taskpane URL', url: document.URL },
  ];
  out.textContent = info + '\ntaskpane lädt von: ' + document.URL + '\n\nTeste URLs…';

  const taskpaneUrl = document.URL;
  let i = 0;
  function next() {
    if (i >= tests.length) {
      out.textContent = info + '\ntaskpane von: ' + taskpaneUrl + '\n\n' + results.join('\n');
      return;
    }
    const t = tests[i++];
    Office.context.ui.displayDialogAsync(
      t.url,
      { height: 1, width: 1, promptBeforeOpen: false },
      (r) => {
        if (r.status === Office.AsyncResultStatus.Failed) {
          results.push(t.label + ': FEHLER code ' + r.error.code);
        } else {
          results.push(t.label + ': OK ✓');
          r.value.close();
        }
        next();
      }
    );
  }
  next();
}
