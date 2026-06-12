/* globals Office, msal */

const CLIENT_ID    = '6a0f7ccb-afe3-4045-9b45-721d2046fafb';
const AUTH_URL     = 'https://gkaufmannzh.github.io/tembu.app/outlook-addin/auth.html';
const SCOPES       = ['User.Read', 'Tasks.ReadWrite', 'Contacts.Read'];
const TEMBU_LIST   = 'Tembu';
const SESSION_KEY  = '@tembu_outlook_session';

let _token = null;
let _account = null;
let _sourceUrl = null;
let _itemType = null;
let _contactEmail = null;
let _appointmentAttendeeNames = [];
let _allRumbles = [];
let _rumbleLoaded = false;
let _messageParticipantNames = [];
let _contactDirectory = []; // { name, phone } from /me/contacts

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

  const pickerMode = new URLSearchParams(window.location.search).get('mode') === 'picker';

  authed ? showForm() : showSignIn();
  wireEvents();
  if (!pickerMode) loadOutlookContext();
  if (authed) loadContactsFromGraph(pickerMode);

  // Always register ItemChanged so that navigating to an email with the add-in
  // already open switches to email context (even when opened in picker mode)
  try {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, function () {
      resetItemContext();
      loadOutlookContext();
    });
  } catch {}
};

// ── Reset per-item state (called on ItemChanged when taskpane is pinned) ──
function resetItemContext() {
  _allRumbles = [];
  _rumbleLoaded = false;
  _messageParticipantNames = [];
  _appointmentAttendeeNames = [];
  _contactEmail = null;
  _sourceUrl = null;
  _itemType = null;

  ['contactName', 'contactPhone', 'rumbleText'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = '';
  });
  ['participantPicker', 'followUpSection', 'briefingSection', 'sourceBadge', 'contactPickerSection'].forEach(id => {
    document.getElementById(id)?.classList.add('hidden');
  });

  // Remove any existing ribbon notification from previous item
  try {
    Office.context.mailbox?.item?.notificationMessages?.removeAsync('tembu-badge', () => {});
  } catch {}

  clearStatus();
  updateBrowseTabLabel();
}

// ── Read context from current Outlook item ────────────────────────────────
function loadOutlookContext() {
  const item = Office.context.mailbox?.item;
  if (!item) {
    // No email/meeting selected (Ribbon-Button mode): show contact picker
    showContactPickerSection();
    return;
  }
  document.getElementById('contactPickerSection')?.classList.add('hidden');

  _itemType = item.itemType;
  const badge = document.getElementById('sourceBadge');
  const typeEl = document.getElementById('sourceType');
  const subjectEl = document.getElementById('sourceSubject');
  const contactInput = document.getElementById('contactName');

  // In read mode item.subject is a string; in compose mode it's an Office.Subject object
  function setSubject(el) {
    if (typeof item.subject === 'string') {
      el.textContent = item.subject;
      setRumbleSuggestion(item.subject);
    } else if (item.subject?.getAsync) {
      item.subject.getAsync(r => {
        if (r.status === Office.AsyncResultStatus.Succeeded) {
          el.textContent = r.value || '';
          setRumbleSuggestion(r.value || '');
        }
      });
    }
  }

  if (item.itemType === Office.MailboxEnums.ItemType.Message) {
    typeEl.textContent = 'E-Mail';
    badge.classList.remove('hidden');
    setSubject(subjectEl);
    _messageParticipantNames = [];

    // Sender
    if (typeof item.from?.displayName === 'string') {
      contactInput.value = item.from.displayName;
      _contactEmail = item.from.emailAddress || null;
      if (item.from.displayName) _messageParticipantNames.push(item.from.displayName);
      triggerPhoneLookup();
    } else if (item.from?.getAsync) {
      item.from.getAsync(r => {
        if (r.status === Office.AsyncResultStatus.Succeeded) {
          contactInput.value = r.value?.displayName || '';
          _contactEmail = r.value?.emailAddress || null;
          if (r.value?.displayName) _messageParticipantNames.push(r.value.displayName);
          updateBrowseTabLabel();
          showParticipantPicker(_messageParticipantNames);
          triggerPhoneLookup();
        }
      });
    }

    // Recipients (To)
    if (Array.isArray(item.to)) {
      _messageParticipantNames.push(...item.to.map(r => r.displayName).filter(Boolean));
    } else if (item.to?.getAsync) {
      item.to.getAsync(r => {
        if (r.status === Office.AsyncResultStatus.Succeeded) {
          _messageParticipantNames.push(...(r.value || []).map(a => a.displayName).filter(Boolean));
          updateBrowseTabLabel();
          showParticipantPicker(_messageParticipantNames);
        }
      });
    }
    updateBrowseTabLabel();
    showParticipantPicker(_messageParticipantNames);

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
      _appointmentAttendeeNames = all.map(a => a.displayName).filter(Boolean);
      updateBrowseTabLabel();
      const first = all.find(a => a.displayName);
      if (first) {
        contactInput.value = first.displayName;
        _contactEmail = first.emailAddress || null;
        triggerPhoneLookup();
      }
      showParticipantPicker(_appointmentAttendeeNames);
      if (_token) loadMeetingBriefing();
    } else if (item.requiredAttendees?.getAsync) {
      item.requiredAttendees.getAsync(r => {
        if (r.status === Office.AsyncResultStatus.Succeeded) {
          _appointmentAttendeeNames = (r.value || []).map(a => a.displayName).filter(Boolean);
          updateBrowseTabLabel();
          const first = (r.value || []).find(a => a.displayName);
          if (first) {
            contactInput.value = first.displayName;
            _contactEmail = first.emailAddress || null;
            triggerPhoneLookup();
          }
          showParticipantPicker(_appointmentAttendeeNames);
          if (_token) loadMeetingBriefing();
        }
      });
    }

    try {
      const restId = Office.context.mailbox.convertToRestId(item.itemId, Office.MailboxEnums.RestVersion.v2_0);
      _sourceUrl = `https://outlook.office.com/calendar/item/${encodeURIComponent(restId)}`;
    } catch {}
  }
}

// ── Contact directory (from /me/contacts) ────────────────────────────────
async function loadContactsFromGraph(forcePickerMode) {
  try {
    let url = '/me/contacts?$select=displayName,mobilePhone,businessPhones&$top=200&$orderby=displayName';
    _contactDirectory = [];
    while (url) {
      const data = await graphFetch('GET', url);
      for (const c of data.value || []) {
        if (c.displayName) {
          _contactDirectory.push({
            name: c.displayName,
            phone: c.mobilePhone || c.businessPhones?.[0] || '',
          });
        }
      }
      url = data['@odata.nextLink']
        ? data['@odata.nextLink'].replace('https://graph.microsoft.com/v1.0', '')
        : null;
    }
    const datalist = document.getElementById('contactSuggestions');
    if (datalist) {
      datalist.innerHTML = _contactDirectory
        .map(c => `<option value="${escapeTp(c.name).replace(/"/g, '&quot;')}">`)
        .join('');
    }
    if (forcePickerMode || !Office.context.mailbox?.item) {
      showContactPickerSection();
    } else {
      // Re-run phone lookup now that directory is populated
      triggerPhoneLookup();
    }
  } catch {}
}

function onContactNameChange(name) {
  const match = _contactDirectory.find(c => c.name.toLowerCase() === name.toLowerCase());
  if (match?.phone) {
    const phoneInput = document.getElementById('contactPhone');
    if (phoneInput && !phoneInput.value) phoneInput.value = match.phone;
  }
}

// ── Phone lookup from contact directory ──────────────────────────────────
function triggerPhoneLookup() {
  if (!_token) return;
  const phoneInput = document.getElementById('contactPhone');
  if (!phoneInput || phoneInput.value) return;
  const contactName = document.getElementById('contactName')?.value?.trim();
  if (!contactName || !_contactDirectory.length) return;
  const match = _contactDirectory.find(c => c.name.toLowerCase() === contactName.toLowerCase());
  if (match?.phone) phoneInput.value = match.phone;
}

// ── Contact picker (Ribbon-Button mode: no email/meeting context) ─────────
function showContactPickerSection() {
  const section = document.getElementById('contactPickerSection');
  if (!section) return;
  section.classList.remove('hidden');
  renderContactPicker('');
}

function renderContactPicker(filter) {
  const list = document.getElementById('contactPickerList');
  if (!list) return;
  if (!_contactDirectory.length) {
    list.innerHTML = '<div class="rumble-empty">Kontakte werden geladen…</div>';
    return;
  }
  const q = filter.toLowerCase().trim();
  const matches = q
    ? _contactDirectory.filter(c => c.name.toLowerCase().includes(q))
    : _contactDirectory.slice(0, 60);
  if (!matches.length) {
    list.innerHTML = '<div class="rumble-empty">Keine Treffer.</div>';
    return;
  }
  list.innerHTML = matches.map(c =>
    `<div class="contact-picker-row" onclick="selectContactFromPicker(${JSON.stringify(c.name)}, ${JSON.stringify(c.phone)})">
      <span class="contact-picker-name">${escapeTp(c.name)}</span>
      ${c.phone ? `<span class="contact-picker-phone">${escapeTp(c.phone)}</span>` : ''}
    </div>`
  ).join('');
}

function filterContactPicker(val) {
  renderContactPicker(val);
}

function selectContactFromPicker(name, phone) {
  document.getElementById('contactName').value = name;
  const phoneInput = document.getElementById('contactPhone');
  if (phoneInput) phoneInput.value = phone || '';
  document.getElementById('contactPickerSection')?.classList.add('hidden');
  const searchEl = document.getElementById('contactPickerSearch');
  if (searchEl) searchEl.value = '';
  document.getElementById('rumbleText')?.focus();
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
            loadContactsFromGraph(new URLSearchParams(window.location.search).get('mode') === 'picker');
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
  _contactDirectory = [];
  showSignIn();
}

// ── ID generation ─────────────────────────────────────────────────────────
function generateRumbleId() {
  if (typeof crypto !== 'undefined' && crypto.randomUUID) return crypto.randomUUID();
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 10);
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

async function saveRumble(contactName, contactPhone, rumbleText) {
  const listId = await getOrCreateTembuList();

  const bodyLines = [
    `RUMBLE_ID:${generateRumbleId()}`,
    `TEXT:${rumbleText}`,
    `CONTACT:${contactName}`,
    `SOURCE_TYPE:${_itemType === Office.MailboxEnums.ItemType.Appointment ? 'appointment' : 'message'}`,
  ];
  if (contactPhone) bodyLines.push(`CONTACT_PHONE:${contactPhone}`);
  if (_contactEmail) bodyLines.push(`CONTACT_EMAIL:${_contactEmail}`);
  if (_sourceUrl) bodyLines.push(`SOURCE_URL:${_sourceUrl}`);
  bodyLines.push(`CREATED:${new Date().toISOString()}`);

  await graphFetch('POST', `/me/todo/lists/${listId}/tasks`, {
    title: `Tembu: ${contactName}`,
    body: { content: bodyLines.join('\n'), contentType: 'text' },
    importance: 'normal',
  });
}

// ── Email suggestion: pre-fill rumble text from subject ───────────────────
function setRumbleSuggestion(subject) {
  const ta = document.getElementById('rumbleText');
  if (!ta || ta.value) return; // don't overwrite manual input
  const cleaned = (subject || '').replace(/^(Re|Fw|FWD|AW|WG|Betreff):\s*/i, '').trim();
  if (cleaned) ta.value = cleaned;
}

// ── Meeting briefing: show existing Rumbles for appointment attendees ──────
function escapeTp(str) {
  return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function parseBodyFields(text) {
  const r = {};
  for (const line of (text || '').split('\n')) {
    const i = line.indexOf(':');
    if (i > 0) r[line.slice(0, i).trim()] = line.slice(i + 1).trim();
  }
  return r;
}

async function loadMeetingBriefing() {
  if (!_token || !_appointmentAttendeeNames.length) return;
  const section = document.getElementById('briefingSection');
  const list = document.getElementById('briefingList');
  if (!section || !list) return;
  try {
    const listId = await getOrCreateTembuList();
    const data = await graphFetch('GET', `/me/todo/lists/${listId}/tasks?$filter=status ne 'completed'&$top=150`);
    const tasks = data.value || [];
    const matches = [];
    const seen = new Set();
    for (const task of tasks) {
      const f = parseBodyFields(task.body?.content);
      const contactName = f.CONTACT || task.title.replace(/^Tembu:\s*/i, '');
      const text = f.TEXT || contactName;
      const dedupKey = f.RUMBLE_ID || f.MEMO_ID ||
        `${contactName.toLowerCase().trim()}||${text.toLowerCase().trim()}`;
      if (seen.has(dedupKey)) continue;
      seen.add(dedupKey);
      const cn = contactName.toLowerCase();
      const hit = _appointmentAttendeeNames.some(n => {
        const mn = n.toLowerCase();
        return mn === cn || mn.includes(cn) || cn.includes(mn);
      });
      if (hit) matches.push({ contactName, text });
    }
    // Show notification badge in the Outlook reading pane (visible even without taskpane open)
    try {
      const nm = Office.context.mailbox?.item?.notificationMessages;
      if (nm) {
        if (matches.length > 0) {
          nm.addAsync('tembu-badge', {
            type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
            message: `Tembu: ${matches.length} Rumble${matches.length !== 1 ? 's' : ''} für dieses Meeting`,
            icon: 'Icon.16',
            persistent: false,
          });
        } else {
          nm.removeAsync('tembu-badge', () => {});
        }
      }
    } catch {}

    if (!matches.length) return;
    section.classList.remove('hidden');
    list.innerHTML = matches.map(m =>
      `<div class="briefing-item"><span class="briefing-name">${escapeTp(m.contactName)}</span><span class="briefing-text">${escapeTp(m.text)}</span></div>`
    ).join('');
  } catch {}
}

// ── Outlook Task aus Rumble erstellen (Nachbereitung) ─────────────────────
async function createFollowUpTask(contactName, rumbleText) {
  try {
    const lists = await graphFetch('GET', '/me/todo/lists');
    const defaultList = (lists.value || []).find(l => l.isDefaultList) || lists.value?.[0];
    if (!defaultList) return false;
    await graphFetch('POST', `/me/todo/lists/${defaultList.id}/tasks`, {
      title: `${contactName}: ${rumbleText.slice(0, 100)}`,
      body: { content: rumbleText, contentType: 'text' },
      importance: 'high',
    });
    return true;
  } catch { return false; }
}

// ── Save handler ──────────────────────────────────────────────────────────
async function handleSave() {
  const contactName  = document.getElementById('contactName').value.trim();
  const contactPhone = document.getElementById('contactPhone').value.trim();
  const rumbleText   = document.getElementById('rumbleText').value.trim();

  if (!contactName) { showStatus('Bitte einen Kontaktnamen eingeben.', 'error'); return; }
  if (!rumbleText)  { showStatus('Bitte einen Rumble-Text eingeben.', 'error'); return; }

  const btn = document.getElementById('btnSave');
  btn.disabled = true;
  btn.textContent = 'Speichern…';
  clearStatus();

  try {
    await saveRumble(contactName, contactPhone, rumbleText);
    showStatus('Rumble gespeichert ✓ Wird beim nächsten App-Start synchronisiert.', 'success');
    document.getElementById('rumbleText').value = '';
    showFollowUpSection(contactName, rumbleText);
  } catch (e) {
    showStatus(`Fehler: ${e.message}`, 'error');
  } finally {
    btn.disabled = false;
    btn.textContent = 'Rumble speichern';
  }
}

// ── Follow-up Task section (Point 3: Erledigt + Diktat) ──────────────────
function showFollowUpSection(contactName, rumbleText) {
  const section = document.getElementById('followUpSection');
  if (!section) return;
  section.classList.remove('hidden');
  const noteInput = document.getElementById('followUpNote');
  if (noteInput) noteInput.value = '';
  const btn = document.getElementById('btnFollowUp');
  if (btn) {
    btn.onclick = async () => {
      const note = noteInput?.value?.trim() || rumbleText;
      btn.disabled = true;
      btn.textContent = 'Erstelle…';
      const ok = await createFollowUpTask(contactName, note);
      if (ok) {
        showStatus('Aufgabe in Outlook erstellt ✓', 'success');
        section.classList.add('hidden');
      } else {
        showStatus('Aufgabe konnte nicht erstellt werden.', 'error');
        btn.disabled = false;
        btn.textContent = 'Als Outlook-Aufgabe speichern';
      }
    };
  }
}

// ── Rumble Browse (Tab: Alle Rumbles / Teilnehmer) ────────────────────────
function showTab(tab) {
  const isCreate = tab === 'create';
  document.getElementById('createPane').classList.toggle('hidden', !isCreate);
  document.getElementById('browsePane').classList.toggle('hidden', isCreate);
  document.getElementById('tabCreate').classList.toggle('active', isCreate);
  document.getElementById('tabBrowse').classList.toggle('active', !isCreate);
  if (!isCreate) {
    const participants = _contextParticipants();
    const searchEl = document.getElementById('rumbleSearch');
    if (searchEl) {
      searchEl.placeholder = participants.length
        ? `Suche in ${participants.length} Teilnehmer-Rumbles…`
        : 'Kontakt suchen…';
    }
    if (!_rumbleLoaded) loadAllRumbles();
    else renderAllRumbles(searchEl?.value || '');
  }
}

function updateBrowseTabLabel() {
  const isAppt = _itemType === Office.MailboxEnums.ItemType.Appointment;
  const isMsg  = _itemType === Office.MailboxEnums.ItemType.Message;
  const tabBrowse = document.getElementById('tabBrowse');
  if (!tabBrowse) return;
  if (isAppt && _appointmentAttendeeNames.length) tabBrowse.textContent = 'Teilnehmer';
  else if (isMsg && _messageParticipantNames.length) tabBrowse.textContent = 'Empfänger';
  else tabBrowse.textContent = 'Alle Rumbles';
}

// ── Participant picker ────────────────────────────────────────────────────
function showParticipantPicker(names) {
  const picker = document.getElementById('participantPicker');
  const chips  = document.getElementById('participantChips');
  if (!picker || !chips) return;
  const unique = [...new Set(names)].filter(Boolean);
  if (unique.length <= 1) { picker.classList.add('hidden'); return; }

  const current = document.getElementById('contactName')?.value || '';
  chips.innerHTML = unique.map(name => {
    const active = name.toLowerCase() === current.toLowerCase() ? ' active' : '';
    return `<button class="participant-chip${active}" onclick="selectParticipant(${JSON.stringify(name)})">${escapeTp(name)}</button>`;
  }).join('');
  picker.classList.remove('hidden');
}

function selectParticipant(name) {
  const contactInput = document.getElementById('contactName');
  if (contactInput) contactInput.value = name;
  document.querySelectorAll('.participant-chip').forEach(c => {
    c.classList.toggle('active', c.textContent === name);
  });
  const phoneInput = document.getElementById('contactPhone');
  if (phoneInput) phoneInput.value = '';
  triggerPhoneLookup();
}

function _contextParticipants() {
  if (_itemType === Office.MailboxEnums.ItemType.Appointment) return _appointmentAttendeeNames;
  if (_itemType === Office.MailboxEnums.ItemType.Message)     return _messageParticipantNames;
  return [];
}

function _filterByNames(rumbles, names) {
  return rumbles.filter(r => {
    const cn = r.contactName.toLowerCase();
    return names.some(n => { const mn = n.toLowerCase(); return mn === cn || mn.includes(cn) || cn.includes(mn); });
  });
}

async function loadAllRumbles() {
  const panel = document.getElementById('rumbleBrowsePanel');
  if (!panel) return;
  panel.innerHTML = '<div class="rumble-empty">Lade…</div>';
  try {
    const listId = await getOrCreateTembuList();
    const data = await graphFetch('GET', `/me/todo/lists/${listId}/tasks?$filter=status ne 'completed'&$top=200`);
    const dedupSeen = new Set();
    _allRumbles = (data.value || []).reduce((acc, task) => {
      const f = parseBodyFields(task.body?.content);
      const contactName = f.CONTACT || task.title.replace(/^Tembu:\s*/i, '');
      const text = f.TEXT || contactName;
      const key = f.RUMBLE_ID || f.MEMO_ID ||
        `${contactName.toLowerCase().trim()}||${text.toLowerCase().trim()}`;
      if (dedupSeen.has(key)) return acc;
      dedupSeen.add(key);
      acc.push({ contactName, text, createdAt: f.CREATED || task.createdDateTime, rumbleId: f.RUMBLE_ID || f.MEMO_ID || null });
      return acc;
    }, []);
    _rumbleLoaded = true;
    renderAllRumbles('');
  } catch (e) {
    if (panel) panel.innerHTML = `<div class="rumble-empty">Fehler: ${escapeTp(e.message)}</div>`;
  }
}

function renderAllRumbles(filter) {
  const panel = document.getElementById('rumbleBrowsePanel');
  if (!panel) return;
  const q = (filter || '').toLowerCase().trim();

  // In appointment/email view: only show Rumbles for participants
  const participants = _contextParticipants();
  let source = participants.length ? _filterByNames(_allRumbles, participants) : _allRumbles;

  const grouped = {};
  for (const r of source) {
    if (q && !r.contactName.toLowerCase().includes(q) && !r.text.toLowerCase().includes(q)) continue;
    if (!grouped[r.contactName]) grouped[r.contactName] = [];
    grouped[r.contactName].push(r);
  }
  const contacts = Object.keys(grouped).sort();
  if (!contacts.length) {
    const msg = participants.length
      ? 'Keine Rumbles für die Teilnehmer.'
      : (q ? 'Keine Treffer.' : 'Keine aktiven Rumbles.');
    panel.innerHTML = `<div class="rumble-empty">${msg}</div>`;
    return;
  }
  panel.innerHTML = contacts.map(name => {
    const rows = grouped[name].map(r =>
      `<div class="rumble-row">${escapeTp(r.text)}</div>`
    ).join('');
    return `<div class="rumble-group"><div class="rumble-group-name">${escapeTp(name)}</div>${rows}</div>`;
  }).join('');
}

function filterRumbles(val) {
  renderAllRumbles(val);
}

// ── UI helpers ────────────────────────────────────────────────────────────
function showForm() {
  document.getElementById('signInSection').classList.add('hidden');
  document.getElementById('formSection').classList.remove('hidden');
  if (_account) {
    document.getElementById('userLabel').textContent = _account.name || _account.username || '';
  }
  // Phone lookup needs _token — trigger after sign-in in case context was already loaded
  triggerPhoneLookup();
  if (_itemType === Office.MailboxEnums.ItemType.Appointment) loadMeetingBriefing();
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
  document.getElementById('contactName').addEventListener('change', e => onContactNameChange(e.target.value));
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
