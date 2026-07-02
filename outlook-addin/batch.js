// Tembu Batch-Analyse v20260701a

const GRAPH = 'https://graph.microsoft.com/v1.0';

const PERSONAL_DOMAINS = new Set([
  'gmail.com','googlemail.com','yahoo.com','yahoo.de','yahoo.co.uk',
  'hotmail.com','hotmail.de','outlook.com','live.com','live.de',
  'icloud.com','me.com','mac.com','gmx.ch','gmx.de','gmx.at','gmx.net',
  'bluewin.ch','sunrise.ch','hispeed.ch','windowslive.com',
  'protonmail.com','proton.me',
]);

// Well-known folder names to exclude by default
const EXCLUDE_DEFAULT = new Set([
  'junkemail','deleteditems','drafts','outbox',
  'recoverableitemsdeletions','recoverableitemspurges',
  'recoverableitemsroot','recoverableitemsversions',
]);

// System-style local parts to skip
const SKIP_LOCAL = [
  'no-reply','noreply','do-not-reply','donotreply','newsletter','notifications',
  'notification','mailer-daemon','bounce','unsubscribe','postmaster','abuse',
  'marketing','automated','system',
];

let _token     = '';
let _serverUrl = '';
let _cancelled = false;
let _running   = false;
let _ownEmail  = '';
let _folders   = []; // [{id, displayName, totalItemCount, checked}]

// ── Init ──────────────────────────────────────────────────────────────────────
window.addEventListener('DOMContentLoaded', async () => {
  TI18n.applyStaticI18n();
  const p = new URLSearchParams(window.location.search);
  _token     = p.get('t')   || '';
  _serverUrl = (p.get('srv') || '').trim().replace(/\/+$/, '');
  if (!_token) {
    document.getElementById('folderList').innerHTML =
      `<span class="folder-loading" style="color:#c00">${escHtml(TI18n.t('batch.noTokenError'))}</span>`;
    return;
  }
  await loadFolders();
});

// ── Load and render folder list ───────────────────────────────────────────────
async function loadFolders() {
  try {
    const data = await gFetch(`${GRAPH}/me/mailFolders?$select=id,displayName,totalItemCount&$top=50`);
    const raw  = data.value || [];

    // Also load child folders (one level deep) for completeness
    const childRequests = raw.map(f =>
      gFetch(`${GRAPH}/me/mailFolders/${f.id}/childFolders?$select=id,displayName,totalItemCount&$top=50`)
        .then(d => d.value || []).catch(() => [])
    );
    const children = await Promise.all(childRequests);

    const all = [...raw];
    children.forEach(arr => all.push(...arr));

    _folders = all
      .filter(f => f.totalItemCount > 0)
      .map(f => {
        const key = (f.displayName || '').toLowerCase().replace(/\s/g, '');
        const checked = !EXCLUDE_DEFAULT.has(key);
        return { id: f.id, displayName: f.displayName, totalItemCount: f.totalItemCount, checked };
      })
      .sort((a, b) => {
        // Inbox + SentItems first, then alphabetical
        const priority = n => n === 'Inbox' || n === 'Posteingang' || n === 'Gesendete Elemente' || n === 'SentItems' ? 0 : 1;
        return priority(a.displayName) - priority(b.displayName) || a.displayName.localeCompare(b.displayName);
      });

    renderFolderList();
    updateStartButton();
  } catch (e) {
    document.getElementById('folderList').innerHTML =
      `<span class="folder-loading" style="color:#c00">${escHtml(TI18n.t('batch.foldersLoadError', { msg: e.message }))}</span>`;
  }
}

function renderFolderList() {
  const list = document.getElementById('folderList');
  list.innerHTML = '';
  for (let i = 0; i < _folders.length; i++) {
    const f     = _folders[i];
    const label = document.createElement('label');
    label.className = 'folder-item';

    const cb = document.createElement('input');
    cb.type    = 'checkbox';
    cb.checked = f.checked;
    cb.addEventListener('change', () => toggleFolder(i, cb.checked));

    const name  = document.createElement('span');
    name.className   = 'folder-name';
    name.textContent = f.displayName;

    const count = document.createElement('span');
    count.className   = 'folder-count';
    count.textContent = TI18n.formatNumber(f.totalItemCount);

    label.appendChild(cb);
    label.appendChild(name);
    label.appendChild(count);
    list.appendChild(label);
  }
}

function toggleFolder(idx, checked) {
  _folders[idx].checked = checked;
  updateStartButton();
}

function selectAllFolders(checked) {
  _folders.forEach(f => f.checked = checked);
  renderFolderList();
  updateStartButton();
}

function updateStartButton() {
  const anySelected = _folders.some(f => f.checked);
  const btn = document.getElementById('btnStart');
  btn.disabled = !anySelected;
  document.getElementById('actionHint').textContent =
    anySelected ? TI18n.tn('batch.foldersSelectedCount', _folders.filter(f => f.checked).length) : TI18n.t('batch.selectAtLeastOne');
}

// ── Filters ───────────────────────────────────────────────────────────────────
function shouldSkip(address) {
  if (!address || !address.includes('@')) return true;
  const email  = address.toLowerCase();
  if (email === _ownEmail) return true;
  const domain = email.split('@')[1] || '';
  if (PERSONAL_DOMAINS.has(domain)) return true;
  const local  = email.split('@')[0] || '';
  if (SKIP_LOCAL.some(s => local === s || local.startsWith(s))) return true;
  return false;
}

// ── Graph helpers ─────────────────────────────────────────────────────────────
async function gFetch(url) {
  const res = await fetch(url, { headers: { Authorization: `Bearer ${_token}` } });
  if (!res.ok) {
    const body = await res.text().catch(() => '');
    throw new Error(`Graph ${res.status}: ${body.slice(0, 100)}`);
  }
  return res.json();
}

async function fetchPaginated(startUrl, limit = 5000) {
  const items = [];
  let url = startUrl;
  while (url && items.length < limit && !_cancelled) {
    let data;
    try { data = await gFetch(url); } catch { break; }
    items.push(...(data.value || []));
    url = data['@odata.nextLink'] || null;
  }
  return items;
}

// ── Phase 1: Scan selected folders, collect unique contacts ───────────────────
async function scanFolders(since) {
  const sel   = '$select=from,toRecipients,receivedDateTime,sentDateTime';
  const s     = encodeURIComponent(since.toISOString());
  const selected = _folders.filter(f => f.checked);
  const contacts = new Map(); // email → displayName

  const add = (addr, name) => {
    const email = (addr || '').toLowerCase().trim();
    if (!email || shouldSkip(email)) return;
    if (!contacts.has(email)) contacts.set(email, name || email);
  };

  for (let i = 0; i < selected.length; i++) {
    if (_cancelled) break;
    const folder = selected[i];
    document.getElementById('progressCurrent').textContent =
      TI18n.t('batch.scanningFolder', { i: i + 1, total: selected.length, name: folder.displayName });

    // Use receivedDateTime filter; for sent folder also try sentDateTime
    const isSent = (folder.displayName || '').toLowerCase().replace(/\s/g,'') === 'sentitems' ||
                   (folder.displayName || '').toLowerCase().includes('gesendet') ||
                   (folder.displayName || '').toLowerCase().includes('sent');

    const filterField = isSent ? 'sentDateTime' : 'receivedDateTime';
    const url = `${GRAPH}/me/mailFolders/${folder.id}/messages?${sel}&$filter=${filterField} ge ${s}&$top=100&$orderby=${filterField} desc`;
    const msgs = await fetchPaginated(url);

    for (const m of msgs) {
      const f = m.from?.emailAddress;
      if (f?.address) add(f.address, f.name);
      for (const r of (m.toRecipients || [])) {
        if (r.emailAddress?.address) add(r.emailAddress.address, r.emailAddress.name);
      }
    }
  }

  return contacts;
}

// ── Phase 2: Fetch emails for a specific contact ──────────────────────────────
async function fetchContactEmails(email, since) {
  const enc = encodeURIComponent;
  const sel = '$select=id,subject,receivedDateTime,sentDateTime,from,toRecipients,bodyPreview,webLink';
  const qFrom = enc(`"from:${email}"`);
  const qTo   = enc(`"to:${email}"`);

  const [inboxR, sentR] = await Promise.allSettled([
    gFetch(`${GRAPH}/me/messages?$search=${qFrom}&${sel}&$top=50`),
    gFetch(`${GRAPH}/me/mailFolders/SentItems/messages?$search=${qTo}&${sel}&$top=50`),
  ]);

  const inboxItems = inboxR.status === 'fulfilled' ? (inboxR.value?.value || []) : [];
  const sentItems  = sentR.status  === 'fulfilled' ? (sentR.value?.value  || []) : [];

  const result = [];
  for (const m of inboxItems) {
    if (m.receivedDateTime && new Date(m.receivedDateTime) < since) continue;
    result.push({
      id: m.id, date: (m.receivedDateTime || '').slice(0, 10),
      type: 'email', direction: 'received',
      subject: m.subject || TI18n.t('common.noSubject'),
      preview: (m.bodyPreview || '').slice(0, 150),
      url: m.webLink || '',
    });
  }
  for (const m of sentItems) {
    if (m.sentDateTime && new Date(m.sentDateTime) < since) continue;
    result.push({
      id: m.id, date: (m.sentDateTime || '').slice(0, 10),
      type: 'email', direction: 'sent',
      subject: m.subject || TI18n.t('common.noSubject'),
      preview: (m.bodyPreview || '').slice(0, 150),
      url: m.webLink || '',
    });
  }
  return result.sort((a, b) => b.date.localeCompare(a.date));
}

// ── Phase 2: Fetch meetings for a specific contact ────────────────────────────
async function fetchContactMeetings(email, since) {
  try {
    const enc = encodeURIComponent;
    const now = new Date().toISOString();
    const sel = '$select=id,subject,start,end,bodyPreview';
    const fil = enc(`attendees/any(a:a/emailAddress/address eq '${email}')`);
    const data = await gFetch(
      `${GRAPH}/me/calendarView?startDateTime=${enc(since.toISOString())}&endDateTime=${enc(now)}&${sel}&$filter=${fil}&$top=50`
    );
    return (data.value || []).map(m => ({
      id: m.id, date: (m.start?.dateTime || '').slice(0, 10),
      type: 'meeting', direction: 'meeting',
      subject: m.subject || TI18n.t('common.noSubject'),
      duration: Math.round((new Date(m.end?.dateTime) - new Date(m.start?.dateTime)) / 60000) || 0,
    }));
  } catch { return []; }
}

// ── AI ────────────────────────────────────────────────────────────────────────
// Läuft server-seitig über tembu-server /api/analyze (gleicher Endpoint wie detail.js).
function isRecent(savedAt) {
  if (!savedAt) return false;
  return Date.now() - new Date(savedAt).getTime() < 24 * 60 * 60 * 1000;
}

// ── Analyse one contact ───────────────────────────────────────────────────────
async function analyzeContact(email, name, since) {
  const existing = await TCore.loadAnalysis(_token, email);
  if (existing && isRecent(existing.savedAt)) return { skipped: true };

  const [emails, meetings] = await Promise.all([
    fetchContactEmails(email, since),
    fetchContactMeetings(email, since),
  ]);

  const res = await fetch(`${_serverUrl}/api/analyze`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${_token}` },
    body: JSON.stringify({
      lang: TI18n.getLang(), contactName: name, contactEmail: email,
      emails:   emails.map(e   => ({ dateStr: e.date, direction: e.direction, subject: e.subject, preview: e.preview })),
      meetings: meetings.map(m => ({ dateStr: m.date, subject: m.subject, durationMin: m.duration })),
      rumbles:  [],
    }),
  });
  if (!res.ok) {
    const body = await res.text().catch(() => '');
    throw new Error(`${res.status} ${body.slice(0, 150)}`);
  }
  const raw      = await res.text();
  const analysis = TCore.parseAIResponse(raw);
  await TCore.saveAnalysis(_token, email, name, email, analysis);

  return { skipped: false, emailCount: emails.length, meetingCount: meetings.length };
}

// ── UI helpers ────────────────────────────────────────────────────────────────
function setProgress(done, total) {
  const pct = total > 0 ? Math.round((done / total) * 100) : 0;
  document.getElementById('progressFill').style.width = pct + '%';
  document.getElementById('progressText').textContent = TI18n.t('batch.progressCount', { done, total, pct });
}

function setStat(id, val) { document.getElementById(id).textContent = val; }

function addLogItem(name, email, icon, cls, statusText) {
  const div = document.createElement('div');
  div.className = 'log-item';
  div.innerHTML = `
    <span class="log-icon">${icon}</span>
    <span class="log-name">${escHtml(name)}</span>
    <span class="log-email">${escHtml(email)}</span>
    <span class="log-status ${cls}">${statusText}</span>`;
  document.getElementById('logList').prepend(div);
}

function escHtml(s) {
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

// ── Start / cancel ────────────────────────────────────────────────────────────
async function startBatch() {
  if (_running) return;
  _running   = true;
  _cancelled = false;

  if (!_serverUrl) {
    alert(TI18n.t('batch.needServerAlert'));
    _running = false;
    return;
  }

  const months = parseInt(document.getElementById('sinceSelect').value || '12');
  const since  = months === 0 ? new Date('2000-01-01') : new Date(Date.now() - months * 30.44 * 24 * 60 * 60 * 1000);

  document.getElementById('btnStart').disabled = true;
  document.getElementById('btnCancel').classList.remove('hidden');
  document.getElementById('progressArea').classList.remove('hidden');
  document.getElementById('doneBox').classList.remove('visible');
  document.getElementById('logList').innerHTML = '';
  document.getElementById('actionHint').textContent = '';
  setStat('statDone', 0); setStat('statSkip', 0); setStat('statErr', 0);

  let done = 0, skipped = 0, errors = 0;

  try {
    // Get own email to exclude self from contact list
    try {
      const me = await gFetch(`${GRAPH}/me?$select=mail,userPrincipalName`);
      _ownEmail = (me.mail || me.userPrincipalName || '').toLowerCase();
    } catch { /* ignore */ }

    // Phase 1: scan selected folders
    document.getElementById('progressText').textContent = TI18n.t('batch.scanningFolders');
    const contacts = await scanFolders(since);
    const total = contacts.size;

    if (total === 0) {
      document.getElementById('progressText').textContent = TI18n.t('batch.noExternalContacts');
      document.getElementById('progressCurrent').textContent = '';
      finish(done, skipped, errors, total, false);
      return;
    }

    // Phase 2: analyse each contact sequentially
    let i = 0;
    for (const [email, name] of contacts) {
      if (_cancelled) break;
      setProgress(i, total);
      document.getElementById('progressCurrent').textContent = TI18n.t('batch.analyzingContact', { name });

      try {
        const result = await analyzeContact(email, name, since);
        if (result.skipped) {
          skipped++; setStat('statSkip', skipped);
          addLogItem(name, email, '⏭', 'skip', TI18n.t('batch.skippedRecent'));
        } else {
          done++; setStat('statDone', done);
          const detail = result.emailCount > 0
            ? TI18n.tn('common.mailCount', result.emailCount) + (result.meetingCount > 0 ? TI18n.t('common.middotSeparator') + TI18n.tn('common.meetingsCount', result.meetingCount) : '')
            : TI18n.t('batch.noMails');
          addLogItem(name, email, '✓', 'ok', detail);
        }
      } catch (e) {
        errors++; setStat('statErr', errors);
        addLogItem(name, email, '✗', 'err', (e.message || TI18n.t('common.genericError')).slice(0, 60));
      }
      i++;
    }

    setProgress(i, total);
    document.getElementById('progressCurrent').textContent = '';
    finish(done, skipped, errors, total, _cancelled);

  } catch (e) {
    document.getElementById('progressText').textContent = TI18n.t('common.errorPrefix', { msg: e.message || String(e) });
    document.getElementById('progressCurrent').textContent = '';
    _running = false;
    document.getElementById('btnStart').disabled = false;
    document.getElementById('btnCancel').classList.add('hidden');
    updateStartButton();
  }
}

function finish(done, skipped, errors, total, cancelled) {
  const doneBox = document.getElementById('doneBox');
  doneBox.classList.add('visible');
  document.getElementById('doneText').textContent =
    TI18n.t('batch.doneSummary', { done, skipped, errors }) +
    (cancelled ? TI18n.t('batch.doneCancelledSuffix') : TI18n.t('batch.doneTotalSuffix', { total }));
  _running = false;
  document.getElementById('btnStart').disabled = false;
  document.getElementById('btnCancel').classList.add('hidden');
  updateStartButton();
}

function cancelBatch() {
  _cancelled = true;
  document.getElementById('btnCancel').disabled = true;
  document.getElementById('progressCurrent').textContent = TI18n.t('batch.cancelling');
}
