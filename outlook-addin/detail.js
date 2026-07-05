/* Tembu Contact Intelligence – detail.js v20260701a */

const SESSION_KEY   = '@tembu_outlook_session';
const DIALOG_TK_KEY = '@tembu_dialog_token';
const DB_NAME       = 'tembu_cache_v1';
const STORE_NAME    = 'contact_analyses';
const TEMBU_LIST    = 'Tembu';
const SINCE_MONTHS_KEY     = 'tembu_since_months';
const SINCE_MONTHS_DEFAULT = 24;
const CACHE_TTL_MS  = 24 * 60 * 60 * 1000;

let _token        = null;
let _serverUrl    = '';
let _contactName  = '';
let _contactEmail = '';
let _cacheKey     = '';
let _rawData      = null;
let _emailSummaries = {}; // conversationId -> { summary, latestDate }

const esc = TCore.esc;

// ── Since-Datum ───────────────────────────────────────────────────────────
function getSinceMonths() {
  return parseInt(localStorage.getItem(SINCE_MONTHS_KEY) || String(SINCE_MONTHS_DEFAULT));
}

function getSinceDate() {
  const m = getSinceMonths();
  if (m === 0) return new Date('2000-01-01');
  return new Date(Date.now() - m * 30.44 * 24 * 60 * 60 * 1000);
}

function getSinceLabel() {
  return TI18n.sinceLabel(getSinceMonths());
}

function updateSinceLabel() {
  const el = document.getElementById('sinceLabel');
  if (el) el.textContent = TI18n.t('detail.interactionsHeading', { label: getSinceLabel() });
}

// ── Init ──────────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', async () => {
  const params    = new URLSearchParams(window.location.search);
  _contactName    = params.get('name')  || '';
  _contactEmail   = params.get('email') || '';
  _serverUrl      = (params.get('srv')  || '').trim().replace(/\/+$/, '');
  _cacheKey       = (_contactEmail || _contactName).toLowerCase().trim();

  document.getElementById('headerName').textContent = _contactName || TI18n.t('detail.headerNameFallback');

  const sinceSelect = document.getElementById('sinceSelect');
  sinceSelect.value = localStorage.getItem(SINCE_MONTHS_KEY) || String(SINCE_MONTHS_DEFAULT);
  sinceSelect.addEventListener('change', () => {
    localStorage.setItem(SINCE_MONTHS_KEY, sinceSelect.value);
    updateSinceLabel();
    loadData(true);
  });
  updateSinceLabel();

  TI18n.applyStaticI18n();

  _token = getStoredToken();
  if (!_token) {
    showError(TI18n.t('common.noTokenError'));
    return;
  }

  await loadData(false);
});

// ── Token ─────────────────────────────────────────────────────────────────
function getStoredToken() {
  // Outlook Desktop isoliert den Dialog-localStorage von der Taskpane → Token per URL-Parameter übergeben
  try {
    const t = new URLSearchParams(window.location.search).get('t');
    if (t && t.length > 10) return t;
  } catch {}
  try {
    const r = localStorage.getItem(DIALOG_TK_KEY);
    if (r) { const d = JSON.parse(r); if (d.exp > Date.now()) return d.token; }
  } catch {}
  try {
    const r = localStorage.getItem(SESSION_KEY);
    if (r) { const d = JSON.parse(r); if (d.expiry > Date.now()) return d.token; }
  } catch {}
  return null;
}

// ── OneDrive cache (dünne Wrapper über TCore.saveAnalysis / loadAnalysis) ─
async function saveToOneDrive(analysis) {
  const _odSteps = [];
  const odLog = (msg) => {
    _odSteps.push(msg);
    const base = (document.getElementById('diagPanel')?.textContent || '').split(' | OD:')[0];
    showDiag(base + ' | OD: ' + _odSteps.join(' → '));
    console.log('[OneDrive]', msg);
  };

  try {
    const mkR1 = await fetch('https://graph.microsoft.com/v1.0/me/drive/root/children', {
      method: 'POST',
      headers: { Authorization: `Bearer ${_token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ name: 'Tembu', folder: {}, '@microsoft.graph.conflictBehavior': 'fail' }),
    });
    odLog('T:' + mkR1.status);

    const mkR2 = await fetch('https://graph.microsoft.com/v1.0/me/drive/root:/Tembu:/children', {
      method: 'POST',
      headers: { Authorization: `Bearer ${_token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ name: 'analysen', folder: {}, '@microsoft.graph.conflictBehavior': 'fail' }),
    });
    odLog('A:' + mkR2.status);

    const putPath = TCore.analysisOneDrivePath(_cacheKey);
    const putRes = await fetch('https://graph.microsoft.com/v1.0' + putPath, {
      method: 'PUT',
      headers: { Authorization: `Bearer ${_token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ contact: _contactName, email: _contactEmail, savedAt: new Date().toISOString(), analysis }),
    });
    const putBody = await putRes.text();
    odLog('PUT:' + putRes.status + (putRes.ok ? '✓' : ' ' + putBody.slice(0, 80)));
  } catch (e) {
    odLog('ERR:' + (e.message || String(e)));
  }
}

async function loadFromOneDrive() {
  return TCore.loadAnalysis(_token, _cacheKey);
}

// ── IndexedDB cache ───────────────────────────────────────────────────────
function openDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = e => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains(STORE_NAME)) db.createObjectStore(STORE_NAME, { keyPath: 'id' });
    };
    req.onsuccess = e => resolve(e.target.result);
    req.onerror   = () => reject(req.error);
  });
}

async function getCached(key) {
  try {
    const db = await openDB();
    return new Promise(resolve => {
      const req = db.transaction(STORE_NAME, 'readonly').objectStore(STORE_NAME).get(key);
      req.onsuccess = () => resolve(req.result || null);
      req.onerror   = () => resolve(null);
    });
  } catch { return null; }
}

// Merged mit dem vorhandenen Datensatz statt ihn zu ersetzen — sonst wuerde z.B. ein
// saveCache({emailSummaries}) das zuvor gespeicherte analysis/rawData ueberschreiben.
async function saveCache(key, data) {
  try {
    const db = await openDB();
    return new Promise(resolve => {
      const tx = db.transaction(STORE_NAME, 'readwrite');
      const store = tx.objectStore(STORE_NAME);
      const getReq = store.get(key);
      getReq.onsuccess = () => {
        const existing = getReq.result || {};
        store.put({ ...existing, id: key, cachedAt: Date.now(), ...data });
      };
      tx.oncomplete = () => resolve();
      tx.onerror    = () => resolve();
    });
  } catch {}
}

// ── Main load ─────────────────────────────────────────────────────────────
async function loadData(force) {
  setLoading(TI18n.t('detail.loadingData'));

  const cached = await getCached(_cacheKey);
  _emailSummaries = cached?.emailSummaries || {};
  const since = getSinceDate();

  setLoading(TI18n.t('detail.loadingEmailsMeetings'));
  let emails = [], meetings = [], rumbles = [];
  try {
    [emails, meetings, rumbles] = await Promise.all([
      fetchEmails(since),
      fetchMeetings(since),
      fetchRumbles(),
    ]);
  } catch (e) {
    const msg = e.message || String(e);
    if (msg.includes('403')) {
      showError(TI18n.t('detail.permissionError403'));
    } else {
      showError(TI18n.t('detail.loadError', { msg }));
    }
    return;
  }

  _rawData = { emails, meetings, rumbles };

  // Merge mit gecachten Daten für inkrementelles Update
  if (cached && !force && cached.rawData) {
    const seenEmails = new Set(cached.rawData.emails.map(e => e.id).filter(Boolean));
    const newEmails  = emails.filter(e => !e.id || !seenEmails.has(e.id));
    _rawData.emails  = [...newEmails, ...cached.rawData.emails].slice(0, 200);

    const seenMeet  = new Set(cached.rawData.meetings.map(m => m.id).filter(Boolean));
    const newMeet   = meetings.filter(m => !m.id || !seenMeet.has(m.id));
    _rawData.meetings = [...newMeet, ...cached.rawData.meetings].slice(0, 100);
  }

  renderStats(_rawData);
  renderTimeline(_rawData);
  updateHeader(_rawData);
  showContent();

  if (cached && !force && cached.analysis) {
    const ageH = Math.round((Date.now() - cached.cachedAt) / 3600000);
    document.getElementById('cacheInfoText').textContent =
      TI18n.tn('common.cacheAgeHours', ageH, { n: ageH < 1 ? '< 1' : ageH });
    document.getElementById('cacheInfo').classList.remove('hidden');
    renderAiAnalysis(cached.analysis);
    renderThemes(cached.analysis.themes || []);
    renderBackground(cached.analysis.background || '');
    return;
  }

  // Kein lokaler Cache → OneDrive versuchen
  if (!force) {
    const cloud = await loadFromOneDrive();
    if (cloud?.analysis) {
      await saveCache(_cacheKey, { analysis: cloud.analysis, rawData: _rawData });
      const ageH = Math.round((Date.now() - new Date(cloud.savedAt).getTime()) / 3600000);
      document.getElementById('cacheInfoText').textContent =
        TI18n.tn('common.cacheAgeHours', ageH, { n: ageH < 1 ? '< 1' : ageH }) + TI18n.t('common.cloudSuffix');
      document.getElementById('cacheInfo').classList.remove('hidden');
      renderAiAnalysis(cloud.analysis);
      renderThemes(cloud.analysis.themes || []);
      renderBackground(cloud.analysis.background || '');
      return;
    }
  }

  await runAI(_rawData);
}

async function forceRefresh() {
  _rawData = null;
  document.getElementById('cacheInfo').classList.add('hidden');
  await loadData(true);
}


// ── Graph helper (thin wrapper mit 403-Spezialbehandlung) ─────────────────
async function gFetch(path) {
  try {
    return await TCore.graphGet(_token, path);
  } catch (e) {
    if (e.message?.startsWith('Graph 403')) {
      throw new Error(TI18n.t('detail.permissionError403Short'));
    }
    throw e;
  }
}

// ── Fetch emails ──────────────────────────────────────────────────────────
async function fetchEmails(since, top = 100) {
  if (!_contactEmail && !_contactName) return [];
  const s   = since.toISOString();
  const enc = encodeURIComponent;
  const selEmail = '$select=id,subject,receivedDateTime,sentDateTime,from,toRecipients,bodyPreview,webLink,conversationId';
  const sinceDate = s.slice(0, 10);
  let diagMode = '', diagRaw = 0, diagFiltered = 0;
  try {
    let inboxReq, sentReq;
    if (_contactEmail) {
      diagMode = 'email-search';
      const qFrom = enc('"from:' + _contactEmail + '"');
      const qTo   = enc('"to:'   + _contactEmail + '"');
      inboxReq = gFetch(`/me/messages?$search=${qFrom}&${selEmail}&$top=${top}`);
      sentReq  = gFetch(`/me/mailFolders/SentItems/messages?$search=${qTo}&${selEmail}&$top=${top}`);
    } else {
      diagMode = 'name-search';
      const cleanName = _contactName.replace(/^(Herr|Frau|Hr\.|Fr\.|Dr\.|Prof\.|Ing\.|Mag\.)\s+/i, '').trim();
      const q = enc('"' + cleanName + '"');
      inboxReq = gFetch(`/me/messages?$search=${q}&${selEmail}&$top=${top}`);
      sentReq  = gFetch(`/me/mailFolders/SentItems/messages?$search=${q}&${selEmail}&$top=${top}`);
    }
    const [inbox, sent] = await Promise.allSettled([inboxReq, sentReq]);
    const result = [];
    const nameLower = _contactName.toLowerCase()
      .replace(/^(herr|frau|hr\.|fr\.|dr\.|prof\.|ing\.|mag\.|dipl\.|lic\.)\s+/gi, '').trim();

    const inboxItems = inbox.status === 'fulfilled' ? (inbox.value?.value || []) : [];
    const sentItems  = sent.status  === 'fulfilled' ? (sent.value?.value  || []) : [];
    diagRaw = inboxItems.length + sentItems.length;

    if (inbox.status === 'rejected') {
      const err = inbox.reason?.message || '';
      if (err.includes('403')) throw inbox.reason;
      diagMode += ' inbox-err:' + err.slice(0, 60);
    }
    if (sent.status === 'rejected') {
      const err = sent.reason?.message || '';
      diagMode += ' sent-err:' + err.slice(0, 60);
    }

    for (const m of inboxItems) {
      const date = (m.receivedDateTime || '').slice(0, 10);
      if (date < sinceDate) continue;
      if (!_contactEmail) {
        const fromName = (m.from?.emailAddress?.name || '').toLowerCase();
        if (!fromName.includes(nameLower)) continue;
      }
      result.push({ id: m.id, date, type: 'email', direction: 'received',
        subject: m.subject || TI18n.t('common.noSubject'), preview: (m.bodyPreview || '').slice(0, 200),
        fromEmail: m.from?.emailAddress?.address || '',
        fromName:  m.from?.emailAddress?.name    || '',
        conversationId: m.conversationId || '',
        url: m.webLink || '' });
    }
    for (const m of sentItems) {
      const date = (m.sentDateTime || '').slice(0, 10);
      if (date < sinceDate) continue;
      if (!_contactEmail) {
        const toNames = (m.toRecipients || []).map(r => (r.emailAddress?.name || '').toLowerCase());
        if (!toNames.some(n => n.includes(nameLower))) continue;
      }
      result.push({ id: m.id, date, type: 'email', direction: 'sent',
        subject: m.subject || TI18n.t('common.noSubject'), preview: (m.bodyPreview || '').slice(0, 200),
        conversationId: m.conversationId || '',
        url: m.webLink || '' });
    }
    diagFiltered = result.length;

    if (!_contactEmail && result.length > 0 && result[0].fromEmail) {
      _contactEmail = result[0].fromEmail;
    }

    showDiag(`js:20260701a | E-Mail: ${diagMode} | roh:${diagRaw} → gefiltert:${diagFiltered} | name="${_contactName}" email="${_contactEmail || '—'}" | seit:${sinceDate}`);

    return result.sort((a, b) => b.date.localeCompare(a.date));
  } catch (e) {
    showDiag(`Email error (${diagMode}, raw:${diagRaw}): ${e.message}`);
    throw e;
  }
}

function showDiag(msg) {
  let el = document.getElementById('diagPanel');
  if (!el) {
    el = document.createElement('div');
    el.id = 'diagPanel';
    el.style.cssText = 'font-size:10px;color:#888;background:#f8f8f8;border-top:1px solid #eee;padding:6px 20px;word-break:break-all;flex-shrink:0;';
    document.body.appendChild(el);
  }
  el.textContent = msg;
}

// ── Fetch meetings ────────────────────────────────────────────────────────
async function fetchMeetings(since, top = 100) {
  if (!_contactEmail && !_contactName) return [];
  const s   = since.toISOString();
  const now = new Date().toISOString();
  const enc = encodeURIComponent;
  const filter = _contactEmail
    ? `attendees/any(a:a/emailAddress/address eq '${_contactEmail}')`
    : `attendees/any(a:a/emailAddress/name eq '${_contactName}')`;
  try {
    const data = await gFetch(
      `/me/calendarView?startDateTime=${enc(s)}&endDateTime=${enc(now)}&$filter=${enc(filter)}&$select=id,subject,start,end,bodyPreview&$top=${top}`
    );
    return (data.value || []).map(e => ({
      id: e.id, type: 'meeting', date: (e.start?.dateTime || '').slice(0, 10),
      subject: e.subject || TI18n.t('common.noTitle'),
      duration: (e.start?.dateTime && e.end?.dateTime)
        ? Math.round((new Date(e.end.dateTime) - new Date(e.start.dateTime)) / 60000) : 0,
      preview: (e.bodyPreview || '').slice(0, 200),
      url: e.webLink || '',
    })).sort((a, b) => b.date.localeCompare(a.date));
  } catch (e) { console.warn('fetchMeetings:', e.message); return []; }
}

// ── Fetch Rumbles ─────────────────────────────────────────────────────────
async function fetchRumbles() {
  if (!_contactName) return [];
  try {
    const lists    = await gFetch('/me/todo/lists');
    const tembuLst = (lists.value || []).find(l => l.displayName === TEMBU_LIST);
    if (!tembuLst) return [];
    const data  = await gFetch(`/me/todo/lists/${tembuLst.id}/tasks?$top=200`);
    const cn    = _contactName.toLowerCase().trim();
    return (data.value || []).filter(t => {
      const f    = TCore.parseBody(t.body?.content || '');
      const name = (f.CONTACT || t.title.replace(/^Tembu:\s*/i, '')).toLowerCase().trim();
      return name === cn || name.includes(cn) || cn.includes(name);
    }).map(t => {
      const f = TCore.parseBody(t.body?.content || '');
      return {
        id: t.id, type: 'rumble',
        date: (f.CREATED || t.createdDateTime || '').slice(0, 10),
        subject: f.TEXT || t.title.replace(/^Tembu:\s*/i, ''),
        preview: '',
      };
    }).sort((a, b) => b.date.localeCompare(a.date));
  } catch (e) { console.warn('fetchRumbles:', e.message); return []; }
}

// ── AI ────────────────────────────────────────────────────────────────────
// Läuft komplett server-seitig über tembu-server /api/analyze — der Admin
// verwaltet Provider/Key zentral in ai-settings.json, der Client braucht keinen Key mehr.
function toContactData(data) {
  return {
    lang:         TI18n.getLang(),
    contactName:  _contactName,
    contactEmail: _contactEmail,
    emails:   data.emails.map(e   => ({ dateStr: e.date, direction: e.direction, subject: e.subject, preview: e.preview })),
    meetings: data.meetings.map(m => ({ dateStr: m.date, subject: m.subject, durationMin: m.duration })),
    rumbles:  data.rumbles.map(r  => ({ dateStr: r.date, subject: r.subject })),
  };
}

function showAiUnavailable(descKey, descArgs) {
  const desc = document.getElementById('noKeyDesc');
  if (desc) desc.innerHTML = TI18n.t(descKey, descArgs);
  document.getElementById('noKeyBox').classList.remove('hidden');
}

async function runAI(data) {
  if (!_serverUrl) {
    showAiUnavailable('detail.noKeyDesc');
    renderThemesEmpty();
    return;
  }

  setLoading(TI18n.t('common.aiAnalyzing'));
  try {
    const res = await fetch(`${_serverUrl}/api/analyze`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${_token}` },
      body: JSON.stringify(toContactData(data)),
    });
    if (!res.ok) {
      const body = await res.text().catch(() => '');
      throw new Error(`${res.status} ${body.slice(0, 150)}`);
    }
    const raw    = await res.text();
    const parsed = TCore.parseAIResponse(raw);
    showDiag(document.getElementById('diagPanel')?.textContent + ` | AI: ${raw.length}ch themes:${parsed.themes?.length ?? '?'} raw:${raw.slice(0, 80).replace(/\n/g, ' ')}`);
    renderAiAnalysis(parsed);
    renderThemes(parsed.themes || []);
    renderBackground(parsed.background || '');
    showContent();
    await saveCache(_cacheKey, { analysis: parsed, rawData: data });
    saveToOneDrive(parsed); // fire and forget — cross-device sync
    document.getElementById('cacheInfoText').textContent = TI18n.t('common.justAnalyzed');
    document.getElementById('cacheInfo').classList.remove('hidden');
  } catch (e) {
    showContent();
    const errMsg = String(e.message || e);
    showDiag((document.getElementById('diagPanel')?.textContent || '') + ` | AI error: ${errMsg.slice(0, 100)}`);
    showAiUnavailable('detail.noKeyDesc');
    const desc = document.getElementById('noKeyDesc');
    if (desc) desc.innerHTML += '<br/><br/>' + TI18n.t('common.errorPrefix', { msg: esc(errMsg) });
    document.getElementById('backgroundContent').innerHTML =
      `<div class="empty-state">${esc(TI18n.t('common.analysisFailed'))}</div>`;
    renderThemesEmpty(TI18n.t('detail.analysisFailedThemes'));
  }
}

// ── Rendering ─────────────────────────────────────────────────────────────
function fmtDate(d) {
  if (!d) return '';
  try { return TI18n.formatDate(new Date(d), { day: '2-digit', month: '2-digit', year: 'numeric' }); }
  catch { return d; }
}

function updateHeader(data) {
  const all   = [...data.emails, ...data.meetings, ...data.rumbles].filter(i => i.date).sort((a, b) => b.date.localeCompare(a.date));
  const parts = [TI18n.t('detail.sinceRangeLabel', { label: getSinceLabel() })];
  if (all.length) {
    const last = all[0];
    const lbl  = last.type === 'email' ? TI18n.t('detail.typeMail') : last.type === 'meeting' ? TI18n.t('detail.typeMeeting') : TI18n.t('detail.typeRumble');
    parts.push(TI18n.t('detail.lastContact', { date: fmtDate(last.date), type: lbl }));
  }
  if (data.meetings.length) parts.push(TI18n.t('detail.lastMeeting', { date: fmtDate(data.meetings[0].date) }));
  document.getElementById('headerMeta').innerHTML = parts.map(p => `<span>${esc(p)}</span>`).join('');
}

function renderStats(data) {
  document.getElementById('statEmails').textContent   = data.emails.length;
  document.getElementById('statMeetings').textContent = data.meetings.length;
  document.getElementById('statRumbles').textContent  = data.rumbles.length;
}

function renderAiAnalysis(a) {
  if (!a) return;
  document.getElementById('aiBox').classList.remove('hidden');
  document.getElementById('noKeyBox').classList.add('hidden');
  document.getElementById('aiSummaryText').textContent = a.summary || '';

  if (a.nextStep) {
    document.getElementById('aiNextStepText').textContent = a.nextStep;
    document.getElementById('aiNextStep').classList.remove('hidden');
  }

  const sentMap = {
    positive: ['s-pos', TI18n.t('common.sentimentPositive')],
    neutral:  ['s-neu', TI18n.t('common.sentimentNeutral')],
    negative: ['s-neg', TI18n.t('common.sentimentNegative')],
  };
  const [cls, lbl] = sentMap[a.sentiment] || sentMap.neutral;
  const meta = document.getElementById('headerMeta');
  meta.querySelectorAll('.sentiment').forEach(el => el.remove());
  meta.insertAdjacentHTML('beforeend', `<span class="sentiment"><span class="s-dot ${cls}"></span>${esc(lbl)}</span>`);

  if (a.openPoints?.length) {
    document.getElementById('openPointsBox').classList.remove('hidden');
    document.getElementById('openPointsList').innerHTML = a.openPoints.map(p => `<div class="open-point">${esc(p)}</div>`).join('');
  }
}

function renderTimeline(data) {
  const items = [
    ...data.emails,
    ...data.meetings,
    ...data.rumbles,
  ].sort((a, b) => (b.date || '').localeCompare(a.date || ''));

  const list = document.getElementById('timelineList');
  if (!items.length) {
    list.innerHTML = `<div class="empty-state">${esc(TI18n.t('detail.interactionsEmpty', { label: getSinceLabel() }))}</div>`;
    return;
  }

  list.innerHTML = items.map(item => {
    let icon, iconCls, badgeCls, badgeTxt;
    if (item.type === 'email') {
      icon     = item.direction === 'received' ? '📧' : '📤';
      iconCls  = item.direction === 'received' ? 'tl-email-in' : 'tl-email-out';
      badgeCls = item.direction === 'received' ? 'b-email-in' : 'b-email-out';
      badgeTxt = item.direction === 'received' ? TI18n.t('detail.receivedBadge') : TI18n.t('detail.sentBadge');
    } else if (item.type === 'meeting') {
      icon = '📅'; iconCls = 'tl-meeting'; badgeCls = 'b-meeting';
      badgeTxt = TI18n.t('detail.meetingBadge') + (item.duration ? ' · ' + item.duration + 'min' : '');
    } else {
      icon = '📝'; iconCls = 'tl-rumble'; badgeCls = 'b-rumble'; badgeTxt = TI18n.t('detail.rumbleBadge');
    }
    return `<div class="timeline-item">
      <div class="tl-icon ${iconCls}">${icon}</div>
      <div class="tl-body">
        <div class="tl-header">
          <span class="tl-date">${fmtDate(item.date)}</span>
          <span class="tl-badge ${badgeCls}">${badgeTxt}</span>
        </div>
        <div class="tl-subject">${esc(item.subject)}</div>
        ${item.preview ? `<div class="tl-preview">${esc(item.preview)}</div>` : ''}
      </div>
    </div>`;
  }).join('');
}

function toggleThemeDetail(el) {
  const detail = el.closest('.theme-card').querySelector('.theme-interactions');
  if (!detail) return;
  const open = detail.classList.toggle('hidden') === false;
  el.querySelector('.theme-toggle-arrow').textContent = open ? '▼' : '▶';
}

function findInteraction(item) {
  if (!_rawData) return null;
  const pool = item.type === 'meeting' ? _rawData.meetings : _rawData.emails;
  const subj = (item.subject || '').toLowerCase().slice(0, 30);
  return pool.find(e => e.date === item.date && (e.subject || '').toLowerCase().includes(subj))
      || pool.find(e => e.date === item.date);
}

function renderThemes(themes) {
  const el = document.getElementById('themesContent');
  if (!themes?.length) { renderThemesEmpty(); return; }
  el.innerHTML = themes.map(t => {
    const sc  = t.status === 'open' ? 's-open' : 's-done';
    const stx = t.status === 'open' ? TI18n.t('common.statusOpen') : TI18n.t('common.statusDone');
    const items = (t.interactions || []);
    const detailHtml = items.map(i => {
      const icon   = i.type === 'meeting' ? '📅' : '✉';
      const match  = findInteraction(i);
      const inner  = `<span class="ti-icon">${icon}</span><span class="ti-date">${esc(i.date || '')}</span><span class="ti-subject">${esc(i.subject || '')}</span>`;
      return match?.url
        ? `<a class="theme-interaction ti-link" href="${esc(match.url)}" target="_blank">${inner}</a>`
        : `<div class="theme-interaction">${inner}</div>`;
    }).join('');
    const countEl = items.length
      ? `<span class="theme-count theme-count-toggle" onclick="toggleThemeDetail(this)">${esc(TI18n.tn('common.interactionCount', items.length))} <span class="theme-toggle-arrow">▶</span></span>`
      : '';
    return `<div class="theme-card">
      <div class="theme-header">
        <span class="theme-name">${esc(t.name)}</span>
        ${countEl}
        <span class="theme-status ${sc}">${stx}</span>
      </div>
      <div class="theme-summary">${esc(t.summary)}</div>
      ${detailHtml ? `<div class="theme-interactions hidden">${detailHtml}</div>` : ''}
    </div>`;
  }).join('');
}

function renderThemesEmpty(msg) {
  const text = msg || (_serverUrl
    ? TI18n.t('detail.themesEmptyHasKey')
    : TI18n.t('detail.themesEmptyNoKey'));
  document.getElementById('themesContent').innerHTML = `<div class="empty-state">${text}</div>`;
}

function renderBackground(text) {
  const el = document.getElementById('backgroundContent');
  if (!text) { el.innerHTML = `<div class="empty-state">${esc(TI18n.t('detail.backgroundEmpty'))}</div>`; return; }
  el.innerHTML = `<div class="bg-box">${esc(text)}</div>`;
}

// ── Tab switching ─────────────────────────────────────────────────────────
function showTab(name) {
  ['Timeline', 'Themes', 'Background', 'Emails'].forEach(t => {
    document.getElementById(`tab${t}`).classList.toggle('hidden', t !== name);
    document.getElementById(`tabBtn${t}`).classList.toggle('active', t === name);
  });
  if (name === 'Emails') ensureEmailsTab();
}

// ── Emails-Tab: nach Unterhaltung gruppiert, KI-Zusammenfassung gecacht ────
// Neu zusammengefasst wird eine Unterhaltung nur, wenn seit der gespeicherten
// Zusammenfassung neuere Mails dazugekommen sind (latestDate-Vergleich) — sonst kein
// erneuter LLM-Aufruf beim naechsten Oeffnen des Kontakts.
const CONV_SUBJECT_PREFIX_RE = /^(AW|RE|FW|FWD|WG)\s*:\s*/i;

function cleanSubject(subject) {
  let s = (subject || '').trim(), prev;
  do { prev = s; s = s.replace(CONV_SUBJECT_PREFIX_RE, '').trim(); } while (s !== prev);
  return s || TI18n.t('common.noSubject');
}

function groupConversations(emails) {
  const groups = new Map();
  for (const e of emails) {
    const key = e.conversationId || ('single:' + e.id);
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(e);
  }
  return [...groups.entries()]
    .map(([id, items]) => {
      const sorted = items.slice().sort((a, b) => b.date.localeCompare(a.date));
      return { id, subject: cleanSubject(sorted[0].subject), emails: sorted, latestDate: sorted[0].date };
    })
    .sort((a, b) => b.latestDate.localeCompare(a.latestDate));
}

function parseConversationSummaries(raw) {
  let text = raw.trim().replace(/^```(?:json)?\n?/, '').replace(/\n?```$/, '').trim();
  try { return JSON.parse(text); } catch {}
  const m = text.match(/\{[\s\S]*\}/);
  if (m) { try { return JSON.parse(m[0]); } catch {} }
  return { summaries: [] };
}

async function ensureEmailsTab() {
  if (!_rawData) return;
  const conversations = groupConversations(_rawData.emails);
  renderEmailsTab(conversations);

  const stale = conversations.filter(c => {
    const cached = _emailSummaries[c.id];
    return !cached || cached.latestDate < c.latestDate;
  });
  if (!stale.length || !_serverUrl) return;

  renderEmailsTab(conversations, /* loadingSummaries */ true, stale);
  try {
    const res = await fetch(`${_serverUrl}/api/analyze/conversations`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${_token}` },
      body: JSON.stringify({
        lang: TI18n.getLang(),
        contactName: _contactName,
        contactEmail: _contactEmail,
        conversations: stale.map(c => ({
          id: c.id,
          subject: c.subject,
          emails: c.emails.slice(0, 15).map(e => ({ dateStr: e.date, direction: e.direction, subject: e.subject, preview: e.preview })),
        })),
      }),
    });
    if (res.ok) {
      const parsed = parseConversationSummaries(await res.text());
      const staleLatest = new Map(stale.map(c => [c.id, c.latestDate]));
      (parsed.summaries || []).forEach(s => {
        if (staleLatest.has(s.id)) _emailSummaries[s.id] = { summary: s.summary, latestDate: staleLatest.get(s.id) };
      });
      await saveCache(_cacheKey, { emailSummaries: _emailSummaries });
    }
  } catch (e) {
    console.warn('[tembu] Unterhaltungs-Zusammenfassung fehlgeschlagen', e);
  }
  renderEmailsTab(conversations);
}

function renderEmailsTab(conversations, loadingSummaries, staleList) {
  const el = document.getElementById('emailsContent');
  if (!conversations.length) {
    el.innerHTML = `<div class="empty-state">${esc(TI18n.t('detail.emailsEmpty'))}</div>`;
    return;
  }
  const stalePending = new Set((staleList || []).map(c => c.id));

  el.innerHTML = conversations.map(c => {
    const cached = _emailSummaries[c.id];
    let summaryHtml = '';
    if (cached) {
      summaryHtml = `<div class="conv-summary">${esc(cached.summary)}</div>`;
    } else if (loadingSummaries && stalePending.has(c.id)) {
      summaryHtml = `<div class="conv-summary loading">${esc(TI18n.t('detail.summaryLoading'))}</div>`;
    }
    const emailRows = c.emails.map(item => `<div class="timeline-item">
        <div class="tl-icon ${item.direction === 'received' ? 'tl-email-in' : 'tl-email-out'}">${item.direction === 'received' ? '📧' : '📤'}</div>
        <div class="tl-body">
          <div class="tl-header">
            <span class="tl-date">${fmtDate(item.date)}</span>
            <span class="tl-badge ${item.direction === 'received' ? 'b-email-in' : 'b-email-out'}">${item.direction === 'received' ? esc(TI18n.t('detail.receivedBadge')) : esc(TI18n.t('detail.sentBadge'))}</span>
          </div>
          <div class="tl-subject">${esc(item.subject)}</div>
          ${item.preview ? `<div class="tl-preview">${esc(item.preview)}</div>` : ''}
        </div>
      </div>`).join('');
    return `<div class="conv-card">
      <div class="conv-header" onclick="toggleConversation(this)">
        <div class="conv-main">
          <div class="conv-subject">${esc(c.subject)}</div>
          ${summaryHtml}
        </div>
        <div class="conv-meta">
          <span class="conv-count">${esc(TI18n.tn('detail.emailCount', c.emails.length))}</span>
          <span class="conv-toggle-arrow">▶</span>
        </div>
      </div>
      <div class="conv-emails hidden">${emailRows}</div>
    </div>`;
  }).join('');
}

function toggleConversation(headerEl) {
  const card = headerEl.closest('.conv-card');
  const body = card.querySelector('.conv-emails');
  const open = body.classList.toggle('hidden') === false;
  headerEl.querySelector('.conv-toggle-arrow').textContent = open ? '▼' : '▶';
}

// ── UI state ──────────────────────────────────────────────────────────────
function setLoading(text) {
  document.getElementById('tlLoading').classList.remove('hidden');
  document.getElementById('tlContent').classList.add('hidden');
  document.getElementById('loadingText').textContent = text;
}

function showContent() {
  document.getElementById('tlLoading').classList.add('hidden');
  document.getElementById('tlContent').classList.remove('hidden');
}

function showError(msg) {
  document.getElementById('tlLoading').innerHTML = `<div class="error-box">${esc(msg)}</div>`;
}
