/* Tembu Contact Intelligence – detail.js */

const SESSION_KEY   = '@tembu_outlook_session';
const DIALOG_TK_KEY = '@tembu_dialog_token';
const SETTINGS_KEY  = '@tembu_detail_settings';
const DB_NAME       = 'tembu_cache_v1';
const STORE_NAME    = 'contact_analyses';
const TEMBU_LIST    = 'Tembu';
const SINCE_MS      = 2 * 365 * 24 * 60 * 60 * 1000; // 2 years
const CACHE_TTL_MS  = 24 * 60 * 60 * 1000;            // 24 hours

let _token        = null;
let _contactName  = '';
let _contactEmail = '';
let _cacheKey     = '';
let _rawData      = null;

// ── Init ──────────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', async () => {
  const params    = new URLSearchParams(window.location.search);
  _contactName    = params.get('name')  || '';
  _contactEmail   = params.get('email') || '';
  _cacheKey       = (_contactEmail || _contactName).toLowerCase().trim();

  document.getElementById('headerName').textContent = _contactName || 'Kontakt';

  const settings = loadSettings();
  if (settings.apiKey) document.getElementById('apiKeyInput').value = settings.apiKey;

  document.getElementById('btnSaveKey').addEventListener('click', onSaveKey);

  _token = getStoredToken();
  if (!_token) {
    showError('Kein Token. Bitte in der Taskpane abmelden und neu anmelden.');
    return;
  }

  await loadData(false);
});

function onSaveKey() {
  const key = document.getElementById('apiKeyInput').value.trim();
  saveSettings({ apiKey: key });
  if (_rawData) runAI(_rawData);
}

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

// ── Settings ──────────────────────────────────────────────────────────────
function loadSettings() {
  try { return JSON.parse(localStorage.getItem(SETTINGS_KEY) || '{}'); } catch { return {}; }
}
function saveSettings(patch) {
  try { localStorage.setItem(SETTINGS_KEY, JSON.stringify({ ...loadSettings(), ...patch })); } catch {}
}
function getApiKey() { return loadSettings().apiKey || ''; }

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

async function saveCache(key, data) {
  try {
    const db = await openDB();
    return new Promise(resolve => {
      const tx = db.transaction(STORE_NAME, 'readwrite');
      tx.objectStore(STORE_NAME).put({ id: key, cachedAt: Date.now(), ...data });
      tx.oncomplete = () => resolve();
      tx.onerror    = () => resolve();
    });
  } catch {}
}

// ── Main load ─────────────────────────────────────────────────────────────
async function loadData(force) {
  setLoading('Daten werden geladen…');

  const cached = await getCached(_cacheKey);
  const since  = (cached && !force)
    ? new Date(cached.cachedAt - 60 * 60 * 1000) // 1h overlap
    : new Date(Date.now() - SINCE_MS);

  setLoading('E-Mails und Termine werden geladen…');
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
      showError('Fehlende Berechtigung (403): Bitte in der Taskpane abmelden und neu anmelden (Mail.Read + Calendars.Read).');
    } else {
      showError('Fehler beim Laden: ' + msg);
    }
    return;
  }

  _rawData = { emails, meetings, rumbles };

  // Merge with cached raw data for incremental update
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
      `KI-Analyse von vor ${ageH < 1 ? '< 1' : ageH} Stunde${ageH !== 1 ? 'n' : ''}`;
    document.getElementById('cacheInfo').classList.remove('hidden');
    renderAiAnalysis(cached.analysis);
    renderThemes(cached.analysis.themes || []);
    renderBackground(cached.analysis.background || '');
    return;
  }

  await runAI(_rawData);
}

async function forceRefresh() {
  _rawData = null;
  document.getElementById('cacheInfo').classList.add('hidden');
  await loadData(true);
}

// ── Graph helpers ─────────────────────────────────────────────────────────
async function gFetch(path) {
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    headers: { Authorization: `Bearer ${_token}` },
  });
  if (res.status === 403) throw new Error('403: Bitte in der Taskpane abmelden und neu anmelden (neue Berechtigungen erforderlich).');
  if (!res.ok) throw new Error(`Graph ${res.status}`);
  return res.json();
}

// ── Fetch emails ──────────────────────────────────────────────────────────
async function fetchEmails(since) {
  if (!_contactEmail && !_contactName) return [];
  const s   = since.toISOString();
  const enc = encodeURIComponent;
  const selEmail = '$select=id,subject,receivedDateTime,sentDateTime,from,toRecipients,bodyPreview';
  const sinceDate = s.slice(0, 10);
  let diagMode = '', diagRaw = 0, diagFiltered = 0;
  try {
    let inboxReq, sentReq;
    if (_contactEmail) {
      diagMode = 'email-filter';
      inboxReq = gFetch(`/me/messages?$filter=${enc(`from/emailAddress/address eq '${_contactEmail}' and receivedDateTime ge ${s}`)}&${selEmail}&$top=100`);
      sentReq  = gFetch(`/me/mailFolders/SentItems/messages?$filter=${enc(`toRecipients/any(r:r/emailAddress/address eq '${_contactEmail}') and sentDateTime ge ${s}`)}&${selEmail}&$top=100`);
    } else {
      diagMode = 'name-search';
      const q = enc('"' + _contactName + '"');
      inboxReq = gFetch(`/me/messages?$search=${q}&${selEmail}&$top=100`);
      sentReq  = gFetch(`/me/mailFolders/SentItems/messages?$search=${q}&${selEmail}&$top=100`);
    }
    const [inbox, sent] = await Promise.allSettled([inboxReq, sentReq]);
    const result = [];
    const nameLower = _contactName.toLowerCase();

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
        subject: m.subject || '(kein Betreff)', preview: (m.bodyPreview || '').slice(0, 200),
        fromEmail: m.from?.emailAddress?.address || '',
        fromName:  m.from?.emailAddress?.name    || '' });
    }
    for (const m of sentItems) {
      const date = (m.sentDateTime || '').slice(0, 10);
      if (date < sinceDate) continue;
      if (!_contactEmail) {
        const toNames = (m.toRecipients || []).map(r => (r.emailAddress?.name || '').toLowerCase());
        if (!toNames.some(n => n.includes(nameLower))) continue;
      }
      result.push({ id: m.id, date, type: 'email', direction: 'sent',
        subject: m.subject || '(kein Betreff)', preview: (m.bodyPreview || '').slice(0, 200) });
    }
    diagFiltered = result.length;

    if (!_contactEmail && result.length > 0 && result[0].fromEmail) {
      _contactEmail = result[0].fromEmail;
    }

    // Diagnose-Info in UI einblenden
    showDiag(`E-Mail: ${diagMode} | roh:${diagRaw} → gefiltert:${diagFiltered} | name="${_contactName}" email="${_contactEmail || '—'}" | seit:${sinceDate}`);

    return result.sort((a, b) => b.date.localeCompare(a.date));
  } catch (e) {
    showDiag(`E-Mail Fehler (${diagMode}, roh:${diagRaw}): ${e.message}`);
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
async function fetchMeetings(since) {
  if (!_contactEmail && !_contactName) return [];
  const s   = since.toISOString();
  const now = new Date().toISOString();
  const enc = encodeURIComponent;
  const filter = _contactEmail
    ? `attendees/any(a:a/emailAddress/address eq '${_contactEmail}')`
    : `attendees/any(a:a/emailAddress/name eq '${_contactName}')`;
  try {
    const data = await gFetch(
      `/me/calendarView?startDateTime=${enc(s)}&endDateTime=${enc(now)}&$filter=${enc(filter)}&$select=id,subject,start,end,bodyPreview&$top=100`
    );
    return (data.value || []).map(e => ({
      id: e.id, type: 'meeting', date: (e.start?.dateTime || '').slice(0, 10),
      subject: e.subject || '(kein Titel)',
      duration: (e.start?.dateTime && e.end?.dateTime)
        ? Math.round((new Date(e.end.dateTime) - new Date(e.start.dateTime)) / 60000) : 0,
      preview: (e.bodyPreview || '').slice(0, 200),
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
      const f    = parseFields(t.body?.content || '');
      const name = (f.CONTACT || t.title.replace(/^Tembu:\s*/i, '')).toLowerCase().trim();
      return name === cn || name.includes(cn) || cn.includes(name);
    }).map(t => {
      const f = parseFields(t.body?.content || '');
      return {
        id: t.id, type: 'rumble',
        date: (f.CREATED || t.createdDateTime || '').slice(0, 10),
        subject: f.TEXT || t.title.replace(/^Tembu:\s*/i, ''),
        preview: '',
      };
    }).sort((a, b) => b.date.localeCompare(a.date));
  } catch (e) { console.warn('fetchRumbles:', e.message); return []; }
}

function parseFields(text) {
  const r = {};
  for (const line of (text || '').split('\n')) {
    const i = line.indexOf(':');
    if (i > 0) r[line.slice(0, i).trim()] = line.slice(i + 1).trim();
  }
  return r;
}

// ── AI ────────────────────────────────────────────────────────────────────
async function runAI(data) {
  const apiKey = getApiKey();
  if (!apiKey) {
    document.getElementById('noKeyBox').classList.remove('hidden');
    renderThemesEmpty();
    return;
  }

  setLoading('KI analysiert…');
  try {
    const raw    = await callAI(buildPrompt(data), apiKey);
    const parsed = parseAIResponse(raw);
    renderAiAnalysis(parsed);
    renderThemes(parsed.themes || []);
    renderBackground(parsed.background || '');
    showContent();
    await saveCache(_cacheKey, { analysis: parsed, rawData: data });
    document.getElementById('cacheInfoText').textContent = 'Gerade analysiert';
    document.getElementById('cacheInfo').classList.remove('hidden');
  } catch (e) {
    showContent();
    document.getElementById('noKeyBox').classList.remove('hidden');
    document.getElementById('noKeyDesc').textContent = `Fehler: ${e.message}. Bitte API-Key prüfen.`;
  }
}

function buildPrompt(data) {
  const { emails, meetings, rumbles } = data;
  const today = new Date().toLocaleDateString('de-CH');

  const eLines = emails.slice(0, 60).map(e =>
    `[${e.date}] EMAIL ${e.direction === 'received' ? 'VON' : 'AN'}: "${e.subject}" – ${e.preview}`
  ).join('\n');
  const mLines = meetings.slice(0, 40).map(m =>
    `[${m.date}] MEETING (${m.duration}min): "${m.subject}"`
  ).join('\n');
  const rLines = rumbles.map(r =>
    `[${r.date}] RUMBLE: "${r.subject}"`
  ).join('\n');

  return `Du bist ein persönlicher Business-Assistent. Heute ist ${today}.
Analysiere alle Interaktionen mit "${_contactName}"${_contactEmail ? ` (${_contactEmail})` : ''}.

${eLines ? `E-MAILS:\n${eLines}\n\n` : ''}${mLines ? `MEETINGS:\n${mLines}\n\n` : ''}${rLines ? `RUMBLES:\n${rLines}\n\n` : ''}${(!eLines && !mLines && !rLines) ? 'Noch keine Interaktionen.\n\n' : ''}Antworte NUR mit validem JSON (kein Markdown):
{
  "summary": "2-3 Sätze zur Beziehung, Häufigkeit, Ton",
  "sentiment": "positiv|neutral|negativ",
  "openPoints": ["Offener Punkt 1", "Offener Punkt 2"],
  "nextStep": "Konkrete Empfehlung für nächstes Gespräch",
  "themes": [
    { "name": "Thema", "count": 3, "status": "offen|abgeschlossen", "summary": "Kurzbeschreibung" }
  ],
  "background": "Öffentlich bekannte Infos zu ${_contactName}: Beruf, Unternehmen, Branche. Falls unbekannt: leer lassen."
}`;
}

function parseAIResponse(raw) {
  let text = raw.trim().replace(/^```(?:json)?\n?/, '').replace(/\n?```$/, '').trim();
  try { return JSON.parse(text); } catch {}
  const m = text.match(/\{[\s\S]*\}/);
  if (m) { try { return JSON.parse(m[0]); } catch {} }
  return { summary: text.slice(0, 400), themes: [], openPoints: [], sentiment: 'neutral', nextStep: '', background: '' };
}

// ── AI provider routing ───────────────────────────────────────────────────
function detectProvider(key) {
  if (!key) return null;
  if (key.startsWith('AIza'))    return 'gemini';
  if (key.startsWith('sk-ant-')) return 'anthropic';
  if (key.startsWith('gsk_'))    return 'groq';
  if (key.startsWith('sk-'))     return 'openai';
  return 'gemini';
}

async function callAI(prompt, apiKey) {
  switch (detectProvider(apiKey)) {
    case 'anthropic': return callAnthropic(prompt, apiKey);
    case 'groq':      return callGroq(prompt, apiKey);
    case 'openai':    return callOpenAI(prompt, apiKey);
    default:          return callGemini(prompt, apiKey);
  }
}

async function callGemini(prompt, apiKey) {
  const res = await fetch(
    `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`,
    { method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }], generationConfig: { temperature: 0.3, maxOutputTokens: 2048 } }) }
  );
  if (!res.ok) { const t = await res.text(); throw new Error(`Gemini ${res.status}: ${t.slice(0, 120)}`); }
  const d = await res.json();
  return d.candidates?.[0]?.content?.parts?.[0]?.text || '';
}

async function callAnthropic(prompt, apiKey) {
  const res = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01', 'Content-Type': 'application/json' },
    body: JSON.stringify({ model: 'claude-haiku-4-5-20251001', max_tokens: 2048, messages: [{ role: 'user', content: prompt }] }),
  });
  if (!res.ok) { const t = await res.text(); throw new Error(`Claude ${res.status}: ${t.slice(0, 120)}`); }
  const d = await res.json();
  return d.content?.[0]?.text || '';
}

async function callOpenAI(prompt, apiKey) {
  const res = await fetch('https://api.openai.com/v1/chat/completions', {
    method: 'POST',
    headers: { Authorization: `Bearer ${apiKey}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ model: 'gpt-4o-mini', max_tokens: 2048, messages: [{ role: 'user', content: prompt }] }),
  });
  if (!res.ok) { const t = await res.text(); throw new Error(`OpenAI ${res.status}: ${t.slice(0, 120)}`); }
  const d = await res.json();
  return d.choices?.[0]?.message?.content || '';
}

async function callGroq(prompt, apiKey) {
  const res = await fetch('https://api.groq.com/openai/v1/chat/completions', {
    method: 'POST',
    headers: { Authorization: `Bearer ${apiKey}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ model: 'llama-3.1-70b-versatile', max_tokens: 2048, messages: [{ role: 'user', content: prompt }] }),
  });
  if (!res.ok) { const t = await res.text(); throw new Error(`Groq ${res.status}: ${t.slice(0, 120)}`); }
  const d = await res.json();
  return d.choices?.[0]?.message?.content || '';
}

// ── Rendering ─────────────────────────────────────────────────────────────
function esc(s) {
  return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function fmtDate(d) {
  if (!d) return '';
  try { return new Date(d).toLocaleDateString('de-CH', { day: '2-digit', month: '2-digit', year: 'numeric' }); }
  catch { return d; }
}

function updateHeader(data) {
  const all    = [...data.emails, ...data.meetings, ...data.rumbles].filter(i => i.date).sort((a, b) => b.date.localeCompare(a.date));
  const parts  = [];
  if (all.length) {
    const last = all[0];
    const lbl  = last.type === 'email' ? 'Mail' : last.type === 'meeting' ? 'Meeting' : 'Rumble';
    parts.push(`Letzter Kontakt: ${fmtDate(last.date)} (${lbl})`);
  }
  if (data.meetings.length) parts.push(`Letztes Meeting: ${fmtDate(data.meetings[0].date)}`);
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

  const sentMap  = { positiv: ['s-pos', 'Positiv'], neutral: ['s-neu', 'Neutral'], negativ: ['s-neg', 'Negativ'] };
  const [cls, lbl] = sentMap[a.sentiment] || sentMap.neutral;
  const meta = document.getElementById('headerMeta');
  meta.innerHTML += `<span class="sentiment"><span class="s-dot ${cls}"></span>${esc(lbl)}</span>`;

  if (a.openPoints?.length) {
    document.getElementById('openPointsBox').classList.remove('hidden');
    document.getElementById('openPointsList').innerHTML = a.openPoints.map(p => `<div class="open-point">${esc(p)}</div>`).join('');
  }

  const key = getApiKey();
  const providerLabels = { gemini: 'Gemini', anthropic: 'Claude', openai: 'GPT-4o', groq: 'Groq/Llama' };
  document.getElementById('providerBadge').textContent = providerLabels[detectProvider(key)] || 'KI';
}

function renderTimeline(data) {
  const items = [
    ...data.emails,
    ...data.meetings,
    ...data.rumbles,
  ].sort((a, b) => (b.date || '').localeCompare(a.date || ''));

  const list = document.getElementById('timelineList');
  if (!items.length) {
    list.innerHTML = '<div class="empty-state">Keine Interaktionen in den letzten 2 Jahren.</div>';
    return;
  }

  list.innerHTML = items.map(item => {
    let icon, iconCls, badgeCls, badgeTxt;
    if (item.type === 'email') {
      icon    = item.direction === 'received' ? '📧' : '📤';
      iconCls = item.direction === 'received' ? 'tl-email-in' : 'tl-email-out';
      badgeCls = item.direction === 'received' ? 'b-email-in' : 'b-email-out';
      badgeTxt = item.direction === 'received' ? 'Empfangen' : 'Gesendet';
    } else if (item.type === 'meeting') {
      icon = '📅'; iconCls = 'tl-meeting'; badgeCls = 'b-meeting';
      badgeTxt = `Meeting${item.duration ? ' · ' + item.duration + 'min' : ''}`;
    } else {
      icon = '📝'; iconCls = 'tl-rumble'; badgeCls = 'b-rumble'; badgeTxt = 'Rumble';
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

function renderThemes(themes) {
  const el = document.getElementById('themesContent');
  if (!themes?.length) { renderThemesEmpty(); return; }
  el.innerHTML = themes.map(t => {
    const sc  = t.status === 'offen' ? 's-open' : 's-done';
    const stx = t.status === 'offen' ? 'offen' : 'abgeschlossen';
    return `<div class="theme-card">
      <div class="theme-header">
        <span class="theme-name">${esc(t.name)}</span>
        <span class="theme-count">${t.count || ''} Interaktionen</span>
        <span class="theme-status ${sc}">${stx}</span>
      </div>
      <div class="theme-summary">${esc(t.summary)}</div>
    </div>`;
  }).join('');
}

function renderThemesEmpty() {
  document.getElementById('themesContent').innerHTML =
    '<div class="empty-state">Kein KI-Key eingetragen.<br/>Gib oben einen API-Key ein um Themen zu erkennen.</div>';
}

function renderBackground(text) {
  const el = document.getElementById('backgroundContent');
  if (!text) { el.innerHTML = '<div class="empty-state">Keine öffentlichen Informationen gefunden.</div>'; return; }
  el.innerHTML = `<div class="bg-box">${esc(text)}</div>`;
}

// ── Tab switching ─────────────────────────────────────────────────────────
function showTab(name) {
  ['Timeline', 'Themes', 'Background'].forEach(t => {
    document.getElementById(`tab${t}`).classList.toggle('hidden', t !== name);
    document.getElementById(`tabBtn${t}`).classList.toggle('active', t === name);
  });
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
