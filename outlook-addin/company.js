/* Tembu Firma-Analyse – company.js v20260629q */

const PERSONAL_DOMAINS = new Set([
  'gmail.com','googlemail.com','hotmail.com','hotmail.de','hotmail.ch',
  'outlook.com','outlook.de','outlook.ch','yahoo.com','yahoo.de','yahoo.ch',
  'live.com','live.de','icloud.com','me.com','protonmail.com','proton.me',
  'aol.com','gmx.ch','gmx.de','gmx.net','bluewin.ch','hispeed.ch','sunrise.ch',
]);

const SESSION_KEY       = '@tembu_outlook_session';
const DIALOG_TK_KEY     = '@tembu_dialog_token';
const SETTINGS_KEY      = '@tembu_detail_settings';
const DB_NAME           = 'tembu_cache_v1';
const STORE_NAME        = 'contact_analyses';
const SINCE_MONTHS_KEY  = 'tembu_since_months';
const HISTORY_MONTHS_KEY = 'tembu_history_months';
const AI_PROVIDER_KEY   = 'tembu_ai_provider';
const AI_MODEL_KEY      = 'tembu_ai_model';
const AI_ENDPOINT_KEY   = 'tembu_ai_endpoint';

const esc = TCore.esc;

let _token       = null;
let _domain      = '';
let _companyName = '';
let _cacheKey    = '';
let _contacts    = [];
let _rawData     = null;

// ── Init ──────────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', async () => {
  const params = new URLSearchParams(window.location.search);
  _domain      = (params.get('domain') || '').toLowerCase().trim();
  _companyName = params.get('company') || deriveCompanyName(_domain);
  _cacheKey    = 'company:' + _domain;

  document.getElementById('headerName').textContent   = _companyName;
  document.getElementById('headerDomain').textContent = _domain;

  const settings = loadSettings();
  if (settings.apiKey) document.getElementById('apiKeyInput').value = settings.apiKey;
  const aiConfig = getAIConfig();
  const providerSelect = document.getElementById('providerSelect');
  providerSelect.value = aiConfig.provider;
  if (aiConfig.endpoint) document.getElementById('localEndpoint').value = aiConfig.endpoint;
  if (aiConfig.model)    document.getElementById('localModel').value    = aiConfig.model;
  updateProviderUI(aiConfig.provider);
  if (isLocalProvider(aiConfig.provider)) loadOllamaModels();
  providerSelect.addEventListener('change', () => {
    const p = providerSelect.value;
    localStorage.setItem(AI_PROVIDER_KEY, p);
    updateProviderUI(p);
    if (isLocalProvider(p)) loadOllamaModels();
  });

  document.getElementById('btnSaveKey').addEventListener('click', () => {
    const key = document.getElementById('apiKeyInput').value.trim();
    saveSettings({ apiKey: key });
    if (_rawData) runAI(_rawData);
  });
  document.getElementById('btnSaveLocal').addEventListener('click', () => {
    const endpoint = document.getElementById('localEndpoint').value.trim();
    const model    = document.getElementById('localModel').value.trim();
    if (endpoint) localStorage.setItem(AI_ENDPOINT_KEY, endpoint);
    if (model)    localStorage.setItem(AI_MODEL_KEY,    model);
    loadOllamaModels();
    if (_rawData) runAI(_rawData);
  });

  const sinceSelect = document.getElementById('sinceSelect');
  sinceSelect.value = localStorage.getItem(SINCE_MONTHS_KEY) || '24';
  sinceSelect.addEventListener('change', () => {
    localStorage.setItem(SINCE_MONTHS_KEY, sinceSelect.value);
    loadData(true);
  });

  _token = getStoredToken();
  if (!_token) { showError('Kein Token. Bitte in der Taskpane abmelden und neu anmelden.'); return; }
  if (!_domain) { showError('Keine Domain angegeben.'); return; }

  await loadData(false);
});

function deriveCompanyName(domain) {
  if (!domain) return 'Unbekannte Firma';
  const part = domain.split('.')[0];
  return part.charAt(0).toUpperCase() + part.slice(1);
}

// ── Token / Settings ──────────────────────────────────────────────────────
function getStoredToken() {
  try { const t = new URLSearchParams(window.location.search).get('t'); if (t?.length > 10) return t; } catch {}
  try { const r = localStorage.getItem(DIALOG_TK_KEY); if (r) { const d = JSON.parse(r); if (d.exp > Date.now()) return d.token; } } catch {}
  try { const r = localStorage.getItem(SESSION_KEY);   if (r) { const d = JSON.parse(r); if (d.expiry > Date.now()) return d.token; } } catch {}
  return null;
}
function loadSettings() { try { return JSON.parse(localStorage.getItem(SETTINGS_KEY) || '{}'); } catch { return {}; } }
function saveSettings(patch) { try { localStorage.setItem(SETTINGS_KEY, JSON.stringify({ ...loadSettings(), ...patch })); } catch {} }
function getAIConfig() {
  const s = loadSettings();
  const provider = localStorage.getItem(AI_PROVIDER_KEY) || 'gemini';
  return { provider, apiKey: s.apiKey || '', model: localStorage.getItem(AI_MODEL_KEY) || '', endpoint: localStorage.getItem(AI_ENDPOINT_KEY) || '' };
}
function isLocalProvider(p) { return p === 'ollama' || p === 'lmstudio'; }

function updateProviderUI(provider) {
  const isLocal = isLocalProvider(provider);
  document.getElementById('apiKeyInput').classList.toggle('hidden', isLocal);
  document.getElementById('btnSaveKey').classList.toggle('hidden', isLocal);
  document.getElementById('geminiLink').classList.toggle('hidden', isLocal || provider !== 'gemini');
  document.getElementById('localEndpoint').classList.toggle('hidden', !isLocal);
  document.getElementById('localModel').classList.toggle('hidden', !isLocal);
  document.getElementById('btnSaveLocal').classList.toggle('hidden', !isLocal);
  const endpointEl = document.getElementById('localEndpoint');
  if (isLocal && !endpointEl.value)
    endpointEl.value = provider === 'lmstudio' ? 'http://localhost:1234' : 'http://localhost:11434';
  const noKeyBtn = document.getElementById('noKeyBtn');
  if (noKeyBtn) noKeyBtn.classList.toggle('hidden', isLocal);
}

async function loadOllamaModels() {
  const endpoint = document.getElementById('localEndpoint')?.value.trim() || localStorage.getItem(AI_ENDPOINT_KEY) || 'http://localhost:11434';
  try {
    const res  = await fetch(`${endpoint}/api/tags`);
    const data = await res.json();
    const list = document.getElementById('modelList');
    if (list && data.models?.length) list.innerHTML = data.models.map(m => `<option value="${m.name}">`).join('');
  } catch {}
}

// ── Since helpers ─────────────────────────────────────────────────────────
function getSinceMonths() { return parseInt(localStorage.getItem(SINCE_MONTHS_KEY) || '24'); }
function getSinceDate() {
  const m = getSinceMonths();
  return m === 0 ? new Date('2000-01-01') : new Date(Date.now() - m * 30.44 * 24 * 60 * 60 * 1000);
}
function getSinceLabel() {
  const m = getSinceMonths();
  if (m === 0) return 'Alles';
  if (m < 12)  return `letzte ${m} Monate`;
  if (m === 12) return 'letztes Jahr';
  return `letzte ${Math.round(m / 12)} Jahre`;
}

// ── IndexedDB cache ───────────────────────────────────────────────────────
function openDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = e => { const db = e.target.result; if (!db.objectStoreNames.contains(STORE_NAME)) db.createObjectStore(STORE_NAME, { keyPath: 'id' }); };
    req.onsuccess = e => resolve(e.target.result);
    req.onerror   = () => reject(req.error);
  });
}
async function getCached(key) {
  try { const db = await openDB(); return new Promise(r => { const req = db.transaction(STORE_NAME,'readonly').objectStore(STORE_NAME).get(key); req.onsuccess = () => r(req.result || null); req.onerror = () => r(null); }); } catch { return null; }
}
async function saveCache(key, data) {
  try { const db = await openDB(); return new Promise(r => { const tx = db.transaction(STORE_NAME,'readwrite'); tx.objectStore(STORE_NAME).put({ id: key, cachedAt: Date.now(), ...data }); tx.oncomplete = () => r(); tx.onerror = () => r(); }); } catch {}
}

// ── Graph helper ──────────────────────────────────────────────────────────
async function gFetch(path) {
  try { return await TCore.graphGet(_token, path); }
  catch (e) { if (e.message?.startsWith('Graph 403')) throw new Error('403: Berechtigung fehlt (Mail.Read + Calendars.Read).'); throw e; }
}

// ── Load data ─────────────────────────────────────────────────────────────
async function loadData(force) {
  setLoading('Kontakte werden gesucht…');
  try {
    _contacts = await findContactsByDomain(_domain);
    const since = getSinceDate();

    setLoading('E-Mails und Meetings werden geladen…');
    const [emails, meetings] = await Promise.all([
      fetchEmailsByDomain(since),
      fetchMeetingsForContacts(_contacts, since),
    ]);

    _rawData = { emails, meetings, contacts: _contacts };
    updateHeader(_rawData);
    renderStats(_rawData);
    renderTimeline(_rawData);
    renderContacts(_rawData);
    showContent();

    const cached = await getCached(_cacheKey);
    if (cached && !force && cached.analysis) {
      const ageH = Math.round((Date.now() - cached.cachedAt) / 3600000);
      document.getElementById('cacheInfoText').textContent = `KI-Analyse von vor ${ageH < 1 ? '< 1' : ageH} Stunde${ageH !== 1 ? 'n' : ''}`;
      document.getElementById('cacheInfo').classList.remove('hidden');
      renderAiAnalysis(cached.analysis);
      renderThemes(cached.analysis.themes || []);
      return;
    }

    if (!force) {
      const cloud = await TCore.loadAnalysis(_token, _cacheKey);
      if (cloud?.analysis) {
        await saveCache(_cacheKey, { analysis: cloud.analysis, rawData: _rawData });
        const ageH = Math.round((Date.now() - new Date(cloud.savedAt).getTime()) / 3600000);
        document.getElementById('cacheInfoText').textContent = `KI-Analyse von vor ${ageH < 1 ? '< 1' : ageH} Stunde${ageH !== 1 ? 'n' : ''} (Cloud)`;
        document.getElementById('cacheInfo').classList.remove('hidden');
        renderAiAnalysis(cloud.analysis);
        renderThemes(cloud.analysis.themes || []);
        return;
      }
    }

    await runAI(_rawData);
  } catch (e) {
    showError('Fehler: ' + (e.message || String(e)));
  }
}

async function forceRefresh() {
  document.getElementById('cacheInfo').classList.add('hidden');
  await loadData(true);
}

async function loadHistory() {
  const btn    = document.getElementById('btnHistory');
  const status = document.getElementById('historyStatus');
  const months = parseInt(document.getElementById('historySelect').value || '60');
  const since  = months === 0 ? new Date('2000-01-01') : new Date(Date.now() - months * 30.44 * 24 * 60 * 60 * 1000);

  btn.disabled = true;
  status.textContent = 'Lade…';
  try {
    const [oldEmails, oldMeetings] = await Promise.all([
      fetchEmailsByDomain(since, 250),
      fetchMeetingsForContacts(_contacts, since, 100),
    ]);
    const seenE = new Set((_rawData?.emails   || []).map(e => e.id).filter(Boolean));
    const seenM = new Set((_rawData?.meetings || []).map(m => m.id).filter(Boolean));
    _rawData = {
      emails:   [...(_rawData?.emails   || []), ...oldEmails.filter(e   => !e.id || !seenE.has(e.id))].sort((a,b)=>b.date.localeCompare(a.date)),
      meetings: [...(_rawData?.meetings || []), ...oldMeetings.filter(m => !m.id || !seenM.has(m.id))].sort((a,b)=>b.date.localeCompare(a.date)),
      contacts: _contacts,
    };
    const added = oldEmails.filter(e => !seenE.has(e.id)).length + oldMeetings.filter(m => !seenM.has(m.id)).length;
    status.textContent = added > 0 ? `+${added} neue gefunden` : 'Keine weiteren gefunden';
    renderStats(_rawData);
    renderTimeline(_rawData);
    updateHeader(_rawData);
    await runAI(_rawData);
  } catch (e) {
    status.textContent = 'Fehler: ' + (e.message || '').slice(0, 50);
  } finally {
    btn.disabled = false;
  }
}

// ── Find contacts by domain ───────────────────────────────────────────────
async function findContactsByDomain(domain) {
  try {
    const data = await gFetch(`/me/contacts?$select=displayName,emailAddresses&$top=200`);
    return (data.value || [])
      .filter(c => (c.emailAddresses || []).some(e => (e.address || '').toLowerCase().endsWith('@' + domain)))
      .map(c => ({
        name:  c.displayName || '',
        email: (c.emailAddresses || []).find(e => (e.address || '').toLowerCase().endsWith('@' + domain))?.address || '',
      }))
      .filter(c => c.email);
  } catch { return []; }
}

// ── Fetch emails by domain ────────────────────────────────────────────────
async function fetchEmailsByDomain(since, top = 100) {
  const sinceDate  = since.toISOString().slice(0, 10);
  const enc        = encodeURIComponent;
  const sel        = '$select=id,subject,receivedDateTime,sentDateTime,from,toRecipients,bodyPreview,webLink';
  const q          = enc('"' + _domain + '"');
  const domainLow  = _domain.toLowerCase();

  const [inbox, sent] = await Promise.allSettled([
    gFetch(`/me/messages?$search=${q}&${sel}&$top=${top}`),
    gFetch(`/me/mailFolders/SentItems/messages?$search=${q}&${sel}&$top=${top}`),
  ]);

  const result = [];
  for (const m of (inbox.status === 'fulfilled' ? inbox.value?.value || [] : [])) {
    const date = (m.receivedDateTime || '').slice(0, 10);
    if (date < sinceDate) continue;
    if (!(m.from?.emailAddress?.address || '').toLowerCase().endsWith('@' + domainLow)) continue;
    result.push({ id: m.id, date, type: 'email', direction: 'received',
      contact: m.from?.emailAddress?.name || m.from?.emailAddress?.address || '',
      contactEmail: m.from?.emailAddress?.address || '',
      subject: m.subject || '(kein Betreff)', preview: (m.bodyPreview || '').slice(0, 200), url: m.webLink || '' });
  }
  for (const m of (sent.status === 'fulfilled' ? sent.value?.value || [] : [])) {
    const date = (m.sentDateTime || '').slice(0, 10);
    if (date < sinceDate) continue;
    const toR = (m.toRecipients || []).find(r => (r.emailAddress?.address || '').toLowerCase().endsWith('@' + domainLow));
    if (!toR) continue;
    result.push({ id: m.id, date, type: 'email', direction: 'sent',
      contact: toR.emailAddress?.name || toR.emailAddress?.address || '',
      contactEmail: toR.emailAddress?.address || '',
      subject: m.subject || '(kein Betreff)', preview: (m.bodyPreview || '').slice(0, 200), url: m.webLink || '' });
  }
  return result.sort((a, b) => b.date.localeCompare(a.date));
}

// ── Fetch meetings for multiple contacts ──────────────────────────────────
async function fetchMeetingsForContacts(contacts, since, topPerContact = 50) {
  if (!contacts.length) return [];
  const s = since.toISOString(), now = new Date().toISOString(), enc = encodeURIComponent;
  const results = await Promise.allSettled(contacts.map(c =>
    gFetch(`/me/calendarView?startDateTime=${enc(s)}&endDateTime=${enc(now)}&$filter=${enc(`attendees/any(a:a/emailAddress/address eq '${c.email}')`)}&$select=id,subject,start,end,webLink&$top=${topPerContact}`)
      .then(data => (data.value || []).map(e => ({
        id: e.id, type: 'meeting', contact: c.name,
        date: (e.start?.dateTime || '').slice(0, 10),
        subject: e.subject || '(kein Titel)',
        duration: Math.round((new Date(e.end?.dateTime) - new Date(e.start?.dateTime)) / 60000) || 0,
        url: e.webLink || '',
      })))
      .catch(() => [])
  ));
  const all = results.flatMap(r => r.status === 'fulfilled' ? r.value : []);
  const seen = new Set();
  return all.filter(m => !seen.has(m.id) && seen.add(m.id)).sort((a, b) => b.date.localeCompare(a.date));
}

// ── AI ────────────────────────────────────────────────────────────────────
async function runAI(data) {
  const config = getAIConfig();
  const hasKey = isLocalProvider(config.provider) || !!config.apiKey;
  if (!hasKey) { document.getElementById('noKeyBox').classList.remove('hidden'); return; }

  const themesEl = document.getElementById('themesContent');
  themesEl.innerHTML = '<div class="loading-box">KI analysiert…</div>';

  try {
    const raw    = await TCore.callAI(buildPrompt(data), config);
    const parsed = TCore.parseAIResponse(raw);
    renderAiAnalysis(parsed);
    renderThemes(parsed.themes || []);
    await saveCache(_cacheKey, { analysis: parsed, rawData: data });
    TCore.saveAnalysis(_token, _cacheKey, _companyName, _domain, parsed).catch(() => {});
    document.getElementById('cacheInfoText').textContent = 'Gerade analysiert';
    document.getElementById('cacheInfo').classList.remove('hidden');
  } catch (e) {
    const errMsg = String(e.message || e);
    const noKeyBox = document.getElementById('noKeyBox');
    const desc     = document.getElementById('noKeyDesc');
    const icon     = document.querySelector('#noKeyBox .no-key-icon');
    const title    = document.querySelector('#noKeyBox .no-key-title');
    if (isLocalProvider(config.provider)) {
      if (icon)  icon.textContent  = '⚠️';
      if (title) title.textContent = 'Verbindungsfehler';
      const isFetch = errMsg.toLowerCase().includes('fetch');
      if (desc) desc.innerHTML = 'Fehler: ' + esc(errMsg)
        + (isFetch ? '<br/><br/>CORS: Ollama muss mit <b>OLLAMA_ORIGINS=*</b> gestartet werden.' : '');
    } else {
      if (icon)  icon.textContent  = '🔑';
      if (title) title.textContent = 'KI-Analyse verfügbar';
      if (desc)  desc.innerHTML    = 'Fehler: ' + esc(errMsg);
    }
    if (noKeyBox) noKeyBox.classList.remove('hidden');
    themesEl.innerHTML = '<div class="empty-state">Analyse fehlgeschlagen.</div>';
  }
}

function buildPrompt(data) {
  const { emails, meetings, contacts } = data;
  const today        = new Date().toLocaleDateString('de-CH');
  const contactNames = contacts.length ? contacts.map(c => c.name).filter(Boolean).join(', ') : 'unbekannte Kontakte';

  const eLines = emails.slice(0, 80).map(e =>
    `[${e.date}] EMAIL ${e.direction === 'received' ? 'VON' : 'AN'} ${e.contact}: "${e.subject}"`
  ).join('\n');
  const mLines = meetings.slice(0, 40).map(m =>
    `[${m.date}] MEETING mit ${m.contact} (${m.duration}min): "${m.subject}"`
  ).join('\n');

  return `Du bist ein Business-Assistent. Heute ist ${today}.
Analysiere meine Geschäftsbeziehung mit der Firma "${_companyName}" (Domain: ${_domain}).
Bekannte Kontakte: ${contactNames}.

${eLines ? `E-MAILS:\n${eLines}\n\n` : ''}${mLines ? `MEETINGS:\n${mLines}\n\n` : ''}${(!eLines && !mLines) ? 'Noch keine Interaktionen gefunden.\n\n' : ''}Antworte NUR mit validem JSON (kein Markdown).
Wichtig: Im "interactions"-Array ALLE zugehörigen Interaktionen auflisten, keine Auswahl.
{
  "summary": "2-3 Sätze zur Gesamtbeziehung mit der Firma",
  "sentiment": "positiv|neutral|negativ",
  "openPoints": ["Offener Punkt 1"],
  "themes": [
    { "name": "Thema", "status": "offen|abgeschlossen", "summary": "Kurzbeschreibung", "contacts": ["Name1","Name2"], "interactions": [{"date":"YYYY-MM-DD","type":"email|meeting","contact":"Name","subject":"Betreff"}] }
  ],
  "nextStep": "Konkrete Empfehlung für nächsten Schritt mit dieser Firma"
}`;
}

// ── Rendering ─────────────────────────────────────────────────────────────
function fmtDate(d) {
  if (!d) return '';
  try { return new Date(d).toLocaleDateString('de-CH', { day:'2-digit', month:'2-digit', year:'numeric' }); }
  catch { return d; }
}

function updateHeader(data) {
  const all = [...data.emails, ...data.meetings].filter(i => i.date).sort((a,b) => b.date.localeCompare(a.date));
  const parts = [`Zeitraum: ${getSinceLabel()}`];
  if (all.length) parts.push(`Letzter Kontakt: ${fmtDate(all[0].date)}`);
  const uniqueContacts = new Set(data.emails.map(e => e.contactEmail).filter(Boolean));
  parts.push(`${uniqueContacts.size || data.contacts.length} Kontakte`);
  document.getElementById('headerMeta').innerHTML = parts.map(p => `<span>${esc(p)}</span>`).join('');
}

function renderStats(data) {
  const unique = new Set(data.emails.map(e => e.contactEmail).filter(Boolean));
  document.getElementById('statContacts').textContent = unique.size || data.contacts.length;
  document.getElementById('statEmails').textContent   = data.emails.length;
  document.getElementById('statMeetings').textContent = data.meetings.length;
}

function renderTimeline(data) {
  const el  = document.getElementById('timelineContent');
  const all = [...data.emails, ...data.meetings].sort((a,b) => b.date.localeCompare(a.date));
  if (!all.length) { el.innerHTML = '<div class="empty-state">Keine Interaktionen im gewählten Zeitraum.</div>'; return; }
  el.innerHTML = all.map(item => {
    const icon    = item.type === 'meeting' ? '📅' : (item.direction === 'received' ? '✉' : '↗');
    const contact = item.contact ? `<span class="tl-contact">${esc(item.contact)}</span>` : '';
    const link    = item.url ? `<a class="tl-link" href="${esc(item.url)}" target="_blank">↗</a>` : '';
    return `<div class="tl-item"><span class="tl-icon">${icon}</span><span class="tl-date">${item.date}</span>${contact}<span class="tl-subject">${esc(item.subject)}</span>${link}</div>`;
  }).join('');
}

function renderAiAnalysis(parsed) {
  const box = document.getElementById('summaryBox');
  if (!box) return;
  document.getElementById('summaryText').textContent = parsed.summary || '';
  const opEl = document.getElementById('openPointsList');
  if (parsed.openPoints?.length) {
    opEl.innerHTML = parsed.openPoints.map(p => `<li>${esc(p)}</li>`).join('');
    opEl.classList.remove('hidden');
  }
  const nsEl = document.getElementById('nextStep');
  if (parsed.nextStep) { nsEl.textContent = '→ ' + parsed.nextStep; nsEl.classList.remove('hidden'); }
  box.classList.remove('hidden');
}

function renderContacts(data) {
  const el = document.getElementById('contactsContent');
  // Aggregate by contactEmail from emails
  const byEmail = {};
  for (const e of data.emails) {
    const key = e.contactEmail?.toLowerCase() || '';
    if (!key) continue;
    if (!byEmail[key]) byEmail[key] = { name: e.contact, email: e.contactEmail, count: 0 };
    byEmail[key].count++;
  }
  // Add known contacts even with 0 interactions
  for (const c of data.contacts) {
    const key = c.email?.toLowerCase() || '';
    if (key && !byEmail[key]) byEmail[key] = { name: c.name, email: c.email, count: 0 };
  }
  const list = Object.values(byEmail).sort((a,b) => b.count - a.count);
  if (!list.length) { el.innerHTML = '<div class="empty-state">Keine Kontakte gefunden.</div>'; return; }

  const detailBase = `detail.html?t=${encodeURIComponent(_token || '')}`;
  el.innerHTML = list.map(c => {
    const initials = (c.name || c.email).split(/\s+/).map(w => w[0]).join('').toUpperCase().slice(0, 2);
    const link = `${detailBase}&name=${encodeURIComponent(c.name || '')}&email=${encodeURIComponent(c.email)}`;
    return `<div class="contact-row">
      <div class="contact-row-avatar">${esc(initials)}</div>
      <div class="contact-row-info">
        <div class="contact-row-name">${esc(c.name || c.email)}</div>
        <div class="contact-row-email">${esc(c.email)}</div>
      </div>
      <div class="contact-row-count">${c.count} Mail${c.count !== 1 ? 's' : ''}</div>
      <a class="contact-row-link" href="${link}" target="_blank">Analyse →</a>
    </div>`;
  }).join('');
}

function findInteraction(item) {
  if (!_rawData) return null;
  const pool = item.type === 'meeting' ? _rawData.meetings : _rawData.emails;
  const subj = (item.subject || '').toLowerCase().slice(0, 30);
  return pool.find(e => e.date === item.date && (e.subject || '').toLowerCase().includes(subj))
      || pool.find(e => e.date === item.date);
}

function toggleThemeDetail(el) {
  const detail = el.closest('.theme-card').querySelector('.theme-interactions');
  if (!detail) return;
  const open = detail.classList.toggle('hidden') === false;
  el.querySelector('.theme-toggle-arrow').textContent = open ? '▼' : '▶';
}

function renderThemes(themes) {
  const el = document.getElementById('themesContent');
  if (!themes?.length) { el.innerHTML = '<div class="empty-state">Keine Themen erkannt.</div>'; return; }
  el.innerHTML = themes.map(t => {
    const sc     = t.status === 'offen' ? 's-open' : 's-done';
    const stx    = t.status === 'offen' ? 'offen'  : 'abgeschlossen';
    const items  = t.interactions || [];
    const cNames = (t.contacts || []).join(' · ');
    const detailHtml = items.map(i => {
      const icon  = i.type === 'meeting' ? '📅' : '✉';
      const match = findInteraction(i);
      const inner = `<span class="ti-icon">${icon}</span><span class="ti-date">${esc(i.date||'')}</span><span class="ti-contact">${esc(i.contact||'')}</span><span class="ti-subject">${esc(i.subject||'')}</span>`;
      return match?.url
        ? `<a class="theme-interaction ti-link" href="${esc(match.url)}" target="_blank">${inner}</a>`
        : `<div class="theme-interaction">${inner}</div>`;
    }).join('');
    const countEl = items.length
      ? `<span class="theme-count theme-count-toggle" onclick="toggleThemeDetail(this)">${items.length} Interaktionen <span class="theme-toggle-arrow">▶</span></span>`
      : '';
    return `<div class="theme-card">
      <div class="theme-header">
        <span class="theme-name">${esc(t.name)}</span>
        ${cNames ? `<span class="theme-contacts">${esc(cNames)}</span>` : ''}
        ${countEl}
        <span class="theme-status ${sc}">${stx}</span>
      </div>
      <div class="theme-summary">${esc(t.summary)}</div>
      ${detailHtml ? `<div class="theme-interactions hidden">${detailHtml}</div>` : ''}
    </div>`;
  }).join('');
}

// ── UI helpers ────────────────────────────────────────────────────────────
function showTab(name) {
  ['Timeline','Themes','Contacts'].forEach(t => {
    document.getElementById('tab' + t)?.classList.toggle('hidden', t !== name);
    document.getElementById('tabBtn' + t)?.classList.toggle('active', t === name);
  });
}
function setLoading(msg) {
  document.getElementById('loadingMsg').textContent = msg;
  document.getElementById('loadingBox').classList.remove('hidden');
  document.getElementById('errorBox').classList.add('hidden');
  document.getElementById('tabTimeline').classList.add('hidden');
  document.getElementById('tabThemes').classList.add('hidden');
  document.getElementById('tabContacts').classList.add('hidden');
}
function showContent() {
  document.getElementById('loadingBox').classList.add('hidden');
  document.getElementById('errorBox').classList.add('hidden');
  document.getElementById('tabTimeline').classList.remove('hidden');
}
function showError(msg) {
  document.getElementById('errorMsg').textContent = msg;
  document.getElementById('errorBox').classList.remove('hidden');
  document.getElementById('loadingBox').classList.add('hidden');
}
