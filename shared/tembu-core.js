/* Tembu Core v1 — gemeinsame Funktionen für Outlook Add-in und Teams Add-in */

const TCore = (() => {

  const TEMBU_LIST = 'Tembu';
  const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

  // ── Graph ─────────────────────────────────────────────────────────────────
  async function graphGet(token, path) {
    const res = await fetch(GRAPH_BASE + path, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!res.ok) throw new Error(`Graph ${res.status} ${path}`);
    return res.json();
  }

  async function graphPost(token, path, body) {
    const res = await fetch(GRAPH_BASE + path, {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    });
    if (!res.ok) throw new Error(`Graph POST ${res.status} ${path}`);
    return res.json();
  }

  async function graphPatch(token, path, body) {
    const res = await fetch(GRAPH_BASE + path, {
      method: 'PATCH',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    });
    if (!res.ok) throw new Error(`Graph PATCH ${res.status} ${path}`);
  }

  async function graphPut(token, path, body) {
    const res = await fetch(GRAPH_BASE + path, {
      method: 'PUT',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: typeof body === 'string' ? body : JSON.stringify(body),
    });
    if (!res.ok) throw new Error(`Graph PUT ${res.status} ${path}`);
    return res.json().catch(() => null);
  }

  // ── Analyse-Cache (OneDrive) ──────────────────────────────────────────────
  // Jeder Nutzer speichert in eigenem OneDrive/Tembu/analysen/ → keine geteilte DB nötig
  function analysisOneDrivePath(cacheKey) {
    const safe = String(cacheKey)
      .replace(/@/g, '-at-')             // @ ist OData-Sonderzeichen → auch %40 reicht nicht
      .replace(/[^a-z0-9._-]/g, '-')
      .replace(/-+/g, '-')
      .slice(0, 80);
    return `/me/drive/root:/Tembu/analysen/${safe}.json:/content`;
  }

  async function ensureOneDriveFolders(token) {
    // OneDrive erstellt Parent-Ordner nicht automatisch → einmalig anlegen
    const mkfolder = async (path, name) => {
      await fetch(GRAPH_BASE + path, {
        method: 'POST',
        headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({ name, folder: {}, '@microsoft.graph.conflictBehavior': 'fail' }),
      }); // 409 = existiert bereits → ignorieren
    };
    await mkfolder('/me/drive/root/children', 'Tembu');
    await mkfolder('/me/drive/root:/Tembu:/children', 'analysen');
  }

  let _foldersEnsured = false;
  async function saveAnalysis(token, cacheKey, contactName, contactEmail, analysis) {
    if (!_foldersEnsured) {
      await ensureOneDriveFolders(token);
      _foldersEnsured = true;
    }
    await graphPut(token, analysisOneDrivePath(cacheKey), JSON.stringify({
      contact: contactName, email: contactEmail,
      savedAt: new Date().toISOString(), analysis,
    }));
  }

  async function loadAnalysis(token, cacheKey) {
    try {
      const res = await fetch(GRAPH_BASE + analysisOneDrivePath(cacheKey), {
        headers: { Authorization: `Bearer ${token}` },
      });
      if (!res.ok) return null;
      return await res.json();
    } catch { return null; }
  }

  // ── To Do / Rumbles ───────────────────────────────────────────────────────
  function parseBody(text) {
    const r = {};
    for (const line of (text || '').split('\n')) {
      const i = line.indexOf(':');
      if (i > 0) r[line.slice(0, i).trim()] = line.slice(i + 1).trim();
    }
    return r;
  }

  async function getOrCreateTembuList(token) {
    const lists = await graphGet(token, '/me/todo/lists');
    const existing = lists.value?.find(l => l.displayName === TEMBU_LIST);
    if (existing) return existing.id;
    const created = await graphPost(token, '/me/todo/lists', { displayName: TEMBU_LIST });
    return created.id;
  }

  // Gibt { listId, tasks } zurück — tasks sind nicht abgeschlossen ($filter ne completed)
  async function fetchRumbleTasks(token) {
    const lists = await graphGet(token, '/me/todo/lists');
    const list = lists.value?.find(l => l.displayName === TEMBU_LIST);
    if (!list) return { listId: null, tasks: [] };
    const data = await graphGet(token, `/me/todo/lists/${list.id}/tasks?$filter=status ne 'completed'&$top=200`);
    return { listId: list.id, tasks: data.value || [] };
  }

  async function markTaskDone(token, listId, taskId) {
    await graphPatch(token, `/me/todo/lists/${listId}/tasks/${taskId}`, { status: 'completed' });
  }

  // ── KI-Provider ───────────────────────────────────────────────────────────
  // config: { provider, apiKey, model, endpoint }
  // provider: 'gemini' | 'anthropic' | 'openai' | 'groq' | 'ollama' | 'lmstudio'
  async function callAI(prompt, config) {
    const { provider, apiKey, model, endpoint } = config || {};
    switch (provider) {
      case 'anthropic':
        return callAnthropic(prompt, apiKey, model);
      case 'openai':
        return callOpenAICompat(prompt, apiKey, 'https://api.openai.com/v1', model || 'gpt-4o-mini');
      case 'groq':
        return callOpenAICompat(prompt, apiKey, 'https://api.groq.com/openai/v1', model || 'llama-3.1-70b-versatile');
      case 'ollama':
        return callOpenAICompat(prompt, null, (endpoint || 'http://localhost:11434') + '/v1', model || 'qwen2.5:14b');
      case 'lmstudio':
        return callOpenAICompat(prompt, null, (endpoint || 'http://localhost:1234') + '/v1', model || '');
      default:
        return callGemini(prompt, apiKey, model);
    }
  }

  async function callGemini(prompt, apiKey, model) {
    const mdl = model || 'gemini-2.5-flash';
    const res = await fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/${mdl}:generateContent?key=${apiKey}`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          contents: [{ parts: [{ text: prompt }] }],
          generationConfig: { temperature: 0.3, maxOutputTokens: 8192 },
        }),
      }
    );
    if (!res.ok) { const t = await res.text(); throw new Error(`Gemini ${res.status}: ${t.slice(0, 120)}`); }
    const d = await res.json();
    const parts = d.candidates?.[0]?.content?.parts || [];
    return parts.filter(p => !p.thought).map(p => p.text || '').join('');
  }

  async function callAnthropic(prompt, apiKey, model) {
    const mdl = model || 'claude-haiku-4-5-20251001';
    const res = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01', 'Content-Type': 'application/json' },
      body: JSON.stringify({ model: mdl, max_tokens: 4096, messages: [{ role: 'user', content: prompt }] }),
    });
    if (!res.ok) { const t = await res.text(); throw new Error(`Claude ${res.status}: ${t.slice(0, 120)}`); }
    const d = await res.json();
    return d.content?.[0]?.text || '';
  }

  // Einheitliche Funktion für OpenAI, Groq, Ollama, LM Studio (alle sprechen OpenAI API)
  async function callOpenAICompat(prompt, apiKey, baseUrl, model) {
    const headers = { 'Content-Type': 'application/json' };
    if (apiKey) headers['Authorization'] = `Bearer ${apiKey}`;
    const res = await fetch(`${baseUrl}/chat/completions`, {
      method: 'POST',
      headers,
      body: JSON.stringify({
        model: model || 'gpt-4o-mini',
        messages: [{ role: 'user', content: prompt }],
        max_tokens: 4096,
        temperature: 0.3,
      }),
    });
    if (!res.ok) { const t = await res.text(); throw new Error(`${model || 'AI'} ${res.status}: ${t.slice(0, 120)}`); }
    const d = await res.json();
    return d.choices?.[0]?.message?.content || '';
  }

  function parseAIResponse(raw) {
    let text = raw.trim().replace(/^```(?:json)?\n?/, '').replace(/\n?```$/, '').trim();
    try { return JSON.parse(text); } catch {}
    const m = text.match(/\{[\s\S]*\}/);
    if (m) { try { return JSON.parse(m[0]); } catch {} }
    return { summary: text.slice(0, 400), themes: [], openPoints: [], sentiment: 'neutral', nextStep: '', background: '' };
  }

  // ── Utilities ─────────────────────────────────────────────────────────────
  function esc(s) {
    return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  function escapeHtml(s) {
    return esc(s).replace(/"/g, '&quot;');
  }

  return {
    TEMBU_LIST,
    graphGet,
    graphPost,
    graphPatch,
    graphPut,
    analysisOneDrivePath,
    saveAnalysis,
    loadAnalysis,
    parseBody,
    getOrCreateTembuList,
    fetchRumbleTasks,
    markTaskDone,
    callAI,
    parseAIResponse,
    esc,
    escapeHtml,
  };
})();
