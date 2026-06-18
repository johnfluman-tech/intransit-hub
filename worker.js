/**
 * Intransit Hub — Cloudflare Worker
 *
 * Routes:
 *   GET/POST     /api/status
 *   GET/POST     /api/logs
 *   GET/POST     /api/drafts
 *   PATCH        /api/drafts/:id
 *   GET/POST     /api/memory          — AI memory store
 *   GET/DELETE   /api/memory/:slug
 *   POST         /api/claude
 *
 * Secrets: HUB_SECRET, CLAUDE_API_KEY   D1 binding: DB
 */

const CORS = {
  'Access-Control-Allow-Origin':  '*',
  'Access-Control-Allow-Methods': 'GET, POST, PATCH, DELETE, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type, Authorization',
};

export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    if (request.method === 'OPTIONS') return new Response(null, { headers: CORS });

    const auth = request.headers.get('Authorization') || '';
    if (auth !== `Bearer ${env.HUB_SECRET}`) return json({ error: 'Unauthorized' }, 401);

    try {
      const p = url.pathname;
      const m = request.method;

      if (p === '/api/status'  && m === 'GET')  return handleStatus(env);
      if (p === '/api/logs'    && m === 'GET')  return handleGetLogs(url, env);
      if (p === '/api/logs'    && m === 'POST') return handlePostLog(request, env);
      if (p === '/api/drafts'  && m === 'GET')  return handleGetDrafts(url, env);
      if (p === '/api/drafts'  && m === 'POST') return handlePostDraft(request, env);
      if (p === '/api/memory'  && m === 'GET')  return handleGetMemory(url, env);
      if (p === '/api/memory'  && m === 'POST') return handlePostMemory(request, env);
      if (p === '/api/claude'  && m === 'POST') return handleClaude(request, env);

      const draftId = p.match(/^\/api\/drafts\/(\d+)$/);
      if (draftId && m === 'PATCH') return handlePatchDraft(request, env, parseInt(draftId[1]));

      const memSlug = p.match(/^\/api\/memory\/([a-zA-Z0-9_-]+)$/);
      if (memSlug && m === 'GET')    return handleGetMemorySingle(env, memSlug[1]);
      if (memSlug && m === 'DELETE') return handleDeleteMemory(env, memSlug[1]);

      return json({ error: 'Not found' }, 404);
    } catch (err) {
      return json({ error: err.message }, 500);
    }
  }
};

// ─── /api/status ────────────────────────────────────
async function handleStatus(env) {
  const apps = ['email_automation', 'tee_time_bot', 'icsource_checker', 'oem_excess'];
  const results = {};
  for (const app of apps) {
    const { results: rows } = await env.DB.prepare(
      'SELECT event_type, created_at, summary FROM app_logs WHERE app_name = ? ORDER BY created_at DESC LIMIT 1'
    ).bind(app).all();
    if (rows && rows.length) {
      const r = rows[0];
      results[app] = { status: r.event_type === 'error' ? 'error' : 'ok', last_run: r.created_at, summary: r.summary };
    } else {
      results[app] = { status: 'unknown', last_run: null, summary: null };
    }
  }
  return json(results);
}

// ─── /api/logs GET ──────────────────────────────────
async function handleGetLogs(url, env) {
  const app   = url.searchParams.get('app')   || '';
  const type  = url.searchParams.get('type')  || '';
  const limit = Math.min(parseInt(url.searchParams.get('limit') || '50'), 200);
  let sql = 'SELECT * FROM app_logs';
  const binds = [], where = [];
  if (app)  { where.push('app_name = ?');   binds.push(app); }
  if (type) { where.push('event_type = ?'); binds.push(type); }
  if (where.length) sql += ' WHERE ' + where.join(' AND ');
  sql += ' ORDER BY created_at DESC LIMIT ?';
  binds.push(limit);
  const { results: rows } = await env.DB.prepare(sql).bind(...binds).all();
  return json({ rows: rows || [] });
}

// ─── /api/logs POST ─────────────────────────────────
async function handlePostLog(request, env) {
  const { app_name, event_type, summary, details } = await request.json();
  if (!app_name || !event_type) return json({ error: 'app_name and event_type are required' }, 400);
  const d = details ? (typeof details === 'string' ? details : JSON.stringify(details)) : null;
  await env.DB.prepare('INSERT INTO app_logs (app_name, event_type, summary, details) VALUES (?, ?, ?, ?)')
    .bind(app_name, event_type, summary || null, d).run();
  return json({ ok: true });
}

// ─── /api/drafts GET ────────────────────────────────
async function handleGetDrafts(url, env) {
  const status = url.searchParams.get('status') || 'pending';
  const limit  = Math.min(parseInt(url.searchParams.get('limit') || '50'), 200);
  const { results: rows } = await env.DB.prepare(
    'SELECT * FROM email_decisions WHERE action = ? ORDER BY created_at DESC LIMIT ?'
  ).bind(status, limit).all();
  return json({ rows: rows || [] });
}

// ─── /api/drafts POST ───────────────────────────────
async function handlePostDraft(request, env) {
  const { thread_id, mpn, sender, subject, draft_content } = await request.json();
  if (!draft_content) return json({ error: 'draft_content is required' }, 400);
  const { meta } = await env.DB.prepare(
    'INSERT INTO email_decisions (thread_id, mpn, sender, subject, action, draft_content) VALUES (?, ?, ?, ?, ?, ?)'
  ).bind(thread_id || null, mpn || null, sender || null, subject || null, 'pending', draft_content).run();
  return json({ ok: true, id: meta.last_row_id });
}

// ─── /api/drafts/:id PATCH ──────────────────────────
async function handlePatchDraft(request, env, id) {
  const { action, sent_content, draft_content } = await request.json();
  if (!action) return json({ error: 'action is required' }, 400);
  if (draft_content !== undefined) {
    await env.DB.prepare('UPDATE email_decisions SET action=?, sent_content=?, draft_content=? WHERE id=?')
      .bind(action, sent_content || null, draft_content, id).run();
  } else {
    await env.DB.prepare('UPDATE email_decisions SET action=?, sent_content=? WHERE id=?')
      .bind(action, sent_content || null, id).run();
  }
  return json({ ok: true });
}

// ─── /api/memory GET (list) ─────────────────────────
async function handleGetMemory(url, env) {
  const type  = url.searchParams.get('type') || '';
  const search = url.searchParams.get('q')   || '';
  let sql = 'SELECT slug, description, type, updated_at FROM ai_memory';
  const binds = [], where = [];
  if (type)   { where.push('type = ?');                binds.push(type); }
  if (search) { where.push('(slug LIKE ? OR description LIKE ? OR body LIKE ?)');
                binds.push(`%${search}%`, `%${search}%`, `%${search}%`); }
  if (where.length) sql += ' WHERE ' + where.join(' AND ');
  sql += ' ORDER BY type, slug';
  const { results: rows } = await env.DB.prepare(sql).bind(...binds).all();
  return json({ rows: rows || [] });
}

// ─── /api/memory/:slug GET ──────────────────────────
async function handleGetMemorySingle(env, slug) {
  const { results: rows } = await env.DB.prepare('SELECT * FROM ai_memory WHERE slug = ?').bind(slug).all();
  if (!rows || !rows.length) return json({ error: 'Not found' }, 404);
  return json(rows[0]);
}

// ─── /api/memory POST (upsert) ──────────────────────
async function handlePostMemory(request, env) {
  const { slug, description, type, body } = await request.json();
  if (!slug || !body) return json({ error: 'slug and body are required' }, 400);
  await env.DB.prepare(
    `INSERT INTO ai_memory (slug, description, type, body, updated_at)
     VALUES (?, ?, ?, ?, datetime('now'))
     ON CONFLICT(slug) DO UPDATE SET description=excluded.description,
       type=excluded.type, body=excluded.body, updated_at=datetime('now')`
  ).bind(slug, description || '', type || 'feedback', body).run();
  return json({ ok: true });
}

// ─── /api/memory/:slug DELETE ───────────────────────
async function handleDeleteMemory(env, slug) {
  await env.DB.prepare('DELETE FROM ai_memory WHERE slug = ?').bind(slug).run();
  return json({ ok: true });
}

// ─── /api/claude ────────────────────────────────────
async function handleClaude(request, env) {
  const body = await request.json();
  const res = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: { 'x-api-key': env.CLAUDE_API_KEY, 'anthropic-version': '2023-06-01', 'content-type': 'application/json' },
    body: JSON.stringify({
      model:      body.model      || 'claude-haiku-4-5-20251001',
      max_tokens: body.max_tokens || 1024,
      system:     body.system     || undefined,
      messages:   body.messages,
    }),
  });
  return json(await res.json(), res.status);
}

// ─── Response helper ────────────────────────────────
function json(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status, headers: { ...CORS, 'Content-Type': 'application/json' },
  });
}
