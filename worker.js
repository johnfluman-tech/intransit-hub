/**
 * Intransit Hub — Cloudflare Worker
 *
 * Routes:
 *   GET  /api/status          — latest log entry per app
 *   GET  /api/logs            — paginated log entries
 *   POST /api/logs            — write a log entry
 *   GET  /api/drafts          — email drafts pending review
 *   POST /api/drafts          — create a draft record (called by email script)
 *   PATCH /api/drafts/:id     — update draft action/content
 *   POST /api/claude          — proxy Claude API
 *
 * Secrets: HUB_SECRET, CLAUDE_API_KEY
 * D1 binding: DB (intransit-hub-db)
 */

const CORS = {
  'Access-Control-Allow-Origin':  '*',
  'Access-Control-Allow-Methods': 'GET, POST, PATCH, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type, Authorization',
};

export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    if (request.method === 'OPTIONS') {
      return new Response(null, { headers: CORS });
    }

    const auth = request.headers.get('Authorization') || '';
    if (auth !== `Bearer ${env.HUB_SECRET}`) {
      return json({ error: 'Unauthorized' }, 401);
    }

    try {
      if (url.pathname === '/api/status' && request.method === 'GET')
        return handleStatus(env);

      if (url.pathname === '/api/logs' && request.method === 'GET')
        return handleGetLogs(url, env);

      if (url.pathname === '/api/logs' && request.method === 'POST')
        return handlePostLog(request, env);

      if (url.pathname === '/api/drafts' && request.method === 'GET')
        return handleGetDrafts(url, env);

      if (url.pathname === '/api/drafts' && request.method === 'POST')
        return handlePostDraft(request, env);

      const patchMatch = url.pathname.match(/^\/api\/drafts\/(\d+)$/);
      if (patchMatch && request.method === 'PATCH')
        return handlePatchDraft(request, env, parseInt(patchMatch[1]));

      if (url.pathname === '/api/claude' && request.method === 'POST')
        return handleClaude(request, env);

      return json({ error: 'Not found' }, 404);
    } catch (err) {
      return json({ error: err.message }, 500);
    }
  }
};

// ─── GET /api/status ────────────────────────────────
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

// ─── GET /api/logs ──────────────────────────────────
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

// ─── POST /api/logs ─────────────────────────────────
async function handlePostLog(request, env) {
  const body = await request.json();
  const { app_name, event_type, summary, details } = body;
  if (!app_name || !event_type) return json({ error: 'app_name and event_type are required' }, 400);
  const detailsStr = details ? (typeof details === 'string' ? details : JSON.stringify(details)) : null;
  await env.DB.prepare(
    'INSERT INTO app_logs (app_name, event_type, summary, details) VALUES (?, ?, ?, ?)'
  ).bind(app_name, event_type, summary || null, detailsStr).run();
  return json({ ok: true });
}

// ─── GET /api/drafts ────────────────────────────────
async function handleGetDrafts(url, env) {
  const status = url.searchParams.get('status') || 'pending';
  const limit  = Math.min(parseInt(url.searchParams.get('limit') || '50'), 200);
  const { results: rows } = await env.DB.prepare(
    'SELECT * FROM email_decisions WHERE action = ? ORDER BY created_at DESC LIMIT ?'
  ).bind(status, limit).all();
  return json({ rows: rows || [] });
}

// ─── POST /api/drafts ───────────────────────────────
async function handlePostDraft(request, env) {
  const body = await request.json();
  const { thread_id, mpn, sender, subject, draft_content } = body;
  if (!draft_content) return json({ error: 'draft_content is required' }, 400);
  const { meta } = await env.DB.prepare(
    'INSERT INTO email_decisions (thread_id, mpn, sender, subject, action, draft_content) VALUES (?, ?, ?, ?, ?, ?)'
  ).bind(thread_id || null, mpn || null, sender || null, subject || null, 'pending', draft_content).run();
  return json({ ok: true, id: meta.last_row_id });
}

// ─── PATCH /api/drafts/:id ──────────────────────────
async function handlePatchDraft(request, env, id) {
  const body = await request.json();
  const { action, sent_content, draft_content } = body;
  if (!action) return json({ error: 'action is required' }, 400);

  if (draft_content !== undefined) {
    await env.DB.prepare(
      'UPDATE email_decisions SET action = ?, sent_content = ?, draft_content = ? WHERE id = ?'
    ).bind(action, sent_content || null, draft_content, id).run();
  } else {
    await env.DB.prepare(
      'UPDATE email_decisions SET action = ?, sent_content = ? WHERE id = ?'
    ).bind(action, sent_content || null, id).run();
  }
  return json({ ok: true });
}

// ─── POST /api/claude ───────────────────────────────
async function handleClaude(request, env) {
  const body = await request.json();
  const res = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'x-api-key':         env.CLAUDE_API_KEY,
      'anthropic-version': '2023-06-01',
      'content-type':      'application/json',
    },
    body: JSON.stringify({
      model:      body.model      || 'claude-haiku-4-5-20251001',
      max_tokens: body.max_tokens || 1024,
      system:     body.system     || undefined,
      messages:   body.messages,
    }),
  });
  const data = await res.json();
  return json(data, res.status);
}

// ─── Response helper ────────────────────────────────
function json(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { ...CORS, 'Content-Type': 'application/json' },
  });
}
