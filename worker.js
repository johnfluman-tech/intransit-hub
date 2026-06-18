/**
 * Intransit Hub — Cloudflare Worker
 * Central API gateway for all Intransit automations.
 *
 * Routes:
 *   GET  /api/status          — latest log entry per app (for dashboard cards)
 *   GET  /api/logs            — paginated log entries (filters: app, type, limit)
 *   POST /api/logs            — write a log entry (called by each app)
 *   POST /api/claude          — proxy Claude API (keeps API key off the front-end)
 *
 * Env vars (set via: wrangler secret put <NAME>):
 *   HUB_SECRET       — bearer token used by the front-end
 *   CLAUDE_API_KEY   — Anthropic API key
 * D1 binding (wrangler.toml):
 *   DB               — intransit-hub-db
 */

const CORS = {
  'Access-Control-Allow-Origin':  '*',
  'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
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
      if (url.pathname === '/api/status' && request.method === 'GET') {
        return handleStatus(env);
      }
      if (url.pathname === '/api/logs' && request.method === 'GET') {
        return handleGetLogs(url, env);
      }
      if (url.pathname === '/api/logs' && request.method === 'POST') {
        return handlePostLog(request, env);
      }
      if (url.pathname === '/api/claude' && request.method === 'POST') {
        return handleClaude(request, env);
      }
      return json({ error: 'Not found' }, 404);
    } catch (err) {
      return json({ error: err.message }, 500);
    }
  }
};

// ─────────────────────────────────────────────────────
//  GET /api/status
// ─────────────────────────────────────────────────────
async function handleStatus(env) {
  const apps = ['email_automation', 'tee_time_bot', 'icsource_checker', 'oem_excess'];
  const results = {};

  for (const app of apps) {
    const { results: rows } = await env.DB.prepare(
      'SELECT event_type, created_at, summary FROM app_logs WHERE app_name = ? ORDER BY created_at DESC LIMIT 1'
    ).bind(app).all();

    if (rows && rows.length) {
      const r = rows[0];
      results[app] = {
        status:   r.event_type === 'error' ? 'error' : 'ok',
        last_run: r.created_at,
        summary:  r.summary,
      };
    } else {
      results[app] = { status: 'unknown', last_run: null, summary: null };
    }
  }

  return json(results);
}

// ─────────────────────────────────────────────────────
//  GET /api/logs?app=&type=&limit=50
// ─────────────────────────────────────────────────────
async function handleGetLogs(url, env) {
  const app   = url.searchParams.get('app')   || '';
  const type  = url.searchParams.get('type')  || '';
  const limit = Math.min(parseInt(url.searchParams.get('limit') || '50'), 200);

  let sql    = 'SELECT * FROM app_logs';
  const binds = [];
  const where = [];

  if (app)  { where.push('app_name = ?');   binds.push(app); }
  if (type) { where.push('event_type = ?'); binds.push(type); }
  if (where.length) sql += ' WHERE ' + where.join(' AND ');
  sql += ' ORDER BY created_at DESC LIMIT ?';
  binds.push(limit);

  const { results: rows } = await env.DB.prepare(sql).bind(...binds).all();
  return json({ rows: rows || [] });
}

// ─────────────────────────────────────────────────────
//  POST /api/logs
//  Body: { app_name, event_type, summary, details }
// ─────────────────────────────────────────────────────
async function handlePostLog(request, env) {
  const body = await request.json();
  const { app_name, event_type, summary, details } = body;

  if (!app_name || !event_type) {
    return json({ error: 'app_name and event_type are required' }, 400);
  }

  const detailsStr = details ? (typeof details === 'string' ? details : JSON.stringify(details)) : null;

  await env.DB.prepare(
    'INSERT INTO app_logs (app_name, event_type, summary, details) VALUES (?, ?, ?, ?)'
  ).bind(app_name, event_type, summary || null, detailsStr).run();

  return json({ ok: true });
}

// ─────────────────────────────────────────────────────
//  POST /api/claude
//  Body: { messages, system, max_tokens, model }
// ─────────────────────────────────────────────────────
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

// ─────────────────────────────────────────────────────
//  Response helper
// ─────────────────────────────────────────────────────
function json(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { ...CORS, 'Content-Type': 'application/json' },
  });
}
