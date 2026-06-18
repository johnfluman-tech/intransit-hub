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
 * Env vars (set in wrangler.toml / Cloudflare dashboard):
 *   HUB_SECRET       — bearer token used by the front-end
 *   SUPABASE_URL     — e.g. https://xxxx.supabase.co
 *   SUPABASE_KEY     — service role key
 *   CLAUDE_API_KEY   — Anthropic API key
 *   APPS_SCRIPT_LOG_URL — (optional) Google Apps Script URL that accepts POST logs
 */

const CORS = {
  'Access-Control-Allow-Origin':  '*',
  'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type, Authorization',
};

export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    // Preflight
    if (request.method === 'OPTIONS') {
      return new Response(null, { headers: CORS });
    }

    // Auth — every request must carry Bearer <HUB_SECRET>
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
//  Returns the latest log entry per app_name.
// ─────────────────────────────────────────────────────
async function handleStatus(env) {
  const apps = ['email_automation', 'tee_time_bot', 'icsource_checker', 'oem_excess'];

  // One query per app — could be optimised with a SQL DISTINCT ON later
  const results = {};
  for (const app of apps) {
    const row = await sbGet(
      env,
      `app_logs?app_name=eq.${app}&order=created_at.desc&limit=1`
    );
    if (row && row.length) {
      const r = row[0];
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

  let query = `app_logs?order=created_at.desc&limit=${limit}`;
  if (app)  query += `&app_name=eq.${app}`;
  if (type) query += `&event_type=eq.${type}`;

  const rows = await sbGet(env, query);
  return json({ rows: rows || [] });
}

// ─────────────────────────────────────────────────────
//  POST /api/logs
//  Body: { app_name, event_type, summary, details }
//  Called by each automation app to record events.
// ─────────────────────────────────────────────────────
async function handlePostLog(request, env) {
  const body = await request.json();
  const { app_name, event_type, summary, details } = body;

  if (!app_name || !event_type) {
    return json({ error: 'app_name and event_type are required' }, 400);
  }

  await sbPost(env, 'app_logs', {
    app_name,
    event_type,
    summary:  summary || null,
    details:  details || null,
  });

  return json({ ok: true });
}

// ─────────────────────────────────────────────────────
//  POST /api/claude
//  Body: { messages, system, max_tokens }
//  Proxies to Claude API — keeps API key server-side.
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
//  Supabase helpers
// ─────────────────────────────────────────────────────
function sbHeaders(env) {
  return {
    'apikey':        env.SUPABASE_KEY,
    'Authorization': `Bearer ${env.SUPABASE_KEY}`,
    'Content-Type':  'application/json',
    'Prefer':        'return=representation',
  };
}

async function sbGet(env, path) {
  const res = await fetch(`${env.SUPABASE_URL}/rest/v1/${path}`, {
    headers: sbHeaders(env),
  });
  if (!res.ok) return null;
  return res.json();
}

async function sbPost(env, table, data) {
  const res = await fetch(`${env.SUPABASE_URL}/rest/v1/${table}`, {
    method:  'POST',
    headers: sbHeaders(env),
    body:    JSON.stringify(data),
  });
  if (!res.ok) {
    const err = await res.text();
    throw new Error(`Supabase error: ${err}`);
  }
  return res.json();
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
