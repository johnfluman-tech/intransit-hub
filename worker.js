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
 *   GET          /api/apps            — rich status for all 6 apps
 *   POST         /api/email-agent     — AI email processing agent
 *   GET          /api/agent-decisions — agent decision history
 *   PATCH        /api/agent-decisions/:id
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

      if (p === '/api/configs' && m === 'GET') return handleGetConfigs(env);
      const cfgMatch = p.match(/^\/api\/configs\/([a-zA-Z0-9_-]+)$/);
      if (cfgMatch && m === 'GET')  return handleGetConfig(env, cfgMatch[1]);
      if (cfgMatch && m === 'POST') return handlePostConfig(request, env, cfgMatch[1]);

      if (p === '/api/inbox' && m === 'GET')  return handleGetInbox(env);
      if (p === '/api/inbox' && m === 'POST') return handlePostInbox(request, env);

      if (p === '/api/apps'             && m === 'GET')  return handleGetApps(env);
      if (p === '/api/email-agent'     && m === 'POST') return handleEmailAgent(request, env);
      if (p === '/api/agent-decisions' && m === 'GET')  return handleGetAgentDecisions(url, env);

      const agentId = p.match(/^\/api\/agent-decisions\/(\d+)$/);
      if (agentId && m === 'PATCH') return handlePatchAgentDecision(request, env, parseInt(agentId[1]));

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
  const body = await request.json();
  const { action, draft_content } = body;
  if (!action) return json({ error: 'action is required' }, 400);
  const hasSent = 'sent_content' in body;
  if (draft_content !== undefined && hasSent) {
    await env.DB.prepare('UPDATE email_decisions SET action=?, sent_content=?, draft_content=? WHERE id=?')
      .bind(action, body.sent_content || null, draft_content, id).run();
  } else if (draft_content !== undefined) {
    await env.DB.prepare('UPDATE email_decisions SET action=?, draft_content=? WHERE id=?')
      .bind(action, draft_content, id).run();
  } else if (hasSent) {
    await env.DB.prepare('UPDATE email_decisions SET action=?, sent_content=? WHERE id=?')
      .bind(action, body.sent_content || null, id).run();
  } else {
    await env.DB.prepare('UPDATE email_decisions SET action=? WHERE id=?')
      .bind(action, id).run();
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

// ─── /api/configs GET all ───────────────────────────
async function handleGetConfigs(env) {
  const { results: rows } = await env.DB.prepare(
    'SELECT app_name, config, updated_at FROM app_configs ORDER BY app_name'
  ).all();
  return json({ rows: rows || [] });
}

// ─── /api/configs/:app GET ───────────────────────────
async function handleGetConfig(env, app) {
  const { results: rows } = await env.DB.prepare(
    'SELECT * FROM app_configs WHERE app_name = ?'
  ).bind(app).all();
  if (!rows || !rows.length) return json({ app_name: app, config: '{}', updated_at: null });
  return json(rows[0]);
}

// ─── /api/configs/:app POST ──────────────────────────
async function handlePostConfig(request, env, app) {
  const body = await request.json();
  const config = typeof body.config === 'string' ? body.config : JSON.stringify(body.config, null, 2);
  await env.DB.prepare(
    `INSERT INTO app_configs (app_name, config, updated_at) VALUES (?, ?, datetime('now'))
     ON CONFLICT(app_name) DO UPDATE SET config=excluded.config, updated_at=datetime('now')`
  ).bind(app, config).run();
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

// ─── /api/inbox GET ─────────────────────────────────
async function handleGetInbox(env) {
  // Return preview rows for threads not yet reviewed or acted on
  const { results: rows } = await env.DB.prepare(`
    SELECT * FROM email_decisions
    WHERE action = 'preview'
    AND (thread_id IS NULL OR thread_id NOT IN (
      SELECT thread_id FROM email_decisions
      WHERE action IN ('correct','wrong','ignored','pending') AND thread_id IS NOT NULL
    ))
    ORDER BY created_at DESC LIMIT 100
  `).all();
  return json({ rows: rows || [] });
}

// ─── /api/inbox POST ─────────────────────────────────
async function handlePostInbox(request, env) {
  const { thread_id, mpn, sender, subject, draft_content } = await request.json();
  if (!thread_id) return json({ error: 'thread_id is required' }, 400);
  await env.DB.prepare("DELETE FROM email_decisions WHERE thread_id = ? AND action = 'preview'")
    .bind(thread_id).run();
  const { meta } = await env.DB.prepare(
    "INSERT INTO email_decisions (thread_id, mpn, sender, subject, action, draft_content) VALUES (?, ?, ?, ?, 'preview', ?)"
  ).bind(thread_id, mpn || null, sender || null, subject || null, draft_content || '').run();
  return json({ ok: true, id: meta.last_row_id });
}

// ─── /api/apps GET ──────────────────────────────────
async function handleGetApps(env) {
  const [{ results: logRows }, { results: cfgRows }] = await Promise.all([
    env.DB.prepare(
      `SELECT app_name, event_type, created_at, summary FROM app_logs
       WHERE id IN (SELECT MAX(id) FROM app_logs GROUP BY app_name)`
    ).all(),
    env.DB.prepare('SELECT app_name, config FROM app_configs').all(),
  ]);

  const logMap = {};
  (logRows || []).forEach(r => { logMap[r.app_name] = r; });
  const cfgMap = {};
  (cfgRows || []).forEach(r => {
    try { cfgMap[r.app_name] = JSON.parse(r.config || '{}'); } catch { cfgMap[r.app_name] = {}; }
  });

  const appNames = ['email_automation', 'tee_time_bot', 'icsource_checker', 'oem_excess', 'sales_app', 'build_results', 'live_monitor'];
  const results = {};
  for (const app of appNames) {
    const log = logMap[app];
    const cfg = cfgMap[app] || {};
    results[app] = {
      status:   log ? (log.event_type === 'error' ? 'error' : 'ok') : 'unknown',
      last_run: log ? log.created_at : null,
      summary:  log ? log.summary : null,
      version:  cfg.version || null,
      enabled:  cfg.enabled !== false,
    };
  }
  return json(results);
}

// ─── /api/email-agent POST ──────────────────────────
const AGENT_SYSTEM_PROMPT = `You are an AI email processing agent for Intransit Technologies, an electronic components distributor specializing in OEM excess inventory (surplus stock from manufacturers and OEMs).

Your job: analyze the incoming email thread, evaluate the provided inventory and Forte data, and return a precise JSON decision. Return ONLY valid JSON — no explanation, no markdown, no extra text.

## ACTIONS (pick exactly one)
- msg_checking: Part IS in OEM excess AND buyer has given a target price → draft checking reply + Forte entry
- request_tp_500: Part IS in OEM excess AND buyer has NOT given a TP → ask for TP ($500 min)
- request_tp_2000: Part IS in OEM excess, OEM notes say "$2,000 MIN" AND buyer has NOT given TP → ask for TP ($2,000 min)
- bill_handle: Part is a standard catalog item (connectors, passives, common memory/CPU not in OEM excess) → route to Bill
- no_bid: Part NOT found in OEM excess → silent, no reply
- no_action: Thread is internal, already has MSG_CHECKING from John, or no actionable request
- forward_deb: Email is a payment advice / remittance notification from a bank or ERP

## DECISION RULES
1. TP (target price) = per-unit price buyer is willing to pay. Look for "TP: 45", "target $2.50", "our budget is $X each", etc.
2. oem_results empty → no_bid (even if buyer gave TP).
3. oem_results present + buyer gave TP → msg_checking.
4. oem_results present + NO TP → request_tp_500 (or _2000 if notes contain "$2,000").
5. Thread already contains "We are checking on it now" from John → no_action.
6. Sender from @intransittech.com → no_action.
7. forte_results shows existing entry within 60 days → set forte_entry to null (no duplicate).
8. forte_entry only populated when action = msg_checking AND qty AND target_price are both known.
9. Never invent a qty or TP — only use what the buyer explicitly stated.
10. Country: extract from buyer's company address (US, CA, NL, DE, GB, JP, etc.).

## STANDARD DRAFT TEXTS
MSG_CHECKING body: "We are checking on it now. If we get a response from the OEM, I will respond to you right away. If we do not respond back to you, please consider this a no bid. Thank you very much for the opportunity."
REQUEST TP $500 body: "We need a target price to proceed. Please note there is a $500 minimum line requirement. Once we have your target we will get back to you right away."
REQUEST TP $2000 body: "We need a target price to proceed. Please note there is a $2,000 minimum line requirement. Once we have your target we will get back to you right away."
BILL body: "Bill will help with this request"

NOTE: Do NOT include a signature in draft_body. The signature is appended automatically with full HTML formatting. draft_body should contain only the message text.

## RESPONSE FORMAT — return exactly this JSON structure
{
  "action": "one of the 7 actions",
  "reasoning": "1-2 sentences explaining the decision",
  "mpn": "exact MPN from thread or null",
  "buyer_email": "buyer reply-to email or null",
  "buyer_country": "2-letter ISO code or null",
  "qty": number or null,
  "target_price": number or null,
  "draft_body": "complete plain-text email body with signature appended, or null",
  "forte_entry": {"mpn":"...","qty":N,"target_price":N,"country":"XX"} or null
}`;

async function handleEmailAgent(request, env) {
  const body = await request.json();
  const { thread_id, last_message_id, subject, sender, thread_content, oem_results, forte_results, current_labels } = body;

  const userMessage =
    `EMAIL THREAD\nSubject: ${subject || '(none)'}\nSender: ${sender || '(unknown)'}\nCurrent labels: ${(current_labels || []).join(', ') || 'none'}\n\n` +
    `THREAD CONTENT:\n${thread_content || '(empty)'}\n\n` +
    `OEM EXCESS RESULTS:\n${JSON.stringify(oem_results || [], null, 2)}\n\n` +
    `FORTE 60-DAY DUPLICATE CHECK (existing entries):\n${JSON.stringify(forte_results || [], null, 2)}\n\n` +
    `Analyze this thread and return your JSON decision.`;

  const claudeRes = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'x-api-key': env.CLAUDE_API_KEY,
      'anthropic-version': '2023-06-01',
      'content-type': 'application/json',
    },
    body: JSON.stringify({
      model: 'claude-haiku-4-5-20251001',
      max_tokens: 1024,
      system: AGENT_SYSTEM_PROMPT,
      messages: [{ role: 'user', content: userMessage }],
    }),
  });

  const claudeData = await claudeRes.json();
  if (!claudeData.content || !claudeData.content[0]) {
    return json({ error: 'Claude API error', raw: claudeData }, 500);
  }

  let decision;
  try {
    const raw = claudeData.content[0].text.replace(/^```(?:json)?\s*/i, '').replace(/\s*```\s*$/, '').trim();
    decision = JSON.parse(raw);
  } catch (e) {
    return json({ error: 'Claude returned non-JSON', raw: claudeData.content[0].text }, 500);
  }

  const { meta } = await env.DB.prepare(
    `INSERT INTO agent_decisions (thread_id, mpn, sender, subject, action, reasoning, draft_body, forte_entry, status)
     VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'pending')`
  ).bind(
    thread_id || null,
    decision.mpn || null,
    sender || null,
    subject || null,
    decision.action,
    decision.reasoning || null,
    decision.draft_body || null,
    decision.forte_entry ? JSON.stringify(decision.forte_entry) : null
  ).run();

  return json({ ...decision, id: meta.last_row_id });
}

// ─── /api/agent-decisions GET ───────────────────────
async function handleGetAgentDecisions(url, env) {
  const status = url.searchParams.get('status') || '';
  const limit  = Math.min(parseInt(url.searchParams.get('limit') || '50'), 200);
  let sql = 'SELECT * FROM agent_decisions';
  const binds = [];
  if (status) { sql += ' WHERE status = ?'; binds.push(status); }
  sql += ' ORDER BY created_at DESC LIMIT ?';
  binds.push(limit);
  const { results: rows } = await env.DB.prepare(sql).bind(...binds).all();
  return json({ rows: rows || [] });
}

// ─── /api/agent-decisions/:id PATCH ────────────────
async function handlePatchAgentDecision(request, env, id) {
  const { status, gmail_draft_id } = await request.json();
  if (!status) return json({ error: 'status required' }, 400);
  if (gmail_draft_id !== undefined) {
    await env.DB.prepare('UPDATE agent_decisions SET status=?, gmail_draft_id=? WHERE id=?')
      .bind(status, gmail_draft_id || null, id).run();
  } else {
    await env.DB.prepare('UPDATE agent_decisions SET status=? WHERE id=?').bind(status, id).run();
  }
  return json({ ok: true });
}

// ─── Response helper ────────────────────────────────
function json(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status, headers: { ...CORS, 'Content-Type': 'application/json' },
  });
}
