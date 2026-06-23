-- Intransit Hub — Cloudflare D1 schema
-- Run with: wrangler d1 execute intransit-hub-db --file=schema.sql

CREATE TABLE IF NOT EXISTS app_logs (
  id          INTEGER PRIMARY KEY AUTOINCREMENT,
  app_name    TEXT NOT NULL,
  event_type  TEXT NOT NULL,
  summary     TEXT,
  details     TEXT,  -- JSON stored as text in SQLite
  created_at  TEXT DEFAULT (datetime('now'))
);

CREATE INDEX IF NOT EXISTS idx_app_logs_app    ON app_logs (app_name);
CREATE INDEX IF NOT EXISTS idx_app_logs_time   ON app_logs (created_at DESC);
CREATE INDEX IF NOT EXISTS idx_app_logs_type   ON app_logs (event_type);

CREATE TABLE IF NOT EXISTS app_configs (
  app_name    TEXT PRIMARY KEY,
  config      TEXT NOT NULL DEFAULT '{}',
  updated_at  TEXT DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS email_decisions (
  id              INTEGER PRIMARY KEY AUTOINCREMENT,
  thread_id       TEXT,
  mpn             TEXT,
  sender          TEXT,
  action          TEXT NOT NULL,
  draft_content   TEXT,
  sent_content    TEXT,
  created_at      TEXT DEFAULT (datetime('now'))
);

CREATE INDEX IF NOT EXISTS idx_email_decisions_mpn ON email_decisions (mpn);

CREATE TABLE IF NOT EXISTS agent_decisions (
  id              INTEGER PRIMARY KEY AUTOINCREMENT,
  thread_id       TEXT,
  mpn             TEXT,
  sender          TEXT,
  subject         TEXT,
  action          TEXT NOT NULL,
  reasoning       TEXT,
  draft_body      TEXT,
  forte_entry     TEXT,
  gmail_draft_id  TEXT,
  status          TEXT DEFAULT 'pending',
  created_at      TEXT DEFAULT (datetime('now'))
);

CREATE INDEX IF NOT EXISTS idx_agent_decisions_thread ON agent_decisions (thread_id);
CREATE INDEX IF NOT EXISTS idx_agent_decisions_status ON agent_decisions (status);
CREATE INDEX IF NOT EXISTS idx_agent_decisions_time   ON agent_decisions (created_at DESC);

CREATE TABLE IF NOT EXISTS pending_issues (
  id              INTEGER PRIMARY KEY AUTOINCREMENT,
  created_at      TEXT DEFAULT (datetime('now')),
  updated_at      TEXT DEFAULT (datetime('now')),
  thread_id       TEXT,
  mpn             TEXT,
  description     TEXT NOT NULL,
  context         TEXT,
  status          TEXT DEFAULT 'pending',
  fix_description TEXT,
  fix_commit      TEXT
);

CREATE INDEX IF NOT EXISTS idx_issues_status ON pending_issues (status);
CREATE INDEX IF NOT EXISTS idx_issues_time   ON pending_issues (created_at DESC);
