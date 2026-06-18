-- Intransit Hub — Supabase schema
-- Run this in the Supabase SQL editor once.

-- ── App event logs ──────────────────────────────────────────────────────────
-- Every automation writes here via POST /api/logs
create table if not exists app_logs (
  id          bigserial primary key,
  app_name    text not null,          -- email_automation | tee_time_bot | icsource_checker | oem_excess
  event_type  text not null,          -- run | error | draft_created | email_sent | booking | no_stock | etc.
  summary     text,                   -- short human-readable message
  details     jsonb,                  -- any extra structured data
  created_at  timestamptz default now()
);

create index if not exists app_logs_app_name_idx    on app_logs (app_name);
create index if not exists app_logs_created_at_idx  on app_logs (created_at desc);
create index if not exists app_logs_event_type_idx  on app_logs (event_type);

-- ── App configs ─────────────────────────────────────────────────────────────
-- Key-value store for per-app configuration (future use)
create table if not exists app_configs (
  app_name    text primary key,
  config      jsonb not null default '{}',
  updated_at  timestamptz default now()
);

-- ── Email decisions ──────────────────────────────────────────────────────────
-- Records what John does with each draft: sent / archived / deleted / edited
-- Used to train the AI reply system over time.
create table if not exists email_decisions (
  id              bigserial primary key,
  thread_id       text,               -- Gmail thread ID
  mpn             text,               -- part number if applicable
  sender          text,               -- who sent the original email
  action          text not null,      -- sent | archived | deleted | edited
  draft_content   text,               -- what Claude drafted
  sent_content    text,               -- what John actually sent (null if archived/deleted)
  edit_diff       text,               -- summary of edits made (future)
  created_at      timestamptz default now()
);

create index if not exists email_decisions_mpn_idx    on email_decisions (mpn);
create index if not exists email_decisions_action_idx on email_decisions (action);

-- ── Row Level Security ───────────────────────────────────────────────────────
-- The Worker uses the service role key and bypasses RLS.
-- Enable RLS so the anon key can't read these tables directly.
alter table app_logs         enable row level security;
alter table app_configs      enable row level security;
alter table email_decisions  enable row level security;
