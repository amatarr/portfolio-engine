-- One-time setup for per-user saved portfolios in the Streamlit app.
-- Run in the Supabase dashboard: SQL Editor > New query > paste > Run.
--
-- Each row is one portfolio (the same CSV the app downloads), keyed by the
-- user name typed at the top of the app and a portfolio name. The app's
-- autosave lives under the name "__autosave__".
--
-- The app reads and writes with the anon key (kept in Streamlit secrets),
-- so these policies let that key read/write this table only.

create table if not exists public.saved_portfolios (
    user_name  text        not null,
    name       text        not null,
    csv        text        not null,
    updated_at timestamptz not null default now(),
    primary key (user_name, name)
);

alter table public.saved_portfolios enable row level security;

create policy "app can read saved portfolios"
    on public.saved_portfolios for select to anon using (true);

create policy "app can add saved portfolios"
    on public.saved_portfolios for insert to anon with check (true);

create policy "app can update saved portfolios"
    on public.saved_portfolios for update to anon using (true) with check (true);

create policy "app can delete saved portfolios"
    on public.saved_portfolios for delete to anon using (true);
