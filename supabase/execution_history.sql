create extension if not exists pgcrypto;

create table if not exists public.execution_history (
  id uuid primary key default gen_random_uuid(),
  user_key text not null,
  execution_group_id uuid,
  plan_name text not null default '',
  item_type text not null,
  title text not null default '',
  subject text not null default '',
  status text not null,
  path text not null,
  recipients jsonb not null default '[]'::jsonb,
  attendees jsonb not null default '[]'::jsonb,
  executed_at timestamptz not null default now(),
  outlook_web_link text,
  teams_join_link text,
  fallback_export_kind text,
  details jsonb not null default '{}'::jsonb
);

create index if not exists execution_history_user_key_executed_at_idx
  on public.execution_history (user_key, executed_at desc);

alter table public.execution_history enable row level security;

do $$
begin
  if not exists (
    select 1
    from pg_policies
    where schemaname = 'public'
      and tablename = 'execution_history'
      and policyname = 'execution_history_anon_all'
  ) then
    create policy "execution_history_anon_all"
      on public.execution_history
      for all
      using (true)
      with check (true);
  end if;
end
$$;
