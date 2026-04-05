create extension if not exists pgcrypto;

create table if not exists public.org_execution_history (
  id uuid primary key default gen_random_uuid(),
  org_id uuid not null references public.organizations(id) on delete cascade,
  user_key text not null default '',
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

create index if not exists org_execution_history_org_id_executed_at_idx
  on public.org_execution_history (org_id, executed_at desc);

alter table public.org_execution_history enable row level security;

do $$
begin
  if not exists (
    select 1
    from pg_policies
    where schemaname = 'public'
      and tablename = 'org_execution_history'
      and policyname = 'org_execution_history_member_all'
  ) then
    create policy "org_execution_history_member_all"
      on public.org_execution_history
      for all
      using (
        exists (
          select 1
          from public.org_members
          where org_members.org_id = org_execution_history.org_id
            and org_members.user_id = auth.uid()
        )
      )
      with check (
        exists (
          select 1
          from public.org_members
          where org_members.org_id = org_execution_history.org_id
            and org_members.user_id = auth.uid()
        )
      );
  end if;
end
$$;

