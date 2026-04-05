create extension if not exists pgcrypto;

create table if not exists public.provider_integrations (
  id uuid primary key default gen_random_uuid(),
  org_id uuid not null references public.organizations(id) on delete cascade,
  provider text not null,
  connection_status text not null default 'not_connected',
  provider_account_id text,
  provider_account_email text,
  provider_display_name text,
  access_token text,
  refresh_token text,
  expires_at timestamptz,
  scope text,
  identity jsonb not null default '{}'::jsonb,
  created_by uuid,
  updated_by uuid,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (org_id, provider)
);

create index if not exists provider_integrations_org_id_provider_idx
  on public.provider_integrations (org_id, provider);

alter table public.provider_integrations enable row level security;

do $$
begin
  if not exists (
    select 1
    from pg_policies
    where schemaname = 'public'
      and tablename = 'provider_integrations'
      and policyname = 'provider_integrations_member_all'
  ) then
    create policy "provider_integrations_member_all"
      on public.provider_integrations
      for all
      using (
        exists (
          select 1
          from public.org_members
          where org_members.org_id = provider_integrations.org_id
            and org_members.user_id = auth.uid()
        )
      )
      with check (
        exists (
          select 1
          from public.org_members
          where org_members.org_id = provider_integrations.org_id
            and org_members.user_id = auth.uid()
        )
      );
  end if;
end
$$;

