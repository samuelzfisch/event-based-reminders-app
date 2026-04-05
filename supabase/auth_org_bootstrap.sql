create extension if not exists pgcrypto;

create table if not exists public.organizations (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  created_by uuid not null,
  created_at timestamptz not null default now()
);

create table if not exists public.org_members (
  id uuid primary key default gen_random_uuid(),
  org_id uuid not null references public.organizations(id) on delete cascade,
  user_id uuid not null,
  role text not null default 'member',
  created_at timestamptz not null default now(),
  unique (org_id, user_id)
);

create index if not exists org_members_user_id_idx
  on public.org_members (user_id, created_at asc);

alter table public.organizations enable row level security;
alter table public.org_members enable row level security;

do $$
begin
  if not exists (
    select 1
    from pg_policies
    where schemaname = 'public'
      and tablename = 'organizations'
      and policyname = 'organizations_select_member'
  ) then
    create policy "organizations_select_member"
      on public.organizations
      for select
      using (
        exists (
          select 1
          from public.org_members
          where org_members.org_id = organizations.id
            and org_members.user_id = auth.uid()
        )
      );
  end if;

  if not exists (
    select 1
    from pg_policies
    where schemaname = 'public'
      and tablename = 'organizations'
      and policyname = 'organizations_insert_owner'
  ) then
    create policy "organizations_insert_owner"
      on public.organizations
      for insert
      with check (created_by = auth.uid());
  end if;

  if not exists (
    select 1
    from pg_policies
    where schemaname = 'public'
      and tablename = 'org_members'
      and policyname = 'org_members_select_self'
  ) then
    create policy "org_members_select_self"
      on public.org_members
      for select
      using (user_id = auth.uid());
  end if;

  if not exists (
    select 1
    from pg_policies
    where schemaname = 'public'
      and tablename = 'org_members'
      and policyname = 'org_members_insert_self'
  ) then
    create policy "org_members_insert_self"
      on public.org_members
      for insert
      with check (user_id = auth.uid());
  end if;
end
$$;

