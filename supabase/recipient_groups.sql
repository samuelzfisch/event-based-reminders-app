create extension if not exists pgcrypto;

create table if not exists public.org_recipient_groups (
  id uuid primary key default gen_random_uuid(),
  org_id uuid not null references public.organizations(id) on delete cascade,
  name text not null,
  emails jsonb not null default '[]'::jsonb,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create index if not exists org_recipient_groups_org_id_name_idx
  on public.org_recipient_groups (org_id, name asc);

alter table public.org_recipient_groups enable row level security;

do $$
begin
  if not exists (
    select 1
    from pg_policies
    where schemaname = 'public'
      and tablename = 'org_recipient_groups'
      and policyname = 'org_recipient_groups_member_all'
  ) then
    create policy "org_recipient_groups_member_all"
      on public.org_recipient_groups
      for all
      using (
        exists (
          select 1
          from public.org_members
          where org_members.org_id = org_recipient_groups.org_id
            and org_members.user_id = auth.uid()
        )
      )
      with check (
        exists (
          select 1
          from public.org_members
          where org_members.org_id = org_recipient_groups.org_id
            and org_members.user_id = auth.uid()
        )
      );
  end if;
end
$$;
