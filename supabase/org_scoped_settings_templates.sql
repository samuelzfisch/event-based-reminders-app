create extension if not exists pgcrypto;

create table if not exists public.org_settings (
  id uuid primary key default gen_random_uuid(),
  org_id uuid not null references public.organizations(id) on delete cascade unique,
  default_reminder_time text,
  default_press_release_time text,
  email_signature_enabled boolean not null default false,
  email_signature_text text not null default '',
  outlook_account_email text not null default '',
  email_handling_mode text not null default 'draft',
  updated_at timestamptz not null default now()
);

create table if not exists public.org_plan_templates (
  id text not null,
  org_id uuid not null references public.organizations(id) on delete cascade,
  name text not null,
  base_type text not null,
  template_mode text,
  is_protected boolean not null default false,
  weekend_rule text not null,
  anchors jsonb not null default '[]'::jsonb,
  items jsonb not null default '[]'::jsonb,
  sort_order integer not null default 0,
  updated_at timestamptz not null default now(),
  primary key (org_id, id)
);

create index if not exists org_plan_templates_org_id_sort_order_idx
  on public.org_plan_templates (org_id, sort_order asc);

alter table public.org_settings enable row level security;
alter table public.org_plan_templates enable row level security;

do $$
begin
  if not exists (
    select 1
    from pg_policies
    where schemaname = 'public'
      and tablename = 'org_settings'
      and policyname = 'org_settings_member_all'
  ) then
    create policy "org_settings_member_all"
      on public.org_settings
      for all
      using (
        exists (
          select 1
          from public.org_members
          where org_members.org_id = org_settings.org_id
            and org_members.user_id = auth.uid()
        )
      )
      with check (
        exists (
          select 1
          from public.org_members
          where org_members.org_id = org_settings.org_id
            and org_members.user_id = auth.uid()
        )
      );
  end if;

  if not exists (
    select 1
    from pg_policies
    where schemaname = 'public'
      and tablename = 'org_plan_templates'
      and policyname = 'org_plan_templates_member_all'
  ) then
    create policy "org_plan_templates_member_all"
      on public.org_plan_templates
      for all
      using (
        exists (
          select 1
          from public.org_members
          where org_members.org_id = org_plan_templates.org_id
            and org_members.user_id = auth.uid()
        )
      )
      with check (
        exists (
          select 1
          from public.org_members
          where org_members.org_id = org_plan_templates.org_id
            and org_members.user_id = auth.uid()
        )
      );
  end if;
end
$$;

