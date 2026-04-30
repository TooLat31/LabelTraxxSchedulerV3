create table if not exists public.app_state (
  id text primary key,
  payload jsonb not null default '{}'::jsonb,
  updated_at timestamptz not null default timezone('utc', now()),
  updated_by text not null default 'system'
);

create or replace function public.set_app_state_updated_at()
returns trigger
language plpgsql
as $$
begin
  new.updated_at = timezone('utc', now());
  return new;
end;
$$;

drop trigger if exists app_state_set_updated_at on public.app_state;

create trigger app_state_set_updated_at
before update on public.app_state
for each row
execute function public.set_app_state_updated_at();

alter table public.app_state enable row level security;

drop policy if exists "app_state_select_all" on public.app_state;
create policy "app_state_select_all"
on public.app_state
for select
using (true);

drop policy if exists "app_state_insert_all" on public.app_state;
create policy "app_state_insert_all"
on public.app_state
for insert
with check (true);

drop policy if exists "app_state_update_all" on public.app_state;
create policy "app_state_update_all"
on public.app_state
for update
using (true)
with check (true);

alter publication supabase_realtime add table public.app_state;

insert into public.app_state (id, payload, updated_by)
values (
  'labeltraxx-shared-state',
  jsonb_build_object(
    'jobs', '[]'::jsonb,
    'assignments', '[]'::jsonb,
    'requests', '[]'::jsonb,
    'pullPaperRequests', '[]'::jsonb,
    'shipmentGroups', '[]'::jsonb,
    'users', jsonb_build_array(
      jsonb_build_object(
        'id', 'user-admin',
        'username', 'Admin',
        'password', '1234',
        'role', 'Management',
        'isAdmin', true,
        'createdAt', timezone('utc', now()),
        'createdBy', 'system'
      )
    ),
    'weekStart', timezone('utc', now())
  ),
  'system'
)
on conflict (id) do nothing;
