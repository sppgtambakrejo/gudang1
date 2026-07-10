-- Jalankan seluruh SQL ini di Supabase Dashboard > SQL Editor.

create table if not exists public.app_storage (
  user_id uuid not null references auth.users(id) on delete cascade,
  key text not null,
  value jsonb not null default '{}'::jsonb,
  updated_at timestamptz not null default now(),
  primary key (user_id, key)
);

alter table public.app_storage enable row level security;

create policy "Owner can read own app data"
on public.app_storage
for select
to authenticated
using ((select auth.uid()) = user_id);

create policy "Owner can insert own app data"
on public.app_storage
for insert
to authenticated
with check ((select auth.uid()) = user_id);

create policy "Owner can update own app data"
on public.app_storage
for update
to authenticated
using ((select auth.uid()) = user_id)
with check ((select auth.uid()) = user_id);

create policy "Owner can delete own app data"
on public.app_storage
for delete
to authenticated
using ((select auth.uid()) = user_id);

create index if not exists app_storage_updated_at_idx
on public.app_storage (updated_at desc);
