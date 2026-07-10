-- Jalankan seluruh SQL ini di Supabase Dashboard > SQL Editor.
-- Versi satu-login: pengguna langsung login memakai akun gudang.

create table if not exists public.app_storage_shared (
  workspace_id text not null,
  key text not null,
  value jsonb not null default '{}'::jsonb,
  updated_at timestamptz not null default now(),
  primary key (workspace_id, key)
);

alter table public.app_storage_shared enable row level security;

drop policy if exists "Gudang shared read" on public.app_storage_shared;
drop policy if exists "Gudang shared insert" on public.app_storage_shared;
drop policy if exists "Gudang shared update" on public.app_storage_shared;
drop policy if exists "Gudang shared delete" on public.app_storage_shared;

create policy "Gudang shared read"
on public.app_storage_shared
for select
to anon, authenticated
using (true);

create policy "Gudang shared insert"
on public.app_storage_shared
for insert
to anon, authenticated
with check (true);

create policy "Gudang shared update"
on public.app_storage_shared
for update
to anon, authenticated
using (true)
with check (true);

create policy "Gudang shared delete"
on public.app_storage_shared
for delete
to anon, authenticated
using (true);

create index if not exists app_storage_shared_updated_at_idx
on public.app_storage_shared (updated_at desc);
