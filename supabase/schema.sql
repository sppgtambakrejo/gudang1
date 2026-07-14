-- ============================================================
-- SPPG Gudang Persediaan — Supabase schema (v2, RELASIONAL)
-- Jalankan seluruh file ini di Supabase Dashboard > SQL Editor
--
-- Perubahan dari v1: semua data (barang, supplier, periode,
-- saldo awal, transaksi masuk/keluar, log aktivitas) sekarang
-- disimpan sebagai baris-baris tabel relasional sungguhan,
-- bukan satu blob JSON per "key" di tabel app_kv. Tabel app_kv
-- lama sudah TIDAK dipakai lagi oleh frontend v2.
-- ============================================================

create extension if not exists pgcrypto;

-- ------------------------------------------------------------
-- 1. PROFILES — menghubungkan Supabase Auth user ke role aplikasi
-- ------------------------------------------------------------
create table if not exists profiles (
  id uuid primary key references auth.users(id) on delete cascade,
  username text unique not null,
  nama text not null,
  role text not null check (role in ('superadmin', 'admin')),
  created_at timestamptz default now()
);

-- ------------------------------------------------------------
-- 2. MASTER_BARANG — daftar barang persediaan
-- ------------------------------------------------------------
create table if not exists master_barang (
  kode text primary key,
  nama text not null,
  satuan text not null,
  created_at timestamptz default now()
);

-- ------------------------------------------------------------
-- 3. SUPPLIER
-- ------------------------------------------------------------
create table if not exists supplier (
  id uuid primary key default gen_random_uuid(),
  nama text not null unique,
  kontak text,
  created_at timestamptz default now()
);

-- ------------------------------------------------------------
-- 4. PERIODE — periode pencatatan stock
-- ------------------------------------------------------------
create table if not exists periode (
  id uuid primary key default gen_random_uuid(),
  nama text not null,
  mulai date not null,
  selesai date not null,
  closed boolean not null default false,
  is_active boolean not null default false,
  created_at timestamptz default now()
);
-- pastikan hanya ada satu periode aktif setiap saat
create unique index if not exists idx_periode_satu_aktif on periode ((is_active)) where is_active;

-- ------------------------------------------------------------
-- 5. SALDO_AWAL — saldo awal & harga per barang, per periode
-- ------------------------------------------------------------
create table if not exists saldo_awal (
  periode_id uuid not null references periode(id) on delete cascade,
  kode text not null references master_barang(kode) on delete cascade,
  saldo numeric not null default 0,
  harga numeric not null default 0,
  primary key (periode_id, kode)
);

-- ------------------------------------------------------------
-- 6. TRANSAKSI_MASUK — penerimaan barang (satu baris = satu transaksi)
-- ------------------------------------------------------------
create table if not exists transaksi_masuk (
  id uuid primary key default gen_random_uuid(),
  periode_id uuid not null references periode(id),
  tanggal date not null,
  kode text not null references master_barang(kode),
  supplier text,
  vol numeric not null check (vol > 0),
  harga numeric not null default 0,
  jumlah numeric generated always as (vol * harga) stored,
  oleh text,
  oleh_id uuid references profiles(id),
  waktu timestamptz not null default now()
);
create index if not exists idx_transaksi_masuk_periode on transaksi_masuk (periode_id);
create index if not exists idx_transaksi_masuk_kode on transaksi_masuk (kode);

-- ------------------------------------------------------------
-- 7. TRANSAKSI_KELUAR — pengeluaran barang (satu baris = satu transaksi)
-- ------------------------------------------------------------
create table if not exists transaksi_keluar (
  id uuid primary key default gen_random_uuid(),
  periode_id uuid not null references periode(id),
  tanggal date not null,
  kode text not null references master_barang(kode),
  vol numeric not null check (vol > 0),
  oleh text,
  oleh_id uuid references profiles(id),
  waktu timestamptz not null default now()
);
create index if not exists idx_transaksi_keluar_periode on transaksi_keluar (periode_id);
create index if not exists idx_transaksi_keluar_kode on transaksi_keluar (kode);

-- ------------------------------------------------------------
-- 8. ACTIVITY_LOG — log aktivitas
-- ------------------------------------------------------------
create table if not exists activity_log (
  id uuid primary key default gen_random_uuid(),
  waktu timestamptz not null default now(),
  oleh text,
  role text,
  aksi text,
  detail text
);
create index if not exists idx_activity_log_waktu on activity_log (waktu desc);

-- ------------------------------------------------------------
-- 9. APP_SETTINGS — pengaturan sensitif (PIN viewer)
-- ------------------------------------------------------------
create table if not exists app_settings (
  key text primary key,
  value text not null
);
insert into app_settings (key, value) values ('viewer_pin', '1234')
on conflict (key) do nothing;

-- ------------------------------------------------------------
-- 10. Helper: role user yang sedang login
-- ------------------------------------------------------------
create or replace function current_profile_role()
returns text
language sql stable
security definer
set search_path = public
as $$
  select role from profiles where id = auth.uid();
$$;

-- ------------------------------------------------------------
-- 11. RLS — semua tabel operasional: staff (admin/superadmin) bisa
--     baca/tulis penuh. Master data & periode hanya dibatasi lebih
--     ketat lagi di level UI (superadmin), tapi RLS tetap mengizinkan
--     admin+superadmin di DB supaya konsisten dengan skema lama.
-- ------------------------------------------------------------
alter table profiles enable row level security;
alter table master_barang enable row level security;
alter table supplier enable row level security;
alter table periode enable row level security;
alter table saldo_awal enable row level security;
alter table transaksi_masuk enable row level security;
alter table transaksi_keluar enable row level security;
alter table activity_log enable row level security;
alter table app_settings enable row level security;

drop policy if exists "profiles_select" on profiles;
create policy "profiles_select" on profiles for select
  using (auth.uid() = id or current_profile_role() = 'superadmin');

drop policy if exists "profiles_superadmin_write" on profiles;
create policy "profiles_superadmin_write" on profiles for all
  using (current_profile_role() = 'superadmin')
  with check (current_profile_role() = 'superadmin');

drop policy if exists "staff_all_master_barang" on master_barang;
create policy "staff_all_master_barang" on master_barang for all
  using (current_profile_role() in ('admin', 'superadmin'))
  with check (current_profile_role() in ('admin', 'superadmin'));

drop policy if exists "staff_all_supplier" on supplier;
create policy "staff_all_supplier" on supplier for all
  using (current_profile_role() in ('admin', 'superadmin'))
  with check (current_profile_role() in ('admin', 'superadmin'));

drop policy if exists "staff_all_periode" on periode;
create policy "staff_all_periode" on periode for all
  using (current_profile_role() in ('admin', 'superadmin'))
  with check (current_profile_role() in ('admin', 'superadmin'));

drop policy if exists "staff_all_saldo_awal" on saldo_awal;
create policy "staff_all_saldo_awal" on saldo_awal for all
  using (current_profile_role() in ('admin', 'superadmin'))
  with check (current_profile_role() in ('admin', 'superadmin'));

drop policy if exists "staff_all_transaksi_masuk" on transaksi_masuk;
create policy "staff_all_transaksi_masuk" on transaksi_masuk for all
  using (current_profile_role() in ('admin', 'superadmin'))
  with check (current_profile_role() in ('admin', 'superadmin'));

drop policy if exists "staff_all_transaksi_keluar" on transaksi_keluar;
create policy "staff_all_transaksi_keluar" on transaksi_keluar for all
  using (current_profile_role() in ('admin', 'superadmin'))
  with check (current_profile_role() in ('admin', 'superadmin'));

drop policy if exists "staff_all_activity_log" on activity_log;
create policy "staff_all_activity_log" on activity_log for all
  using (current_profile_role() in ('admin', 'superadmin'))
  with check (current_profile_role() in ('admin', 'superadmin'));

drop policy if exists "app_settings_superadmin" on app_settings;
create policy "app_settings_superadmin" on app_settings for all
  using (current_profile_role() = 'superadmin')
  with check (current_profile_role() = 'superadmin');

-- ------------------------------------------------------------
-- 12. Pembuatan & penghapusan akun Admin/Super Admin ditangani oleh
--     Edge Functions `create-user` & `delete-user` (lihat folder
--     supabase/functions) — tidak berubah dari skema lama.
-- ------------------------------------------------------------

-- ------------------------------------------------------------
-- 13. RPC — akses Viewer lewat PIN (tanpa akun Auth). SECURITY DEFINER
--     supaya bisa melewati RLS hanya lewat jalur PIN yang benar.
--     Mengembalikan bentuk JSON yang sama seperti skema lama supaya
--     frontend tidak perlu berubah untuk mode Viewer.
-- ------------------------------------------------------------
create or replace function verify_viewer_pin(p_pin text)
returns boolean
language plpgsql
security definer
set search_path = public
as $$
declare
  stored text;
begin
  select value into stored from app_settings where key = 'viewer_pin';
  return stored is not null and stored = p_pin;
end;
$$;
grant execute on function verify_viewer_pin(text) to anon, authenticated;

create or replace function get_viewer_data(p_pin text)
returns jsonb
language plpgsql
security definer
set search_path = public
as $$
declare
  result jsonb;
  j_master jsonb;
  j_supplier jsonb;
  j_periode jsonb;
  j_saldo jsonb;
  j_masuk jsonb;
  j_keluar jsonb;
begin
  if not verify_viewer_pin(p_pin) then
    raise exception 'PIN salah.';
  end if;

  select coalesce(jsonb_agg(jsonb_build_object('kode', kode, 'nama', nama, 'satuan', satuan) order by kode), '[]'::jsonb)
    into j_master from master_barang;

  select coalesce(jsonb_agg(jsonb_build_object('id', id, 'nama', nama, 'kontak', kontak)), '[]'::jsonb)
    into j_supplier from supplier;

  select jsonb_build_object(
    'list', coalesce(jsonb_agg(jsonb_build_object(
      'id', id, 'nama', nama, 'mulai', mulai, 'selesai', selesai, 'closed', closed
    ) order by mulai), '[]'::jsonb),
    'activeId', (select id from periode where is_active limit 1)
  ) into j_periode from periode;

  select coalesce(jsonb_object_agg(periode_id, per_periode), '{}'::jsonb) into j_saldo
  from (
    select periode_id, jsonb_object_agg(kode, jsonb_build_object('saldo', saldo, 'harga', harga)) as per_periode
    from saldo_awal
    group by periode_id
  ) s;

  select coalesce(jsonb_agg(jsonb_build_object(
    'id', tm.id, 'tanggal', tm.tanggal, 'supplier', tm.supplier, 'kode', tm.kode,
    'nama', mb.nama, 'satuan', mb.satuan, 'vol', tm.vol, 'harga', tm.harga, 'jumlah', tm.jumlah,
    'oleh', tm.oleh, 'waktu', tm.waktu, 'periodeId', tm.periode_id
  ) order by tm.waktu desc), '[]'::jsonb)
  into j_masuk from transaksi_masuk tm join master_barang mb on mb.kode = tm.kode;

  select coalesce(jsonb_agg(jsonb_build_object(
    'id', tk.id, 'tanggal', tk.tanggal, 'kode', tk.kode,
    'nama', mb.nama, 'satuan', mb.satuan, 'vol', tk.vol,
    'oleh', tk.oleh, 'waktu', tk.waktu, 'periodeId', tk.periode_id
  ) order by tk.waktu desc), '[]'::jsonb)
  into j_keluar from transaksi_keluar tk join master_barang mb on mb.kode = tk.kode;

  result := jsonb_build_object(
    'master', j_master, 'supplier', j_supplier, 'periode', j_periode,
    'saldo', j_saldo, 'masuk', j_masuk, 'keluar', j_keluar
  );
  return result;
end;
$$;
grant execute on function get_viewer_data(text) to anon, authenticated;

create or replace function set_viewer_pin(p_new_pin text)
returns void
language plpgsql
security definer
set search_path = public
as $$
begin
  if current_profile_role() <> 'superadmin' then
    raise exception 'Hanya Super Admin yang boleh mengubah PIN Viewer.';
  end if;
  update app_settings set value = p_new_pin where key = 'viewer_pin';
end;
$$;
grant execute on function set_viewer_pin(text) to authenticated;

-- ------------------------------------------------------------
-- 14. RPC — tutup periode aktif & buka periode baru dalam satu
--     transaksi atomik (hitung saldo akhir → jadi saldo awal baru).
-- ------------------------------------------------------------
create or replace function tutup_periode(
  p_nama_baru text, p_mulai_baru date, p_selesai_baru date
)
returns uuid
language plpgsql
security definer
set search_path = public
as $$
declare
  v_active_id uuid;
  v_new_id uuid;
begin
  if current_profile_role() not in ('admin', 'superadmin') then
    raise exception 'Tidak diizinkan.';
  end if;

  select id into v_active_id from periode where is_active limit 1;
  if v_active_id is null then
    raise exception 'Tidak ada periode aktif.';
  end if;

  update periode set closed = true, is_active = false where id = v_active_id;

  insert into periode (nama, mulai, selesai, closed, is_active)
  values (p_nama_baru, p_mulai_baru, p_selesai_baru, false, true)
  returning id into v_new_id;

  -- saldo akhir periode lama = saldo_awal + total masuk - total keluar, jadi saldo awal periode baru
  insert into saldo_awal (periode_id, kode, saldo, harga)
  select
    v_new_id,
    mb.kode,
    coalesce(sa.saldo, 0)
      + coalesce((select sum(vol) from transaksi_masuk where periode_id = v_active_id and kode = mb.kode), 0)
      - coalesce((select sum(vol) from transaksi_keluar where periode_id = v_active_id and kode = mb.kode), 0),
    coalesce((select harga from transaksi_masuk where periode_id = v_active_id and kode = mb.kode order by waktu desc limit 1), sa.harga, 0)
  from master_barang mb
  left join saldo_awal sa on sa.periode_id = v_active_id and sa.kode = mb.kode;

  return v_new_id;
end;
$$;
grant execute on function tutup_periode(text, date, date) to authenticated;

-- ------------------------------------------------------------
-- 15. Setelah menjalankan schema ini:
--     a. Buat Super Admin pertama lewat Dashboard > Authentication > Users,
--        lalu masukkan ke tabel profiles (lihat README).
--     b. Buat periode pertama secara manual, misalnya:
--
--        insert into periode (nama, mulai, selesai, is_active)
--        values ('Periode 1 - Februari 2026', '2026-02-16', '2026-02-28', true);
--
--     c. Isi master_barang & saldo_awal awal (bisa lewat menu "Master
--        Barang" di aplikasi setelah login, satu per satu, atau lewat
--        SQL Editor untuk isi massal).
--     d. Deploy Edge Functions `create-user` & `delete-user` (lihat README).
-- ------------------------------------------------------------
