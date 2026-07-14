# Gudang Persediaan — SPPG

Web app inventory bahan pangan (React + Vite), database Supabase, deploy otomatis ke GitHub Pages.

## 1. Setup Supabase

1. Buka project Supabase kamu → **SQL Editor**.
2. Jalankan seluruh isi `supabase/schema.sql`. Ini membuat tabel relasional (`master_barang`, `supplier`, `periode`, `saldo_awal`, `transaksi_masuk`, `transaksi_keluar`, `activity_log`, `profiles`, `app_settings`), RLS, dan fungsi RPC (`get_viewer_data`, `set_viewer_pin`, `tutup_periode`, dst). Data barang & saldo awal contoh **tidak** diisi lewat SQL — aplikasi akan mengisinya otomatis sekali saja saat login pertama kali jika tabel masih kosong (lihat bagian Arsitektur).
3. Buat akun Super Admin pertama:
   - **Authentication → Users → Add user**
     Email: `superadmin@sppg.local`, Password: bebas (simpan baik-baik).
   - Salin **UUID** user itu, lalu jalankan di SQL Editor (ganti `<UUID>`):
     ```sql
     insert into profiles (id, username, nama, role)
     values ('<UUID>', 'superadmin', 'Super Admin', 'superadmin');
     ```
4. Ambil **Project URL** dan **anon public key** dari **Settings → API** — dipakai di langkah 3.
5. Deploy dua Edge Function yang menangani tambah/hapus akun (perlu [Supabase CLI](https://supabase.com/docs/guides/cli)):
   ```bash
   supabase login
   supabase link --project-ref <project-ref-kamu>
   supabase functions deploy create-user
   supabase functions deploy delete-user
   ```
   Tidak perlu set secret tambahan — `SUPABASE_URL` dan `SUPABASE_SERVICE_ROLE_KEY` sudah otomatis tersedia di environment Edge Function.

Setelah Super Admin pertama login, akun Admin/Super Admin berikutnya bisa langsung ditambah dari menu **Manajemen User** di aplikasi — cukup nama, username, password, dan role. Tidak perlu email, tidak perlu buka Supabase Dashboard lagi.

## 2. Setup repo GitHub

1. Push folder ini ke repo GitHub kamu (branch `main`).
2. **Settings → Pages** → Source: **GitHub Actions**.
3. **Settings → Secrets and variables → Actions → New repository secret**, tambahkan:
   - `VITE_SUPABASE_URL`
   - `VITE_SUPABASE_ANON_KEY`
4. Push ke `main` → workflow `.github/workflows/deploy.yml` otomatis build & deploy. URL app: `https://<username>.github.io/<nama-repo>/`.

## 3. Jalankan lokal (opsional)

```bash
npm install
cp .env.example .env   # isi dengan URL & anon key Supabase kamu
npm run dev
```

## Arsitektur singkat

- **Login Admin/Super Admin**: Supabase Auth (email sintetis `username@sppg.local`, tidak pernah ditampilkan), role dibaca dari tabel `profiles`.
- **Tambah/hapus akun Admin**: lewat Edge Function `create-user` / `delete-user` (jalan dengan `service_role` key di server Supabase, bukan di browser) — supaya Super Admin bisa membuat login lengkap (username + password) tanpa email dan tanpa sesi Super Admin ter-switch ke akun baru.
- **Login Viewer**: PIN saja (tanpa akun Auth), divalidasi lewat RPC `get_viewer_data` yang menyusun data read-only (stok, riwayat, dsb.) langsung dari tabel-tabel relasional.
- **Data operasional disimpan sebagai baris tabel relasional sungguhan**, bukan blob JSON:
  - `master_barang` — daftar barang (kode, nama, satuan)
  - `supplier` — daftar supplier
  - `periode` — daftar periode pencatatan (hanya satu yang `is_active`)
  - `saldo_awal` — saldo & harga awal per barang, per periode
  - `transaksi_masuk` / `transaksi_keluar` — **satu baris per transaksi** (bukan array JSON), dengan `jumlah` dihitung otomatis oleh database (`generated column`)
  - `activity_log` — satu baris per aksi
  
  Setiap tambah/ubah/hapus dari UI langsung melakukan `insert`/`update`/`delete` ke baris yang relevan (lihat objek `persist` di `src/App.jsx`), bukan menimpa seluruh tabel — ini menghilangkan risiko race condition/data-loss yang pernah terjadi di versi `app_kv` sebelumnya.
  - Saat login pertama kali dan tabel `master_barang`/`periode` masih kosong, aplikasi otomatis mengisi data contoh (dari `ITEMS`/`SALDO_AWAL` di `src/App.jsx`) — sekali saja, dan hanya jika query berhasil mengembalikan hasil kosong (bukan karena error jaringan/RLS).
  - **Tutup periode** dijalankan lewat RPC atomik `tutup_periode` di database: menutup periode aktif, membuka periode baru, dan menyalin saldo akhir → saldo awal baru dalam satu transaksi SQL.
- **RLS** membatasi baca/tulis tabel operasional & `profiles` hanya untuk role `admin`/`superadmin`. `master_barang`, `periode`, dan `activity_log` dibatasi lebih ketat lagi di level UI (menu hanya muncul untuk Super Admin), meski RLS di DB tetap mengizinkan admin+superadmin agar konsisten. PIN viewer tidak pernah tersimpan di tabel yang bisa dibaca klien — hanya lewat RPC `SECURITY DEFINER`.
- Menghapus barang di **Master Barang** yang sudah punya riwayat transaksi masuk/keluar akan ditolak oleh database (foreign key `restrict`) — ini disengaja, supaya riwayat transaksi lama tidak pernah kehilangan rujukan barangnya.

## Batasan yang perlu diketahui

- Super Admin **pertama** tetap harus dibuat manual lewat Supabase Dashboard + SQL insert (langkah 3 di atas) — ini satu-satunya proses manual, karena butuh titik awal sebelum ada Super Admin yang bisa memanggil Edge Function.
- Export Excel (`xlsx`) tetap jalan seperti sebelumnya karena murni proses di browser.
- Skema ini (v2) menggantikan skema `app_kv` lama sepenuhnya. Karena project ini belum pernah di-deploy/dipakai produksi, **tidak ada migrasi data lama** yang perlu dilakukan — cukup jalankan `schema.sql` yang baru di project Supabase yang masih kosong/baru.
