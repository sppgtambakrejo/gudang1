# Gudang Persediaan SPPG — GitHub Pages + Supabase

Aplikasi React + Vite dengan database Supabase dan autentikasi email pemilik.

## 1. Buat project Supabase

1. Buka Supabase dan buat project baru.
2. Masuk ke **SQL Editor**.
3. Salin dan jalankan seluruh isi `supabase-schema.sql`.
4. Buka **Authentication > Users > Add user**.
5. Buat akun email dan password pemilik.
6. Buka **Project Settings > API Keys** lalu salin Project URL dan Publishable key. Project lama dapat menggunakan anon key.

## 2. Menjalankan di komputer

1. Salin `.env.example` menjadi `.env`.
2. Isi URL dan key Supabase.
3. Jalankan:

```bash
npm install
npm run dev
```

## 3. Deploy ke GitHub Pages

1. Upload seluruh isi folder project ke repository GitHub.
2. Buka **Settings > Secrets and variables > Actions**.
3. Tambahkan repository secrets:
   - `VITE_SUPABASE_URL`
   - `VITE_SUPABASE_PUBLISHABLE_KEY`
4. Buka **Settings > Pages**, pilih **GitHub Actions** sebagai source.
5. Push ke branch `main`. Workflow akan build dan deploy otomatis.

## Login

Urutannya:

1. Login pemilik memakai email/password Supabase.
2. Login internal aplikasi:
   - Super Admin: `superadmin` / `super123`
   - Admin: `admin1` / `admin123`
   - Viewer PIN: `1234`

Segera ubah kredensial internal setelah login pertama.

## Penyimpanan data

Semua master barang, periode, transaksi masuk/keluar, saldo, supplier, user internal, dan log aktivitas disimpan dalam tabel `app_storage`. Row Level Security memastikan setiap akun pemilik hanya bisa membaca dan mengubah datanya sendiri.
