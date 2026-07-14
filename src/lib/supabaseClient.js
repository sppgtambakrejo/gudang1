import { createClient } from "@supabase/supabase-js";

const url = import.meta.env.VITE_SUPABASE_URL;
const anonKey = import.meta.env.VITE_SUPABASE_ANON_KEY;

if (!url || !anonKey) {
  // Fails loudly in dev/build if secrets weren't wired up correctly.
  console.error(
    "VITE_SUPABASE_URL / VITE_SUPABASE_ANON_KEY belum diatur. Cek file .env (lokal) atau GitHub Secrets (deploy)."
  );
}

export const supabase = createClient(url, anonKey);
