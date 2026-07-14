// supabase/functions/create-user/index.ts
// Dipanggil dari menu "Manajemen User" oleh Super Admin.
// Membuat akun Supabase Auth (email sintetis username@sppg.local, tidak pernah
// ditampilkan ke user) + baris profiles, tanpa mengganggu sesi Super Admin
// yang sedang login (karena pakai Admin API, bukan auth.signUp di client).

import { createClient } from "https://esm.sh/@supabase/supabase-js@2.45.4";
import { corsHeaders, json } from "../_shared/cors.ts";

const SUPABASE_URL = Deno.env.get("SUPABASE_URL");
const SERVICE_ROLE_KEY = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY");

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response(null, { headers: corsHeaders });
  if (req.method !== "POST") return json({ error: "Method not allowed" }, 405);

  try {
    const authHeader = req.headers.get("Authorization");
    if (!authHeader) return json({ error: "Unauthorized" }, 401);

    // Klien yang "berbicara sebagai" pemanggil, untuk memverifikasi identitasnya.
    const callerClient = createClient(SUPABASE_URL, SERVICE_ROLE_KEY, {
      global: { headers: { Authorization: authHeader } },
    });
    const { data: callerAuth, error: callerErr } = await callerClient.auth.getUser();
    if (callerErr || !callerAuth?.user) return json({ error: "Unauthorized" }, 401);

    // Klien service_role untuk operasi admin (bypass RLS).
    const admin = createClient(SUPABASE_URL, SERVICE_ROLE_KEY);

    const { data: callerProfile } = await admin
      .from("profiles")
      .select("role")
      .eq("id", callerAuth.user.id)
      .maybeSingle();

    if (!callerProfile || callerProfile.role !== "superadmin") {
      return json({ error: "Hanya Super Admin yang boleh menambah akun." }, 403);
    }

    const body = await req.json().catch(() => ({}));
    const { username, password, nama, role } = body;

    if (!username || !password || !nama || !role) {
      return json({ error: "Lengkapi username, password, nama, dan role." }, 400);
    }
    if (!["admin", "superadmin"].includes(role)) {
      return json({ error: "Role tidak valid." }, 400);
    }
    if (String(password).length < 6) {
      return json({ error: "Password minimal 6 karakter." }, 400);
    }

    const cleanUsername = String(username).trim().toLowerCase();
    if (!/^[a-z0-9._-]{3,32}$/.test(cleanUsername)) {
      return json({ error: "Username 3-32 karakter: huruf kecil, angka, titik, underscore, atau strip." }, 400);
    }

    const { data: existing } = await admin.from("profiles").select("id").eq("username", cleanUsername).maybeSingle();
    if (existing) return json({ error: "Username sudah dipakai." }, 400);

    const email = `${cleanUsername}@sppg.local`;

    const { data: created, error: createErr } = await admin.auth.admin.createUser({
      email,
      password,
      email_confirm: true,
    });
    if (createErr || !created?.user) {
      return json({ error: createErr?.message || "Gagal membuat akun." }, 400);
    }

    const { error: profileErr } = await admin.from("profiles").insert({
      id: created.user.id,
      username: cleanUsername,
      nama,
      role,
    });
    if (profileErr) {
      // Rollback supaya tidak ada akun Auth "yatim" tanpa profil.
      await admin.auth.admin.deleteUser(created.user.id);
      return json({ error: profileErr.message }, 400);
    }

    return json({ ok: true, id: created.user.id, username: cleanUsername });
  } catch (e) {
    return json({ error: String(e?.message || e) }, 500);
  }
});
