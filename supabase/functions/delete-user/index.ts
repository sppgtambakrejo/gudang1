// supabase/functions/delete-user/index.ts
// Dipanggil dari menu "Manajemen User" oleh Super Admin untuk menghapus
// akun sepenuhnya: baris profiles DAN login Supabase Auth-nya.

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

    const callerClient = createClient(SUPABASE_URL, SERVICE_ROLE_KEY, {
      global: { headers: { Authorization: authHeader } },
    });
    const { data: callerAuth, error: callerErr } = await callerClient.auth.getUser();
    if (callerErr || !callerAuth?.user) return json({ error: "Unauthorized" }, 401);

    const admin = createClient(SUPABASE_URL, SERVICE_ROLE_KEY);

    const { data: callerProfile } = await admin
      .from("profiles")
      .select("role")
      .eq("id", callerAuth.user.id)
      .maybeSingle();

    if (!callerProfile || callerProfile.role !== "superadmin") {
      return json({ error: "Hanya Super Admin yang boleh menghapus akun." }, 403);
    }

    const body = await req.json().catch(() => ({}));
    const { id } = body;
    if (!id) return json({ error: "ID akun tidak ada." }, 400);
    if (id === callerAuth.user.id) return json({ error: "Tidak bisa menghapus akun sendiri." }, 400);

    await admin.from("profiles").delete().eq("id", id);
    const { error: delErr } = await admin.auth.admin.deleteUser(id);
    if (delErr) return json({ error: delErr.message }, 400);

    return json({ ok: true });
  } catch (e) {
    return json({ error: String(e?.message || e) }, 500);
  }
});
