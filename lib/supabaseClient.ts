import { createClient, type SupabaseClient } from "@supabase/supabase-js";

let browserClient: SupabaseClient | null | undefined;

export function isSupabaseConfigured() {
  return Boolean(
    String(process.env.NEXT_PUBLIC_SUPABASE_URL ?? "").trim() &&
      String(process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY ?? "").trim()
  );
}

export function getSupabaseBrowserClient() {
  if (browserClient !== undefined) return browserClient;

  const url = String(process.env.NEXT_PUBLIC_SUPABASE_URL ?? "").trim();
  const anonKey = String(process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY ?? "").trim();

  if (!url || !anonKey) {
    browserClient = null;
    return browserClient;
  }

  browserClient = createClient(url, anonKey, {
    auth: {
      persistSession: false,
      autoRefreshToken: false,
      detectSessionInUrl: false,
    },
  });
  return browserClient;
}

