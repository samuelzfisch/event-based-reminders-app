"use client";

import { useEffect, useState } from "react";
import { useRouter } from "next/navigation";

import { useAuthContext } from "../../components/auth-provider";
import { getSupabaseBrowserClient } from "../../../lib/supabaseClient";

type EmailAuthCallbackType = "email" | "magiclink" | "signup" | "recovery" | "invite" | "email_change";

export default function AuthCallbackPage() {
  const router = useRouter();
  const { authEnabled, authBypassEnabled, currentUser, refreshAuthContext } = useAuthContext();
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    if (!authEnabled || authBypassEnabled) return;

    if (currentUser) {
      router.replace("/");
      return;
    }

    let cancelled = false;

    async function processAuthReturn() {
      if (typeof window === "undefined") return;

      const supabase = getSupabaseBrowserClient();
      if (!supabase) return;

      const currentUrl = new URL(window.location.href);
      const hashParams = new URLSearchParams(currentUrl.hash.replace(/^#/, ""));
      const accessToken = hashParams.get("access_token");
      const refreshToken = hashParams.get("refresh_token");
      const code = currentUrl.searchParams.get("code");
      const tokenHash = currentUrl.searchParams.get("token_hash");
      const type = currentUrl.searchParams.get("type");

      const cleanupUrl = () => {
        window.history.replaceState({}, document.title, currentUrl.pathname);
      };

      try {
        if (accessToken && refreshToken) {
          const { error: setSessionError } = await supabase.auth.setSession({
            access_token: accessToken,
            refresh_token: refreshToken,
          });
          if (setSessionError) throw setSessionError;
          cleanupUrl();
          if (!cancelled) {
            await refreshAuthContext();
            router.replace("/");
          }
          return;
        }

        if (code) {
          const { error: exchangeError } = await supabase.auth.exchangeCodeForSession(code);
          if (exchangeError) throw exchangeError;
          cleanupUrl();
          if (!cancelled) {
            await refreshAuthContext();
            router.replace("/");
          }
          return;
        }

        if (tokenHash && type) {
          const { error: verifyError } = await supabase.auth.verifyOtp({
            token_hash: tokenHash,
            type: type as EmailAuthCallbackType,
          });
          if (verifyError) throw verifyError;
          cleanupUrl();
          if (!cancelled) {
            await refreshAuthContext();
            router.replace("/");
          }
          return;
        }

        if (!cancelled) {
          setError("This sign-in link is missing required auth details.");
        }
      } catch (processingError) {
        if (!cancelled) {
          setError(processingError instanceof Error ? processingError.message : "Could not complete sign-in.");
        }
      }
    }

    void processAuthReturn();

    return () => {
      cancelled = true;
    };
  }, [authBypassEnabled, authEnabled, currentUser, refreshAuthContext, router]);

  return (
    <div className="flex min-h-screen items-center justify-center bg-gray-50 px-4 py-8">
      <div className="w-full max-w-md rounded-2xl border bg-white p-6 shadow-sm">
        <h1 className="text-2xl font-semibold text-gray-900">Completing sign-in</h1>
        {error ? (
          <p className="mt-3 text-sm text-red-700">{error}</p>
        ) : (
          <p className="mt-3 text-sm text-gray-600">Finalizing your sign-in and opening the app…</p>
        )}
      </div>
    </div>
  );
}
