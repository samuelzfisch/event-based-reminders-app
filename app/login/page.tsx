"use client";

import Link from "next/link";
import { useEffect, useState, type FormEvent } from "react";

import { useAuthContext } from "../components/auth-provider";
import { getSupabaseBrowserClient } from "../../lib/supabaseClient";

export default function LoginPage() {
  const { authEnabled, authBypassEnabled, loading, currentUser, refreshAuthContext, sendMagicLink } = useAuthContext();
  const [email, setEmail] = useState("");
  const [message, setMessage] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [sending, setSending] = useState(false);

  useEffect(() => {
    if (!authEnabled || authBypassEnabled || currentUser) return;

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
        const cleanedUrl = `${currentUrl.pathname}${currentUrl.search}`;
        window.history.replaceState({}, document.title, cleanedUrl);
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
          }
          return;
        }

        if (code) {
          const { error: exchangeError } = await supabase.auth.exchangeCodeForSession(code);
          if (exchangeError) throw exchangeError;
          currentUrl.searchParams.delete("code");
          currentUrl.searchParams.delete("type");
          currentUrl.searchParams.delete("next");
          cleanupUrl();
          if (!cancelled) {
            await refreshAuthContext();
          }
          return;
        }

        if (tokenHash && type) {
          const { error: verifyError } = await supabase.auth.verifyOtp({
            token_hash: tokenHash,
            type: type as "email" | "magiclink" | "signup" | "recovery" | "invite" | "email_change",
          });
          if (verifyError) throw verifyError;
          currentUrl.searchParams.delete("token_hash");
          currentUrl.searchParams.delete("type");
          currentUrl.searchParams.delete("next");
          cleanupUrl();
          if (!cancelled) {
            await refreshAuthContext();
          }
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
  }, [authBypassEnabled, authEnabled, currentUser, refreshAuthContext]);

  useEffect(() => {
    console.info("[loginPage] mount", {
      authEnabled,
      loading,
      hasCurrentUser: Boolean(currentUser),
      userId: currentUser?.id ?? null,
    });
  }, [authEnabled, currentUser, loading]);

  useEffect(() => {
    console.info("[loginPage] auth state", {
      loading,
      hasCurrentUser: Boolean(currentUser),
      userId: currentUser?.id ?? null,
    });
    if (currentUser) {
      console.info("[loginPage] redirecting because authenticated", { userId: currentUser.id });
    }
  }, [currentUser, loading]);

  async function handleSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    const nextEmail = email.trim();
    if (!nextEmail) {
      setError("Please enter your email address.");
      return;
    }

    setSending(true);
    setError(null);
    setMessage(null);

    try {
      await sendMagicLink(nextEmail);
      setMessage("Check your email for a sign-in link.");
    } catch (submitError) {
      setError(submitError instanceof Error ? submitError.message : "Could not send sign-in link.");
    } finally {
      setSending(false);
    }
  }

  return (
    <div className="flex min-h-screen items-center justify-center bg-gray-50 px-4 py-8">
      <div className="w-full max-w-md rounded-2xl border bg-white p-6 shadow-sm">
        <div>
          <h1 className="text-2xl font-semibold text-gray-900">Sign in</h1>
          <p className="mt-2 text-sm text-gray-600">
            Use your email to access Event-Based Reminders.
          </p>
        </div>

        {!authEnabled ? (
          <div className="mt-6 rounded-xl border border-amber-200 bg-amber-50 px-4 py-3 text-sm text-amber-800">
            Supabase auth is not configured for this environment yet. The app will continue using legacy local mode.
          </div>
        ) : authBypassEnabled ? (
          <div className="mt-6 space-y-4 rounded-xl border border-blue-200 bg-blue-50 px-4 py-3 text-sm text-blue-800">
            <div>Auth bypass is enabled for this environment.</div>
            <div>
              <Link href="/" className="font-medium underline underline-offset-2">
                Continue to the app
              </Link>
            </div>
          </div>
        ) : currentUser ? (
          <div className="mt-6 rounded-xl border border-blue-200 bg-blue-50 px-4 py-3 text-sm text-blue-800">
            Signed in. Redirecting to your workspace…
          </div>
        ) : loading ? (
          <div className="mt-6 rounded-xl border border-gray-200 bg-gray-50 px-4 py-3 text-sm text-gray-700">
            Checking your sign-in…
          </div>
        ) : (
          <form onSubmit={handleSubmit} className="mt-6 space-y-4">
            <div>
              <label htmlFor="email" className="block text-xs font-semibold uppercase tracking-wide text-gray-500">
                Email
              </label>
              <input
                id="email"
                type="email"
                autoComplete="email"
                value={email}
                onChange={(event) => setEmail(event.target.value)}
                className="mt-2 w-full rounded-lg border bg-white px-3 py-2 text-sm text-gray-900"
                placeholder="you@company.com"
                disabled={loading || sending}
              />
            </div>

            <button
              type="submit"
              disabled={loading || sending}
              className="inline-flex w-full items-center justify-center rounded-lg bg-blue-600 px-4 py-2 text-sm font-medium text-white hover:bg-blue-700 disabled:cursor-not-allowed disabled:bg-blue-300"
            >
              {sending ? "Sending link..." : "Send magic link"}
            </button>

            {message ? <p className="text-sm text-green-700">{message}</p> : null}
            {error ? <p className="text-sm text-red-700">{error}</p> : null}
          </form>
        )}
      </div>
    </div>
  );
}
