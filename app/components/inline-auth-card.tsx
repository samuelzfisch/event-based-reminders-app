"use client";

import Link from "next/link";
import { useState, type FormEvent } from "react";

import { useAuthContext } from "./auth-provider";

type InlineAuthCardProps = {
  title?: string;
  description?: string;
  showContinueLink?: boolean;
  compact?: boolean;
};

export function InlineAuthCard({
  title = "Sign in",
  description = "Use your email to access Event-Based Reminders.",
  showContinueLink = true,
  compact = false,
}: InlineAuthCardProps) {
  const { authEnabled, authBypassEnabled, loading, currentUser, sendMagicLink } = useAuthContext();
  const [email, setEmail] = useState("");
  const [message, setMessage] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [sending, setSending] = useState(false);

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
    <div className="rounded-2xl border bg-white p-6 shadow-sm">
      <div>
        <h1 className={`${compact ? "text-xl" : "text-2xl"} font-semibold text-gray-900`}>{title}</h1>
        <p className="mt-2 text-sm text-gray-600">{description}</p>
      </div>

      {!authEnabled ? (
        <div className="mt-6 rounded-xl border border-amber-200 bg-amber-50 px-4 py-3 text-sm text-amber-800">
          Supabase auth is not configured for this environment yet. The app will continue using legacy local mode.
        </div>
      ) : authBypassEnabled ? (
        <div className="mt-6 space-y-4 rounded-xl border border-blue-200 bg-blue-50 px-4 py-3 text-sm text-blue-800">
          <div>Auth bypass is enabled for this environment.</div>
          {showContinueLink ? (
            <div>
              <Link href="/" className="font-medium underline underline-offset-2">
                Continue to the app
              </Link>
            </div>
          ) : null}
        </div>
      ) : currentUser ? (
        <div className="mt-6 rounded-xl border border-blue-200 bg-blue-50 px-4 py-3 text-sm text-blue-800">
          Signed in. Opening your workspace…
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

          {message ? (
            <div className="rounded-xl border border-green-200 bg-green-50 px-4 py-3 text-sm text-green-800">
              {message}
            </div>
          ) : null}
          {error ? (
            <div className="rounded-xl border border-red-200 bg-red-50 px-4 py-3 text-sm text-red-800">
              {error}
            </div>
          ) : null}
        </form>
      )}
    </div>
  );
}
