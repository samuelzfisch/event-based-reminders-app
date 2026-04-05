"use client";

import { useState, type FormEvent } from "react";

import { useAuthContext } from "../components/auth-provider";

export default function LoginPage() {
  const { authEnabled, loading, sendMagicLink } = useAuthContext();
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
