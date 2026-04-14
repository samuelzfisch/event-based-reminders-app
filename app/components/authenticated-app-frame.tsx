"use client";

import { useEffect, useSyncExternalStore, type ReactNode } from "react";
import { usePathname, useRouter } from "next/navigation";

import { useAuthContext } from "./auth-provider";
import { AppShell } from "./app-shell";
import { InlineAuthCard } from "./inline-auth-card";

function LoadingScreen({ label }: { label: string }) {
  return (
    <div className="flex min-h-screen items-center justify-center bg-gray-50 px-4">
      <div className="rounded-2xl border bg-white px-6 py-5 text-sm text-gray-600 shadow-sm">{label}</div>
    </div>
  );
}

export function AuthenticatedAppFrame({ children }: { children: ReactNode }) {
  const pathname = usePathname();
  const router = useRouter();
  const { authEnabled, authBypassEnabled, loading, currentUser } = useAuthContext();
  const hasHydrated = useSyncExternalStore(
    () => () => {},
    () => true,
    () => false
  );

  const isLoginRoute = pathname === "/login";
  const isAuthCallbackRoute = pathname === "/auth/callback";

  useEffect(() => {
    if (!hasHydrated) return;
    if (authBypassEnabled) return;
    if (!authEnabled) return;

    if (currentUser && isLoginRoute) {
      router.replace("/");
    }
  }, [authBypassEnabled, authEnabled, currentUser, hasHydrated, isLoginRoute, router]);

  if (!hasHydrated && !authBypassEnabled && authEnabled) {
    return <LoadingScreen label="Loading workspace…" />;
  }

  if (!authBypassEnabled && authEnabled && loading) {
    return <LoadingScreen label="Loading workspace…" />;
  }

  if (!authBypassEnabled && authEnabled && !loading && currentUser && isLoginRoute) {
    return <LoadingScreen label="Redirecting to your workspace…" />;
  }

  if (!authBypassEnabled && authEnabled && !loading && !currentUser && isAuthCallbackRoute) {
    return <>{children}</>;
  }

  if (isLoginRoute) {
    return <>{children}</>;
  }

  if (!authBypassEnabled && authEnabled && !loading && !currentUser) {
    return (
      <div className="min-h-screen bg-gray-50">
        <div className="mx-auto flex min-h-screen w-full max-w-7xl items-center px-4 py-8">
          <div className="grid w-full gap-6 lg:grid-cols-[minmax(0,1.2fr)_minmax(320px,420px)] lg:items-center">
            <div className="space-y-6">
              <div className="space-y-3">
                <div className="text-xs font-semibold uppercase tracking-wide text-gray-500">Event-Based Reminders</div>
                <h1 className="text-3xl font-bold text-gray-900 sm:text-4xl">Build reminder plans without a separate login detour.</h1>
                <p className="max-w-2xl text-sm text-gray-600 sm:text-base">
                  Sign in with your email to open plans, settings, and history in the same product surface.
                </p>
              </div>

              <div className="grid gap-4 sm:grid-cols-3">
                <div className="rounded-2xl border bg-white p-4 shadow-sm">
                  <div className="text-xs font-semibold uppercase tracking-wide text-gray-500">Plans</div>
                  <div className="mt-2 text-sm text-gray-600">Build repeatable event timelines, reminders, and drafts.</div>
                </div>
                <div className="rounded-2xl border bg-white p-4 shadow-sm">
                  <div className="text-xs font-semibold uppercase tracking-wide text-gray-500">History</div>
                  <div className="mt-2 text-sm text-gray-600">Review what ran and make safe follow-up adjustments.</div>
                </div>
                <div className="rounded-2xl border bg-white p-4 shadow-sm">
                  <div className="text-xs font-semibold uppercase tracking-wide text-gray-500">Settings</div>
                  <div className="mt-2 text-sm text-gray-600">Manage account connections and workspace defaults.</div>
                </div>
              </div>
            </div>

            <InlineAuthCard
              title="Sign in to continue"
              description="Enter your email and we’ll send a magic link. You’ll come back through the hosted callback and land right in the app."
              showContinueLink={false}
            />
          </div>
        </div>
      </div>
    );
  }

  return <AppShell>{children}</AppShell>;
}
