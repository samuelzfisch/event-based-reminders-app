"use client";

import type { ReactNode } from "react";

import { AppShellNav } from "./app-shell-nav";
import { useAuthContext } from "./auth-provider";

export function AppShell({ children }: { children: ReactNode }) {
  const { authEnabled, currentUser, signOut } = useAuthContext();

  return (
    <div className="min-h-screen bg-gray-50">
      <div className="mx-auto flex w-full max-w-7xl gap-6 px-4 py-8">
        <aside className="hidden w-64 shrink-0 lg:block">
          <div className="sticky top-8 space-y-4">
            <div className="rounded-2xl border bg-white shadow-sm">
              <div className="space-y-4 p-4">
                <div>
                  <div className="text-xs font-semibold uppercase tracking-wide text-gray-500">App</div>
                  <div className="mt-1 text-lg font-semibold text-gray-900">Event-Based Reminders</div>
                  <p className="mt-2 text-sm text-gray-600">
                    Familiar multi-page shell for plans, settings, and future reminders tooling.
                  </p>
                </div>
                <AppShellNav />
                {authEnabled && currentUser ? (
                  <div className="rounded-xl border bg-gray-50 p-3">
                    <div className="text-xs font-semibold uppercase tracking-wide text-gray-500">Signed In</div>
                    <div className="mt-1 break-all text-sm font-medium text-gray-900">{currentUser.email || "Authenticated user"}</div>
                    <button
                      type="button"
                      onClick={() => void signOut()}
                      className="mt-3 inline-flex rounded-lg border px-3 py-2 text-sm font-medium text-gray-700 hover:bg-white"
                    >
                      Sign Out
                    </button>
                  </div>
                ) : null}
              </div>
            </div>
          </div>
        </aside>

        <main className="min-w-0 flex-1">{children}</main>
      </div>
    </div>
  );
}
