import type { ReactNode } from "react";

import { AppShellNav } from "./app-shell-nav";

export function AppShell({ children }: { children: ReactNode }) {
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
              </div>
            </div>
          </div>
        </aside>

        <main className="min-w-0 flex-1">{children}</main>
      </div>
    </div>
  );
}
