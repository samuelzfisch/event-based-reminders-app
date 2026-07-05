"use client";

import { createContext, type ReactNode, useContext, useEffect, useState } from "react";
import { usePathname } from "next/navigation";

import { AppShellNav } from "./app-shell-nav";
import { useAuthContext } from "./auth-provider";

const AppShellSidebarContentContext = createContext<((content: ReactNode | null) => void) | null>(null);

export function AppShellSidebarContent({ children }: { children: ReactNode }) {
  const setSidebarContent = useContext(AppShellSidebarContentContext);

  useEffect(() => {
    if (!setSidebarContent) return;
    setSidebarContent(children);
    return () => {
      setSidebarContent(null);
    };
  }, [children, setSidebarContent]);

  return null;
}

export function AppShell({ children }: { children: ReactNode }) {
  const { authEnabled, currentUser, signOut } = useAuthContext();
  const [sidebarContent, setSidebarContent] = useState<ReactNode | null>(null);
  const pathname = usePathname();
  const isPlansPage = pathname === "/plans" || pathname.startsWith("/plans/");

  return (
    <AppShellSidebarContentContext.Provider value={setSidebarContent}>
      <div className="min-h-screen bg-[var(--app-bg)]">
        <div
          className={`mx-auto flex w-full px-2 pb-2.5 sm:px-2.5 ${
            isPlansPage ? "max-w-[1040px] gap-10 pt-10 lg:gap-14" : "max-w-[960px] gap-4 pt-5 lg:gap-5"
          }`}
        >
          <aside className="hidden w-[150px] shrink-0 md:block lg:w-[160px]">
            <div>
              <div className="rounded-[18px] border border-[var(--app-border-soft)] bg-white/95 p-4 shadow-[var(--app-shadow)] backdrop-blur">
                <div className="space-y-5">
                  <div className="text-[1rem] font-semibold tracking-[-0.02em] text-slate-950">Event-Based Reminders</div>
                  <AppShellNav />
                  {authEnabled && currentUser ? (
                    <div className="border-t border-slate-200 pt-4">
                      <div className="text-[11px] font-semibold uppercase tracking-[0.14em] text-slate-500">Signed in</div>
                      <div className="mt-2 break-all text-[11px] font-medium leading-[1.5] text-slate-900">{currentUser.email || "Authenticated user"}</div>
                      <button
                        type="button"
                        onClick={() => void signOut()}
                        className="mt-2.5 inline-flex rounded-lg border border-slate-300 bg-white px-2 py-1 text-[11px] font-medium text-slate-700 transition-colors duration-200 hover:border-slate-400 hover:bg-slate-50"
                      >
                        Sign Out
                      </button>
                    </div>
                  ) : null}
                  {sidebarContent ? <div className="border-t border-slate-200 pt-4">{sidebarContent}</div> : null}
                </div>
              </div>
            </div>
          </aside>

          <main className="min-w-0 flex-1">{children}</main>
        </div>
      </div>
    </AppShellSidebarContentContext.Provider>
  );
}
