"use client";

import Link from "next/link";
import { usePathname } from "next/navigation";

const NAV_ITEMS = [
  { href: "/", label: "Dashboard" },
  { href: "/plans", label: "Plans" },
  { href: "/history", label: "History" },
  { href: "/settings", label: "Settings" },
] as const;

const PLANS_SIDEBAR_NEUTRAL_STORAGE_KEY = "event-reminders:plans-sidebar-neutral";
const PLANS_SIDEBAR_NEUTRAL_EVENT = "event-reminders:plans-sidebar-neutral";

type AppShellNavVariant = "sidebar" | "mobile";

function isActivePath(pathname: string, href: string) {
  if (href === "/") return pathname === "/";
  return pathname === href || pathname.startsWith(`${href}/`);
}

export function AppShellNav({ variant = "sidebar" }: { variant?: AppShellNavVariant }) {
  const pathname = usePathname();

  function handleNavClick(href: string) {
    if (href !== "/plans") return;
    if (typeof window === "undefined") return;
    window.sessionStorage.setItem(PLANS_SIDEBAR_NEUTRAL_STORAGE_KEY, "1");
    window.dispatchEvent(new Event(PLANS_SIDEBAR_NEUTRAL_EVENT));
  }

  if (variant === "mobile") {
    return (
      <nav className="grid grid-cols-4 gap-[4px] rounded-xl border border-[var(--app-border-soft)] bg-white/95 p-[4px] shadow-[0_12px_28px_-24px_rgba(38,72,104,0.3)]" aria-label="Primary navigation">
        {NAV_ITEMS.map((item) => {
          const isActive = isActivePath(pathname, item.href);
          return (
            <Link
              key={item.href}
              href={item.href}
              onClick={() => handleNavClick(item.href)}
              className={`relative flex min-h-[36px] min-w-0 items-center justify-center rounded-lg px-[8px] text-center text-[12px] font-medium leading-none transition-colors duration-200 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--app-focus)] focus-visible:ring-offset-1 focus-visible:ring-offset-[var(--app-bg)] ${
                isActive
                  ? "bg-[var(--app-primary-soft)] text-[#315f92] shadow-[inset_0_0_0_1px_rgba(79,127,184,0.16)]"
                  : "text-slate-700 hover:bg-[var(--app-surface-muted)] hover:text-slate-950"
              }`}
              aria-current={isActive ? "page" : undefined}
            >
              <span className="min-w-0 truncate whitespace-nowrap">{item.label}</span>
            </Link>
          );
        })}
      </nav>
    );
  }

  return (
    <nav className="space-y-1">
      {NAV_ITEMS.map((item) => {
        const isActive = isActivePath(pathname, item.href);
        return (
          <Link
            key={item.href}
            href={item.href}
            onClick={() => handleNavClick(item.href)}
            className={`group relative flex min-h-[26px] items-center rounded-lg px-2.5 py-1.5 text-[11px] font-medium transition-colors duration-200 ${
              isActive
                ? "bg-[var(--app-primary-soft)] text-[#315f92]"
                : "text-slate-700 hover:bg-[var(--app-surface-muted)] hover:text-slate-950"
            }`}
            aria-current={isActive ? "page" : undefined}
          >
            {isActive ? <span className="absolute left-0 top-1/2 h-4 -translate-y-1/2 rounded-full border-l-2 border-[var(--app-primary)]" /> : null}
            <span>{item.label}</span>
          </Link>
        );
      })}
    </nav>
  );
}
