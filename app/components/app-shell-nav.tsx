"use client";

import Link from "next/link";
import { usePathname } from "next/navigation";

const NAV_ITEMS = [
  { href: "/", label: "Home" },
  { href: "/plans", label: "Plans" },
  { href: "/history", label: "History" },
  { href: "/settings", label: "Settings" },
] as const;

const PLANS_SIDEBAR_NEUTRAL_STORAGE_KEY = "event-reminders:plans-sidebar-neutral";
const PLANS_SIDEBAR_NEUTRAL_EVENT = "event-reminders:plans-sidebar-neutral";

function isActivePath(pathname: string, href: string) {
  if (href === "/") return pathname === "/";
  return pathname === href || pathname.startsWith(`${href}/`);
}

export function AppShellNav() {
  const pathname = usePathname();

  function handleNavClick(href: string) {
    if (href !== "/plans") return;
    if (typeof window === "undefined") return;
    window.sessionStorage.setItem(PLANS_SIDEBAR_NEUTRAL_STORAGE_KEY, "1");
    window.dispatchEvent(new Event(PLANS_SIDEBAR_NEUTRAL_EVENT));
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
