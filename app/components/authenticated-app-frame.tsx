"use client";

import { useEffect, type ReactNode } from "react";
import { usePathname, useRouter } from "next/navigation";

import { useAuthContext } from "./auth-provider";
import { AppShell } from "./app-shell";

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
  const { authEnabled, loading, currentUser } = useAuthContext();

  const isLoginRoute = pathname === "/login";

  useEffect(() => {
    if (!authEnabled || loading) return;

    if (!currentUser && !isLoginRoute) {
      router.replace("/login");
      return;
    }

    if (currentUser && isLoginRoute) {
      router.replace("/");
    }
  }, [authEnabled, currentUser, isLoginRoute, loading, router]);

  if (authEnabled && loading) {
    return <LoadingScreen label="Loading workspace…" />;
  }

  if (authEnabled && !currentUser && !isLoginRoute) {
    return <LoadingScreen label="Redirecting to login…" />;
  }

  if (authEnabled && currentUser && isLoginRoute) {
    return <LoadingScreen label="Redirecting to your workspace…" />;
  }

  if (isLoginRoute) {
    return <>{children}</>;
  }

  return <AppShell>{children}</AppShell>;
}
