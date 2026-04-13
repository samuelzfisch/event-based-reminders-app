"use client";

import { useEffect, useSyncExternalStore, type ReactNode } from "react";
import { usePathname, useRouter } from "next/navigation";

import { useAuthContext } from "./auth-provider";
import { getCachedOrgContext } from "../../lib/orgBootstrap";
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
  const { authEnabled, authBypassEnabled, loading, currentUser, currentOrgId } = useAuthContext();
  const hasHydrated = useSyncExternalStore(
    () => () => {},
    () => true,
    () => false
  );
  const cachedOrgContext = getCachedOrgContext();
  const hasCachedWorkspaceContext = Boolean(cachedOrgContext?.orgId);

  const isLoginRoute = pathname === "/login";

  useEffect(() => {
    console.info("[appFrame] render state", {
      pathname,
      authEnabled,
      authBypassEnabled,
      loading,
      hasCurrentUser: Boolean(currentUser),
      currentOrgId,
      hasHydrated,
    });
  }, [authBypassEnabled, authEnabled, currentOrgId, currentUser, hasHydrated, loading, pathname]);

  useEffect(() => {
    if (!hasHydrated) return;
    console.info("[appFrame] client mounted", {
      pathname,
    });
  }, [hasHydrated, pathname]);

  useEffect(() => {
    if (!hasHydrated) return;
    if (authBypassEnabled) return;
    if (!authEnabled) return;

    if (!currentUser && !isLoginRoute) {
      if (loading) return;
      router.replace("/login");
      return;
    }

    if (currentUser && isLoginRoute) {
      router.replace("/");
    }
  }, [authBypassEnabled, authEnabled, currentUser, hasHydrated, isLoginRoute, loading, router]);

  useEffect(() => {
    console.info("[appFrame] auth gate state", {
      pathname,
      loading,
      hasCurrentUser: Boolean(currentUser),
      currentOrgId,
      cachedOrgId: cachedOrgContext?.orgId ?? null,
      hasHydrated,
    });
  }, [cachedOrgContext?.orgId, currentOrgId, currentUser, hasHydrated, loading, pathname]);

  useEffect(() => {
    if (authEnabled && currentUser && !currentOrgId) {
      console.info("[appFrame] rendering with authenticated user and no currentOrgId yet", {
        userId: currentUser.id,
        pathname,
      });
    }
  }, [authEnabled, currentOrgId, currentUser, pathname]);

  useEffect(() => {
    if (!hasHydrated) return;
    const branch = !authBypassEnabled && authEnabled && loading && !currentUser && !hasCachedWorkspaceContext
      ? "loading_workspace"
      : !authBypassEnabled && authEnabled && !currentUser && !isLoginRoute
        ? "redirecting_to_login"
        : !authBypassEnabled && authEnabled && currentUser && isLoginRoute
          ? "redirecting_to_workspace"
          : isLoginRoute
            ? "login"
            : "workspace";
    console.info("[appFrame] final branch chosen", {
      pathname,
      branch,
      loading,
      hasCurrentUser: Boolean(currentUser),
      currentOrgId,
      cachedOrgId: cachedOrgContext?.orgId ?? null,
    });
  }, [
    authBypassEnabled,
    authEnabled,
    cachedOrgContext?.orgId,
    currentOrgId,
    currentUser,
    hasCachedWorkspaceContext,
    hasHydrated,
    isLoginRoute,
    loading,
    pathname,
  ]);

  if (!hasHydrated && !authBypassEnabled && authEnabled) {
    return <LoadingScreen label="Loading workspace…" />;
  }

  if (!authBypassEnabled && authEnabled && loading) {
    return <LoadingScreen label="Loading workspace…" />;
  }

  if (!authBypassEnabled && authEnabled && !loading && !currentUser && !isLoginRoute) {
    return <LoadingScreen label="Redirecting to login…" />;
  }

  if (!authBypassEnabled && authEnabled && !loading && currentUser && isLoginRoute) {
    return <LoadingScreen label="Redirecting to your workspace…" />;
  }

  if (isLoginRoute) {
    return <>{children}</>;
  }

  return <AppShell>{children}</AppShell>;
}
