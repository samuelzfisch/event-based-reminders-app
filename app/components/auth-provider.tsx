"use client";

import {
  createContext,
  useContext,
  useEffect,
  useEffectEvent,
  useRef,
  useState,
  type ReactNode,
} from "react";
import type { Session, User } from "@supabase/supabase-js";

import {
  clearCachedOrgContext,
  getCachedOrgContext,
  type BootstrappedOrgContext,
  bootstrapCurrentOrgForUser,
} from "../../lib/orgBootstrap";
import { getSupabaseBrowserClient, isSupabaseConfigured } from "../../lib/supabaseClient";

type AuthContextValue = {
  authEnabled: boolean;
  loading: boolean;
  currentSession: Session | null;
  currentUser: User | null;
  currentOrgId: string | null;
  currentOrgRole: string | null;
  refreshAuthContext: () => Promise<void>;
  sendMagicLink: (email: string) => Promise<void>;
  signOut: () => Promise<void>;
};

const AuthContext = createContext<AuthContextValue | null>(null);
const AUTH_RESOLUTION_TIMEOUT_MS = 2000;
const ORG_CONTEXT_RETRY_DELAY_MS = 750;
export function AuthProvider({ children }: { children: ReactNode }) {
  const [authEnabled] = useState(() => isSupabaseConfigured());
  const [loading, setLoading] = useState(true);
  const [currentSession, setCurrentSession] = useState<Session | null>(null);
  const [currentUser, setCurrentUser] = useState<User | null>(null);
  const [orgContext, setOrgContext] = useState<BootstrappedOrgContext | null>(() => getCachedOrgContext());
  const mountedRef = useRef(true);
  const currentSessionRef = useRef<Session | null>(null);
  const currentUserIdRef = useRef<string | null>(null);
  const authResolutionRef = useRef<Promise<void> | null>(null);
  const orgBootstrapStartedRef = useRef<string | null>(null);
  const orgBootstrapRetriedRef = useRef<string | null>(null);

  useEffect(() => {
    currentSessionRef.current = currentSession;
  }, [currentSession]);

  useEffect(() => {
    currentUserIdRef.current = currentUser?.id ?? null;
  }, [currentUser]);

  async function resolveOrgContextInBackground(user: User, source: string) {
    const cachedContext = getCachedOrgContext();
    if (mountedRef.current) {
      if (cachedContext?.userId === user.id) {
        console.info("[auth] currentOrgId set from cache", { source, orgId: cachedContext.orgId });
        setOrgContext(cachedContext);
      } else {
        setOrgContext(null);
        console.info("[auth] currentOrgId is null", { source, reason: "no_matching_cached_org_context" });
      }
    }

    if (orgBootstrapStartedRef.current === user.id) {
      console.info("[auth] org bootstrap already in progress", { source, userId: user.id });
      return;
    }

    orgBootstrapStartedRef.current = user.id;
    try {
      const nextOrgContext = await bootstrapCurrentOrgForUser({
        userId: user.id,
        email: user.email,
      });
      if (!mountedRef.current) return;
      setOrgContext(nextOrgContext);
      if (nextOrgContext?.orgId) {
        console.info("[auth] currentOrgId set", { source, orgId: nextOrgContext.orgId });
        orgBootstrapRetriedRef.current = null;
      } else {
        console.info("[auth] currentOrgId is null", { source, reason: "org_bootstrap_returned_null" });
        if (orgBootstrapRetriedRef.current !== user.id) {
          orgBootstrapRetriedRef.current = user.id;
          window.setTimeout(() => {
            if (!mountedRef.current || currentUserIdRef.current !== user.id) return;
            console.info("[auth] retrying org bootstrap", { source, userId: user.id });
            void resolveOrgContextInBackground(user, `${source}:retry`);
          }, ORG_CONTEXT_RETRY_DELAY_MS);
        }
      }
    } catch (error) {
      if (!mountedRef.current) return;
      setOrgContext(null);
      console.error("[auth] background org bootstrap failed", { source, error });
      console.info("[auth] currentOrgId is null", { source, reason: "org_bootstrap_failed" });
    } finally {
      if (orgBootstrapStartedRef.current === user.id) {
        orgBootstrapStartedRef.current = null;
      }
    }
  }

  async function applyResolvedSession(session: Session | null, source: string) {
    if (!mountedRef.current) return;

    console.info("[auth] applying session", {
      source,
      hasSession: Boolean(session),
      userId: session?.user?.id ?? null,
    });

    setCurrentSession(session ?? null);
    setCurrentUser(session?.user ?? null);

    if (!session?.user) {
      console.info("[auth] signed-out state applied", { source });
      clearCachedOrgContext();
      setOrgContext(null);
      orgBootstrapStartedRef.current = null;
      setLoading(false);
      console.info("[auth] loading cleared", { source });
      return;
    }

    setLoading(false);
    console.info("[auth] loading cleared", { source });
    void resolveOrgContextInBackground(session.user, source);
  }

  async function resolveWithSingleFlight(source: string, resolver: () => Promise<Session | null>) {
    const existingResolution = authResolutionRef.current;
    if (existingResolution) {
      console.info("[auth] auth resolution reused", { source });
      await existingResolution;
      return;
    }

    const resolutionPromise = (async () => {
      console.info("[auth] auth resolution start", { source });

      try {
        const session = await resolver();
        await applyResolvedSession(session, source);
        console.info("[auth] auth resolution success", {
          source,
          hasSession: Boolean(session),
          userId: session?.user?.id ?? null,
        });
      } catch (error) {
        console.error("[auth] auth resolution failed", { source, error });
        if (mountedRef.current) {
          setCurrentSession(null);
          setCurrentUser(null);
          setOrgContext(null);
          setLoading(false);
          console.info("[auth] loading cleared", { source });
        }
      }
    });

    const guardedResolutionPromise: Promise<void> = Promise.race([
      (async () => {
        await resolutionPromise;
      })(),
      new Promise<void>((resolve) => {
        setTimeout(() => {
          console.warn("[auth] auth resolution timed out");
          if (mountedRef.current) {
            setLoading(false);
            console.log("[auth] loading cleared");
          }
          resolve();
        }, AUTH_RESOLUTION_TIMEOUT_MS);
      }),
    ]);

    authResolutionRef.current = guardedResolutionPromise;
    await guardedResolutionPromise;
    authResolutionRef.current = null;
  }

  async function refreshAuthContext() {
    if (!authEnabled) {
      if (mountedRef.current) {
        setCurrentSession(null);
        setCurrentUser(null);
        setOrgContext(null);
        setLoading(false);
      }
      return;
    }

    await resolveWithSingleFlight("mount_getSession", async () => {
      const supabase = getSupabaseBrowserClient();
      if (!supabase) return null;

      const {
        data: { session },
        error,
      } = await supabase.auth.getSession();

      if (error) {
        console.error("[auth] session restore failed on mount", error);
        return null;
      }

      if (session) {
        console.info("[auth] session restored on mount", { userId: session.user.id });
      } else {
        console.info("[auth] no session found on mount");
      }

      return session ?? null;
    });
  }

  const refreshAuthContextEffect = useEffectEvent(async () => {
    await refreshAuthContext();
  });

  const resolveWithSingleFlightEffect = useEffectEvent(async (source: string, resolver: () => Promise<Session | null>) => {
    await resolveWithSingleFlight(source, resolver);
  });

  useEffect(() => {
    mountedRef.current = true;

    if (!authEnabled) {
      setLoading(false);
      return () => {
        mountedRef.current = false;
      };
    }

    const supabase = getSupabaseBrowserClient();
    if (!supabase) {
      setLoading(false);
      console.info("[auth] loading cleared", { source: "mount_no_client" });
      return () => {
        mountedRef.current = false;
      };
    }

    const {
      data: { subscription },
    } = supabase.auth.onAuthStateChange(async (event, session) => {
      if (!mountedRef.current) return;
      console.info("[auth] auth event received", { event });

      if (event === "INITIAL_SESSION") {
        if (session && !currentSessionRef.current) {
          await resolveWithSingleFlightEffect("INITIAL_SESSION", async () => session);
        }
        return;
      }

      if (event === "SIGNED_IN") {
        await resolveWithSingleFlightEffect("SIGNED_IN", async () => session ?? null);
        return;
      }

      if (event === "TOKEN_REFRESHED") {
        if (session) {
          console.info("[auth] session refresh success", { userId: session.user.id });
        } else {
          console.warn("[auth] session refresh returned no session");
        }
        return;
      }

      if (event === "SIGNED_OUT") {
        await resolveWithSingleFlightEffect("SIGNED_OUT", async () => null);
        return;
      }
    });

    console.info("[auth] mount auth check start");
    void refreshAuthContextEffect();

    return () => {
      mountedRef.current = false;
      subscription.unsubscribe();
    };
  }, [authEnabled]);

  async function sendMagicLink(email: string) {
    const supabase = getSupabaseBrowserClient();
    if (!supabase) {
      throw new Error("Supabase auth is not configured.");
    }

    const redirectTo =
      typeof window !== "undefined" ? `${window.location.origin}/login` : undefined;

    const { error } = await supabase.auth.signInWithOtp({
      email,
      options: {
        emailRedirectTo: redirectTo,
      },
    });

    if (error) {
      throw error;
    }
  }

  async function signOut() {
    const supabase = getSupabaseBrowserClient();
    clearCachedOrgContext();
    if (!supabase) {
      setCurrentSession(null);
      setCurrentUser(null);
      setOrgContext(null);
      return;
    }

    await supabase.auth.signOut();
    if (!mountedRef.current) return;
    setCurrentSession(null);
    setCurrentUser(null);
    setOrgContext(null);
  }

  return (
    <AuthContext.Provider
      value={{
        authEnabled,
        loading,
        currentSession,
        currentUser,
        currentOrgId: orgContext?.orgId ?? null,
        currentOrgRole: orgContext?.role ?? null,
        refreshAuthContext,
        sendMagicLink,
        signOut,
      }}
    >
      {children}
    </AuthContext.Provider>
  );
}

export function useAuthContext() {
  const context = useContext(AuthContext);
  if (!context) {
    throw new Error("useAuthContext must be used inside AuthProvider.");
  }
  return context;
}
