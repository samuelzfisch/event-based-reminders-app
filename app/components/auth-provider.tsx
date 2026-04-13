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
import { isAuthBypassEnabled } from "../../lib/authBypass";
import { getSupabaseBrowserClient, isSupabaseConfigured } from "../../lib/supabaseClient";

type AuthContextValue = {
  authEnabled: boolean;
  authBypassEnabled: boolean;
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
  const [authBypassEnabled] = useState(() => isAuthBypassEnabled());
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
    const matchingCachedContext = cachedContext?.userId === user.id ? cachedContext : null;
    if (mountedRef.current) {
      if (matchingCachedContext) {
        setOrgContext(matchingCachedContext);
      } else {
        setOrgContext(null);
      }
    }

    if (orgBootstrapStartedRef.current === user.id) {
      return;
    }

    orgBootstrapStartedRef.current = user.id;
    try {
      const nextOrgContext = await bootstrapCurrentOrgForUser({
        userId: user.id,
        email: user.email,
      });
      if (!mountedRef.current) return;
      if (nextOrgContext?.orgId) {
        setOrgContext(nextOrgContext);
        orgBootstrapRetriedRef.current = null;
      } else {
        if (matchingCachedContext) {
          setOrgContext(matchingCachedContext);
        } else {
          setOrgContext(null);
        }
        if (orgBootstrapRetriedRef.current !== user.id) {
          orgBootstrapRetriedRef.current = user.id;
          window.setTimeout(() => {
            if (!mountedRef.current || currentUserIdRef.current !== user.id) return;
            void resolveOrgContextInBackground(user, `${source}:retry`);
          }, ORG_CONTEXT_RETRY_DELAY_MS);
        }
      }
    } catch (error) {
      if (!mountedRef.current) return;
      if (matchingCachedContext) {
        setOrgContext(matchingCachedContext);
      } else {
        setOrgContext(null);
      }
      console.error("[auth] background org bootstrap failed", { source, error });
    } finally {
      if (orgBootstrapStartedRef.current === user.id) {
        orgBootstrapStartedRef.current = null;
      }
    }
  }

  async function applyResolvedSession(session: Session | null, source: string) {
    if (!mountedRef.current) return;

    setCurrentSession(session ?? null);
    setCurrentUser(session?.user ?? null);

    if (!session?.user) {
      clearCachedOrgContext();
      setOrgContext(null);
      orgBootstrapStartedRef.current = null;
      setLoading(false);
      return;
    }

    setLoading(false);
    void resolveOrgContextInBackground(session.user, source);
  }

  async function resolveWithSingleFlight(source: string, resolver: () => Promise<Session | null>) {
    const existingResolution = authResolutionRef.current;
    if (existingResolution) {
      await existingResolution;
      return;
    }

    const resolutionPromise = (async () => {
      try {
        const session = await resolver();
        if (!session && currentSessionRef.current?.user) {
          if (mountedRef.current) {
            setLoading(false);
          }
          return;
        }
        await applyResolvedSession(session, source);
      } catch (error) {
        console.error("[auth] auth resolution failed", { source, error });
        if (mountedRef.current) {
          setCurrentSession(null);
          setCurrentUser(null);
          setOrgContext(null);
          setLoading(false);
        }
      }
    });

    let timeoutHandle: ReturnType<typeof setTimeout> | null = null;

    const guardedResolutionPromise: Promise<void> = Promise.race([
      (async () => {
        await resolutionPromise;
      })(),
      new Promise<void>((resolve) => {
        timeoutHandle = setTimeout(() => {
          console.warn("[auth] auth resolution timed out");
          if (mountedRef.current) {
            setLoading(false);
          }
          resolve();
        }, AUTH_RESOLUTION_TIMEOUT_MS);
      }),
    ]);

    authResolutionRef.current = guardedResolutionPromise;
    await guardedResolutionPromise;
    if (timeoutHandle) {
      clearTimeout(timeoutHandle);
    }
    authResolutionRef.current = null;
  }

  async function refreshAuthContext() {
    if (authBypassEnabled) {
      if (mountedRef.current) {
        setCurrentSession(null);
        setCurrentUser(null);
        setOrgContext(null);
        setLoading(false);
      }
      return;
    }

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

      return session ?? null;
    });
  }

  const refreshAuthContextEffect = useEffectEvent(async () => {
    await refreshAuthContext();
  });

  const applyResolvedSessionEffect = useEffectEvent(async (session: Session | null, source: string) => {
    await applyResolvedSession(session, source);
  });

  const resolveWithSingleFlightEffect = useEffectEvent(async (source: string, resolver: () => Promise<Session | null>) => {
    await resolveWithSingleFlight(source, resolver);
  });

  useEffect(() => {
    mountedRef.current = true;

    if (authBypassEnabled) {
      setLoading(false);
      return () => {
        mountedRef.current = false;
      };
    }

    if (!authEnabled) {
      setLoading(false);
      return () => {
        mountedRef.current = false;
      };
    }

    const supabase = getSupabaseBrowserClient();
    if (!supabase) {
      setLoading(false);
      return () => {
        mountedRef.current = false;
      };
    }

    const {
      data: { subscription },
    } = supabase.auth.onAuthStateChange(async (event, session) => {
      if (!mountedRef.current) return;

      if (event === "INITIAL_SESSION") {
        await applyResolvedSessionEffect(session ?? null, "INITIAL_SESSION");
        return;
      }

      if (event === "SIGNED_IN") {
        await applyResolvedSessionEffect(session ?? null, "SIGNED_IN");
        return;
      }

      if (event === "TOKEN_REFRESHED") {
        if (!session) {
          console.warn("[auth] session refresh returned no session");
        }
        await applyResolvedSessionEffect(session ?? null, "TOKEN_REFRESHED");
        return;
      }

      if (event === "SIGNED_OUT") {
        await resolveWithSingleFlightEffect("SIGNED_OUT", async () => null);
        return;
      }
    });

    void refreshAuthContextEffect();

    return () => {
      mountedRef.current = false;
      subscription.unsubscribe();
    };
  }, [authBypassEnabled, authEnabled]);

  async function sendMagicLink(email: string) {
    const supabase = getSupabaseBrowserClient();
    if (!supabase) {
      throw new Error("Supabase auth is not configured.");
    }

    const redirectTo =
      typeof window !== "undefined" ? `${window.location.origin}/auth/callback` : undefined;

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
        authBypassEnabled,
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
