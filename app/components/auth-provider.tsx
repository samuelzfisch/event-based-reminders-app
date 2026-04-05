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

import { clearCachedOrgContext, type BootstrappedOrgContext, bootstrapCurrentOrgForUser } from "../../lib/orgBootstrap";
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
const ORG_BOOTSTRAP_TIMEOUT_MS = 2000;

function hasAuthParamsInUrl() {
  if (typeof window === "undefined") return false;

  const search = new URLSearchParams(window.location.search);
  const hash = new URLSearchParams(window.location.hash.replace(/^#/, ""));

  return (
    search.has("code") ||
    search.has("access_token") ||
    search.has("refresh_token") ||
    hash.has("access_token") ||
    hash.has("refresh_token")
  );
}

export function AuthProvider({ children }: { children: ReactNode }) {
  const [authEnabled] = useState(() => isSupabaseConfigured());
  const [loading, setLoading] = useState(true);
  const [currentSession, setCurrentSession] = useState<Session | null>(null);
  const [currentUser, setCurrentUser] = useState<User | null>(null);
  const [orgContext, setOrgContext] = useState<BootstrappedOrgContext | null>(null);
  const mountedRef = useRef(true);
  const currentUserRef = useRef<User | null>(null);
  const authResolutionRef = useRef<Promise<void> | null>(null);

  useEffect(() => {
    currentUserRef.current = currentUser;
  }, [currentUser]);

  async function bootstrapOrgContextWithTimeout(user: User, source: string) {
    return await Promise.race([
      bootstrapCurrentOrgForUser({
        userId: user.id,
        email: user.email,
      }),
      new Promise<null>((resolve) => {
        window.setTimeout(() => {
          console.warn("[auth] org bootstrap timed out", { source, userId: user.id });
          resolve(null);
        }, ORG_BOOTSTRAP_TIMEOUT_MS);
      }),
    ]);
  }

  async function applyResolvedSession(session: Session | null, source: string) {
    if (!mountedRef.current) return;

    setCurrentSession(session ?? null);
    setCurrentUser(session?.user ?? null);

    if (!session?.user) {
      console.info("[auth] signed out", { source });
      clearCachedOrgContext();
      setOrgContext(null);
      setLoading(false);
      return;
    }

    console.info("[auth] signed in", { source, userId: session.user.id });
    setLoading(true);
    try {
      const nextOrgContext = await bootstrapOrgContextWithTimeout(session.user, source);
      if (!mountedRef.current) return;
      setOrgContext(nextOrgContext);
    } finally {
      if (mountedRef.current) {
        setLoading(false);
      }
    }
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

    const existingResolution = authResolutionRef.current;
    if (existingResolution) {
      console.info("[auth] auth resolution reused", { source: "refresh" });
      await existingResolution;
      return;
    }

    const resolutionPromise = (async () => {
      console.info("[auth] auth resolution start", { source: "refresh" });
      setLoading(true);

      try {
        const supabase = getSupabaseBrowserClient();
        if (!supabase) {
          if (mountedRef.current) {
            setCurrentSession(null);
            setCurrentUser(null);
            setOrgContext(null);
            setLoading(false);
          }
          return;
        }

        const {
          data: { session },
        } = await supabase.auth.getSession();

        await applyResolvedSession(session ?? null, "refresh");
        console.info("[auth] auth resolution success", {
          source: "refresh",
          hasSession: Boolean(session),
          userId: session?.user?.id ?? null,
        });
      } finally {
        authResolutionRef.current = null;
      }
    })();

    authResolutionRef.current = resolutionPromise;
    await resolutionPromise;
  }

  async function resolveSessionFromEvent(source: string, session: Session | null) {
    const existingResolution = authResolutionRef.current;
    if (existingResolution) {
      console.info("[auth] auth resolution reused", { source });
      await existingResolution;
      return;
    }

    const resolutionPromise = (async () => {
      console.info("[auth] auth resolution start", { source });
      await applyResolvedSession(session, source);
      console.info("[auth] auth resolution success", {
        source,
        hasSession: Boolean(session),
        userId: session?.user?.id ?? null,
      });
    })().finally(() => {
      authResolutionRef.current = null;
    });

    authResolutionRef.current = resolutionPromise;
    await resolutionPromise;
  }

  const refreshAuthContextEffect = useEffectEvent(async () => {
    await refreshAuthContext();
  });

  const resolveSessionFromEventEffect = useEffectEvent(async (source: string, session: Session | null) => {
    await resolveSessionFromEvent(source, session);
  });

  useEffect(() => {
    mountedRef.current = true;

    if (!authEnabled) {
      setLoading(false);
      return () => {
        mountedRef.current = false;
      };
    }

    void refreshAuthContextEffect();

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
        await resolveSessionFromEventEffect("INITIAL_SESSION", session ?? null);
        return;
      }

      if (event === "SIGNED_IN") {
        await resolveSessionFromEventEffect("SIGNED_IN", session ?? null);
        return;
      }

      if (event === "SIGNED_OUT") {
        await resolveSessionFromEventEffect("SIGNED_OUT", null);
        return;
      }
    });

    const shouldRetrySessionLoad = hasAuthParamsInUrl();
    const retryTimer = shouldRetrySessionLoad
      ? window.setTimeout(() => {
          if (!authResolutionRef.current && !currentUserRef.current) {
            void refreshAuthContextEffect();
          }
        }, 500)
      : null;

    return () => {
      mountedRef.current = false;
      if (retryTimer) {
        window.clearTimeout(retryTimer);
      }
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
