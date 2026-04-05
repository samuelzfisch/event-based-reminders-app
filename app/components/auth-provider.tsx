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

export function AuthProvider({ children }: { children: ReactNode }) {
  const [authEnabled] = useState(() => isSupabaseConfigured());
  const [loading, setLoading] = useState(true);
  const [currentSession, setCurrentSession] = useState<Session | null>(null);
  const [currentUser, setCurrentUser] = useState<User | null>(null);
  const [orgContext, setOrgContext] = useState<BootstrappedOrgContext | null>(null);
  const mountedRef = useRef(true);
  const authResolutionRef = useRef<Promise<void> | null>(null);

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
      setLoading(false);
      console.info("[auth] loading cleared", { source });
      return;
    }

    setLoading(true);
    try {
      const nextOrgContext = await bootstrapOrgContextWithTimeout(session.user, source);
      if (!mountedRef.current) return;
      setOrgContext(nextOrgContext);
    } finally {
      if (mountedRef.current) {
        setLoading(false);
        console.info("[auth] loading cleared", { source });
      }
    }
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
    })().finally(() => {
      authResolutionRef.current = null;
    });

    authResolutionRef.current = resolutionPromise;
    await resolutionPromise;
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
      } = await supabase.auth.getSession();

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

    void refreshAuthContextEffect();

    const {
      data: { subscription },
    } = supabase.auth.onAuthStateChange(async (event, session) => {
      if (!mountedRef.current) return;
      console.info("[auth] auth event received", { event });

      if (event === "INITIAL_SESSION") {
        return;
      }

      if (event === "SIGNED_IN") {
        await resolveWithSingleFlightEffect("SIGNED_IN", async () => session ?? null);
        return;
      }

      if (event === "SIGNED_OUT") {
        await resolveWithSingleFlightEffect("SIGNED_OUT", async () => null);
        return;
      }
    });

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
