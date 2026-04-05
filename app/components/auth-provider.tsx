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

async function resolveAuthState() {
  if (!isSupabaseConfigured()) {
    return {
      session: null,
      user: null,
      orgContext: null,
    };
  }

  const supabase = getSupabaseBrowserClient();
  if (!supabase) {
    return {
      session: null,
      user: null,
      orgContext: null,
    };
  }

  const {
    data: { session },
  } = await supabase.auth.getSession();
  const user = session?.user ?? null;
  const orgContext = user
    ? await bootstrapCurrentOrgForUser({
        userId: user.id,
        email: user.email,
      })
    : null;

  return { session: session ?? null, user, orgContext };
}

export function AuthProvider({ children }: { children: ReactNode }) {
  const [authEnabled] = useState(() => isSupabaseConfigured());
  const [loading, setLoading] = useState(true);
  const [currentSession, setCurrentSession] = useState<Session | null>(null);
  const [currentUser, setCurrentUser] = useState<User | null>(null);
  const [orgContext, setOrgContext] = useState<BootstrappedOrgContext | null>(null);
  const mountedRef = useRef(true);

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

    setLoading(true);
    try {
      const nextState = await resolveAuthState();
      if (!mountedRef.current) return;
      setCurrentSession(nextState.session);
      setCurrentUser(nextState.user);
      setOrgContext(nextState.orgContext);
    } finally {
      if (mountedRef.current) {
        setLoading(false);
      }
    }
  }

  const refreshAuthContextEffect = useEffectEvent(async () => {
    await refreshAuthContext();
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
    } = supabase.auth.onAuthStateChange(async (_event, session) => {
      if (!mountedRef.current) return;

      setCurrentSession(session ?? null);
      setCurrentUser(session?.user ?? null);

      if (!session?.user) {
        clearCachedOrgContext();
        setOrgContext(null);
        setLoading(false);
        return;
      }

      setLoading(true);
      try {
        const nextOrgContext = await bootstrapCurrentOrgForUser({
          userId: session.user.id,
          email: session.user.email,
        });
        if (!mountedRef.current) return;
        setOrgContext(nextOrgContext);
      } finally {
        if (mountedRef.current) {
          setLoading(false);
        }
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
