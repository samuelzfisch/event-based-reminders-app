import { migrateLegacyPersistedValue, readPersistedValue, removePersistedValue, writePersistedValue } from "./browserStorage";
import { getSupabaseBrowserClient, isSupabaseConfigured } from "./supabaseClient";

const ORG_CONTEXT_STORAGE_KEY = "event-based-reminders-app:org-context";
const LEGACY_ORG_CONTEXT_STORAGE_KEYS = ["standalone-plans:org-context"];
const activeBootstrapPromises = new Map<string, Promise<BootstrappedOrgContext | null>>();

export type BootstrappedOrgContext = {
  userId: string;
  orgId: string;
  role: string;
};

function normalizeString(value: unknown) {
  return typeof value === "string" ? value.trim() : "";
}

export function getCachedOrgContext() {
  if (typeof window === "undefined") return null;

  migrateLegacyPersistedValue("localStorage", ORG_CONTEXT_STORAGE_KEY, LEGACY_ORG_CONTEXT_STORAGE_KEYS);
  const raw = readPersistedValue("localStorage", ORG_CONTEXT_STORAGE_KEY, LEGACY_ORG_CONTEXT_STORAGE_KEYS);
  if (!raw) return null;

  try {
    const parsed = JSON.parse(raw) as Record<string, unknown>;
    const userId = normalizeString(parsed.userId);
    const orgId = normalizeString(parsed.orgId);
    const role = normalizeString(parsed.role) || "member";
    if (!userId || !orgId) return null;
    return { userId, orgId, role } satisfies BootstrappedOrgContext;
  } catch {
    return null;
  }
}

function cacheOrgContext(context: BootstrappedOrgContext | null) {
  if (typeof window === "undefined") return;
  if (!context) {
    writePersistedValue("localStorage", ORG_CONTEXT_STORAGE_KEY, "");
    return;
  }
  writePersistedValue("localStorage", ORG_CONTEXT_STORAGE_KEY, JSON.stringify(context));
}

function deriveOrganizationName(email: string | null | undefined) {
  const safeEmail = normalizeString(email);
  if (!safeEmail) return "My Organization";
  const localPart = safeEmail.split("@")[0]?.trim();
  if (!localPart) return "My Organization";
  const normalized = localPart.replace(/[._-]+/g, " ").trim();
  const label = normalized ? normalized.charAt(0).toUpperCase() + normalized.slice(1) : "My";
  return `${label}'s Organization`;
}

async function runBootstrapCurrentOrgForUser(input: {
  userId: string;
  email?: string | null;
}) {
  try {
    console.info("[orgBootstrap] bootstrap start", { userId: input.userId, email: input.email ?? null });

    if (!isSupabaseConfigured()) return null;

    const supabase = getSupabaseBrowserClient();
    if (!supabase) return null;

    const {
      data: { user: authenticatedUser },
      error: getUserError,
    } = await supabase.auth.getUser();

    if (getUserError) {
      console.error("[orgBootstrap] auth.getUser failed", getUserError);
      return null;
    }

    if (!authenticatedUser?.id || authenticatedUser.id !== input.userId) {
      console.error("[orgBootstrap] authenticated user mismatch", {
        expectedUserId: input.userId,
        actualUserId: authenticatedUser?.id ?? null,
      });
      return null;
    }

    const cached = getCachedOrgContext();
    if (cached?.userId === input.userId) {
      return cached;
    }

    const membershipQuery = await supabase
      .from("org_members")
      .select("org_id,role,created_at")
      .eq("user_id", input.userId)
      .order("created_at", { ascending: true })
      .limit(1)
      .maybeSingle();

    if (membershipQuery.error) {
      console.error("[orgBootstrap] membership query failed", membershipQuery.error);
      return null;
    }

    if (membershipQuery.data?.org_id) {
      const context = {
        userId: input.userId,
        orgId: membershipQuery.data.org_id,
        role: normalizeString(membershipQuery.data.role) || "member",
      } satisfies BootstrappedOrgContext;
      cacheOrgContext(context);
      return context;
    }

    console.info("[orgBootstrap] creating organization", { userId: input.userId });
    const createOrg = await supabase
      .from("organizations")
      .insert({
        name: deriveOrganizationName(input.email),
        created_by: input.userId,
      })
      .select("id")
      .single();

    if (createOrg.error || !createOrg.data?.id) {
      console.error("[orgBootstrap] organization insert failed", createOrg.error ?? null);
      return null;
    }

    console.info("[orgBootstrap] creating org membership", {
      userId: input.userId,
      orgId: createOrg.data.id,
    });
    const membershipInsert = await supabase.from("org_members").insert({
      org_id: createOrg.data.id,
      user_id: input.userId,
      role: "owner",
    });

    if (membershipInsert.error) {
      console.error("[orgBootstrap] org_members insert failed", membershipInsert.error);
      return null;
    }

    const context = {
      userId: input.userId,
      orgId: createOrg.data.id,
      role: "owner",
    } satisfies BootstrappedOrgContext;
    cacheOrgContext(context);
    return context;
  } catch (error) {
    console.error("[orgBootstrap] bootstrap failed", error);
    return null;
  }
}

export async function bootstrapCurrentOrgForUser(input: {
  userId: string;
  email?: string | null;
}) {
  const cached = getCachedOrgContext();
  if (cached?.userId === input.userId) {
    return cached;
  }

  const existingPromise = activeBootstrapPromises.get(input.userId);
  if (existingPromise) {
    console.info("[orgBootstrap] bootstrap reused", { userId: input.userId });
    return existingPromise;
  }

  const bootstrapPromise = runBootstrapCurrentOrgForUser(input).finally(() => {
    activeBootstrapPromises.delete(input.userId);
  });

  activeBootstrapPromises.set(input.userId, bootstrapPromise);
  return bootstrapPromise;
}

export function clearCachedOrgContext() {
  if (typeof window === "undefined") return;
  removePersistedValue("localStorage", ORG_CONTEXT_STORAGE_KEY, LEGACY_ORG_CONTEXT_STORAGE_KEYS);
}
