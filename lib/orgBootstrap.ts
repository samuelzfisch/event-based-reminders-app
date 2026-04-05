import { migrateLegacyPersistedValue, readPersistedValue, removePersistedValue, writePersistedValue } from "./browserStorage";
import { getSupabaseBrowserClient, isSupabaseConfigured } from "./supabaseClient";

const ORG_CONTEXT_STORAGE_KEY = "event-based-reminders-app:org-context";
const LEGACY_ORG_CONTEXT_STORAGE_KEYS = ["standalone-plans:org-context"];

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

export async function bootstrapCurrentOrgForUser(input: {
  userId: string;
  email?: string | null;
}) {
  if (!isSupabaseConfigured()) return null;

  const supabase = getSupabaseBrowserClient();
  if (!supabase) return null;

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

  if (membershipQuery.data?.org_id) {
    const context = {
      userId: input.userId,
      orgId: membershipQuery.data.org_id,
      role: normalizeString(membershipQuery.data.role) || "member",
    } satisfies BootstrappedOrgContext;
    cacheOrgContext(context);
    return context;
  }

  const createOrg = await supabase
    .from("organizations")
    .insert({
      name: deriveOrganizationName(input.email),
      created_by: input.userId,
    })
    .select("id")
    .single();

  if (createOrg.error || !createOrg.data?.id) {
    throw new Error(createOrg.error?.message || "Could not create organization.");
  }

  const membershipInsert = await supabase.from("org_members").insert({
    org_id: createOrg.data.id,
    user_id: input.userId,
    role: "owner",
  });

  if (membershipInsert.error) {
    throw new Error(membershipInsert.error.message || "Could not create organization membership.");
  }

  const context = {
    userId: input.userId,
    orgId: createOrg.data.id,
    role: "owner",
  } satisfies BootstrappedOrgContext;
  cacheOrgContext(context);
  return context;
}

export function clearCachedOrgContext() {
  if (typeof window === "undefined") return;
  removePersistedValue("localStorage", ORG_CONTEXT_STORAGE_KEY, LEGACY_ORG_CONTEXT_STORAGE_KEYS);
}
