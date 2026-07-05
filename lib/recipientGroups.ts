import { readPersistedValue, writePersistedValue } from "./browserStorage";
import { getScopedStorageKey } from "./clientPersistence";
import { getCachedOrgContext } from "./orgBootstrap";
import { getSupabaseBrowserClient, isSupabaseConfigured } from "./supabaseClient";

export type RecipientGroup = {
  id: string;
  name: string;
  emails: string[];
  createdAt: string;
  updatedAt: string;
};

export type RecipientEntry =
  | {
      type: "email";
      email: string;
      name: string;
    }
  | {
      type: "group";
      groupId: string;
      name: string;
    };

type PersistedRecipientGroupsRecord = {
  groups: RecipientGroup[];
  updatedAt: string;
};

type RecipientGroupSaveInput = {
  id?: string;
  name: string;
  emails: string[];
};

const RECIPIENT_GROUPS_STORAGE_KEY = "event-based-reminders-app:recipient-groups";
export const RECIPIENT_GROUPS_UPDATED_EVENT = "event-based-reminders-app:recipient-groups-updated";

function getRecipientGroupsStorageKey() {
  return getScopedStorageKey(RECIPIENT_GROUPS_STORAGE_KEY);
}

function isObject(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

export function normalizeRecipientGroupEmail(value: string) {
  return value.trim().replace(/^mailto:/i, "").trim().toLowerCase();
}

export function isValidRecipientGroupEmail(value: string) {
  const normalized = normalizeRecipientGroupEmail(value);
  return normalized.includes("@");
}

export function normalizeRecipientGroupEmails(values: unknown) {
  const rawValues = Array.isArray(values)
    ? values
    : typeof values === "string"
      ? values.split(/[,\n;]/)
      : [];

  return Array.from(
    new Set(
      rawValues
        .map((entry) => normalizeRecipientGroupEmail(String(entry ?? "")))
        .filter((entry) => entry && isValidRecipientGroupEmail(entry))
    )
  );
}

function isRecipientEntry(value: unknown): value is RecipientEntry {
  if (!isObject(value)) return false;
  if (value.type === "email") {
    return typeof value.email === "string" && typeof value.name === "string";
  }
  if (value.type === "group") {
    return typeof value.groupId === "string" && typeof value.name === "string";
  }
  return false;
}

export function normalizeRecipientEntry(value: unknown): RecipientEntry | null {
  if (typeof value === "string") {
    const normalizedEmail = normalizeRecipientGroupEmail(value);
    if (!normalizedEmail || !isValidRecipientGroupEmail(normalizedEmail)) return null;
    return {
      type: "email",
      email: normalizedEmail,
      name: normalizedEmail,
    };
  }

  if (!isRecipientEntry(value)) return null;

  if (value.type === "email") {
    const normalizedEmail = normalizeRecipientGroupEmail(value.email);
    if (!normalizedEmail || !isValidRecipientGroupEmail(normalizedEmail)) return null;
    return {
      type: "email",
      email: normalizedEmail,
      name: typeof value.name === "string" && value.name.trim() ? value.name.trim() : normalizedEmail,
    };
  }

  const groupId = value.groupId.trim();
  const name = value.name.trim();
  if (!groupId || !name) return null;
  return {
    type: "group",
    groupId,
    name,
  };
}

export function normalizeRecipientEntries(values: unknown): RecipientEntry[] {
  const rawValues = Array.isArray(values)
    ? values
    : typeof values === "string"
      ? values.split(/[,\n;]/)
      : [];

  const normalizedEntries = rawValues
    .map((entry) => normalizeRecipientEntry(entry))
    .filter((entry): entry is RecipientEntry => Boolean(entry));

  return dedupeRecipientEntries(normalizedEntries);
}

export function createEmailRecipientEntry(value: string): RecipientEntry | null {
  return normalizeRecipientEntry(value);
}

export function createGroupRecipientEntry(group: Pick<RecipientGroup, "id" | "name">): RecipientEntry {
  return {
    type: "group",
    groupId: group.id,
    name: group.name.trim(),
  };
}

export function dedupeRecipientEntries(entries: RecipientEntry[]) {
  const seenEmails = new Set<string>();
  const seenGroups = new Set<string>();
  const deduped: RecipientEntry[] = [];

  entries.forEach((entry) => {
    if (entry.type === "email") {
      const normalizedEmail = normalizeRecipientGroupEmail(entry.email);
      if (!normalizedEmail || seenEmails.has(normalizedEmail)) return;
      seenEmails.add(normalizedEmail);
      deduped.push({
        type: "email",
        email: normalizedEmail,
        name: entry.name.trim() || normalizedEmail,
      });
      return;
    }

    const normalizedGroupId = entry.groupId.trim();
    if (!normalizedGroupId || seenGroups.has(normalizedGroupId)) return;
    seenGroups.add(normalizedGroupId);
    deduped.push({
      type: "group",
      groupId: normalizedGroupId,
      name: entry.name.trim() || "Recipient Group",
    });
  });

  return deduped;
}

export function mergeRecipientEntries(existing: RecipientEntry[], additions: RecipientEntry[]) {
  return dedupeRecipientEntries([...existing, ...additions]);
}

export function resolveRecipientEntries(entries: RecipientEntry[], groups: RecipientGroup[]) {
  const groupsById = new Map(groups.map((group) => [group.id, group]));
  return Array.from(
    new Set(
      entries.flatMap((entry) => {
        if (entry.type === "email") {
          return [normalizeRecipientGroupEmail(entry.email)];
        }
        const group = groupsById.get(entry.groupId);
        if (!group) return [];
        return group.emails.map((email) => normalizeRecipientGroupEmail(email));
      }).filter((email) => email && isValidRecipientGroupEmail(email))
    )
  );
}

export function getRecipientGroupFromEntry(entry: RecipientEntry, groups: RecipientGroup[]) {
  if (entry.type !== "group") return null;
  return groups.find((group) => group.id === entry.groupId) ?? null;
}

function normalizeRecipientGroup(value: unknown): RecipientGroup | null {
  if (!isObject(value)) return null;
  const id = typeof value.id === "string" ? value.id : "";
  const name = typeof value.name === "string" ? value.name.trim() : "";
  if (!id || !name) return null;

  const createdAt = typeof value.createdAt === "string" ? value.createdAt : typeof value.created_at === "string" ? value.created_at : "";
  const updatedAt = typeof value.updatedAt === "string" ? value.updatedAt : typeof value.updated_at === "string" ? value.updated_at : "";

  return {
    id,
    name,
    emails: normalizeRecipientGroupEmails(value.emails),
    createdAt: createdAt || updatedAt || new Date().toISOString(),
    updatedAt: updatedAt || createdAt || new Date().toISOString(),
  };
}

function sortRecipientGroups(groups: RecipientGroup[]) {
  return [...groups].sort((left, right) => left.name.localeCompare(right.name, undefined, { sensitivity: "base" }));
}

function readLocalRecipientGroupsRecord() {
  if (typeof window === "undefined") return null;
  const raw = readPersistedValue("localStorage", getRecipientGroupsStorageKey());
  if (!raw) return null;

  try {
    const parsed = JSON.parse(raw) as Record<string, unknown>;
    const groups = Array.isArray(parsed.groups)
      ? parsed.groups.map((entry) => normalizeRecipientGroup(entry)).filter((entry): entry is RecipientGroup => Boolean(entry))
      : [];
    return {
      groups: sortRecipientGroups(groups),
      updatedAt: typeof parsed.updatedAt === "string" ? parsed.updatedAt : "",
    } satisfies PersistedRecipientGroupsRecord;
  } catch {
    return null;
  }
}

function cacheRecipientGroups(groups: RecipientGroup[], updatedAt = new Date().toISOString()) {
  if (typeof window === "undefined") return;
  const normalizedGroups = sortRecipientGroups(groups).map((group) => ({
    id: group.id,
    name: group.name,
    emails: normalizeRecipientGroupEmails(group.emails),
    createdAt: group.createdAt,
    updatedAt: group.updatedAt,
  }));
  writePersistedValue(
    "localStorage",
    getRecipientGroupsStorageKey(),
    JSON.stringify({ groups: normalizedGroups, updatedAt })
  );
  window.dispatchEvent(
    new CustomEvent(RECIPIENT_GROUPS_UPDATED_EVENT, {
      detail: normalizedGroups,
    })
  );
}

export function loadRecipientGroups() {
  return readLocalRecipientGroupsRecord()?.groups ?? [];
}

export async function hydrateRecipientGroupsFromSupabase() {
  if (typeof window === "undefined" || !isSupabaseConfigured()) {
    return loadRecipientGroups();
  }

  const supabase = getSupabaseBrowserClient();
  const orgId = getCachedOrgContext()?.orgId ?? "";

  if (!supabase || !orgId) {
    return loadRecipientGroups();
  }

  const { data, error } = await supabase
    .from("org_recipient_groups")
    .select("id,name,emails,created_at,updated_at")
    .eq("org_id", orgId)
    .order("name", { ascending: true });

  if (error || !Array.isArray(data)) {
    return loadRecipientGroups();
  }

  const groups = data
    .map((entry) => normalizeRecipientGroup(entry))
    .filter((entry): entry is RecipientGroup => Boolean(entry));

  cacheRecipientGroups(groups);
  return groups;
}

export async function saveRecipientGroup(input: RecipientGroupSaveInput) {
  const now = new Date().toISOString();
  const nextGroup: RecipientGroup = {
    id: input.id?.trim() || crypto.randomUUID(),
    name: input.name.trim(),
    emails: normalizeRecipientGroupEmails(input.emails),
    createdAt: now,
    updatedAt: now,
  };

  const localGroups = loadRecipientGroups();
  const existingGroup = localGroups.find((group) => group.id === nextGroup.id) ?? null;
  const mergedGroup = existingGroup
    ? {
        ...existingGroup,
        name: nextGroup.name,
        emails: nextGroup.emails,
        updatedAt: now,
      }
    : nextGroup;

  const nextGroups = sortRecipientGroups(
    existingGroup
      ? localGroups.map((group) => (group.id === mergedGroup.id ? mergedGroup : group))
      : [...localGroups, mergedGroup]
  );

  cacheRecipientGroups(nextGroups, now);

  if (!isSupabaseConfigured()) {
    return mergedGroup;
  }

  const supabase = getSupabaseBrowserClient();
  const orgId = getCachedOrgContext()?.orgId ?? "";

  if (!supabase || !orgId) {
    return mergedGroup;
  }

  const { data, error } = await supabase
    .from("org_recipient_groups")
    .upsert(
      {
        id: mergedGroup.id,
        org_id: orgId,
        name: mergedGroup.name,
        emails: mergedGroup.emails,
        updated_at: now,
      },
      { onConflict: "id" }
    )
    .select("id,name,emails,created_at,updated_at")
    .single();

  if (error || !data) {
    return mergedGroup;
  }

  const normalizedSavedGroup = normalizeRecipientGroup(data) ?? mergedGroup;
  const syncedGroups = sortRecipientGroups(
    nextGroups.map((group) => (group.id === normalizedSavedGroup.id ? normalizedSavedGroup : group))
  );
  cacheRecipientGroups(syncedGroups, normalizedSavedGroup.updatedAt);
  return normalizedSavedGroup;
}

export async function deleteRecipientGroup(id: string) {
  const targetId = id.trim();
  if (!targetId) return;

  const now = new Date().toISOString();
  const nextGroups = loadRecipientGroups().filter((group) => group.id !== targetId);
  cacheRecipientGroups(nextGroups, now);

  if (!isSupabaseConfigured()) {
    return;
  }

  const supabase = getSupabaseBrowserClient();
  const orgId = getCachedOrgContext()?.orgId ?? "";

  if (!supabase || !orgId) {
    return;
  }

  await supabase.from("org_recipient_groups").delete().eq("org_id", orgId).eq("id", targetId);
}
