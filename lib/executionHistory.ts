import { migrateLegacyPersistedValue, readPersistedValue, writePersistedValue } from "./browserStorage";
import { getSupabaseBrowserClient, isSupabaseConfigured } from "./supabaseClient";
import { getLocalUserKey } from "./userKey";

export type ExecutionHistoryStatus =
  | "success"
  | "fallback"
  | "failed"
  | "modified"
  | "modify_failed"
  | "recalled"
  | "recall_failed"
  | "already_removed"
  | "already_canceled";
export type ExecutionHistoryPath = "graph" | "fallback";
export type ExecutionHistoryItemType = "email" | "reminder" | "meeting" | "teams_meeting";
export type ExecutionHistoryFallbackExportKind = "eml" | "ics" | null;
export type ExecutionHistoryProvider = "outlook" | "local_export";
export type ExecutionHistoryProviderObjectType = "message" | "event" | "file" | null;

export type ExecutionHistoryRecord = {
  id: string;
  userKey: string;
  executionGroupId: string | null;
  planName: string;
  itemType: ExecutionHistoryItemType;
  title: string;
  subject: string;
  status: ExecutionHistoryStatus;
  path: ExecutionHistoryPath;
  recipients: string[];
  attendees: string[];
  executedAt: string;
  outlookWebLink: string | null;
  teamsJoinLink: string | null;
  fallbackExportKind: ExecutionHistoryFallbackExportKind;
  provider: ExecutionHistoryProvider;
  providerObjectId: string | null;
  providerObjectType: ExecutionHistoryProviderObjectType;
  canRecall: boolean;
  canModify: boolean;
  recallImplemented: boolean;
  modifyImplemented: boolean;
  recallReason: string | null;
  modifyReason: string | null;
  scheduledFor: string | null;
  endsAt: string | null;
  isAllDay: boolean;
  details: Record<string, unknown>;
};

export type ExecutionHistoryInsert = Omit<ExecutionHistoryRecord, "id" | "userKey" | "executedAt"> & {
  id?: string;
  executedAt?: string;
};

export const EXECUTION_HISTORY_UPDATED_EVENT = "event-based-reminders-app:execution-history-updated";
const EXECUTION_HISTORY_STORAGE_KEY = "event-based-reminders-app:execution-history";
const LEGACY_EXECUTION_HISTORY_STORAGE_KEYS = ["standalone-plans:execution-history"];
const MAX_LOCAL_HISTORY_RECORDS = 400;
const HISTORY_DEBUG_PREFIX = "[executionHistory]";

function isObject(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function normalizeStringArray(value: unknown) {
  if (!Array.isArray(value)) return [];
  return value.filter((entry): entry is string => typeof entry === "string").map((entry) => entry.trim()).filter(Boolean);
}

function normalizeHistoryStatus(value: unknown): ExecutionHistoryStatus {
  return value === "fallback" ||
    value === "failed" ||
    value === "modified" ||
    value === "modify_failed" ||
    value === "recalled" ||
    value === "recall_failed" ||
    value === "already_removed" ||
    value === "already_canceled"
    ? value
    : "success";
}

function normalizeHistoryPath(value: unknown): ExecutionHistoryPath {
  return value === "fallback" ? "fallback" : "graph";
}

function normalizeHistoryItemType(value: unknown): ExecutionHistoryItemType {
  if (value === "reminder" || value === "meeting" || value === "teams_meeting") return value;
  return "email";
}

function normalizeFallbackExportKind(value: unknown): ExecutionHistoryFallbackExportKind {
  if (value === "eml" || value === "ics") return value;
  return null;
}

function readString(value: unknown, fallback = "") {
  return typeof value === "string" ? value : fallback;
}

function readNullableString(value: unknown) {
  return typeof value === "string" && value.trim() ? value : null;
}

function readBoolean(value: unknown, fallback = false) {
  return typeof value === "boolean" ? value : fallback;
}

function normalizeProvider(value: unknown): ExecutionHistoryProvider {
  return value === "local_export" ? "local_export" : "outlook";
}

function normalizeProviderObjectType(value: unknown): ExecutionHistoryProviderObjectType {
  return value === "message" || value === "event" || value === "file" ? value : null;
}

function getLegacyDetails(value: Record<string, unknown>) {
  return isObject(value.details) ? value.details : {};
}

export function getExecutionHistoryRecallState(record: Pick<
  ExecutionHistoryRecord,
  "provider" | "providerObjectId" | "providerObjectType" | "status" | "details"
>) {
  const action = typeof record.details.action === "string" ? record.details.action : "";

  if (record.status === "recalled" || record.status === "already_removed" || record.status === "already_canceled") {
    return {
      canRecall: false,
      recallImplemented: true,
      recallReason:
        record.status === "already_removed"
          ? "Already removed."
          : record.status === "already_canceled"
            ? "Already canceled."
            : "Already recalled.",
    };
  }

  if (record.provider !== "outlook" || !record.providerObjectId) {
    return {
      canRecall: false,
      recallImplemented: false,
      recallReason: "This item cannot be recalled.",
    };
  }

  if (record.providerObjectType === "event") {
    return {
      canRecall: true,
      recallImplemented: true,
      recallReason: null,
    };
  }

  if (record.providerObjectType === "message" && action === "draft_created") {
    return {
      canRecall: true,
      recallImplemented: true,
      recallReason: null,
    };
  }

  if (record.providerObjectType === "message" && action === "email_scheduled") {
    return {
      canRecall: false,
      recallImplemented: false,
      recallReason: "Scheduled emails cannot be recalled from History.",
    };
  }

  if (record.providerObjectType === "message" && action === "email_sent") {
    return {
      canRecall: false,
      recallImplemented: false,
      recallReason: "Sent emails cannot be recalled from History.",
    };
  }

  return {
    canRecall: false,
    recallImplemented: false,
    recallReason: record.status === "recall_failed" ? "Last recall attempt failed. You can try again." : "This item cannot be recalled.",
  };
}

export function getExecutionHistoryModifyState(record: Pick<
  ExecutionHistoryRecord,
  "provider" | "providerObjectId" | "providerObjectType" | "status" | "details"
>) {
  const action = typeof record.details.action === "string" ? record.details.action : "";
  const scheduledEmailState = typeof record.details.scheduledEmailState === "string" ? record.details.scheduledEmailState : "";

  if (record.provider !== "outlook" || !record.providerObjectId) {
    return {
      canModify: false,
      modifyImplemented: false,
      modifyReason: "This item cannot be modified.",
    };
  }

  if (record.providerObjectType === "event") {
    return {
      canModify: true,
      modifyImplemented: true,
      modifyReason: null,
    };
  }

  if (record.providerObjectType === "message" && action === "draft_created") {
    return {
      canModify: true,
      modifyImplemented: true,
      modifyReason: null,
    };
  }

  if (record.providerObjectType === "message" && action === "email_scheduled") {
    if (scheduledEmailState === "sent") {
      return {
        canModify: false,
        modifyImplemented: true,
        modifyReason: "This scheduled email has already been sent and can't be changed.",
      };
    }
    return {
      canModify: true,
      modifyImplemented: true,
      modifyReason: null,
    };
  }

  if (record.providerObjectType === "message" && action === "email_sent") {
    return {
      canModify: false,
      modifyImplemented: false,
      modifyReason: "Sent emails cannot be modified from History.",
    };
  }

  return {
    canModify: false,
    modifyImplemented: false,
    modifyReason: record.status === "modify_failed" ? "Last edit attempt failed. You can try again." : "This item cannot be modified.",
  };
}

function normalizeRecord(value: unknown): ExecutionHistoryRecord | null {
  if (!isObject(value) || typeof value.id !== "string") return null;

  const details = getLegacyDetails(value);

  const record = {
    id: value.id,
    userKey: readString(value.user_key ?? value.userKey),
    executionGroupId: readNullableString(value.execution_group_id ?? value.executionGroupId),
    planName: readString(value.plan_name ?? value.planName),
    itemType: normalizeHistoryItemType(value.item_type ?? value.itemType),
    title: readString(value.title),
    subject: readString(value.subject),
    status: normalizeHistoryStatus(value.status),
    path: normalizeHistoryPath(value.path),
    recipients: normalizeStringArray(value.recipients),
    attendees: normalizeStringArray(value.attendees),
    executedAt: readString(value.executed_at ?? value.executedAt, new Date(0).toISOString()),
    outlookWebLink: readNullableString(value.outlook_web_link ?? value.outlookWebLink),
    teamsJoinLink: readNullableString(value.teams_join_link ?? value.teamsJoinLink),
    fallbackExportKind: normalizeFallbackExportKind(value.fallback_export_kind ?? value.fallbackExportKind),
    provider: normalizeProvider(value.provider ?? details.provider ?? (value.path === "fallback" ? "local_export" : "outlook")),
    providerObjectId: readNullableString(value.provider_object_id ?? value.providerObjectId ?? details.providerObjectId),
    providerObjectType: normalizeProviderObjectType(
      value.provider_object_type ?? value.providerObjectType ?? details.providerObjectType
    ),
    canRecall: readBoolean(value.can_recall ?? value.canRecall ?? details.canRecall),
    canModify: readBoolean(value.can_modify ?? value.canModify ?? details.canModify),
    recallImplemented: readBoolean(value.recall_implemented ?? value.recallImplemented ?? details.recallImplemented),
    modifyImplemented: readBoolean(value.modify_implemented ?? value.modifyImplemented ?? details.modifyImplemented),
    recallReason: readNullableString(value.recall_reason ?? value.recallReason ?? details.recallReason),
    modifyReason: readNullableString(value.modify_reason ?? value.modifyReason ?? details.modifyReason),
    scheduledFor: readNullableString(value.scheduled_for ?? value.scheduledFor ?? details.scheduledFor),
    endsAt: readNullableString(value.ends_at ?? value.endsAt ?? details.endsAt),
    isAllDay: readBoolean(value.is_all_day ?? value.isAllDay ?? details.isAllDay),
    details,
  };

  const recallState = getExecutionHistoryRecallState(record);
  const modifyState = getExecutionHistoryModifyState(record);
  return {
    ...record,
    canRecall: recallState.canRecall,
    recallImplemented: recallState.recallImplemented,
    recallReason: recallState.recallReason,
    canModify: modifyState.canModify,
    modifyImplemented: modifyState.modifyImplemented,
    modifyReason: modifyState.modifyReason,
  };
}

function loadLocalExecutionHistory() {
  if (typeof window === "undefined") return [];
  migrateLegacyPersistedValue("localStorage", EXECUTION_HISTORY_STORAGE_KEY, LEGACY_EXECUTION_HISTORY_STORAGE_KEYS);
  const raw = readPersistedValue("localStorage", EXECUTION_HISTORY_STORAGE_KEY, LEGACY_EXECUTION_HISTORY_STORAGE_KEYS);
  if (!raw) return [];

  try {
    const parsed = JSON.parse(raw) as unknown[];
    if (!Array.isArray(parsed)) return [];
    return parsed.map(normalizeRecord).filter((entry): entry is ExecutionHistoryRecord => Boolean(entry));
  } catch {
    return [];
  }
}

function saveLocalExecutionHistory(records: ExecutionHistoryRecord[]) {
  if (typeof window === "undefined") return;
  const nextRecords = records.slice(0, MAX_LOCAL_HISTORY_RECORDS);
  writePersistedValue("localStorage", EXECUTION_HISTORY_STORAGE_KEY, JSON.stringify(nextRecords));
  console.info(HISTORY_DEBUG_PREFIX, "saved local records", {
    storageKey: EXECUTION_HISTORY_STORAGE_KEY,
    count: nextRecords.length,
    firstRecordId: nextRecords[0]?.id ?? null,
  });
  window.dispatchEvent(new CustomEvent(EXECUTION_HISTORY_UPDATED_EVENT, { detail: nextRecords[0] ?? null }));
}

function cacheExecutionHistoryRecord(record: ExecutionHistoryRecord) {
  const nextRecords = [record, ...loadLocalExecutionHistory().filter((entry) => entry.id !== record.id)];
  saveLocalExecutionHistory(nextRecords);
}

function mergeExecutionHistoryRecords(records: ExecutionHistoryRecord[]) {
  const seen = new Set<string>();
  return records
    .filter((record) => {
      if (seen.has(record.id)) return false;
      seen.add(record.id);
      return true;
    })
    .sort((left, right) => right.executedAt.localeCompare(left.executedAt));
}

function updateLocalExecutionHistoryRecord(
  recordId: string,
  updater: (record: ExecutionHistoryRecord) => ExecutionHistoryRecord
) {
  const nextRecords = loadLocalExecutionHistory().map((record) => (record.id === recordId ? updater(record) : record));
  saveLocalExecutionHistory(nextRecords);
  return nextRecords.find((record) => record.id === recordId) ?? null;
}

export async function writeExecutionHistory(entry: ExecutionHistoryInsert) {
  if (typeof window === "undefined") return;

  const userKey = getLocalUserKey();
  if (!userKey) return;

  const record: ExecutionHistoryRecord = {
    ...entry,
    id: entry.id ?? crypto.randomUUID(),
    userKey,
    executionGroupId: entry.executionGroupId ?? null,
    executedAt: entry.executedAt ?? new Date().toISOString(),
  };

  console.info(HISTORY_DEBUG_PREFIX, "writeExecutionHistory called", {
    storageKey: EXECUTION_HISTORY_STORAGE_KEY,
    userKey,
    executionGroupId: record.executionGroupId,
    itemType: record.itemType,
    status: record.status,
    path: record.path,
    executedAt: record.executedAt,
    providerObjectId: record.providerObjectId,
  });

  cacheExecutionHistoryRecord(record);

  if (!isSupabaseConfigured()) return;

  const supabase = getSupabaseBrowserClient();
  if (!supabase || !userKey) return;

  await supabase.from("execution_history").insert({
    id: record.id,
    user_key: userKey,
    execution_group_id: record.executionGroupId,
    plan_name: record.planName,
    item_type: record.itemType,
    title: record.title,
    subject: record.subject,
    status: record.status,
    path: record.path,
    recipients: record.recipients,
    attendees: record.attendees,
    executed_at: record.executedAt,
    outlook_web_link: record.outlookWebLink,
    teams_join_link: record.teamsJoinLink,
    fallback_export_kind: record.fallbackExportKind,
    details: {
      ...record.details,
      provider: record.provider,
      providerObjectId: record.providerObjectId,
      providerObjectType: record.providerObjectType,
      canRecall: record.canRecall,
      canModify: record.canModify,
      recallImplemented: record.recallImplemented,
      modifyImplemented: record.modifyImplemented,
      recallReason: record.recallReason,
      modifyReason: record.modifyReason,
      scheduledFor: record.scheduledFor,
      endsAt: record.endsAt,
      isAllDay: record.isAllDay,
    },
  });
}

export async function updateExecutionHistoryRecord(
  recordId: string,
  updates: Partial<ExecutionHistoryRecord> & { details?: Record<string, unknown> }
) {
  const userKey = getLocalUserKey();
  if (!userKey) return null;

  const nextRecord = updateLocalExecutionHistoryRecord(recordId, (record) => ({
    ...record,
    ...updates,
    details: updates.details ? { ...record.details, ...updates.details } : record.details,
  }));

  if (!nextRecord || !isSupabaseConfigured()) return nextRecord;

  const supabase = getSupabaseBrowserClient();
  if (!supabase) return nextRecord;

  await supabase
    .from("execution_history")
    .update({
      status: nextRecord.status,
      details: {
        ...nextRecord.details,
        provider: nextRecord.provider,
        providerObjectId: nextRecord.providerObjectId,
        providerObjectType: nextRecord.providerObjectType,
        canRecall: nextRecord.canRecall,
        canModify: nextRecord.canModify,
        recallImplemented: nextRecord.recallImplemented,
        modifyImplemented: nextRecord.modifyImplemented,
        recallReason: nextRecord.recallReason,
        modifyReason: nextRecord.modifyReason,
        scheduledFor: nextRecord.scheduledFor,
        endsAt: nextRecord.endsAt,
        isAllDay: nextRecord.isAllDay,
      },
    })
    .eq("id", recordId)
    .eq("user_key", userKey);

  return nextRecord;
}

export async function listExecutionHistory(limit = 200) {
  const localRecords = loadLocalExecutionHistory();
  const userKey = getLocalUserKey();
  const matchingLocalRecords = localRecords.filter((record) => !record.userKey || record.userKey === userKey);
  console.info(HISTORY_DEBUG_PREFIX, "listExecutionHistory local load", {
    storageKey: EXECUTION_HISTORY_STORAGE_KEY,
    userKey,
    localCount: localRecords.length,
    matchingLocalCount: matchingLocalRecords.length,
  });
  if (typeof window === "undefined" || !isSupabaseConfigured()) return matchingLocalRecords.slice(0, limit);

  const supabase = getSupabaseBrowserClient();

  if (!supabase || !userKey) return matchingLocalRecords.slice(0, limit);

  const { data, error } = await supabase
    .from("execution_history")
    .select(
      "id,user_key,execution_group_id,plan_name,item_type,title,subject,status,path,recipients,attendees,executed_at,outlook_web_link,teams_join_link,fallback_export_kind,details"
    )
    .eq("user_key", userKey)
    .order("executed_at", { ascending: false })
    .limit(limit);

  if (error || !data) {
    console.info(HISTORY_DEBUG_PREFIX, "listExecutionHistory supabase fallback", {
      userKey,
      hadError: Boolean(error),
      localCount: matchingLocalRecords.length,
    });
    return matchingLocalRecords.slice(0, limit);
  }

  const merged = mergeExecutionHistoryRecords([
    ...data.map(normalizeRecord).filter((entry): entry is ExecutionHistoryRecord => Boolean(entry)),
    ...matchingLocalRecords,
  ]).slice(0, limit);
  console.info(HISTORY_DEBUG_PREFIX, "listExecutionHistory merged load", {
    userKey,
    remoteCount: data.length,
    localCount: matchingLocalRecords.length,
    mergedCount: merged.length,
  });
  return merged;
}
