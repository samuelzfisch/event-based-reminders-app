import { getSupabaseBrowserClient, isSupabaseConfigured } from "./supabaseClient";
import { getLocalUserKey } from "./userKey";

export type ExecutionHistoryStatus = "success" | "fallback" | "failed";
export type ExecutionHistoryPath = "graph" | "fallback";
export type ExecutionHistoryItemType = "email" | "reminder" | "meeting" | "teams_meeting";
export type ExecutionHistoryFallbackExportKind = "eml" | "ics" | null;

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
  details: Record<string, unknown>;
};

export type ExecutionHistoryInsert = Omit<ExecutionHistoryRecord, "id" | "userKey" | "executedAt"> & {
  executedAt?: string;
};

function isObject(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function normalizeStringArray(value: unknown) {
  if (!Array.isArray(value)) return [];
  return value.filter((entry): entry is string => typeof entry === "string").map((entry) => entry.trim()).filter(Boolean);
}

function normalizeHistoryStatus(value: unknown): ExecutionHistoryStatus {
  return value === "fallback" || value === "failed" ? value : "success";
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

function normalizeRecord(value: unknown): ExecutionHistoryRecord | null {
  if (!isObject(value) || typeof value.id !== "string") return null;

  return {
    id: value.id,
    userKey: typeof value.user_key === "string" ? value.user_key : "",
    executionGroupId: typeof value.execution_group_id === "string" ? value.execution_group_id : null,
    planName: typeof value.plan_name === "string" ? value.plan_name : "",
    itemType: normalizeHistoryItemType(value.item_type),
    title: typeof value.title === "string" ? value.title : "",
    subject: typeof value.subject === "string" ? value.subject : "",
    status: normalizeHistoryStatus(value.status),
    path: normalizeHistoryPath(value.path),
    recipients: normalizeStringArray(value.recipients),
    attendees: normalizeStringArray(value.attendees),
    executedAt: typeof value.executed_at === "string" ? value.executed_at : new Date(0).toISOString(),
    outlookWebLink: typeof value.outlook_web_link === "string" ? value.outlook_web_link : null,
    teamsJoinLink: typeof value.teams_join_link === "string" ? value.teams_join_link : null,
    fallbackExportKind: normalizeFallbackExportKind(value.fallback_export_kind),
    details: isObject(value.details) ? value.details : {},
  };
}

export async function writeExecutionHistory(entry: ExecutionHistoryInsert) {
  if (typeof window === "undefined" || !isSupabaseConfigured()) return;

  const supabase = getSupabaseBrowserClient();
  const userKey = getLocalUserKey();

  if (!supabase || !userKey) return;

  await supabase.from("execution_history").insert({
    user_key: userKey,
    execution_group_id: entry.executionGroupId,
    plan_name: entry.planName,
    item_type: entry.itemType,
    title: entry.title,
    subject: entry.subject,
    status: entry.status,
    path: entry.path,
    recipients: entry.recipients,
    attendees: entry.attendees,
    executed_at: entry.executedAt ?? new Date().toISOString(),
    outlook_web_link: entry.outlookWebLink,
    teams_join_link: entry.teamsJoinLink,
    fallback_export_kind: entry.fallbackExportKind,
    details: entry.details,
  });
}

export async function listExecutionHistory(limit = 200) {
  if (typeof window === "undefined" || !isSupabaseConfigured()) return [];

  const supabase = getSupabaseBrowserClient();
  const userKey = getLocalUserKey();

  if (!supabase || !userKey) return [];

  const { data, error } = await supabase
    .from("execution_history")
    .select(
      "id,user_key,execution_group_id,plan_name,item_type,title,subject,status,path,recipients,attendees,executed_at,outlook_web_link,teams_join_link,fallback_export_kind,details"
    )
    .eq("user_key", userKey)
    .order("executed_at", { ascending: false })
    .limit(limit);

  if (error || !data) return [];

  return data.map(normalizeRecord).filter((entry): entry is ExecutionHistoryRecord => Boolean(entry));
}
