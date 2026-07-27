import type { ExecutionHistoryRecord } from "./executionHistory";
import type { PersistedPlanTemplate } from "./templateStore";

export type DashboardActionType = "reminder" | "email" | "meeting";

export type DashboardUpcomingItem = {
  id: string;
  type: DashboardActionType;
  title: string;
  sourceName: string | null;
  scheduledAt: string;
  providerLabel: string | null;
  statusLabel: string | null;
};

export type DashboardActivityItem = {
  id: string;
  type: DashboardActionType | "system";
  outcome: string;
  sourceName: string | null;
  occurredAt: string;
  providerLabel: string | null;
};

export type DashboardWorkflowSummary = {
  id: string;
  name: string;
};

export type DashboardMetrics = {
  today: number;
  nextSevenDays: number;
  savedWorkflows: number;
  recentActivity: number;
};

export type DashboardUpcomingGroupKey = "today" | "tomorrow" | "this_week" | "later";

export type DashboardUpcomingGroup = {
  key: DashboardUpcomingGroupKey;
  label: "Today" | "Tomorrow" | "This week" | "Later";
  items: DashboardUpcomingItem[];
};

const unavailableStatuses = new Set<ExecutionHistoryRecord["status"]>([
  "failed",
  "modify_failed",
  "recalled",
  "recall_failed",
  "already_removed",
  "already_canceled",
]);

const attentionStatuses = new Set<ExecutionHistoryRecord["status"]>([
  "failed",
  "modify_failed",
  "recall_failed",
]);

function parseTimestamp(value: string | null | undefined) {
  const timestamp = Date.parse(String(value ?? ""));
  return Number.isFinite(timestamp) ? timestamp : null;
}

function getAction(record: ExecutionHistoryRecord) {
  return typeof record.details.action === "string" ? record.details.action : "";
}

function getScheduledEmailState(record: ExecutionHistoryRecord) {
  return typeof record.details.scheduledEmailState === "string" ? record.details.scheduledEmailState : "";
}

function getProviderLabel(record: ExecutionHistoryRecord) {
  if (record.provider === "outlook") return "Outlook";
  if (record.provider === "gmail") return "Google";
  if (record.provider === "local_export") return "Local export";
  return null;
}

function getActionType(record: ExecutionHistoryRecord): DashboardActionType {
  if (record.itemType === "email") return "email";
  if (record.itemType === "meeting" || record.itemType === "teams_meeting") return "meeting";
  return "reminder";
}

function getFallbackTitle(type: DashboardActionType) {
  if (type === "email") return "Email";
  if (type === "meeting") return "Meeting";
  return "Reminder";
}

function getRecordTitle(record: ExecutionHistoryRecord) {
  const type = getActionType(record);
  return record.subject.trim() || record.title.trim() || getFallbackTitle(type);
}

function getSourceName(record: ExecutionHistoryRecord) {
  return record.planName.trim() || null;
}

function getUpcomingStatusLabel(record: ExecutionHistoryRecord) {
  if (record.status === "modified") return "Updated";
  if (record.itemType === "email" && getAction(record) === "email_scheduled") return "Scheduled";
  return null;
}

function isReliableFutureRecord(record: ExecutionHistoryRecord, now: Date) {
  if (!record.scheduledFor) return false;
  if (unavailableStatuses.has(record.status)) return false;
  if (record.provider === "local_export") return false;
  if (!record.providerObjectId) return false;

  const scheduledTimestamp = parseTimestamp(record.scheduledFor);
  if (scheduledTimestamp == null || scheduledTimestamp <= now.getTime()) return false;

  if (record.itemType === "email") {
    return (
      record.providerObjectType === "message" &&
      getAction(record) === "email_scheduled" &&
      getScheduledEmailState(record) !== "sent"
    );
  }

  if (record.itemType === "meeting" || record.itemType === "teams_meeting" || record.itemType === "reminder") {
    return record.providerObjectType === "event";
  }

  return false;
}

function startOfLocalDay(value: Date) {
  return new Date(value.getFullYear(), value.getMonth(), value.getDate());
}

function endOfLocalDay(value: Date) {
  return new Date(value.getFullYear(), value.getMonth(), value.getDate(), 23, 59, 59, 999);
}

function addLocalDays(value: Date, days: number) {
  return new Date(value.getFullYear(), value.getMonth(), value.getDate() + days);
}

function getGroupKeyForDate(value: Date, now: Date): DashboardUpcomingGroupKey {
  const todayStart = startOfLocalDay(now);
  const todayEnd = endOfLocalDay(now);
  const tomorrowStart = addLocalDays(todayStart, 1);
  const tomorrowEnd = endOfLocalDay(tomorrowStart);
  const endOfWeek = endOfLocalDay(addLocalDays(todayStart, 6 - todayStart.getDay()));
  const timestamp = value.getTime();

  if (timestamp >= todayStart.getTime() && timestamp <= todayEnd.getTime()) return "today";
  if (timestamp >= tomorrowStart.getTime() && timestamp <= tomorrowEnd.getTime()) return "tomorrow";
  if (timestamp > tomorrowEnd.getTime() && timestamp <= endOfWeek.getTime()) return "this_week";
  return "later";
}

function getRecipientCount(record: ExecutionHistoryRecord) {
  const uniqueRecipients = new Set([...record.recipients, ...record.attendees].map((entry) => entry.trim()).filter(Boolean));
  return uniqueRecipients.size;
}

function formatCountedNoun(count: number, singular: string, plural = `${singular}s`) {
  return `${count} ${count === 1 ? singular : plural}`;
}

function getActivityOutcome(record: ExecutionHistoryRecord) {
  const action = getAction(record);
  const scheduledEmailState = getScheduledEmailState(record);
  const type = getActionType(record);

  if (record.status === "failed") return "Export failed";
  if (record.status === "modify_failed") return "Update failed";
  if (record.status === "recall_failed") return "Recall failed";
  if (record.status === "recalled") return "Item recalled";
  if (record.status === "already_removed") return "Item already removed";
  if (record.status === "already_canceled") return "Item already canceled";
  if (record.status === "modified") return "Item updated";

  if (record.path === "fallback") {
    if (record.fallbackExportKind === "eml") return "Email draft exported";
    if (record.fallbackExportKind === "ics") return "Calendar file exported";
    return "Local export created";
  }

  if (type === "email") {
    if (action === "email_sent" || scheduledEmailState === "sent") {
      const recipientCount = getRecipientCount(record);
      return recipientCount > 0 ? `Email sent to ${formatCountedNoun(recipientCount, "recipient")}` : "Email sent";
    }
    if (action === "email_scheduled") return "Email scheduled";
    if (action === "draft_created") return "Email draft created";
    return "Email action completed";
  }

  if (type === "meeting") return "Meeting created";
  return "Reminder scheduled";
}

export function buildDashboardUpcomingItems(records: ExecutionHistoryRecord[], now: Date): DashboardUpcomingItem[] {
  const seen = new Set<string>();

  return records
    .map((record, originalIndex) => ({ record, originalIndex }))
    .filter(({ record }) => isReliableFutureRecord(record, now))
    .map(({ record, originalIndex }) => {
      const scheduledTimestamp = parseTimestamp(record.scheduledFor);
      const dedupeKey = record.providerObjectId
        ? `${record.provider}:${record.providerObjectType ?? "object"}:${record.providerObjectId}`
        : record.id;
      return {
        item: {
          id: record.id,
          type: getActionType(record),
          title: getRecordTitle(record),
          sourceName: getSourceName(record),
          scheduledAt: record.scheduledFor ?? "",
          providerLabel: getProviderLabel(record),
          statusLabel: getUpcomingStatusLabel(record),
        } satisfies DashboardUpcomingItem,
        sortTimestamp: scheduledTimestamp ?? Number.MAX_SAFE_INTEGER,
        originalIndex,
        dedupeKey,
      };
    })
    .filter(({ dedupeKey }) => {
      if (seen.has(dedupeKey)) return false;
      seen.add(dedupeKey);
      return true;
    })
    .sort((left, right) => {
      if (left.sortTimestamp !== right.sortTimestamp) return left.sortTimestamp - right.sortTimestamp;
      return left.originalIndex - right.originalIndex;
    })
    .map(({ item }) => item);
}

export function groupDashboardUpcomingItems(
  items: DashboardUpcomingItem[],
  now: Date
): DashboardUpcomingGroup[] {
  const groups: DashboardUpcomingGroup[] = [
    { key: "today", label: "Today", items: [] },
    { key: "tomorrow", label: "Tomorrow", items: [] },
    { key: "this_week", label: "This week", items: [] },
    { key: "later", label: "Later", items: [] },
  ];
  const groupsByKey = new Map(groups.map((group) => [group.key, group]));

  items.forEach((item) => {
    const timestamp = parseTimestamp(item.scheduledAt);
    if (timestamp == null) return;
    const groupKey = getGroupKeyForDate(new Date(timestamp), now);
    groupsByKey.get(groupKey)?.items.push(item);
  });

  return groups.filter((group) => group.items.length > 0);
}

export function buildDashboardActivityItems(records: ExecutionHistoryRecord[], limit = 8): DashboardActivityItem[] {
  return records
    .map((record, originalIndex) => ({
      record,
      originalIndex,
      timestamp: parseTimestamp(record.executedAt) ?? 0,
    }))
    .filter(({ timestamp }) => timestamp > 0)
    .sort((left, right) => {
      if (left.timestamp !== right.timestamp) return right.timestamp - left.timestamp;
      return left.originalIndex - right.originalIndex;
    })
    .slice(0, limit)
    .map(({ record }) => ({
      id: record.id,
      type: getActionType(record),
      outcome: getActivityOutcome(record),
      sourceName: getSourceName(record),
      occurredAt: record.executedAt,
      providerLabel: getProviderLabel(record),
    }));
}

export function buildDashboardWorkflowSummaries(templates: PersistedPlanTemplate[]): DashboardWorkflowSummary[] {
  return templates
    .filter((template) => !template.id.startsWith("seed:"))
    .map((template) => ({
      id: template.id,
      name: template.name.trim() || "Untitled workflow",
    }));
}

export function buildDashboardMetrics(input: {
  records: ExecutionHistoryRecord[];
  upcomingItems: DashboardUpcomingItem[];
  workflows: DashboardWorkflowSummary[];
  now: Date;
}): DashboardMetrics {
  const todayStart = startOfLocalDay(input.now);
  const todayEnd = endOfLocalDay(input.now);
  const nextSevenEnd = endOfLocalDay(addLocalDays(todayStart, 7));
  const recentActivityStart = addLocalDays(todayStart, -29);

  return {
    today: input.upcomingItems.filter((item) => {
      const timestamp = parseTimestamp(item.scheduledAt);
      return timestamp != null && timestamp >= todayStart.getTime() && timestamp <= todayEnd.getTime();
    }).length,
    nextSevenDays: input.upcomingItems.filter((item) => {
      const timestamp = parseTimestamp(item.scheduledAt);
      return timestamp != null && timestamp > todayEnd.getTime() && timestamp <= nextSevenEnd.getTime();
    }).length,
    savedWorkflows: input.workflows.length,
    recentActivity: input.records.filter((record) => {
      const timestamp = parseTimestamp(record.executedAt);
      return timestamp != null && timestamp >= recentActivityStart.getTime() && timestamp <= input.now.getTime();
    }).length,
  };
}

export function countDashboardAttentionRecords(records: ExecutionHistoryRecord[]) {
  return records.filter((record) => attentionStatuses.has(record.status)).length;
}
