"use client";

import Link from "next/link";
import { useEffect, useMemo, useState } from "react";

import {
  EXECUTION_HISTORY_UPDATED_EVENT,
  listExecutionHistory,
  type ExecutionHistoryRecord,
} from "../../lib/executionHistory";

type PlanExecutionGroup = {
  key: string;
  planName: string;
  executionGroupId: string | null;
  items: ExecutionHistoryRecord[];
  latestExecutedAt: string;
};

type DayExecutionGroup = {
  day: string;
  plans: PlanExecutionGroup[];
};

type HistoryEmailDraftDetails = {
  to: string[];
  cc: string[];
  bcc: string[];
  subject: string;
  body: string;
};

type HistoryMeetingDetails = {
  attendees: string[];
  location: string;
  title: string;
  body: string;
};

function formatDayLabel(value: string) {
  return new Intl.DateTimeFormat("en-US", {
    weekday: "long",
    month: "long",
    day: "numeric",
    year: "numeric",
  }).format(new Date(value));
}

function formatDateTime(value: string | null) {
  if (!value) return "Not available";
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return "Not available";
  return new Intl.DateTimeFormat("en-US", {
    month: "short",
    day: "numeric",
    hour: "numeric",
    minute: "2-digit",
  }).format(parsed);
}

function formatDateOnly(value: string | null) {
  if (!value) return "Not available";
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return "Not available";
  return new Intl.DateTimeFormat("en-US", {
    month: "2-digit",
    day: "2-digit",
    year: "numeric",
  }).format(parsed);
}

function formatTimeOnly(value: string | null) {
  if (!value) return "Not available";
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return "Not available";
  return new Intl.DateTimeFormat("en-US", {
    hour: "numeric",
    minute: "2-digit",
  }).format(parsed);
}

function formatItemTypeLabel(type: ExecutionHistoryRecord["itemType"]) {
  if (type === "teams_meeting") return "Teams meeting";
  return type.charAt(0).toUpperCase() + type.slice(1);
}

function formatTimelineItemType(record: ExecutionHistoryRecord) {
  if (record.itemType === "teams_meeting") return "Teams Meeting";
  if (record.itemType === "meeting") return "Meeting";
  if (record.itemType === "reminder") return "Reminder";
  if (record.itemType === "email") {
    const action = typeof record.details.action === "string" ? record.details.action : "";
    if (action === "draft_created") return "Email Draft";
    if (action === "email_scheduled") return "Scheduled Email";
    if (action === "email_sent") return "Email";
    if (record.path === "fallback" && record.fallbackExportKind === "eml") return "Email Draft";
    return "Email";
  }
  return formatItemTypeLabel(record.itemType);
}

function getViewFullLabel(record: ExecutionHistoryRecord) {
  const label = formatTimelineItemType(record);
  return `View Full ${label}`;
}

function getTypeAccentClasses(label: string) {
  if (label === "Reminder") return "text-blue-600";
  if (label === "Meeting" || label === "Teams Meeting") return "text-violet-600";
  return "text-green-600";
}

function readStringArray(value: unknown) {
  if (!Array.isArray(value)) return [];
  return value.filter((entry): entry is string => typeof entry === "string");
}

function getHistoryBody(record: ExecutionHistoryRecord) {
  return typeof record.details.body === "string" ? record.details.body : "";
}

function getHistoryEmailDraftDetails(record: ExecutionHistoryRecord): HistoryEmailDraftDetails | null {
  const value = record.details.emailDraft;
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const draft = value as Record<string, unknown>;
  return {
    to: readStringArray(draft.to),
    cc: readStringArray(draft.cc),
    bcc: readStringArray(draft.bcc),
    subject: typeof draft.subject === "string" ? draft.subject : "",
    body: typeof draft.body === "string" ? draft.body : "",
  };
}

function getHistoryMeetingDetails(record: ExecutionHistoryRecord): HistoryMeetingDetails | null {
  const value = record.details.meetingDraft;
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const meeting = value as Record<string, unknown>;
  return {
    attendees: readStringArray(meeting.attendees),
    location: typeof meeting.location === "string" ? meeting.location : "",
    title: typeof meeting.title === "string" ? meeting.title : record.title,
    body: typeof meeting.body === "string" ? meeting.body : "",
  };
}

export default function HistoryPage() {
  const [records, setRecords] = useState<ExecutionHistoryRecord[]>([]);
  const [loading, setLoading] = useState(true);
  const [expandedDays, setExpandedDays] = useState<Record<string, boolean>>({});
  const [expandedPlans, setExpandedPlans] = useState<Record<string, boolean>>({});
  const [expandedItems, setExpandedItems] = useState<Record<string, boolean>>({});

  useEffect(() => {
    let cancelled = false;

    async function load() {
      setLoading(true);
      const nextRecords = await listExecutionHistory();
      console.info("[historyPage] loaded records", {
        count: nextRecords.length,
        days: Array.from(new Set(nextRecords.map((record) => record.executedAt.slice(0, 10)))),
        firstRecord: nextRecords[0]
          ? {
              id: nextRecords[0].id,
              userKey: nextRecords[0].userKey,
              executionGroupId: nextRecords[0].executionGroupId,
              planName: nextRecords[0].planName,
              executedAt: nextRecords[0].executedAt,
            }
          : null,
      });
      if (cancelled) return;
      setRecords(nextRecords);
      setExpandedDays((current) => {
        if (Object.keys(current).length > 0) return current;
        return Object.fromEntries(nextRecords.map((record) => [record.executedAt.slice(0, 10), true]));
      });
      setExpandedPlans((current) => {
        if (Object.keys(current).length > 0) return current;
        return {};
      });
      setLoading(false);
    }

    void load();

    function refresh() {
      void load();
    }

    window.addEventListener(EXECUTION_HISTORY_UPDATED_EVENT, refresh as EventListener);

    return () => {
      cancelled = true;
      window.removeEventListener(EXECUTION_HISTORY_UPDATED_EVENT, refresh as EventListener);
    };
  }, []);

  const groupedRecords = useMemo<DayExecutionGroup[]>(() => {
    const dayMap = new Map<string, Map<string, PlanExecutionGroup>>();

    for (const record of records) {
      const day = record.executedAt.slice(0, 10);
      const dayGroup = dayMap.get(day) ?? new Map<string, PlanExecutionGroup>();
      if (!dayMap.has(day)) dayMap.set(day, dayGroup);

      const groupKey = record.executionGroupId || `legacy:${record.planName || "Unnamed plan"}:${day}`;
      const existing = dayGroup.get(groupKey);

      if (existing) {
        existing.items.push(record);
        if (record.executedAt > existing.latestExecutedAt) {
          existing.latestExecutedAt = record.executedAt;
        }
        continue;
      }

      dayGroup.set(groupKey, {
        key: groupKey,
        planName: record.planName || record.subject || record.title || "Unnamed plan",
        executionGroupId: record.executionGroupId,
        items: [record],
        latestExecutedAt: record.executedAt,
      });
    }

    return Array.from(dayMap.entries()).map(([day, planMap]) => ({
      day,
      plans: Array.from(planMap.values()).sort((left, right) => right.latestExecutedAt.localeCompare(left.latestExecutedAt)),
    }));
  }, [records]);

  return (
    <div className="space-y-8 text-gray-900">
      <section>
        <div>
          <h1 className="text-3xl font-bold text-gray-900">History</h1>
        </div>
      </section>

      <section className="rounded-2xl border bg-white shadow-sm">
        <div className="border-b px-6 py-4">
          <h2 className="text-lg font-semibold text-gray-900">Execution History</h2>
        </div>
        <div className="p-6">
          {loading ? (
            <p className="text-sm text-gray-600">Loading history…</p>
          ) : groupedRecords.length === 0 ? (
            <div className="space-y-3 rounded-2xl border border-dashed bg-gray-50 p-6 text-sm text-gray-600">
              <p>No execution history yet.</p>
              <p>History now persists locally as well as to Supabase when configured, so newly exported plans should show up here right away.</p>
              <Link href="/plans" className="inline-flex rounded-lg bg-blue-600 px-4 py-2 font-medium text-white hover:bg-blue-700">
                Go to Plans
              </Link>
            </div>
          ) : (
            <div className="space-y-6">
              {groupedRecords.map((dayGroup) => {
                const isDayExpanded = expandedDays[dayGroup.day] ?? true;

                return (
                  <section key={dayGroup.day} className="rounded-2xl border border-gray-200">
                    <button
                      type="button"
                      onClick={() => setExpandedDays((current) => ({ ...current, [dayGroup.day]: !isDayExpanded }))}
                      className="flex w-full items-center justify-between gap-4 px-5 py-4 text-left hover:bg-gray-50"
                    >
                      <div className="text-lg font-semibold text-gray-900">{formatDayLabel(dayGroup.day)}</div>
                      <div className="text-sm font-medium text-gray-500">{isDayExpanded ? "Collapse" : "Expand"}</div>
                    </button>

                    {isDayExpanded ? (
                      <div className="space-y-4 border-t bg-gray-50/60 p-4">
                        {dayGroup.plans.map((planGroup) => {
                          const isPlanExpanded = expandedPlans[planGroup.key] ?? false;
                          return (
                            <section key={planGroup.key} className="rounded-2xl border bg-white shadow-sm">
                              <button
                                type="button"
                                onClick={() => setExpandedPlans((current) => ({ ...current, [planGroup.key]: !isPlanExpanded }))}
                                className="flex w-full items-center justify-between gap-4 px-5 py-4 text-left hover:bg-gray-50"
                              >
                                <div className="min-w-0">
                                  <div className="text-lg font-semibold text-gray-900">Event Name: {planGroup.planName}</div>
                                  <div className="mt-1 text-sm text-gray-600">
                                    {planGroup.items.length} created item{planGroup.items.length === 1 ? "" : "s"}
                                  </div>
                                  <div className="mt-1 text-sm text-gray-600">Created at: {formatDateTime(planGroup.latestExecutedAt)}</div>
                                </div>
                                <div className="text-sm font-medium text-gray-500">{isPlanExpanded ? "Collapse" : "Expand"}</div>
                              </button>

                              {isPlanExpanded ? (
                                <div className="border-t p-5">
                                  <div className="hidden items-center gap-x-3 border-b pb-3 text-xs font-semibold uppercase tracking-wide text-gray-600 md:grid md:grid-cols-[140px_minmax(0,1.25fr)_190px_130px_170px]">
                                    <div>Type</div>
                                    <div>Title</div>
                                    <div className="text-center">Date Disseminated</div>
                                    <div className="text-center">Time</div>
                                    <div className="text-right">Action</div>
                                  </div>
                                  <div className="divide-y">
                                    {planGroup.items.map((item) => {
                                      const itemTypeLabel = formatTimelineItemType(item);
                                      const itemDateTime = item.scheduledFor || item.executedAt;
                                      const isItemExpanded = expandedItems[item.id] ?? false;
                                      const reminderBody = getHistoryBody(item);
                                      const emailDraft = getHistoryEmailDraftDetails(item);
                                      const meetingDetails = getHistoryMeetingDetails(item);
                                      return (
                                        <article key={item.id} className="py-4">
                                          <div className="grid items-center gap-3 md:grid-cols-[140px_minmax(0,1.25fr)_190px_130px_170px]">
                                            <div className={`text-sm font-medium ${getTypeAccentClasses(itemTypeLabel)}`}>{itemTypeLabel}</div>
                                            <div className="min-w-0 truncate text-base text-gray-900">{item.subject || item.title || "Untitled item"}</div>
                                            <div className="rounded-xl border bg-white px-4 py-3 text-center text-sm text-gray-900">
                                              {formatDateOnly(itemDateTime)}
                                            </div>
                                            <div className="rounded-xl border bg-white px-4 py-3 text-center text-sm text-gray-900">
                                              {formatTimeOnly(itemDateTime)}
                                            </div>
                                            <div className="flex justify-end">
                                              <button
                                                type="button"
                                                onClick={() => setExpandedItems((current) => ({ ...current, [item.id]: !isItemExpanded }))}
                                                className="rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900 hover:bg-gray-50"
                                              >
                                                {isItemExpanded ? `Hide ${itemTypeLabel}` : getViewFullLabel(item)}
                                              </button>
                                            </div>
                                          </div>
                                          {isItemExpanded ? (
                                            <div className="mt-4 space-y-4 rounded-xl border bg-gray-50 p-4">
                                              <div className="text-sm text-gray-700">Event Name: {planGroup.planName}</div>
                                              {item.itemType === "reminder" ? (
                                                <div className="bg-blue-50 px-4 py-3">
                                                  <div className="grid grid-cols-1 gap-3">
                                                    <div>
                                                      <label className="mb-1 block text-sm font-medium text-blue-950">Reminder Body</label>
                                                      <textarea
                                                        rows={5}
                                                        readOnly
                                                        className="w-full rounded-lg border border-blue-200 bg-white px-3 py-2 text-sm text-gray-900"
                                                        value={reminderBody}
                                                      />
                                                    </div>
                                                  </div>
                                                </div>
                                              ) : null}
                                              {(item.itemType === "meeting" || item.itemType === "teams_meeting") && meetingDetails ? (
                                                <div className="bg-violet-50 px-4 py-3">
                                                  <div className="grid grid-cols-1 gap-3 md:grid-cols-2">
                                                    <div className="md:col-span-2">
                                                      <label className="mb-1 block text-sm font-medium text-violet-950">To</label>
                                                      <textarea
                                                        rows={2}
                                                        readOnly
                                                        className="w-full rounded-lg border border-violet-200 bg-white px-3 py-2 text-sm text-gray-900"
                                                        value={meetingDetails.attendees.join(", ")}
                                                      />
                                                    </div>
                                                    <div className="md:col-span-2">
                                                      <label className="mb-1 block text-sm font-medium text-violet-950">Subject</label>
                                                      <input
                                                        readOnly
                                                        className="w-full rounded-lg border border-violet-200 bg-white px-3 py-2 text-sm text-gray-900"
                                                        value={meetingDetails.title}
                                                      />
                                                    </div>
                                                    <div className="md:col-span-2">
                                                      <label className="mb-1 block text-sm font-medium text-violet-950">Message</label>
                                                      <textarea
                                                        rows={6}
                                                        readOnly
                                                        className="w-full rounded-lg border border-violet-200 bg-white px-3 py-2 text-sm text-gray-900"
                                                        value={meetingDetails.body}
                                                      />
                                                    </div>
                                                  </div>
                                                </div>
                                              ) : null}
                                              {item.itemType === "email" && emailDraft ? (
                                                <div className="bg-amber-50 px-4 py-3">
                                                  <div className="grid grid-cols-1 gap-3">
                                                    <div>
                                                      <label className="mb-1 block text-sm font-medium text-amber-950">To</label>
                                                      <textarea
                                                        rows={2}
                                                        readOnly
                                                        className="w-full rounded-lg border border-amber-200 bg-white px-3 py-2 text-sm text-gray-900"
                                                        value={emailDraft.to.join(", ")}
                                                      />
                                                    </div>
                                                    <div>
                                                      <label className="mb-1 block text-sm font-medium text-amber-950">Cc</label>
                                                      <textarea
                                                        rows={2}
                                                        readOnly
                                                        className="w-full rounded-lg border border-amber-200 bg-white px-3 py-2 text-sm text-gray-900"
                                                        value={emailDraft.cc.join(", ")}
                                                      />
                                                    </div>
                                                    <div>
                                                      <label className="mb-1 block text-sm font-medium text-amber-950">Bcc</label>
                                                      <textarea
                                                        rows={2}
                                                        readOnly
                                                        className="w-full rounded-lg border border-amber-200 bg-white px-3 py-2 text-sm text-gray-900"
                                                        value={emailDraft.bcc.join(", ")}
                                                      />
                                                    </div>
                                                    <div>
                                                      <label className="mb-1 block text-sm font-medium text-amber-950">Subject</label>
                                                      <input
                                                        readOnly
                                                        className="w-full rounded-lg border border-amber-200 bg-white px-3 py-2 text-sm text-gray-900"
                                                        value={emailDraft.subject}
                                                      />
                                                    </div>
                                                    <div>
                                                      <label className="mb-1 block text-sm font-medium text-amber-950">Message</label>
                                                      <textarea
                                                        rows={8}
                                                        readOnly
                                                        className="w-full rounded-lg border border-amber-200 bg-white px-3 py-2 text-sm text-gray-900"
                                                        value={emailDraft.body}
                                                      />
                                                    </div>
                                                  </div>
                                                </div>
                                              ) : null}
                                            </div>
                                          ) : null}
                                        </article>
                                      );
                                    })}
                                  </div>
                                </div>
                              ) : null}
                            </section>
                          );
                        })}
                      </div>
                    ) : null}
                  </section>
                );
              })}
            </div>
          )}
        </div>
      </section>
    </div>
  );
}
