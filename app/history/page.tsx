"use client";

import Link from "next/link";
import { useEffect, useMemo, useState } from "react";

import {
  EXECUTION_HISTORY_UPDATED_EVENT,
  getExecutionHistoryModifyState,
  getExecutionHistoryRecallState,
  listExecutionHistory,
  updateExecutionHistoryRecord,
  type ExecutionHistoryRecord,
} from "../../lib/executionHistory";
import {
  deleteOutlookCalendarEvent,
  deleteOutlookMessage,
  createOutlookCalendarEvent,
  replaceOutlookScheduledEmail,
  updateOutlookCalendarEvent,
  updateOutlookMessageDraft,
  type OutlookRecallResult,
} from "../../lib/outlookClient";
import { createPlan, type TemplateItem } from "../../lib/planEngine";
import { addDaysISO } from "../../lib/dateUtils";
import { buildAnchorMap, classifyPlanRow, normalizeAnchorKey, resolvePlanAnchors } from "../../lib/plansRuntime";
import type { PlanDateBasis, PlanItem, PlanRowType, PlanType, WeekendRule } from "../../types/plan";

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

type HistoryEditDraft = {
  subject: string;
  body: string;
  to: string;
  cc: string;
  bcc: string;
  date: string;
  time: string;
};

type SnapshotAnchorValue = {
  key: string;
  value: string;
  displayValue?: string;
  locked?: boolean;
};

type SnapshotRowDefinition = {
  id: string;
  title: string;
  body?: string;
  offsetDays: number;
  dateBasis?: PlanDateBasis;
  rowType?: PlanRowType;
  reminderTime?: string;
  emailDraft?: {
    to?: string[];
    cc?: string[];
    bcc?: string[];
    subject?: string;
    body?: string;
  } | null;
  durationDraft?: {
    durationMinutes?: number;
    useCustomEnd?: boolean;
    endDate?: string;
    endTime?: string;
    isAllDay?: boolean;
  } | null;
  meetingDraft?: {
    attendees?: string[];
    location?: string;
    durationMinutes?: number;
    useCustomEnd?: boolean;
    endDate?: string;
    endTime?: string;
    isAllDay?: boolean;
    teamsMeeting?: boolean;
  } | null;
};

type ExecutionPlanSnapshot = {
  templateBaseType: PlanType;
  templateMode?: string;
  templateId?: string | null;
  templateName?: string;
  eventName?: string;
  anchorDate: string;
  noEventDate?: boolean;
  weekendRule: WeekendRule;
  anchorValues: SnapshotAnchorValue[];
  originalRowDefinitions: SnapshotRowDefinition[];
};

type PlanReschedulePreviewAction = "Update" | "Replace" | "Unchanged" | "Locked" | "Unsupported";

type PlanReschedulePreviewItem = {
  record: ExecutionHistoryRecord;
  nextItem: PlanItem | null;
  action: PlanReschedulePreviewAction;
  reason: string | null;
  oldDateTime: string | null;
  newDateTime: string | null;
  isOverridden: boolean;
};

function formatDayLabel(value: string) {
  return new Intl.DateTimeFormat("en-US", {
    weekday: "long",
    month: "long",
    day: "numeric",
    year: "numeric",
  }).format(new Date(value));
}

function getLocalDayKey(value: string) {
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return value.slice(0, 10);
  return `${parsed.getFullYear()}-${String(parsed.getMonth() + 1).padStart(2, "0")}-${String(parsed.getDate()).padStart(2, "0")}`;
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

function formatDateInputValue(value: string | null) {
  if (!value) return "";
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return "";
  return `${parsed.getFullYear()}-${String(parsed.getMonth() + 1).padStart(2, "0")}-${String(parsed.getDate()).padStart(2, "0")}`;
}

function formatTimeInputValue(value: string | null) {
  if (!value) return "";
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return "";
  return `${String(parsed.getHours()).padStart(2, "0")}:${String(parsed.getMinutes()).padStart(2, "0")}`;
}

function formatItemTypeLabel(type: ExecutionHistoryRecord["itemType"]) {
  if (type === "teams_meeting") return "Meeting";
  return type.charAt(0).toUpperCase() + type.slice(1);
}

function formatTimelineItemType(record: ExecutionHistoryRecord) {
  if (record.itemType === "teams_meeting") return "Meeting";
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

function getItemTypeDisplayLabel(record: ExecutionHistoryRecord) {
  const primaryLabel = formatTimelineItemType(record);
  if (record.itemType === "teams_meeting") {
    return `${primaryLabel} • Teams`;
  }
  return primaryLabel;
}

function getViewFullLabel(record: ExecutionHistoryRecord) {
  const label = formatTimelineItemType(record);
  return `View Full ${label}`;
}

function getHistoryAction(record: ExecutionHistoryRecord) {
  return typeof record.details.action === "string" ? record.details.action : "";
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

function isObject(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function readString(value: unknown, fallback = "") {
  return typeof value === "string" ? value : fallback;
}

function joinAddresses(value: string[]) {
  return value.join(", ");
}

function splitAddresses(value: string) {
  return value
    .split(",")
    .map((entry) => entry.trim())
    .filter(Boolean);
}

function normalizeReminderTimeInput(value: string) {
  return value.trim();
}

function normalizeEmailDraftValue(value: SnapshotRowDefinition["emailDraft"]) {
  if (!value) return undefined;
  return {
    to: readStringArray(value.to),
    cc: readStringArray(value.cc),
    bcc: readStringArray(value.bcc),
    subject: readString(value.subject),
    body: readString(value.body),
  };
}

function normalizeMeetingDraftValue(value: SnapshotRowDefinition["meetingDraft"]) {
  if (!value) return undefined;
  return {
    attendees: readStringArray(value.attendees),
    location: readString(value.location),
    durationMinutes: typeof value.durationMinutes === "number" && value.durationMinutes > 0 ? value.durationMinutes : 30,
    useCustomEnd: Boolean(value.useCustomEnd),
    endDate: readString(value.endDate),
    endTime: readString(value.endTime),
    isAllDay: Boolean(value.isAllDay),
    teamsMeeting: Boolean(value.teamsMeeting),
  };
}

function normalizeDurationDraftValue(value: SnapshotRowDefinition["durationDraft"]) {
  if (!value) return undefined;
  return {
    durationMinutes: typeof value.durationMinutes === "number" && value.durationMinutes > 0 ? value.durationMinutes : 30,
    useCustomEnd: Boolean(value.useCustomEnd),
    endDate: readString(value.endDate),
    endTime: readString(value.endTime),
    isAllDay: Boolean(value.isAllDay),
  };
}

function parseTimeInput(value: string) {
  const normalized = value.trim().toUpperCase();
  if (!normalized) return null;

  const twelveHourMatch = normalized.match(/^(\d{1,2}):(\d{2})\s*([AP]M)$/);
  if (twelveHourMatch) {
    const hours = Number(twelveHourMatch[1]);
    const minutes = Number(twelveHourMatch[2]);
    const meridiem = twelveHourMatch[3];
    if (hours < 1 || hours > 12 || minutes < 0 || minutes > 59) return null;
    const normalizedHours = meridiem === "AM" ? (hours === 12 ? 0 : hours) : hours === 12 ? 12 : hours + 12;
    return `${String(normalizedHours).padStart(2, "0")}:${String(minutes).padStart(2, "0")}`;
  }

  const twentyFourHourMatch = normalized.match(/^([01]?\d|2[0-3]):([0-5]\d)$/);
  if (twentyFourHourMatch) {
    return `${String(Number(twentyFourHourMatch[1])).padStart(2, "0")}:${twentyFourHourMatch[2]}`;
  }

  return null;
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

function createEditDraft(record: ExecutionHistoryRecord): HistoryEditDraft {
  const emailDraft = getHistoryEmailDraftDetails(record);
  const meetingDetails = getHistoryMeetingDetails(record);
  return {
    subject:
      record.itemType === "email"
        ? emailDraft?.subject ?? record.subject
        : record.itemType === "meeting" || record.itemType === "teams_meeting"
          ? meetingDetails?.title ?? record.subject
          : record.subject,
    body:
      record.itemType === "email"
        ? emailDraft?.body ?? ""
        : record.itemType === "meeting" || record.itemType === "teams_meeting"
          ? meetingDetails?.body ?? ""
          : getHistoryBody(record),
    to:
      record.itemType === "email"
        ? joinAddresses(emailDraft?.to ?? [])
        : joinAddresses(meetingDetails?.attendees ?? []),
    cc: joinAddresses(emailDraft?.cc ?? []),
    bcc: joinAddresses(emailDraft?.bcc ?? []),
    date: formatDateInputValue(record.scheduledFor || record.executedAt),
    time: record.isAllDay ? "" : formatTimeInputValue(record.scheduledFor || record.executedAt),
  };
}

function buildLocalIso(date: string, time: string) {
  const normalizedTime = time.trim() ? `${time.trim()}:00` : "00:00:00";
  return `${date}T${normalizedTime}`;
}

function addMinutesToIso(iso: string, minutes: number) {
  const parsed = new Date(iso);
  if (Number.isNaN(parsed.getTime())) return iso;
  parsed.setMinutes(parsed.getMinutes() + minutes);
  return `${parsed.getFullYear()}-${String(parsed.getMonth() + 1).padStart(2, "0")}-${String(parsed.getDate()).padStart(2, "0")}T${String(parsed.getHours()).padStart(2, "0")}:${String(parsed.getMinutes()).padStart(2, "0")}:00`;
}

function getDurationMinutes(record: ExecutionHistoryRecord) {
  if (!record.scheduledFor || !record.endsAt) {
    return record.isAllDay ? 24 * 60 : 30;
  }
  const start = new Date(record.scheduledFor);
  const end = new Date(record.endsAt);
  if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime())) {
    return record.isAllDay ? 24 * 60 : 30;
  }
  const diffMinutes = Math.max(1, Math.round((end.getTime() - start.getTime()) / 60000));
  return diffMinutes;
}

function getOverrideChangedFields(record: ExecutionHistoryRecord, draft: HistoryEditDraft) {
  const fields = new Set<string>();
  const emailDraft = getHistoryEmailDraftDetails(record);
  const meetingDetails = getHistoryMeetingDetails(record);

  if ((draft.subject.trim() || "") !== (record.itemType === "email" ? emailDraft?.subject ?? record.subject : record.itemType === "meeting" || record.itemType === "teams_meeting" ? meetingDetails?.title ?? record.subject : record.subject)) {
    fields.add("subject");
  }
  if ((draft.body ?? "") !== (record.itemType === "email" ? emailDraft?.body ?? "" : record.itemType === "meeting" || record.itemType === "teams_meeting" ? meetingDetails?.body ?? "" : getHistoryBody(record))) {
    fields.add("body");
  }
  if (draft.date !== formatDateInputValue(record.scheduledFor || record.executedAt)) {
    fields.add("date");
  }
  if (!record.isAllDay && draft.time !== formatTimeInputValue(record.scheduledFor || record.executedAt)) {
    fields.add("time");
  }
  if (record.itemType === "email") {
    if (draft.to !== joinAddresses(emailDraft?.to ?? [])) fields.add("to");
    if (draft.cc !== joinAddresses(emailDraft?.cc ?? [])) fields.add("cc");
    if (draft.bcc !== joinAddresses(emailDraft?.bcc ?? [])) fields.add("bcc");
  }
  if (record.itemType === "meeting" || record.itemType === "teams_meeting") {
    if (draft.to !== joinAddresses(meetingDetails?.attendees ?? [])) fields.add("attendees");
  }
  if (record.itemType === "reminder" && (draft.subject.trim() || "") !== (record.subject || record.title)) {
    fields.add("text");
  }

  return Array.from(fields);
}

function formatPreviewDateTime(value: string | null) {
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

function getExecutionPlanSnapshot(record: ExecutionHistoryRecord): ExecutionPlanSnapshot | null {
  const rawSnapshot = record.details.executionPlanSnapshot;
  if (!isObject(rawSnapshot)) return null;
  const anchorValues = Array.isArray(rawSnapshot.anchorValues)
    ? rawSnapshot.anchorValues
        .filter((entry): entry is Record<string, unknown> => isObject(entry))
        .map((entry) => ({
          key: readString(entry.key),
          value: readString(entry.value),
          displayValue: typeof entry.displayValue === "string" ? entry.displayValue : undefined,
          locked: Boolean(entry.locked),
        }))
        .filter((entry) => entry.key)
    : [];
  const originalRowDefinitions = Array.isArray(rawSnapshot.originalRowDefinitions)
    ? rawSnapshot.originalRowDefinitions
        .filter((entry): entry is Record<string, unknown> => isObject(entry))
        .map((entry) => {
          const dateBasis: PlanDateBasis = entry.dateBasis === "today" ? "today" : "event";
          const rowType: PlanRowType =
            entry.rowType === "email" || entry.rowType === "calendar_event" ? entry.rowType : "reminder";
          return {
            id: readString(entry.id),
            title: readString(entry.title),
            body: typeof entry.body === "string" ? entry.body : "",
            offsetDays: typeof entry.offsetDays === "number" ? entry.offsetDays : 0,
            dateBasis,
            rowType,
            reminderTime: typeof entry.reminderTime === "string" ? entry.reminderTime : "",
            emailDraft: isObject(entry.emailDraft) ? (entry.emailDraft as SnapshotRowDefinition["emailDraft"]) : null,
            durationDraft: isObject(entry.durationDraft) ? (entry.durationDraft as SnapshotRowDefinition["durationDraft"]) : null,
            meetingDraft: isObject(entry.meetingDraft) ? (entry.meetingDraft as SnapshotRowDefinition["meetingDraft"]) : null,
          };
        })
        .filter((entry) => entry.id)
    : [];

  if (!rawSnapshot.templateBaseType || !rawSnapshot.anchorDate || !rawSnapshot.weekendRule || originalRowDefinitions.length === 0) {
    return null;
  }

  return {
    templateBaseType: rawSnapshot.templateBaseType as PlanType,
    templateMode: typeof rawSnapshot.templateMode === "string" ? rawSnapshot.templateMode : undefined,
    templateId: typeof rawSnapshot.templateId === "string" ? rawSnapshot.templateId : null,
    templateName: typeof rawSnapshot.templateName === "string" ? rawSnapshot.templateName : undefined,
    eventName: typeof rawSnapshot.eventName === "string" ? rawSnapshot.eventName : undefined,
    anchorDate: readString(rawSnapshot.anchorDate),
    noEventDate: Boolean(rawSnapshot.noEventDate),
    weekendRule: rawSnapshot.weekendRule === "none" ? "none" : "prior_business_day",
    anchorValues,
    originalRowDefinitions,
  };
}

function getOverrideState(record: ExecutionHistoryRecord) {
  const rawOverride = record.details.overrideTracking;
  if (!isObject(rawOverride)) {
    return { isOverridden: false };
  }
  return {
    isOverridden: Boolean(rawOverride.isOverridden),
  };
}

function diffDays(fromDate: string, toDate: string) {
  const [fromY, fromM, fromD] = fromDate.split("-").map(Number);
  const [toY, toM, toD] = toDate.split("-").map(Number);
  const from = new Date(fromY ?? 2000, (fromM ?? 1) - 1, fromD ?? 1);
  const to = new Date(toY ?? 2000, (toM ?? 1) - 1, toD ?? 1);
  return Math.round((to.getTime() - from.getTime()) / (24 * 60 * 60 * 1000));
}

function buildUpdatedSnapshotAnchors(snapshot: ExecutionPlanSnapshot, nextEventDate: string) {
  const deltaDays = diffDays(snapshot.anchorDate, nextEventDate);
  return snapshot.anchorValues.map((anchor) => {
    const normalizedKey = normalizeAnchorKey(anchor.key);
    if (normalizedKey === normalizeAnchorKey("Event Date")) {
      return { ...anchor, value: nextEventDate };
    }
    if (snapshot.templateBaseType === "press_release" && normalizedKey === normalizeAnchorKey("Dissemination Date")) {
      return { ...anchor, value: nextEventDate };
    }
    if (snapshot.templateBaseType === "conference" && normalizedKey === normalizeAnchorKey("Conference Start Date")) {
      return { ...anchor, value: nextEventDate };
    }
    if (snapshot.templateBaseType === "conference" && normalizedKey === normalizeAnchorKey("Conference End Date") && anchor.value) {
      return { ...anchor, value: addDaysISO(anchor.value, deltaDays) };
    }
    if (snapshot.templateBaseType === "earnings" && normalizedKey === normalizeAnchorKey("Earnings Call Date")) {
      return { ...anchor, value: nextEventDate };
    }
    return anchor;
  });
}

function buildTemplateItemsFromSnapshot(snapshot: ExecutionPlanSnapshot): TemplateItem[] {
  return snapshot.originalRowDefinitions.map((row) => ({
    id: row.id,
    title: row.title,
    body: row.body || undefined,
    offsetDays: row.offsetDays,
    dateBasis: row.dateBasis ?? "event",
    rowType: row.rowType ?? "reminder",
    reminderTime: row.reminderTime ? normalizeReminderTimeInput(row.reminderTime) : undefined,
    emailDraft: normalizeEmailDraftValue(row.emailDraft),
    durationDraft: normalizeDurationDraftValue(row.durationDraft),
    meetingDraft: normalizeMeetingDraftValue(row.meetingDraft),
  }));
}

function buildRescheduledPlan(snapshot: ExecutionPlanSnapshot, nextEventDate: string) {
  const templateItems = buildTemplateItemsFromSnapshot(snapshot);
  const plan = createPlan({
    name: snapshot.eventName || snapshot.templateName || "Untitled plan",
    type: snapshot.templateBaseType,
    anchorDate: nextEventDate,
    weekendRule: snapshot.weekendRule,
    template: templateItems,
  });
  const anchorMap = buildAnchorMap(buildUpdatedSnapshotAnchors(snapshot, nextEventDate));
  return resolvePlanAnchors(plan, anchorMap);
}

function getComputedPlanItemTiming(item: PlanItem) {
  const rowKind = classifyPlanRow(item);
  const isAllDay = Boolean(item.meetingDraft?.isAllDay || item.durationDraft?.isAllDay);
  const resolvedTime = parseTimeInput(item.reminderTime ?? "");

  if ((rowKind === "meeting" || rowKind === "reminder") && (isAllDay || !resolvedTime)) {
    return {
      scheduledFor: `${item.customDueDate ?? item.dueDate}T00:00:00`,
      endsAt: `${addDaysISO(item.customDueDate ?? item.dueDate, 1)}T00:00:00`,
      isAllDay: true,
    };
  }

  const baseDate = item.customDueDate ?? item.dueDate;
  const baseTime = resolvedTime ?? "09:00";
  const scheduledFor = `${baseDate}T${baseTime}:00`;

  if (rowKind === "meeting" || rowKind === "reminder") {
    if ((item.meetingDraft?.useCustomEnd || item.durationDraft?.useCustomEnd) && (item.meetingDraft?.endDate || item.durationDraft?.endDate) && (item.meetingDraft?.endTime || item.durationDraft?.endTime)) {
      const endDate = item.meetingDraft?.endDate || item.durationDraft?.endDate || baseDate;
      const endTime = parseTimeInput(item.meetingDraft?.endTime || item.durationDraft?.endTime || "") || "09:30";
      return {
        scheduledFor,
        endsAt: `${endDate}T${endTime}:00`,
        isAllDay: false,
      };
    }

    const durationMinutes = item.meetingDraft?.durationMinutes ?? item.durationDraft?.durationMinutes ?? 30;
    return {
      scheduledFor,
      endsAt: addMinutesToIso(scheduledFor, durationMinutes),
      isAllDay: false,
    };
  }

  return {
    scheduledFor,
    endsAt: null,
    isAllDay: false,
  };
}

function getPlanModifyPreview(planGroup: PlanExecutionGroup, nextEventDate: string): { snapshot: ExecutionPlanSnapshot | null; items: PlanReschedulePreviewItem[] } {
  const firstRecord = planGroup.items[0];
  const snapshot = firstRecord ? getExecutionPlanSnapshot(firstRecord) : null;
  if (!snapshot || !nextEventDate) {
    return { snapshot, items: [] };
  }

  const rescheduledPlan = buildRescheduledPlan(snapshot, nextEventDate);
  const nextItemsByRowId = new Map(rescheduledPlan.items.map((item) => [item.id, item]));

  const items = planGroup.items.map((record) => {
    const sourceRowId = typeof record.details.sourceRowId === "string" ? record.details.sourceRowId : record.id;
    const nextItem = nextItemsByRowId.get(sourceRowId) ?? null;
    const modifyState = getExecutionHistoryModifyState(record);
    const overrideState = getOverrideState(record);
    const oldDateTime = record.scheduledFor || record.executedAt;
    const nextTiming = nextItem ? getComputedPlanItemTiming(nextItem) : null;
    const newDateTime = nextTiming?.scheduledFor ?? oldDateTime;
    const actionName = getHistoryAction(record);

    if (overrideState.isOverridden) {
      return {
        record,
        nextItem,
        action: "Locked" as const,
        reason: "Skipped because this item was edited separately.",
        oldDateTime,
        newDateTime,
        isOverridden: true,
      };
    }

    if (actionName === "email_sent") {
      return {
        record,
        nextItem,
        action: "Locked" as const,
        reason: "Already sent.",
        oldDateTime,
        newDateTime,
        isOverridden: false,
      };
    }

    if (!nextItem) {
      return {
        record,
        nextItem,
        action: "Unsupported" as const,
        reason: "Could not match this item back to the original plan row.",
        oldDateTime,
        newDateTime: oldDateTime,
        isOverridden: false,
      };
    }

    const contentChanged =
      (record.subject || record.title) !== (nextItem.customTitle || nextItem.title) ||
      getHistoryBody(record) !== (nextItem.body ?? "");
    const timingChanged = oldDateTime !== newDateTime;
    const recipientsChanged =
      record.itemType === "email"
        ? joinAddresses(record.recipients) !== joinAddresses([...(nextItem.emailDraft?.to ?? []), ...(nextItem.emailDraft?.cc ?? []), ...(nextItem.emailDraft?.bcc ?? [])])
        : record.itemType === "meeting" || record.itemType === "teams_meeting"
          ? joinAddresses(record.attendees) !== joinAddresses(nextItem.meetingDraft?.attendees ?? [])
          : false;

    if (!timingChanged && !contentChanged && !recipientsChanged) {
      return {
        record,
        nextItem,
        action: "Unchanged" as const,
        reason: null,
        oldDateTime,
        newDateTime,
        isOverridden: false,
      };
    }

    if (record.itemType === "reminder") {
      return {
        record,
        nextItem,
        action: "Replace" as const,
        reason: "Will recreate this reminder on the new timeline.",
        oldDateTime,
        newDateTime,
        isOverridden: false,
      };
    }

    if (record.itemType === "email" && actionName === "email_scheduled") {
      return {
        record,
        nextItem,
        action: "Replace" as const,
        reason: "Will replace the unsent scheduled email.",
        oldDateTime,
        newDateTime,
        isOverridden: false,
      };
    }

    if (modifyState.canModify && modifyState.modifyImplemented) {
      return {
        record,
        nextItem,
        action: "Update" as const,
        reason: null,
        oldDateTime,
        newDateTime,
        isOverridden: false,
      };
    }

    return {
      record,
      nextItem,
      action: "Unsupported" as const,
      reason: modifyState.modifyReason || "This item can't be changed from History.",
      oldDateTime,
      newDateTime,
      isOverridden: false,
    };
  });

  return { snapshot, items };
}

export default function HistoryPage() {
  const [records, setRecords] = useState<ExecutionHistoryRecord[]>([]);
  const [loading, setLoading] = useState(true);
  const [expandedDays, setExpandedDays] = useState<Record<string, boolean>>({});
  const [expandedPlans, setExpandedPlans] = useState<Record<string, boolean>>({});
  const [expandedItems, setExpandedItems] = useState<Record<string, boolean>>({});
  const [editingItems, setEditingItems] = useState<Record<string, boolean>>({});
  const [editDrafts, setEditDrafts] = useState<Record<string, HistoryEditDraft>>({});
  const [modifyingPlans, setModifyingPlans] = useState<Record<string, boolean>>({});
  const [planModifyDates, setPlanModifyDates] = useState<Record<string, string>>({});
  const [pendingPlanRecalls, setPendingPlanRecalls] = useState<Record<string, boolean>>({});
  const [pendingItemRecalls, setPendingItemRecalls] = useState<Record<string, boolean>>({});
  const [pendingItemModifies, setPendingItemModifies] = useState<Record<string, boolean>>({});
  const [pendingPlanModifies, setPendingPlanModifies] = useState<Record<string, boolean>>({});
  const [planMessages, setPlanMessages] = useState<Record<string, { tone: "success" | "warning" | "error"; text: string }>>({});
  const [itemMessages, setItemMessages] = useState<Record<string, { tone: "success" | "error"; text: string }>>({});

  useEffect(() => {
    let cancelled = false;

    async function load() {
      setLoading(true);
      const nextRecords = await listExecutionHistory();
      console.info("[historyPage] loaded records", {
        count: nextRecords.length,
        days: Array.from(new Set(nextRecords.map((record) => getLocalDayKey(record.executedAt)))),
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
        return Object.fromEntries(nextRecords.map((record) => [getLocalDayKey(record.executedAt), true]));
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
      const day = getLocalDayKey(record.executedAt);
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

  function toggleItemEdit(record: ExecutionHistoryRecord, nextOpen?: boolean) {
    const shouldOpen = nextOpen ?? !(editingItems[record.id] ?? false);
    setExpandedItems((current) => ({ ...current, [record.id]: shouldOpen || current[record.id] || false }));
    setEditingItems((current) => ({ ...current, [record.id]: shouldOpen }));
    if (shouldOpen) {
      setEditDrafts((current) => ({
        ...current,
        [record.id]: current[record.id] ?? createEditDraft(record),
      }));
    }
  }

  function handleEditDraftChange(recordId: string, field: keyof HistoryEditDraft, value: string) {
    setEditDrafts((current) => ({
      ...current,
      [recordId]: {
        ...(current[recordId] ?? {
          subject: "",
          body: "",
          to: "",
          cc: "",
          bcc: "",
          date: "",
          time: "",
        }),
        [field]: value,
      },
    }));
  }

  function togglePlanModify(planGroup: PlanExecutionGroup, nextOpen?: boolean) {
    const shouldOpen = nextOpen ?? !(modifyingPlans[planGroup.key] ?? false);
    setExpandedPlans((current) => ({ ...current, [planGroup.key]: shouldOpen || current[planGroup.key] || false }));
    setModifyingPlans((current) => ({ ...current, [planGroup.key]: shouldOpen }));
    if (shouldOpen) {
      const snapshot = getExecutionPlanSnapshot(planGroup.items[0]);
      setPlanModifyDates((current) => ({
        ...current,
        [planGroup.key]: current[planGroup.key] ?? snapshot?.anchorDate ?? "",
      }));
    }
  }

  async function recallHistoryItem(record: ExecutionHistoryRecord): Promise<OutlookRecallResult> {
    const recallState = getExecutionHistoryRecallState(record);
    if (!recallState.canRecall || !recallState.recallImplemented || !record.providerObjectId) {
      throw new Error(recallState.recallReason || "This item cannot be recalled.");
    }

    if (record.providerObjectType === "message") {
      return await deleteOutlookMessage({
        messageId: record.providerObjectId,
      });
    } else if (record.providerObjectType === "event") {
      return await deleteOutlookCalendarEvent({
        eventId: record.providerObjectId,
        sendCancellation: record.itemType === "meeting" || record.itemType === "teams_meeting",
      });
    } else {
      throw new Error("This item cannot be recalled.");
    }
  }

  function getRecallStatusMessage(result: OutlookRecallResult) {
    if (result === "already_removed") {
      return { status: "already_removed" as const, text: "Already removed." };
    }
    if (result === "already_canceled") {
      return { status: "already_canceled" as const, text: "Already canceled." };
    }
    return { status: "recalled" as const, text: "Recalled." };
  }

  async function handleRecallItem(record: ExecutionHistoryRecord) {
    const recallState = getExecutionHistoryRecallState(record);
    if (!recallState.canRecall || !recallState.recallImplemented) return;
    const confirmed = window.confirm(`Recall "${record.subject || record.title || "this item"}"?`);
    if (!confirmed) return;

    setPendingItemRecalls((current) => ({ ...current, [record.id]: true }));
    setItemMessages((current) => {
      const next = { ...current };
      delete next[record.id];
      return next;
    });

    try {
      const result = await recallHistoryItem(record);
      const statusMessage = getRecallStatusMessage(result);
      await updateExecutionHistoryRecord(record.id, {
        status: statusMessage.status,
        details: {
          recalledAt: new Date().toISOString(),
          recallError: null,
        },
      });
      setItemMessages((current) => ({
        ...current,
        [record.id]: { tone: "success", text: statusMessage.text },
      }));
    } catch (error) {
      const message = error instanceof Error ? error.message : "Recall failed.";
      await updateExecutionHistoryRecord(record.id, {
        status: "recall_failed",
        details: {
          recallError: message,
        },
      });
      setItemMessages((current) => ({
        ...current,
        [record.id]: { tone: "error", text: message },
      }));
    } finally {
      setPendingItemRecalls((current) => ({ ...current, [record.id]: false }));
    }
  }

  async function handleRecallPlan(planGroup: PlanExecutionGroup) {
    const recallableItems = planGroup.items.filter((item) => {
      const recallState = getExecutionHistoryRecallState(item);
      return recallState.canRecall && recallState.recallImplemented;
    });

    if (recallableItems.length === 0) return;

    const confirmed = window.confirm(`Recall all supported items for "${planGroup.planName}"?`);
    if (!confirmed) return;

    setPendingPlanRecalls((current) => ({ ...current, [planGroup.key]: true }));
    setPlanMessages((current) => {
      const next = { ...current };
      delete next[planGroup.key];
      return next;
    });

    let successCount = 0;
    let alreadyHandledCount = 0;
    let failedCount = 0;

    for (const item of recallableItems) {
      try {
        const result = await recallHistoryItem(item);
        const statusMessage = getRecallStatusMessage(result);
        await updateExecutionHistoryRecord(item.id, {
          status: statusMessage.status,
          details: {
            recalledAt: new Date().toISOString(),
            recallError: null,
          },
        });
        if (result === "recalled") {
          successCount += 1;
        } else {
          alreadyHandledCount += 1;
        }
      } catch (error) {
        failedCount += 1;
        await updateExecutionHistoryRecord(item.id, {
          status: "recall_failed",
          details: {
            recallError: error instanceof Error ? error.message : "Recall failed.",
          },
        });
      }
    }

    setPendingPlanRecalls((current) => ({ ...current, [planGroup.key]: false }));
    setPlanMessages((current) => ({
      ...current,
      [planGroup.key]:
        failedCount === 0
          ? {
              tone: "success",
              text:
                alreadyHandledCount > 0 && successCount === 0
                  ? "Everything was already gone."
                  : alreadyHandledCount > 0
                    ? `Event recalled. ${successCount} removed, ${alreadyHandledCount} already gone.`
                    : "Event recalled.",
            }
          : successCount > 0 || alreadyHandledCount > 0
            ? {
                tone: "warning",
                text: `Partially recalled. ${successCount + alreadyHandledCount} handled, ${failedCount} failed.`,
              }
            : { tone: "error", text: "Recall failed." },
    }));
  }

  async function handleModifyItem(record: ExecutionHistoryRecord) {
    const modifyState = getExecutionHistoryModifyState(record);
    if (!modifyState.canModify || !modifyState.modifyImplemented || !record.providerObjectId) return;

    const draft = editDrafts[record.id] ?? createEditDraft(record);
    const action = getHistoryAction(record);
    setPendingItemModifies((current) => ({ ...current, [record.id]: true }));
    setItemMessages((current) => {
      const next = { ...current };
      delete next[record.id];
      return next;
    });

    try {
      const changedFields = getOverrideChangedFields(record, draft);
      if (record.providerObjectType === "message" && record.itemType === "email") {
        const nextEmailDraft = {
          to: splitAddresses(draft.to),
          cc: splitAddresses(draft.cc),
          bcc: splitAddresses(draft.bcc),
          subject: draft.subject.trim(),
          body: draft.body,
        };

        if (action === "email_scheduled") {
          if (!draft.date || !draft.time) {
            throw new Error("Add a date and time before saving.");
          }
          const scheduledSendISO = buildLocalIso(draft.date, draft.time);
          const replacedProviderObjectId = record.providerObjectId;
          const replacementResult = await replaceOutlookScheduledEmail({
            messageId: record.providerObjectId,
            draft: nextEmailDraft,
            fallbackSubject: draft.subject.trim() || record.subject || record.title || "Email draft",
            scheduledSendISO,
          });

          await updateExecutionHistoryRecord(record.id, {
            status: "modified",
            title: draft.subject.trim() || record.title,
            subject: draft.subject.trim() || record.subject,
            recipients: [...nextEmailDraft.to, ...nextEmailDraft.cc, ...nextEmailDraft.bcc],
            providerObjectId: replacementResult.id,
            outlookWebLink: replacementResult.webLink,
            scheduledFor: scheduledSendISO,
            details: {
              modifiedAt: new Date().toISOString(),
              modifyError: null,
              body: draft.body,
              emailDraft: nextEmailDraft,
              scheduledSendAt: scheduledSendISO,
              scheduledEmailState: "scheduled",
              replacedProviderObjectId,
              overrideTracking: {
                isOverridden: true,
                overriddenAt: new Date().toISOString(),
                overrideSource: "item_modify",
                changedFields,
              },
              replacementHistory: [
                ...(((Array.isArray(record.details.replacementHistory)
                  ? record.details.replacementHistory
                  : []) as unknown[]).filter(
                  (entry): entry is Record<string, unknown> => Boolean(entry) && typeof entry === "object" && !Array.isArray(entry)
                )),
                {
                  replacedProviderObjectId,
                  replacementProviderObjectId: replacementResult.id,
                  replacedAt: new Date().toISOString(),
                },
              ],
            },
          });
        } else {
          await updateOutlookMessageDraft({
            messageId: record.providerObjectId,
            draft: nextEmailDraft,
            fallbackSubject: draft.subject.trim() || record.subject || record.title || "Email draft",
          });

          await updateExecutionHistoryRecord(record.id, {
            status: "modified",
            title: draft.subject.trim() || record.title,
            subject: draft.subject.trim() || record.subject,
            recipients: [...nextEmailDraft.to, ...nextEmailDraft.cc, ...nextEmailDraft.bcc],
            details: {
              modifiedAt: new Date().toISOString(),
              modifyError: null,
              body: draft.body,
              emailDraft: nextEmailDraft,
              overrideTracking: {
                isOverridden: true,
                overriddenAt: new Date().toISOString(),
                overrideSource: "item_modify",
                changedFields,
              },
            },
          });
        }
      } else if (record.providerObjectType === "event") {
        if (!draft.date) {
          throw new Error("Add a date before saving.");
        }

        const startISO = record.isAllDay ? `${draft.date}T00:00:00` : buildLocalIso(draft.date, draft.time || "09:00");
        const endISO = addMinutesToIso(startISO, getDurationMinutes(record));

        await updateOutlookCalendarEvent({
          eventId: record.providerObjectId,
          subject: draft.subject.trim() || record.subject || record.title || "Event",
          bodyText: draft.body,
          startISO,
          endISO,
          timeZone: "America/New_York",
          isAllDay: record.isAllDay,
          attendees:
            record.itemType === "meeting" || record.itemType === "teams_meeting"
              ? splitAddresses(draft.to)
              : undefined,
        });

        const nextDetails =
          record.itemType === "meeting" || record.itemType === "teams_meeting"
            ? {
                meetingDraft: {
                  ...(getHistoryMeetingDetails(record) ?? {
                    attendees: [],
                    location: "",
                    title: "",
                    body: "",
                  }),
                  attendees: splitAddresses(draft.to),
                  title: draft.subject.trim() || record.subject,
                  body: draft.body,
                },
              }
            : {
                body: draft.body,
              };

        await updateExecutionHistoryRecord(record.id, {
          status: "modified",
          title: draft.subject.trim() || record.title,
          subject: draft.subject.trim() || record.subject,
          attendees:
            record.itemType === "meeting" || record.itemType === "teams_meeting" ? splitAddresses(draft.to) : record.attendees,
          scheduledFor: startISO,
          endsAt: endISO,
          details: {
            modifiedAt: new Date().toISOString(),
            modifyError: null,
            body: draft.body,
            overrideTracking: {
              isOverridden: true,
              overriddenAt: new Date().toISOString(),
              overrideSource: "item_modify",
              changedFields,
            },
            ...nextDetails,
          },
        });
      } else {
        throw new Error("This item cannot be modified.");
      }

      setItemMessages((current) => ({
        ...current,
        [record.id]: { tone: "success", text: "Saved changes." },
      }));
      setEditingItems((current) => ({ ...current, [record.id]: false }));
    } catch (error) {
      const message = error instanceof Error ? error.message : "Could not save changes.";
      await updateExecutionHistoryRecord(record.id, {
        status: "modify_failed",
        details: {
          modifyError: message,
          scheduledEmailState:
            action === "email_scheduled" && message.includes("already been sent") ? "sent" : record.details.scheduledEmailState ?? null,
        },
      });
      setItemMessages((current) => ({
        ...current,
        [record.id]: { tone: "error", text: message },
      }));
    } finally {
      setPendingItemModifies((current) => ({ ...current, [record.id]: false }));
    }
  }

  async function handleModifyPlan(planGroup: PlanExecutionGroup) {
    const nextEventDate = planModifyDates[planGroup.key] ?? "";
    const { snapshot, items } = getPlanModifyPreview(planGroup, nextEventDate);
    if (!snapshot || !nextEventDate) return;

    setPendingPlanModifies((current) => ({ ...current, [planGroup.key]: true }));
    setPlanMessages((current) => {
      const next = { ...current };
      delete next[planGroup.key];
      return next;
    });

    let updatedCount = 0;
    let replacedCount = 0;
    let skippedCount = 0;
    let failedCount = 0;

    for (const previewItem of items) {
      const { record, nextItem, action } = previewItem;
      if (!nextItem || action === "Locked" || action === "Unsupported" || action === "Unchanged") {
        skippedCount += 1;
        continue;
      }

      try {
        const nextTiming = getComputedPlanItemTiming(nextItem);
        const rescheduleDetails = {
          appliedAt: new Date().toISOString(),
          fromEventDate: snapshot.anchorDate,
          toEventDate: nextEventDate,
          action,
        };

        if (record.itemType === "meeting" || record.itemType === "teams_meeting") {
          await updateOutlookCalendarEvent({
            eventId: record.providerObjectId || "",
            subject: nextItem.customTitle || nextItem.title,
            bodyText: nextItem.body ?? "",
            startISO: nextTiming.scheduledFor,
            endISO: nextTiming.endsAt || addMinutesToIso(nextTiming.scheduledFor, 30),
            timeZone: "America/New_York",
            isAllDay: nextTiming.isAllDay,
            attendees: nextItem.meetingDraft?.attendees ?? [],
          });

          await updateExecutionHistoryRecord(record.id, {
            status: "modified",
            title: nextItem.customTitle || nextItem.title,
            subject: nextItem.customTitle || nextItem.title,
            attendees: nextItem.meetingDraft?.attendees ?? [],
            scheduledFor: nextTiming.scheduledFor,
            endsAt: nextTiming.endsAt,
            isAllDay: nextTiming.isAllDay,
            details: {
              body: nextItem.body ?? "",
              meetingDraft: {
                attendees: nextItem.meetingDraft?.attendees ?? [],
                location: nextItem.meetingDraft?.location ?? "",
                title: nextItem.customTitle || nextItem.title,
                body: nextItem.body ?? "",
              },
              latestPlanReschedule: rescheduleDetails,
            },
          });
          updatedCount += 1;
          continue;
        }

        if (record.itemType === "reminder") {
          const createResult = await createOutlookCalendarEvent({
            subject: nextItem.customTitle || nextItem.title,
            bodyText: nextItem.body ?? "",
            startISO: nextTiming.scheduledFor,
            endISO: nextTiming.endsAt || addMinutesToIso(nextTiming.scheduledFor, 30),
            timeZone: "America/New_York",
            isAllDay: nextTiming.isAllDay,
          });

          if (record.providerObjectId) {
            await deleteOutlookCalendarEvent({
              eventId: record.providerObjectId,
            });
          }

          const replacedProviderObjectId = record.providerObjectId;
          await updateExecutionHistoryRecord(record.id, {
            status: "modified",
            title: nextItem.customTitle || nextItem.title,
            subject: nextItem.customTitle || nextItem.title,
            providerObjectId: createResult.id,
            outlookWebLink: createResult.webLink,
            teamsJoinLink: createResult.joinUrl || null,
            scheduledFor: nextTiming.scheduledFor,
            endsAt: nextTiming.endsAt,
            isAllDay: nextTiming.isAllDay,
            details: {
              body: nextItem.body ?? "",
              replacedProviderObjectId,
              replacementHistory: [
                ...(((Array.isArray(record.details.replacementHistory) ? record.details.replacementHistory : []) as unknown[]).filter(
                  (entry): entry is Record<string, unknown> => Boolean(entry) && typeof entry === "object" && !Array.isArray(entry)
                )),
                {
                  replacedProviderObjectId,
                  replacementProviderObjectId: createResult.id,
                  replacedAt: new Date().toISOString(),
                  reason: "plan_reschedule",
                },
              ],
              latestPlanReschedule: rescheduleDetails,
            },
          });
          replacedCount += 1;
          continue;
        }

        if (record.itemType === "email" && getHistoryAction(record) === "email_scheduled") {
          const nextEmailDraft = {
            to: nextItem.emailDraft?.to ?? [],
            cc: nextItem.emailDraft?.cc ?? [],
            bcc: nextItem.emailDraft?.bcc ?? [],
            subject: nextItem.emailDraft?.subject ?? nextItem.customTitle ?? nextItem.title,
            body: nextItem.emailDraft?.body ?? "",
          };
          const replacementResult = await replaceOutlookScheduledEmail({
            messageId: record.providerObjectId || "",
            draft: nextEmailDraft,
            fallbackSubject: nextEmailDraft.subject,
            scheduledSendISO: nextTiming.scheduledFor,
          });
          const replacedProviderObjectId = record.providerObjectId;
          await updateExecutionHistoryRecord(record.id, {
            status: "modified",
            title: nextItem.customTitle || nextItem.title,
            subject: nextEmailDraft.subject,
            recipients: [...nextEmailDraft.to, ...nextEmailDraft.cc, ...nextEmailDraft.bcc],
            providerObjectId: replacementResult.id,
            outlookWebLink: replacementResult.webLink,
            scheduledFor: nextTiming.scheduledFor,
            details: {
              body: nextEmailDraft.body,
              emailDraft: nextEmailDraft,
              scheduledSendAt: nextTiming.scheduledFor,
              scheduledEmailState: "scheduled",
              replacedProviderObjectId,
              replacementHistory: [
                ...(((Array.isArray(record.details.replacementHistory) ? record.details.replacementHistory : []) as unknown[]).filter(
                  (entry): entry is Record<string, unknown> => Boolean(entry) && typeof entry === "object" && !Array.isArray(entry)
                )),
                {
                  replacedProviderObjectId,
                  replacementProviderObjectId: replacementResult.id,
                  replacedAt: new Date().toISOString(),
                  reason: "plan_reschedule",
                },
              ],
              latestPlanReschedule: rescheduleDetails,
            },
          });
          replacedCount += 1;
          continue;
        }

        if (record.itemType === "email") {
          const nextEmailDraft = {
            to: nextItem.emailDraft?.to ?? [],
            cc: nextItem.emailDraft?.cc ?? [],
            bcc: nextItem.emailDraft?.bcc ?? [],
            subject: nextItem.emailDraft?.subject ?? nextItem.customTitle ?? nextItem.title,
            body: nextItem.emailDraft?.body ?? "",
          };

          await updateOutlookMessageDraft({
            messageId: record.providerObjectId || "",
            draft: nextEmailDraft,
            fallbackSubject: nextEmailDraft.subject,
          });

          await updateExecutionHistoryRecord(record.id, {
            status: "modified",
            title: nextItem.customTitle || nextItem.title,
            subject: nextEmailDraft.subject,
            recipients: [...nextEmailDraft.to, ...nextEmailDraft.cc, ...nextEmailDraft.bcc],
            scheduledFor: nextTiming.scheduledFor,
            details: {
              body: nextEmailDraft.body,
              emailDraft: nextEmailDraft,
              latestPlanReschedule: rescheduleDetails,
            },
          });
          updatedCount += 1;
          continue;
        }

        skippedCount += 1;
      } catch (error) {
        failedCount += 1;
        await updateExecutionHistoryRecord(record.id, {
          details: {
            planRescheduleError: error instanceof Error ? error.message : "Could not update this item.",
          },
        });
      }
    }

    setPendingPlanModifies((current) => ({ ...current, [planGroup.key]: false }));
    setPlanMessages((current) => ({
      ...current,
      [planGroup.key]:
        failedCount === 0 && updatedCount + replacedCount > 0
          ? {
              tone: skippedCount > 0 ? "warning" : "success",
              text:
                skippedCount > 0
                  ? `Partially updated. ${updatedCount} updated, ${replacedCount} replaced, ${skippedCount} skipped.`
                  : `Event updated. ${updatedCount} updated, ${replacedCount} replaced.`,
            }
          : updatedCount + replacedCount > 0
            ? {
                tone: "warning",
                text: `Partially updated. ${updatedCount} updated, ${replacedCount} replaced, ${skippedCount} skipped, ${failedCount} failed.`,
              }
            : {
                tone: failedCount > 0 ? "error" : "warning",
                text: failedCount > 0 ? `Update failed for ${failedCount} item${failedCount === 1 ? "" : "s"}.` : "Nothing changed.",
              },
    }));
  }

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
                          const isPlanModifying = modifyingPlans[planGroup.key] ?? false;
                          const planSnapshot = getExecutionPlanSnapshot(planGroup.items[0]);
                          const planModifyDate = planModifyDates[planGroup.key] ?? planSnapshot?.anchorDate ?? "";
                          const planModifyPreview = planModifyDate ? getPlanModifyPreview(planGroup, planModifyDate) : { snapshot: planSnapshot, items: [] };
                          const recallablePlanItems = planGroup.items.filter((item) => {
                            const recallState = getExecutionHistoryRecallState(item);
                            return recallState.canRecall && recallState.recallImplemented;
                          });
                          const canModifyPlan = Boolean(planSnapshot && !planSnapshot.noEventDate);
                          const planMessage = planMessages[planGroup.key] ?? null;
                          return (
                            <section key={planGroup.key} className="rounded-2xl border bg-white shadow-sm">
                              <div className="flex items-center justify-between gap-4 px-5 py-4">
                                <div className="min-w-0">
                                  <div className="text-lg font-semibold text-gray-900">Event Name: {planGroup.planName}</div>
                                  <div className="mt-1 text-sm text-gray-600">
                                    {planGroup.items.length} created item{planGroup.items.length === 1 ? "" : "s"}
                                  </div>
                                  <div className="mt-1 text-sm text-gray-600">Created at: {formatDateTime(planGroup.latestExecutedAt)}</div>
                                  {planMessage ? (
                                    <div
                                      className={`mt-2 text-sm ${
                                        planMessage.tone === "success"
                                          ? "text-green-700"
                                          : planMessage.tone === "warning"
                                            ? "text-amber-700"
                                            : "text-red-700"
                                      }`}
                                    >
                                      {planMessage.text}
                                    </div>
                                  ) : null}
                                </div>
                                <div className="flex items-center gap-3">
                                  <button
                                    type="button"
                                    onClick={() => togglePlanModify(planGroup)}
                                    disabled={!canModifyPlan || pendingPlanModifies[planGroup.key]}
                                    className="rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900 hover:bg-gray-50 disabled:cursor-not-allowed disabled:text-gray-400"
                                  >
                                    {pendingPlanModifies[planGroup.key] ? "Updating..." : "Modify Event"}
                                  </button>
                                  <button
                                    type="button"
                                    onClick={() => void handleRecallPlan(planGroup)}
                                    disabled={recallablePlanItems.length === 0 || pendingPlanRecalls[planGroup.key]}
                                    className="rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900 hover:bg-gray-50 disabled:cursor-not-allowed disabled:text-gray-400"
                                  >
                                    {pendingPlanRecalls[planGroup.key] ? "Recalling..." : "Recall Event"}
                                  </button>
                                  <button
                                    type="button"
                                    onClick={() => setExpandedPlans((current) => ({ ...current, [planGroup.key]: !isPlanExpanded }))}
                                    className="text-sm font-medium text-gray-500 hover:text-gray-700"
                                  >
                                    {isPlanExpanded ? "Collapse" : "Expand"}
                                  </button>
                                </div>
                              </div>

                              {isPlanModifying ? (
                                <div className="border-t bg-gray-50/70 px-5 py-4">
                                  <div className="grid gap-4 md:grid-cols-[220px_220px_auto] md:items-end">
                                    <div>
                                      <div className="text-xs font-semibold uppercase tracking-wide text-gray-500">Current Event Date</div>
                                      <div className="mt-2 rounded-lg border bg-white px-3 py-2 text-sm text-gray-900">
                                        {planSnapshot?.anchorDate ? formatDateOnly(`${planSnapshot.anchorDate}T00:00:00`) : "Not available"}
                                      </div>
                                    </div>
                                    <div>
                                      <label className="text-xs font-semibold uppercase tracking-wide text-gray-500">New Event Date</label>
                                      <input
                                        type="date"
                                        value={planModifyDate}
                                        onChange={(event) => setPlanModifyDates((current) => ({ ...current, [planGroup.key]: event.target.value }))}
                                        className="mt-2 w-full rounded-lg border bg-white px-3 py-2 text-sm text-gray-900"
                                      />
                                    </div>
                                    <div className="flex justify-end gap-2">
                                      <button
                                        type="button"
                                        onClick={() => setModifyingPlans((current) => ({ ...current, [planGroup.key]: false }))}
                                        className="rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900 hover:bg-gray-50"
                                      >
                                        Close
                                      </button>
                                      <button
                                        type="button"
                                        onClick={() => void handleModifyPlan(planGroup)}
                                        disabled={!planModifyDate || planModifyPreview.items.every((item) => item.action === "Unchanged" || item.action === "Locked" || item.action === "Unsupported")}
                                        className="rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900 hover:bg-gray-50 disabled:cursor-not-allowed disabled:text-gray-400"
                                      >
                                        Apply Changes
                                      </button>
                                    </div>
                                  </div>

                                  <div className="mt-4 overflow-hidden rounded-xl border bg-white">
                                    <div className="hidden grid-cols-[140px_minmax(0,1fr)_180px_180px_150px] gap-3 border-b px-4 py-3 text-xs font-semibold uppercase tracking-wide text-gray-600 md:grid">
                                      <div>Type</div>
                                      <div>Title</div>
                                      <div>Old Date / Time</div>
                                      <div>New Date / Time</div>
                                      <div>Planned Action</div>
                                    </div>
                                    <div className="divide-y">
                                      {planModifyPreview.items.map((previewItem) => (
                                        <div key={`modify-preview:${planGroup.key}:${previewItem.record.id}`} className="grid gap-3 px-4 py-3 md:grid-cols-[140px_minmax(0,1fr)_180px_180px_150px] md:items-center">
                                          <div className={`text-sm font-medium ${getTypeAccentClasses(formatTimelineItemType(previewItem.record))}`}>
                                            {getItemTypeDisplayLabel(previewItem.record)}
                                          </div>
                                          <div className="text-sm text-gray-900">{previewItem.record.subject || previewItem.record.title}</div>
                                          <div className="text-sm text-gray-600">{formatPreviewDateTime(previewItem.oldDateTime)}</div>
                                          <div className="text-sm text-gray-900">{formatPreviewDateTime(previewItem.newDateTime)}</div>
                                          <div className="text-sm">
                                            <div
                                              className={
                                                previewItem.action === "Update"
                                                  ? "text-blue-700"
                                                  : previewItem.action === "Replace"
                                                    ? "text-amber-700"
                                                    : previewItem.action === "Locked"
                                                      ? "text-gray-700"
                                                      : previewItem.action === "Unsupported"
                                                        ? "text-red-700"
                                                        : "text-gray-500"
                                              }
                                            >
                                              {previewItem.action}
                                            </div>
                                            {previewItem.reason ? <div className="mt-1 text-xs text-gray-500">{previewItem.reason}</div> : null}
                                          </div>
                                        </div>
                                      ))}
                                    </div>
                                  </div>
                                </div>
                              ) : null}

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
                                      const itemAction = getHistoryAction(item);
                                      const itemTypeLabel = formatTimelineItemType(item);
                                      const itemTypeDisplayLabel = getItemTypeDisplayLabel(item);
                                      const itemDateTime = item.scheduledFor || item.executedAt;
                                      const isItemExpanded = expandedItems[item.id] ?? false;
                                      const isItemEditing = editingItems[item.id] ?? false;
                                      const reminderBody = getHistoryBody(item);
                                      const emailDraft = getHistoryEmailDraftDetails(item);
                                      const meetingDetails = getHistoryMeetingDetails(item);
                                      const recallState = getExecutionHistoryRecallState(item);
                                      const modifyState = getExecutionHistoryModifyState(item);
                                      const editDraft = editDrafts[item.id] ?? createEditDraft(item);
                                      const itemMessage = itemMessages[item.id] ?? null;
                                      return (
                                        <article key={item.id} className="py-4">
                                          <div className="grid items-center gap-3 md:grid-cols-[140px_minmax(0,1.25fr)_190px_130px_220px]">
                                            <div className={`text-sm font-medium ${getTypeAccentClasses(itemTypeLabel)}`}>
                                              {itemTypeDisplayLabel}
                                              {item.status === "modified" ? <div className="mt-1 text-xs text-green-700">Modified</div> : null}
                                              {item.status === "recalled" ? <div className="mt-1 text-xs text-green-700">Recalled</div> : null}
                                              {item.status === "already_removed" ? <div className="mt-1 text-xs text-green-700">Already removed</div> : null}
                                              {item.status === "already_canceled" ? <div className="mt-1 text-xs text-green-700">Already canceled</div> : null}
                                              {item.status === "modify_failed" ? <div className="mt-1 text-xs text-red-700">Edit failed</div> : null}
                                              {item.status === "recall_failed" ? <div className="mt-1 text-xs text-red-700">Recall failed</div> : null}
                                            </div>
                                            <div className="min-w-0 truncate text-base text-gray-900">{item.subject || item.title || "Untitled item"}</div>
                                            <div className="rounded-xl border bg-white px-4 py-3 text-center text-sm text-gray-900">
                                              {formatDateOnly(itemDateTime)}
                                            </div>
                                            <div className="rounded-xl border bg-white px-4 py-3 text-center text-sm text-gray-900">
                                              {formatTimeOnly(itemDateTime)}
                                            </div>
                                            <div className="flex justify-end gap-2">
                                              <button
                                                type="button"
                                                onClick={() => toggleItemEdit(item)}
                                                disabled={!modifyState.canModify || !modifyState.modifyImplemented || pendingItemModifies[item.id]}
                                                className="rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900 hover:bg-gray-50 disabled:cursor-not-allowed disabled:text-gray-400"
                                              >
                                                {pendingItemModifies[item.id] ? "Saving..." : "Modify"}
                                              </button>
                                              <button
                                                type="button"
                                                onClick={() => void handleRecallItem(item)}
                                                disabled={!recallState.canRecall || !recallState.recallImplemented || pendingItemRecalls[item.id]}
                                                className="rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900 hover:bg-gray-50 disabled:cursor-not-allowed disabled:text-gray-400"
                                              >
                                                {pendingItemRecalls[item.id] ? "Recalling..." : "Recall"}
                                              </button>
                                              <button
                                                type="button"
                                                onClick={() => {
                                                  setExpandedItems((current) => ({ ...current, [item.id]: !isItemExpanded }));
                                                  if (isItemExpanded) {
                                                    setEditingItems((current) => ({ ...current, [item.id]: false }));
                                                  }
                                                }}
                                                className="rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900 hover:bg-gray-50"
                                              >
                                                {isItemExpanded ? `Hide ${itemTypeLabel}` : getViewFullLabel(item)}
                                              </button>
                                            </div>
                                          </div>
                                          {isItemExpanded ? (
                                            <div className="mt-4 space-y-4 rounded-xl border bg-gray-50 p-4">
                                              {itemMessage ? (
                                                <div className={`text-sm ${itemMessage.tone === "success" ? "text-green-700" : "text-red-700"}`}>
                                                  {itemMessage.text}
                                                </div>
                                              ) : null}
                                              {!recallState.canRecall && recallState.recallReason ? (
                                                <div className="text-sm text-gray-600">{recallState.recallReason}</div>
                                              ) : null}
                                              {!modifyState.canModify && modifyState.modifyReason ? (
                                                <div className="text-sm text-gray-600">{modifyState.modifyReason}</div>
                                              ) : null}
                                              <div className="text-sm text-gray-700">Event Name: {planGroup.planName}</div>
                                              {isItemEditing && item.itemType === "reminder" ? (
                                                <div className="bg-blue-50 px-4 py-3">
                                                  <div className="grid grid-cols-1 gap-3 md:grid-cols-2">
                                                    <div className="md:col-span-2">
                                                      <label className="mb-1 block text-sm font-medium text-blue-950">Reminder Text</label>
                                                      <input
                                                        className="w-full rounded-lg border border-blue-200 bg-white px-3 py-2 text-sm text-gray-900"
                                                        value={editDraft.subject}
                                                        onChange={(event) => handleEditDraftChange(item.id, "subject", event.target.value)}
                                                      />
                                                    </div>
                                                    <div>
                                                      <label className="mb-1 block text-sm font-medium text-blue-950">Date Disseminated</label>
                                                      <input
                                                        type="date"
                                                        className="w-full rounded-lg border border-blue-200 bg-white px-3 py-2 text-sm text-gray-900"
                                                        value={editDraft.date}
                                                        onChange={(event) => handleEditDraftChange(item.id, "date", event.target.value)}
                                                      />
                                                    </div>
                                                    {!item.isAllDay ? (
                                                      <div>
                                                        <label className="mb-1 block text-sm font-medium text-blue-950">Time</label>
                                                        <input
                                                          type="time"
                                                          className="w-full rounded-lg border border-blue-200 bg-white px-3 py-2 text-sm text-gray-900"
                                                          value={editDraft.time}
                                                          onChange={(event) => handleEditDraftChange(item.id, "time", event.target.value)}
                                                        />
                                                      </div>
                                                    ) : null}
                                                    <div className="md:col-span-2">
                                                      <label className="mb-1 block text-sm font-medium text-blue-950">Reminder Body</label>
                                                      <textarea
                                                        rows={5}
                                                        className="w-full rounded-lg border border-blue-200 bg-white px-3 py-2 text-sm text-gray-900"
                                                        value={editDraft.body}
                                                        onChange={(event) => handleEditDraftChange(item.id, "body", event.target.value)}
                                                      />
                                                    </div>
                                                  </div>
                                                </div>
                                              ) : item.itemType === "reminder" ? (
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
                                              {isItemEditing && (item.itemType === "meeting" || item.itemType === "teams_meeting") ? (
                                                <div className="bg-violet-50 px-4 py-3">
                                                  <div className="grid grid-cols-1 gap-3 md:grid-cols-2">
                                                    <div className="md:col-span-2">
                                                      <label className="mb-1 block text-sm font-medium text-violet-950">To</label>
                                                      <textarea
                                                        rows={2}
                                                        className="w-full rounded-lg border border-violet-200 bg-white px-3 py-2 text-sm text-gray-900"
                                                        value={editDraft.to}
                                                        onChange={(event) => handleEditDraftChange(item.id, "to", event.target.value)}
                                                      />
                                                    </div>
                                                    <div className="md:col-span-2">
                                                      <label className="mb-1 block text-sm font-medium text-violet-950">Subject</label>
                                                      <input
                                                        className="w-full rounded-lg border border-violet-200 bg-white px-3 py-2 text-sm text-gray-900"
                                                        value={editDraft.subject}
                                                        onChange={(event) => handleEditDraftChange(item.id, "subject", event.target.value)}
                                                      />
                                                    </div>
                                                    <div>
                                                      <label className="mb-1 block text-sm font-medium text-violet-950">Date Disseminated</label>
                                                      <input
                                                        type="date"
                                                        className="w-full rounded-lg border border-violet-200 bg-white px-3 py-2 text-sm text-gray-900"
                                                        value={editDraft.date}
                                                        onChange={(event) => handleEditDraftChange(item.id, "date", event.target.value)}
                                                      />
                                                    </div>
                                                    {!item.isAllDay ? (
                                                      <div>
                                                        <label className="mb-1 block text-sm font-medium text-violet-950">Time</label>
                                                        <input
                                                          type="time"
                                                          className="w-full rounded-lg border border-violet-200 bg-white px-3 py-2 text-sm text-gray-900"
                                                          value={editDraft.time}
                                                          onChange={(event) => handleEditDraftChange(item.id, "time", event.target.value)}
                                                        />
                                                      </div>
                                                    ) : null}
                                                    <div className="md:col-span-2">
                                                      <label className="mb-1 block text-sm font-medium text-violet-950">Message</label>
                                                      <textarea
                                                        rows={6}
                                                        className="w-full rounded-lg border border-violet-200 bg-white px-3 py-2 text-sm text-gray-900"
                                                        value={editDraft.body}
                                                        onChange={(event) => handleEditDraftChange(item.id, "body", event.target.value)}
                                                      />
                                                    </div>
                                                  </div>
                                                </div>
                                              ) : (item.itemType === "meeting" || item.itemType === "teams_meeting") && meetingDetails ? (
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
                                              {isItemEditing && item.itemType === "email" ? (
                                                <div className="bg-amber-50 px-4 py-3">
                                                  <div className="grid grid-cols-1 gap-3">
                                                    {itemAction === "email_scheduled" ? (
                                                      <div className="grid grid-cols-1 gap-3 md:grid-cols-2">
                                                        <div>
                                                          <label className="mb-1 block text-sm font-medium text-amber-950">Date Disseminated</label>
                                                          <input
                                                            type="date"
                                                            className="w-full rounded-lg border border-amber-200 bg-white px-3 py-2 text-sm text-gray-900"
                                                            value={editDraft.date}
                                                            onChange={(event) => handleEditDraftChange(item.id, "date", event.target.value)}
                                                          />
                                                        </div>
                                                        <div>
                                                          <label className="mb-1 block text-sm font-medium text-amber-950">Time</label>
                                                          <input
                                                            type="time"
                                                            className="w-full rounded-lg border border-amber-200 bg-white px-3 py-2 text-sm text-gray-900"
                                                            value={editDraft.time}
                                                            onChange={(event) => handleEditDraftChange(item.id, "time", event.target.value)}
                                                          />
                                                        </div>
                                                      </div>
                                                    ) : null}
                                                    <div>
                                                      <label className="mb-1 block text-sm font-medium text-amber-950">To</label>
                                                      <textarea
                                                        rows={2}
                                                        className="w-full rounded-lg border border-amber-200 bg-white px-3 py-2 text-sm text-gray-900"
                                                        value={editDraft.to}
                                                        onChange={(event) => handleEditDraftChange(item.id, "to", event.target.value)}
                                                      />
                                                    </div>
                                                    <div>
                                                      <label className="mb-1 block text-sm font-medium text-amber-950">Cc</label>
                                                      <textarea
                                                        rows={2}
                                                        className="w-full rounded-lg border border-amber-200 bg-white px-3 py-2 text-sm text-gray-900"
                                                        value={editDraft.cc}
                                                        onChange={(event) => handleEditDraftChange(item.id, "cc", event.target.value)}
                                                      />
                                                    </div>
                                                    <div>
                                                      <label className="mb-1 block text-sm font-medium text-amber-950">Bcc</label>
                                                      <textarea
                                                        rows={2}
                                                        className="w-full rounded-lg border border-amber-200 bg-white px-3 py-2 text-sm text-gray-900"
                                                        value={editDraft.bcc}
                                                        onChange={(event) => handleEditDraftChange(item.id, "bcc", event.target.value)}
                                                      />
                                                    </div>
                                                    <div>
                                                      <label className="mb-1 block text-sm font-medium text-amber-950">Subject</label>
                                                      <input
                                                        className="w-full rounded-lg border border-amber-200 bg-white px-3 py-2 text-sm text-gray-900"
                                                        value={editDraft.subject}
                                                        onChange={(event) => handleEditDraftChange(item.id, "subject", event.target.value)}
                                                      />
                                                    </div>
                                                    <div>
                                                      <label className="mb-1 block text-sm font-medium text-amber-950">Message</label>
                                                      <textarea
                                                        rows={8}
                                                        className="w-full rounded-lg border border-amber-200 bg-white px-3 py-2 text-sm text-gray-900"
                                                        value={editDraft.body}
                                                        onChange={(event) => handleEditDraftChange(item.id, "body", event.target.value)}
                                                      />
                                                    </div>
                                                  </div>
                                                </div>
                                              ) : item.itemType === "email" && emailDraft ? (
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
                                              {isItemEditing ? (
                                                <div className="flex justify-end gap-2">
                                                  <button
                                                    type="button"
                                                    onClick={() => {
                                                      setEditDrafts((current) => ({ ...current, [item.id]: createEditDraft(item) }));
                                                      setEditingItems((current) => ({ ...current, [item.id]: false }));
                                                    }}
                                                    className="rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900 hover:bg-gray-50"
                                                  >
                                                    Cancel
                                                  </button>
                                                  <button
                                                    type="button"
                                                    onClick={() => void handleModifyItem(item)}
                                                    disabled={pendingItemModifies[item.id]}
                                                    className="rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900 hover:bg-gray-50 disabled:cursor-not-allowed disabled:text-gray-400"
                                                  >
                                                    {pendingItemModifies[item.id] ? "Saving..." : "Save Changes"}
                                                  </button>
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
