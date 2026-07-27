"use client";

import Link from "next/link";
import { useEffect, useMemo, useRef, useState, type KeyboardEvent as ReactKeyboardEvent, type MouseEvent as ReactMouseEvent } from "react";
import { createPortal } from "react-dom";

import {
  clearExecutionHistory,
  deleteExecutionHistoryRecord,
  deleteExecutionHistoryRecords,
  EXECUTION_HISTORY_UPDATED_EVENT,
  getExecutionHistoryModifyState,
  getExecutionHistoryRecallState,
  listExecutionHistory,
  readCachedExecutionHistorySnapshot,
  updateExecutionHistoryRecord,
  type ExecutionHistoryRecord,
} from "../../lib/executionHistory";
import {
  createGoogleCalendarEvent,
  deleteGoogleCalendarEvent,
  updateGoogleCalendarEvent,
  type GoogleCalendarRecallResult,
} from "../../lib/gmailClient";
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
import { useAuthContext } from "../components/auth-provider";

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
    addGoogleMeet?: boolean;
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

type HistoryRecallResult = OutlookRecallResult | GoogleCalendarRecallResult;

type ConfirmationTone = "destructive" | "provider";

type ConfirmationDialogState = {
  title: string;
  body: string;
  confirmLabel: string;
  tone: ConfirmationTone;
  onConfirm: () => Promise<void>;
};

type HistoryActionMenuKind = "plan" | "item";

type HistoryActionMenuPosition = {
  top: number;
  left: number;
  width: number;
};

type HistoryActionMenuState = {
  kind: HistoryActionMenuKind;
  id: string;
  position: HistoryActionMenuPosition;
};

type HistoryActionMenuAction = {
  key: string;
  label: string;
  tone?: "default" | "provider" | "history";
  disabled?: boolean;
  title?: string;
  onSelect: () => void;
};

const HISTORY_ACTION_MENU_WIDTH = 216;
const HISTORY_ACTION_MENU_GAP = 6;
const HISTORY_ACTION_MENU_MARGIN = 12;
const HISTORY_ACTION_MENU_ROW_HEIGHT = 40;
const HISTORY_ACTION_MENU_VERTICAL_PADDING = 12;
const LOCAL_PROVIDER_RECALL_UNAVAILABLE_COPY =
  "This action did not create an item in a connected provider, so provider recall is unavailable.";

function formatDayLabel(value: string) {
  const parsed = new Date(`${value}T00:00:00`);
  if (!Number.isNaN(parsed.getTime())) {
    const today = new Date();
    const todayKey = getLocalDayKey(today.toISOString());
    const yesterday = new Date(today);
    yesterday.setDate(today.getDate() - 1);
    const yesterdayKey = getLocalDayKey(yesterday.toISOString());

    if (value === todayKey) return "Today";
    if (value === yesterdayKey) return "Yesterday";

    const sameYear = parsed.getFullYear() === today.getFullYear();
    return new Intl.DateTimeFormat("en-US", {
      weekday: sameYear ? "long" : undefined,
      month: "long",
      day: "numeric",
      year: sameYear ? undefined : "numeric",
    }).format(parsed);
  }

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

function formatReadableDate(value: string | null) {
  if (!value) return null;
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return null;
  return new Intl.DateTimeFormat("en-US", {
    month: "short",
    day: "numeric",
    year: "numeric",
  }).format(parsed);
}

function formatReadableTime(value: string | null) {
  if (!value) return null;
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return null;
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

function IconTrash() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true" className="h-4 w-4" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <path d="M3 6h18" />
      <path d="M8 6V4.5h8V6" />
      <path d="M6.5 6l1 13.5h9L17.5 6" />
      <path d="M10 10.5v5" />
      <path d="M14 10.5v5" />
    </svg>
  );
}

function IconEllipsis() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true" className="h-4 w-4" fill="currentColor">
      <circle cx="5" cy="12" r="1.7" />
      <circle cx="12" cy="12" r="1.7" />
      <circle cx="19" cy="12" r="1.7" />
    </svg>
  );
}

function getItemTypeDisplayLabel(record: ExecutionHistoryRecord) {
  return formatTimelineItemType(record);
}

function getHistoryAction(record: ExecutionHistoryRecord) {
  return typeof record.details.action === "string" ? record.details.action : "";
}

function getScheduledEmailState(record: ExecutionHistoryRecord) {
  return typeof record.details.scheduledEmailState === "string" ? record.details.scheduledEmailState : "";
}

function getProviderLabel(record: ExecutionHistoryRecord) {
  if (record.provider === "outlook") return "Outlook";
  if (record.provider === "local_export") return "Local export";
  if (record.provider === "gmail") {
    return record.itemType === "email" || record.providerObjectType === "message" ? "Gmail" : "Google Calendar";
  }
  return null;
}

function getConsistentGroupProviderLabel(planGroup: PlanExecutionGroup) {
  const labels = Array.from(
    new Set(planGroup.items.map(getProviderLabel).filter((label): label is NonNullable<ReturnType<typeof getProviderLabel>> => Boolean(label)))
  );
  return labels.length === 1 ? labels[0] : null;
}

function getPlanDisplayName(planGroup: PlanExecutionGroup) {
  const name = planGroup.planName.trim();
  const normalized = name.toLowerCase();
  if (!name || normalized === "unnamed plan" || normalized === "unknown event" || normalized === "untitled history group" || normalized === "null" || normalized === "undefined") {
    return "Workflow activity";
  }
  return name;
}

function formatCountedNoun(count: number, singular: string, plural = `${singular}s`) {
  return `${count} ${count === 1 ? singular : plural}`;
}

function getRecipientCount(record: ExecutionHistoryRecord) {
  const uniqueRecipients = new Set([...record.recipients, ...record.attendees].map((entry) => entry.trim()).filter(Boolean));
  return uniqueRecipients.size;
}

function getHistoryOutcome(record: ExecutionHistoryRecord) {
  const action = getHistoryAction(record);
  const scheduledEmailState = getScheduledEmailState(record);

  if (record.status === "failed") return { text: "Export failed", tone: "error" as const };
  if (record.status === "modify_failed") return { text: "Update failed", tone: "error" as const };
  if (record.status === "recall_failed") return { text: "Recall failed", tone: "error" as const };
  if (record.status === "recalled") return { text: "Item recalled", tone: "neutral" as const };
  if (record.status === "already_removed") return { text: "Item already removed", tone: "neutral" as const };
  if (record.status === "already_canceled") return { text: "Item already canceled", tone: "neutral" as const };
  if (record.status === "modified") {
    if (record.itemType === "reminder") return { text: "Reminder updated", tone: "success" as const };
    if (record.itemType === "email") return { text: "Email updated", tone: "success" as const };
    return { text: "Meeting updated", tone: "success" as const };
  }

  if (record.path === "fallback") {
    if (record.fallbackExportKind === "eml") return { text: "Email draft exported", tone: "neutral" as const };
    if (record.fallbackExportKind === "ics") return { text: "Calendar file exported", tone: "neutral" as const };
    return { text: "Local export created", tone: "neutral" as const };
  }

  if (record.itemType === "email") {
    if (action === "email_sent" || scheduledEmailState === "sent") {
      const recipientCount = getRecipientCount(record);
      return {
        text: recipientCount > 0 ? `Email sent to ${formatCountedNoun(recipientCount, "recipient")}` : "Email sent",
        tone: "success" as const,
      };
    }
    if (action === "email_scheduled") return { text: "Email scheduled", tone: "success" as const };
    if (action === "draft_created") return { text: "Email draft created", tone: "success" as const };
    return { text: "Email action completed", tone: "success" as const };
  }

  if (record.itemType === "meeting" || record.itemType === "teams_meeting") return { text: "Meeting created", tone: "success" as const };
  return { text: "Reminder scheduled", tone: "success" as const };
}

function getOutcomeTextClasses(tone: ReturnType<typeof getHistoryOutcome>["tone"]) {
  if (tone === "error") return "text-red-700";
  if (tone === "neutral") return "text-slate-700";
  return "text-slate-950";
}

function getItemTitle(record: ExecutionHistoryRecord) {
  return (record.subject || record.title || "").trim();
}

function getItemAccessibleName(record: ExecutionHistoryRecord) {
  return getItemTitle(record) || getItemTypeDisplayLabel(record).toLowerCase();
}

function getActivityMetadata(record: ExecutionHistoryRecord) {
  const parts: string[] = [];
  const scheduledDate = formatReadableDate(record.scheduledFor);
  const scheduledTime = record.isAllDay ? "All day" : formatReadableTime(record.scheduledFor);
  const createdDate = formatReadableDate(record.executedAt);
  const createdTime = formatReadableTime(record.executedAt);
  const provider = getProviderLabel(record);

  if (scheduledDate) {
    parts.push(`Scheduled ${scheduledDate}`);
    if (scheduledTime) parts.push(scheduledTime);
  } else if (createdDate) {
    parts.push(`Created ${createdDate}`);
    if (createdTime) parts.push(createdTime);
  }

  if (provider) parts.push(provider);
  return parts.join(" · ");
}

function getGroupMetadata(planGroup: PlanExecutionGroup) {
  const parts = [`${planGroup.items.length} ${planGroup.items.length === 1 ? "action" : "actions"}`];
  const provider = getConsistentGroupProviderLabel(planGroup);
  const latestTime = formatReadableTime(planGroup.latestExecutedAt);
  if (provider && provider !== "Local export") parts.push(provider);
  if (latestTime) parts.push(latestTime);
  return parts.join(" · ");
}

function getAttentionCount(records: ExecutionHistoryRecord[]) {
  return records.filter((record) => record.status === "failed" || record.status === "modify_failed" || record.status === "recall_failed").length;
}

function getProviderRecallActionLabel(record: ExecutionHistoryRecord) {
  if (record.provider === "gmail" && record.providerObjectType === "event") return "Remove from Google Calendar";
  if (record.provider === "outlook" && record.providerObjectType === "message" && getHistoryAction(record) === "email_scheduled") {
    return "Cancel scheduled email";
  }
  if (record.provider === "outlook") return "Recall from Outlook";
  return "Change connected provider item";
}

function getPlanRecallActionLabel(planGroup: PlanExecutionGroup) {
  const provider = getConsistentGroupProviderLabel(planGroup);
  if (provider === "Outlook") return "Recall from Outlook";
  if (provider === "Google Calendar") return "Remove from Google Calendar";
  if (provider === "Gmail") return "Cancel Gmail items";
  return "Change connected provider items";
}

function getProviderRecallConfirmation(record: ExecutionHistoryRecord) {
  if (record.provider === "gmail" && record.providerObjectType === "event") {
    return {
      title: "Remove this event from Google Calendar?",
      body: "This attempts to remove the calendar event. The History record will remain and update with the result.",
      confirmLabel: "Remove event",
    };
  }

  if (record.provider === "outlook" && record.providerObjectType === "message" && getHistoryAction(record) === "email_scheduled") {
    return {
      title: "Cancel this scheduled email?",
      body: "This attempts to remove the scheduled email from Outlook. The History record will remain and update with the result.",
      confirmLabel: "Cancel email",
    };
  }

  return {
    title: "Recall this item from Outlook?",
    body: "This attempts to remove the item from Outlook. The History record will remain and update with the result.",
    confirmLabel: "Recall item",
  };
}

function getHistoryActionMenuPosition(trigger: HTMLElement, itemCount: number): HistoryActionMenuPosition {
  const rect = trigger.getBoundingClientRect();
  const width = HISTORY_ACTION_MENU_WIDTH;
  const estimatedHeight = HISTORY_ACTION_MENU_VERTICAL_PADDING + Math.max(1, itemCount) * HISTORY_ACTION_MENU_ROW_HEIGHT;
  const maxLeft = window.innerWidth - width - HISTORY_ACTION_MENU_MARGIN;
  const left = Math.min(Math.max(rect.right - width, HISTORY_ACTION_MENU_MARGIN), Math.max(HISTORY_ACTION_MENU_MARGIN, maxLeft));
  const belowTop = rect.bottom + HISTORY_ACTION_MENU_GAP;
  const aboveTop = rect.top - HISTORY_ACTION_MENU_GAP - estimatedHeight;
  const fitsBelow = belowTop + estimatedHeight <= window.innerHeight - HISTORY_ACTION_MENU_MARGIN;
  const fitsAbove = aboveTop >= HISTORY_ACTION_MENU_MARGIN;
  const unclampedTop = fitsBelow || !fitsAbove ? belowTop : aboveTop;
  const maxTop = window.innerHeight - estimatedHeight - HISTORY_ACTION_MENU_MARGIN;
  const top = Math.min(Math.max(unclampedTop, HISTORY_ACTION_MENU_MARGIN), Math.max(HISTORY_ACTION_MENU_MARGIN, maxTop));

  return { top, left, width };
}

function hasExpandableHistoryItemDetails(record: ExecutionHistoryRecord) {
  if (record.itemType === "reminder" && getHistoryBody(record)) return true;
  if ((record.itemType === "meeting" || record.itemType === "teams_meeting") && getHistoryMeetingDetails(record)) return true;
  if (record.itemType === "email" && getHistoryEmailDraftDetails(record)) return true;
  return false;
}

function getRecallUnavailableCopy(record: ExecutionHistoryRecord, recallReason: string | null) {
  if (!record.providerObjectId && record.provider === "local_export") {
    return LOCAL_PROVIDER_RECALL_UNAVAILABLE_COPY;
  }

  return recallReason;
}

function getItemRemovalConfirmationBody(record: ExecutionHistoryRecord) {
  const provider = getProviderLabel(record);
  if (!provider || provider === "Local export") {
    return "This removes the record from History only. It does not change items in connected providers.";
  }

  return `This removes the record from History only. It does not change the item in ${provider}.`;
}

function getGroupRemovalConfirmationTitle(planGroup: PlanExecutionGroup) {
  const count = planGroup.items.length;
  return `Remove ${count} ${count === 1 ? "action" : "actions"} from History?`;
}

function getGroupRemovalConfirmationBody(planGroup: PlanExecutionGroup) {
  const groupName = getPlanDisplayName(planGroup);
  const provider = getConsistentGroupProviderLabel(planGroup);

  if (provider && provider !== "Local export") {
    return `This removes all History records for "${groupName}" only. It does not change the items in ${provider}.`;
  }

  return `This removes all History records for "${groupName}" only. It does not change items in connected providers.`;
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
    addGoogleMeet: Boolean(value.addGoogleMeet),
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

function addMinutesToIso(iso: string, minutes: number) {
  const parsed = new Date(iso);
  if (Number.isNaN(parsed.getTime())) return iso;
  parsed.setMinutes(parsed.getMinutes() + minutes);
  return `${parsed.getFullYear()}-${String(parsed.getMonth() + 1).padStart(2, "0")}-${String(parsed.getDate()).padStart(2, "0")}T${String(parsed.getHours()).padStart(2, "0")}:${String(parsed.getMinutes()).padStart(2, "0")}:00`;
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

function getLatestPlanRescheduleState(planGroup: PlanExecutionGroup) {
  let latestMatch: { appliedAt: string; toEventDate: string; toEventTime: string | null } | null = null;

  for (const item of planGroup.items) {
    const value = item.details.latestPlanReschedule;
    if (!value || typeof value !== "object" || Array.isArray(value)) continue;

    const candidate = value as Record<string, unknown>;
    const appliedAt = readString(candidate.appliedAt);
    const toEventDate = readString(candidate.toEventDate);
    const toEventTime = typeof candidate.toEventTime === "string" ? candidate.toEventTime : null;
    if (!appliedAt || !toEventDate) continue;

    if (!latestMatch || appliedAt > latestMatch.appliedAt) {
      latestMatch = {
        appliedAt,
        toEventDate,
        toEventTime,
      };
    }
  }

  return latestMatch;
}

function getRepresentativeScheduledRecord(planGroup: PlanExecutionGroup, snapshot: ExecutionPlanSnapshot | null) {
  const eventRowIds = new Set(
    (snapshot?.originalRowDefinitions ?? [])
      .filter((row) => (row.dateBasis ?? "event") === "event" && row.offsetDays === 0)
      .map((row) => row.id)
      .filter(Boolean)
  );

  const byScheduledTime = (left: ExecutionHistoryRecord, right: ExecutionHistoryRecord) =>
    (left.scheduledFor || left.executedAt || "").localeCompare(right.scheduledFor || right.executedAt || "");

  const matchingEventRows = planGroup.items
    .filter((item) => {
      const sourceRowId = typeof item.details.sourceRowId === "string" ? item.details.sourceRowId : item.id;
      return eventRowIds.has(sourceRowId) && Boolean(item.scheduledFor || item.executedAt);
    })
    .sort(byScheduledTime);

  if (matchingEventRows.length > 0) return matchingEventRows[0];

  const scheduledItems = planGroup.items.filter((item) => Boolean(item.scheduledFor || item.executedAt)).sort(byScheduledTime);
  return scheduledItems[0] ?? null;
}

function getRepresentativeSourceRow(planGroup: PlanExecutionGroup, snapshot: ExecutionPlanSnapshot | null) {
  if (!snapshot) return null;
  const representativeRecord = getRepresentativeScheduledRecord(planGroup, snapshot);
  if (!representativeRecord) return null;
  const sourceRowId =
    typeof representativeRecord.details.sourceRowId === "string" ? representativeRecord.details.sourceRowId : representativeRecord.id;
  return snapshot.originalRowDefinitions.find((row) => row.id === sourceRowId) ?? null;
}

function getAnchorDateFromDisplayedEventDate(
  planGroup: PlanExecutionGroup,
  snapshot: ExecutionPlanSnapshot | null,
  displayedEventDate: string
) {
  if (!snapshot || !displayedEventDate) return displayedEventDate;

  const representativeSourceRow = getRepresentativeSourceRow(planGroup, snapshot);
  if (!representativeSourceRow || (representativeSourceRow.dateBasis ?? "event") !== "event") {
    return displayedEventDate;
  }

  return addDaysISO(displayedEventDate, -representativeSourceRow.offsetDays);
}

function getCurrentEventDateValue(planGroup: PlanExecutionGroup, snapshot: ExecutionPlanSnapshot | null) {
  const latestReschedule = getLatestPlanRescheduleState(planGroup);
  if (latestReschedule?.toEventDate) return latestReschedule.toEventDate;

  const representativeRecord = getRepresentativeScheduledRecord(planGroup, snapshot);
  const representativeDate = formatDateInputValue(representativeRecord?.scheduledFor || representativeRecord?.executedAt || null);
  if (representativeDate) return representativeDate;

  return snapshot?.anchorDate ?? "";
}

function getCurrentEventTimeValue(planGroup: PlanExecutionGroup, snapshot: ExecutionPlanSnapshot | null) {
  const latestReschedule = getLatestPlanRescheduleState(planGroup);
  if (latestReschedule?.toEventTime) {
    const parsedRescheduleTime = parseTimeInput(latestReschedule.toEventTime);
    if (parsedRescheduleTime) {
      return parsedRescheduleTime;
    }
  }

  const representativeRecord = getRepresentativeScheduledRecord(planGroup, snapshot);
  const representativeTime = formatTimeInputValue(representativeRecord?.scheduledFor || representativeRecord?.executedAt || null);
  if (representativeTime) return representativeTime;

  if (!snapshot?.anchorDate) return "";

  const anchorTime =
    snapshot.anchorValues.find((anchor) => normalizeAnchorKey(anchor.key) === normalizeAnchorKey("Dissemination Time"))?.value ||
    snapshot.anchorValues.find((anchor) => normalizeAnchorKey(anchor.key) === normalizeAnchorKey("Earnings Call Time"))?.value ||
    "";
  const parsedAnchorTime = parseTimeInput(anchorTime);
  if (parsedAnchorTime) {
    return parsedAnchorTime;
  }

  const matchingEventDayRecord = planGroup.items.find((item) => {
    const itemDate = formatDateInputValue(item.scheduledFor || item.executedAt);
    return itemDate === snapshot.anchorDate;
  });

  return matchingEventDayRecord ? formatTimeInputValue(matchingEventDayRecord.scheduledFor || matchingEventDayRecord.executedAt) : "";
}

function isUnavailableHistoryItem(record: ExecutionHistoryRecord) {
  return record.status === "recalled" || record.status === "already_removed" || record.status === "already_canceled";
}

function isUnavailablePlanGroup(planGroup: PlanExecutionGroup) {
  return planGroup.items.length > 0 && planGroup.items.every((item) => isUnavailableHistoryItem(item));
}

function shouldShowPlanMessage(planMessage: { tone: "success" | "warning" | "error"; text: string; helperText?: string } | null) {
  if (!planMessage) return false;
  return true;
}

function shouldShowPlanHelperText(planMessage: { tone: "success" | "warning" | "error"; text: string; helperText?: string } | null) {
  if (!planMessage?.helperText) return false;
  return true;
}

function buildUpdatedSnapshotAnchors(snapshot: ExecutionPlanSnapshot, nextEventDate: string, nextEventTime?: string) {
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
    if (nextEventTime && normalizedKey === normalizeAnchorKey("Dissemination Time")) {
      return { ...anchor, value: nextEventTime };
    }
    if (nextEventTime && normalizedKey === normalizeAnchorKey("Earnings Call Time")) {
      return { ...anchor, value: nextEventTime };
    }
    return anchor;
  });
}

function buildTemplateItemsFromSnapshot(
  snapshot: ExecutionPlanSnapshot,
  options?: { representativeRowId?: string | null; eventTimeOverride?: string }
): TemplateItem[] {
  const normalizedOverrideTime = options?.eventTimeOverride ? parseTimeInput(options.eventTimeOverride) : null;

  return snapshot.originalRowDefinitions.map((row) => ({
    id: row.id,
    title: row.title,
    body: row.body || undefined,
    offsetDays: row.offsetDays,
    dateBasis: row.dateBasis ?? "event",
    rowType: row.rowType ?? "reminder",
    reminderTime:
      normalizedOverrideTime && options?.representativeRowId === row.id
        ? normalizedOverrideTime
        : row.reminderTime
          ? normalizeReminderTimeInput(row.reminderTime)
          : undefined,
    emailDraft: normalizeEmailDraftValue(row.emailDraft),
    durationDraft: normalizeDurationDraftValue(row.durationDraft),
    meetingDraft: normalizeMeetingDraftValue(row.meetingDraft),
  }));
}

function buildRescheduledPlan(
  snapshot: ExecutionPlanSnapshot,
  nextEventDate: string,
  nextEventTime?: string,
  weekendRule?: WeekendRule,
  options?: { representativeRowId?: string | null }
) {
  const templateItems = buildTemplateItemsFromSnapshot(snapshot, {
    representativeRowId: options?.representativeRowId ?? null,
    eventTimeOverride: nextEventTime,
  });
  const plan = createPlan({
    name: snapshot.eventName || snapshot.templateName || "Untitled plan",
    type: snapshot.templateBaseType,
    anchorDate: nextEventDate,
    weekendRule: weekendRule ?? snapshot.weekendRule,
    template: templateItems,
  });
  const anchorMap = buildAnchorMap(buildUpdatedSnapshotAnchors(snapshot, nextEventDate, nextEventTime));
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

function getPlanModifyPreview(
  planGroup: PlanExecutionGroup,
  nextEventDate: string,
  nextEventTime?: string,
  weekendRule?: WeekendRule
): { snapshot: ExecutionPlanSnapshot | null; items: PlanReschedulePreviewItem[] } {
  const firstRecord = planGroup.items[0];
  const snapshot = firstRecord ? getExecutionPlanSnapshot(firstRecord) : null;
  if (!snapshot || !nextEventDate) {
    return { snapshot, items: [] };
  }

  const currentEventDate = getCurrentEventDateValue(planGroup, snapshot);
  const currentEventTime = getCurrentEventTimeValue(planGroup, snapshot);
  const representativeSourceRow = getRepresentativeSourceRow(planGroup, snapshot);
  const currentAnchorDate = getAnchorDateFromDisplayedEventDate(planGroup, snapshot, currentEventDate);
  const nextAnchorDate = getAnchorDateFromDisplayedEventDate(planGroup, snapshot, nextEventDate);
  const currentRescheduledPlan = currentEventDate
    ? buildRescheduledPlan(snapshot, currentAnchorDate, currentEventTime, weekendRule ?? snapshot.weekendRule, {
        representativeRowId: representativeSourceRow?.id ?? null,
      })
    : null;
  const currentItemsByRowId = new Map((currentRescheduledPlan?.items ?? []).map((item) => [item.id, item]));
  const rescheduledPlan = buildRescheduledPlan(snapshot, nextAnchorDate, nextEventTime, weekendRule, {
    representativeRowId: representativeSourceRow?.id ?? null,
  });
  const nextItemsByRowId = new Map(rescheduledPlan.items.map((item) => [item.id, item]));

  const items = planGroup.items.map((record) => {
    const sourceRowId = typeof record.details.sourceRowId === "string" ? record.details.sourceRowId : record.id;
    const nextItem = nextItemsByRowId.get(sourceRowId) ?? null;
    const currentItem = currentItemsByRowId.get(sourceRowId) ?? null;
    const modifyState = getExecutionHistoryModifyState(record);
    const overrideState = getOverrideState(record);
    const currentTiming = currentItem ? getComputedPlanItemTiming(currentItem) : null;
    const oldDateTime = currentTiming?.scheduledFor ?? record.scheduledFor ?? record.executedAt;
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

    if (record.itemType === "meeting" || record.itemType === "teams_meeting") {
      const meetingEnd = new Date(record.endsAt || record.scheduledFor || "");
      if (!Number.isNaN(meetingEnd.getTime()) && meetingEnd.getTime() < Date.now()) {
        return {
          record,
          nextItem,
          action: "Locked" as const,
          reason: "This meeting already happened, so it will stay as-is.",
          oldDateTime,
          newDateTime: oldDateTime,
          isOverridden: false,
        };
      }
    }

    if (actionName === "email_sent") {
      return {
        record,
        nextItem,
        action: "Locked" as const,
        reason: "This email has already been sent, so it will stay as-is.",
        oldDateTime,
        newDateTime,
        isOverridden: false,
      };
    }

    if (record.itemType === "email" && actionName === "email_scheduled") {
      const scheduledEmailState = typeof record.details.scheduledEmailState === "string" ? record.details.scheduledEmailState : "";
      if (scheduledEmailState === "sent") {
        return {
          record,
          nextItem,
          action: "Locked" as const,
          reason: "This email has already been sent, so it will stay as-is.",
          oldDateTime,
          newDateTime: oldDateTime,
          isOverridden: false,
        };
      }
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
        reason: "Reminder will be moved to the new scheduled date.",
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

function getPlanModifyEligibility(planGroup: PlanExecutionGroup) {
  const firstRecord = planGroup.items[0];
  const snapshot = firstRecord ? getExecutionPlanSnapshot(firstRecord) : null;
  if (!snapshot || !snapshot.anchorDate) {
    return {
      canModifyPlan: false,
      snapshot,
      baselinePreview: [] as PlanReschedulePreviewItem[],
    };
  }

  const baselinePreview = getPlanModifyPreview(
    planGroup,
    getCurrentEventDateValue(planGroup, snapshot),
    getCurrentEventTimeValue(planGroup, snapshot),
    snapshot.weekendRule
  ).items;
  const canModifyPlan = planGroup.items.some((item) => {
    const modifyState = getExecutionHistoryModifyState(item);
    return modifyState.canModify && modifyState.modifyImplemented;
  });

  return {
    canModifyPlan,
    snapshot,
    baselinePreview,
  };
}

function getPlanModifyResultMessage(
  items: PlanReschedulePreviewItem[],
  counts: { updatedCount: number; replacedCount: number; failedCount: number }
) {
  const actionableCount = counts.updatedCount + counts.replacedCount;
  if (counts.failedCount === 0 && actionableCount > 0) {
    const skippedCount = items.filter((item) => item.action === "Locked" || item.action === "Unsupported" || item.action === "Unchanged").length;
    return {
      tone: skippedCount > 0 ? ("warning" as const) : ("success" as const),
      text: "Event updated.",
    };
  }

  if (actionableCount > 0) {
    return {
      tone: "warning" as const,
      text: "Event updated.",
    };
  }

  if (counts.failedCount > 0) {
    return {
      tone: "error" as const,
      text: `Update failed for ${counts.failedCount} item${counts.failedCount === 1 ? "" : "s"}.`,
    };
  }

  const unchangedCount = items.filter((item) => item.action === "Unchanged").length;
  const lockedCount = items.filter((item) => item.action === "Locked").length;
  const unsupportedCount = items.filter((item) => item.action === "Unsupported").length;

  if (unchangedCount > 0 && lockedCount === 0 && unsupportedCount === 0) {
    return {
      tone: "warning" as const,
      text: "Nothing changed.",
    };
  }

  const helperParts: string[] = [];
  if (unsupportedCount > 0) {
    helperParts.push(`${unsupportedCount} item${unsupportedCount === 1 ? "" : "s"} can't be modified from History`);
  }
  if (lockedCount > 0) {
    helperParts.push(`${lockedCount} item${lockedCount === 1 ? "" : "s"} must stay as-is`);
  }
  if (unchangedCount > 0) {
    helperParts.push(`${unchangedCount} item${unchangedCount === 1 ? "" : "s"} already match the current schedule`);
  }

  return {
    tone: "warning" as const,
    text: "No changes were applied.",
    helperText: helperParts.length > 0 ? `${helperParts.join(". ")}.` : undefined,
  };
}

function getPlanRecallUnavailableMessage(planGroup: PlanExecutionGroup) {
  const reasons = Array.from(
    new Set(
      planGroup.items
        .map((item) => getExecutionHistoryRecallState(item).recallReason)
        .filter((reason): reason is string => Boolean(reason))
    )
  );

  if (reasons.length === 0) {
    return {
      tone: "warning" as const,
      text: "Nothing to recall.",
    };
  }

  return {
    tone: "warning" as const,
    text: reasons[0],
    helperText: reasons.length > 1 ? reasons.slice(1).join(" ") : undefined,
  };
}

function getPlanModifyAvailabilityMessage(items: PlanReschedulePreviewItem[]) {
  const actionableCount = items.filter((item) => item.action === "Update" || item.action === "Replace").length;
  if (actionableCount > 0) return null;

  const reasons = Array.from(new Set(items.map((item) => item.reason).filter((reason): reason is string => Boolean(reason))));
  if (reasons.length === 0) return null;

  return {
    tone: "warning" as const,
    text: reasons[0],
    helperText: reasons.length > 1 ? reasons.slice(1).join(" ") : undefined,
  };
}

export default function HistoryPage() {
  const exposeModifyUI = true;
  const { authEnabled, currentUser, currentOrgId } = useAuthContext();
  const [records, setRecords] = useState<ExecutionHistoryRecord[]>([]);
  const [hasHistorySnapshot, setHasHistorySnapshot] = useState(false);
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState(false);
  const [staleRefreshFailed, setStaleRefreshFailed] = useState(false);
  const [refreshKey, setRefreshKey] = useState(0);
  const [expandedPlans, setExpandedPlans] = useState<Record<string, boolean>>({});
  const [expandedItems, setExpandedItems] = useState<Record<string, boolean>>({});
  const [openActionMenu, setOpenActionMenu] = useState<HistoryActionMenuState | null>(null);
  const [modifyingPlans, setModifyingPlans] = useState<Record<string, boolean>>({});
  const [planModifyDates, setPlanModifyDates] = useState<Record<string, string>>({});
  const [planModifyTimes, setPlanModifyTimes] = useState<Record<string, string>>({});
  const [planModifyWeekendRules, setPlanModifyWeekendRules] = useState<Record<string, WeekendRule>>({});
  const [pendingPlanRecalls, setPendingPlanRecalls] = useState<Record<string, boolean>>({});
  const [pendingItemRecalls, setPendingItemRecalls] = useState<Record<string, boolean>>({});
  const [pendingPlanDeletes, setPendingPlanDeletes] = useState<Record<string, boolean>>({});
  const [pendingItemDeletes, setPendingItemDeletes] = useState<Record<string, boolean>>({});
  const [pendingPlanModifies, setPendingPlanModifies] = useState<Record<string, boolean>>({});
  const [planMessages, setPlanMessages] = useState<Record<string, { tone: "success" | "warning" | "error"; text: string; helperText?: string }>>({});
  const [itemMessages, setItemMessages] = useState<Record<string, { tone: "success" | "error" | "neutral"; text: string }>>({});
  const recordsRef = useRef(records);
  const hasHistorySnapshotRef = useRef(hasHistorySnapshot);
  const actionMenuRef = useRef<HTMLDivElement | null>(null);
  const actionMenuTriggerRef = useRef<HTMLButtonElement | null>(null);
  const focusActionMenuOnOpenRef = useRef(false);
  const ignoreNextActionMenuClickRef = useRef(false);
  const confirmationDialogRef = useRef<HTMLDivElement | null>(null);
  const confirmationCancelRef = useRef<HTMLButtonElement | null>(null);
  const confirmationOpenerRef = useRef<HTMLElement | null>(null);
  const [confirmationDialog, setConfirmationDialog] = useState<ConfirmationDialogState | null>(null);
  const [confirmationPending, setConfirmationPending] = useState(false);
  const [mounted, setMounted] = useState(false);

  useEffect(() => {
    recordsRef.current = records;
  }, [records]);

  useEffect(() => {
    hasHistorySnapshotRef.current = hasHistorySnapshot;
  }, [hasHistorySnapshot]);

  useEffect(() => {
    setMounted(true);
  }, []);

  useEffect(() => {
    let cancelled = false;

    async function load() {
      console.info("[historyPage] fetch start", {
        authEnabled,
        userId: currentUser?.id ?? null,
        orgId: currentOrgId ?? null,
      });
      setLoadError(false);
      setStaleRefreshFailed(false);

      const cachedSnapshot = readCachedExecutionHistorySnapshot();
      if (!cancelled) {
        if (cachedSnapshot.hasSnapshot || recordsRef.current.length === 0) {
          recordsRef.current = cachedSnapshot.records;
          hasHistorySnapshotRef.current = cachedSnapshot.hasSnapshot;
          setRecords(cachedSnapshot.records);
          setHasHistorySnapshot(cachedSnapshot.hasSnapshot);
          setLoading(cachedSnapshot.records.length === 0 && !cachedSnapshot.hasSnapshot);
        } else {
          setLoading(false);
        }
      }

      try {
        const nextRecords = await listExecutionHistory();
        console.info("[historyPage] fetch success", {
          count: nextRecords.length,
          days: Array.from(new Set(nextRecords.map((record) => getLocalDayKey(record.executedAt)))),
          orgId: currentOrgId ?? null,
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
        setHasHistorySnapshot(true);
        setExpandedPlans((current) => {
          if (Object.keys(current).length > 0) return current;
          return {};
        });
      } catch (error) {
        console.error("[historyPage] fetch failed", error);
        if (cancelled) return;
        if (recordsRef.current.length > 0) {
          setStaleRefreshFailed(true);
        } else if (!hasHistorySnapshotRef.current) {
          setLoadError(true);
          setRecords([]);
        }
      } finally {
        if (!cancelled) {
          setLoading(false);
        }
      }
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
  }, [authEnabled, currentUser?.id, currentOrgId, refreshKey]);

  useEffect(() => {
    if (!openActionMenu) return;

    function handleClickOutside(event: MouseEvent) {
      const target = event.target as Node;
      if (actionMenuRef.current?.contains(target)) return;
      if (actionMenuTriggerRef.current?.contains(target)) return;
      setOpenActionMenu(null);
      actionMenuTriggerRef.current = null;
    }

    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, [openActionMenu]);

  useEffect(() => {
    if (!openActionMenu) return;

    function closeForViewportChange() {
      setOpenActionMenu(null);
      actionMenuTriggerRef.current = null;
    }

    window.addEventListener("resize", closeForViewportChange);
    window.addEventListener("scroll", closeForViewportChange, true);
    return () => {
      window.removeEventListener("resize", closeForViewportChange);
      window.removeEventListener("scroll", closeForViewportChange, true);
    };
  }, [openActionMenu]);

  useEffect(() => {
    if (!openActionMenu || !focusActionMenuOnOpenRef.current) return;

    const focusTimer = window.setTimeout(() => {
      const firstItem = actionMenuRef.current?.querySelector<HTMLButtonElement>('button[role="menuitem"]:not(:disabled)');
      firstItem?.focus();
      focusActionMenuOnOpenRef.current = false;
    }, 0);

    return () => window.clearTimeout(focusTimer);
  }, [openActionMenu]);

  useEffect(() => {
    if (!confirmationDialog) return;

    const previousOverflow = document.body.style.overflow;
    document.body.style.overflow = "hidden";

    const focusTimer = window.setTimeout(() => {
      const dialogNode = confirmationDialogRef.current;
      const firstFocusable = dialogNode?.querySelector<HTMLElement>(
        'button:not([disabled]), [href], input:not([disabled]), select:not([disabled]), textarea:not([disabled]), [tabindex]:not([tabindex="-1"])'
      );
      (confirmationCancelRef.current ?? firstFocusable)?.focus();
    }, 0);

    function handleKeyDown(event: KeyboardEvent) {
      const dialogNode = confirmationDialogRef.current;
      if (!dialogNode) return;

      if (event.key === "Escape") {
        if (!confirmationPending) {
          event.preventDefault();
          setConfirmationDialog(null);
        }
        return;
      }

      if (event.key !== "Tab") return;

      const focusable = Array.from(
        dialogNode.querySelectorAll<HTMLElement>(
          'button:not([disabled]), [href], input:not([disabled]), select:not([disabled]), textarea:not([disabled]), [tabindex]:not([tabindex="-1"])'
        )
      );
      if (focusable.length === 0) return;

      const first = focusable[0];
      const last = focusable[focusable.length - 1];
      if (event.shiftKey && document.activeElement === first) {
        event.preventDefault();
        last.focus();
      } else if (!event.shiftKey && document.activeElement === last) {
        event.preventDefault();
        first.focus();
      }
    }

    document.addEventListener("keydown", handleKeyDown);
    return () => {
      window.clearTimeout(focusTimer);
      document.body.style.overflow = previousOverflow;
      document.removeEventListener("keydown", handleKeyDown);
    };
  }, [confirmationDialog, confirmationPending]);

  useEffect(() => {
    if (confirmationDialog) return;
    const opener = confirmationOpenerRef.current;
    confirmationOpenerRef.current = null;
    if (!opener) return;
    window.setTimeout(() => {
      if (document.contains(opener)) opener.focus();
    }, 0);
  }, [confirmationDialog]);

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

  const attentionCount = useMemo(() => getAttentionCount(records), [records]);

  function retryHistoryLoad() {
    setRefreshKey((current) => current + 1);
  }

  function closeActionMenu(options?: { restoreFocus?: boolean }) {
    const trigger = actionMenuTriggerRef.current;
    setOpenActionMenu(null);
    actionMenuTriggerRef.current = null;

    if (options?.restoreFocus && trigger && document.contains(trigger)) {
      window.setTimeout(() => {
        if (document.contains(trigger)) trigger.focus();
      }, 0);
    }
  }

  function openActionMenuFromTrigger(
    trigger: HTMLButtonElement,
    kind: HistoryActionMenuKind,
    id: string,
    itemCount: number,
    focusFirstItem: boolean
  ) {
    const isAlreadyOpen = openActionMenu?.kind === kind && openActionMenu.id === id;
    if (isAlreadyOpen) {
      closeActionMenu();
      return;
    }

    actionMenuTriggerRef.current = trigger;
    focusActionMenuOnOpenRef.current = focusFirstItem;
    setOpenActionMenu({
      kind,
      id,
      position: getHistoryActionMenuPosition(trigger, itemCount),
    });
  }

  function handleActionMenuTriggerClick(
    event: ReactMouseEvent<HTMLButtonElement>,
    kind: HistoryActionMenuKind,
    id: string,
    itemCount: number
  ) {
    if (ignoreNextActionMenuClickRef.current) {
      ignoreNextActionMenuClickRef.current = false;
      return;
    }

    openActionMenuFromTrigger(event.currentTarget, kind, id, itemCount, false);
  }

  function handleActionMenuTriggerKeyDown(
    event: ReactKeyboardEvent<HTMLButtonElement>,
    kind: HistoryActionMenuKind,
    id: string,
    itemCount: number
  ) {
    if (event.key === "ArrowDown" || event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      if (event.key === "Enter" || event.key === " ") {
        ignoreNextActionMenuClickRef.current = true;
      }
      openActionMenuFromTrigger(event.currentTarget, kind, id, itemCount, true);
    } else if (event.key === "Escape" && openActionMenu?.kind === kind && openActionMenu.id === id) {
      event.preventDefault();
      closeActionMenu({ restoreFocus: true });
    }
  }

  function handleActionMenuKeyDown(event: ReactKeyboardEvent<HTMLDivElement>) {
    const menuNode = actionMenuRef.current;
    const items = Array.from(menuNode?.querySelectorAll<HTMLButtonElement>('button[role="menuitem"]:not(:disabled)') ?? []);
    const currentIndex = items.findIndex((item) => item === document.activeElement);

    if (event.key === "Escape") {
      event.preventDefault();
      closeActionMenu({ restoreFocus: true });
      return;
    }

    if (event.key === "Tab") {
      setOpenActionMenu(null);
      actionMenuTriggerRef.current = null;
      return;
    }

    if (items.length === 0) return;

    if (event.key === "ArrowDown") {
      event.preventDefault();
      items[(currentIndex + 1 + items.length) % items.length]?.focus();
    } else if (event.key === "ArrowUp") {
      event.preventDefault();
      items[(currentIndex - 1 + items.length) % items.length]?.focus();
    } else if (event.key === "Home") {
      event.preventDefault();
      items[0]?.focus();
    } else if (event.key === "End") {
      event.preventDefault();
      items[items.length - 1]?.focus();
    }
  }

  function getMenuItemClasses(tone: HistoryActionMenuAction["tone"]) {
    if (tone === "history") {
      return "text-red-700 hover:bg-red-50 focus:ring-red-500/25 disabled:text-slate-400";
    }
    if (tone === "provider") {
      return "text-amber-700 hover:bg-amber-50 focus:ring-amber-500/25 disabled:text-slate-400";
    }
    return "text-slate-700 hover:bg-slate-50 focus:ring-slate-500/25 disabled:text-slate-400";
  }

  function renderActionMenu(kind: HistoryActionMenuKind, id: string, label: string, actions: HistoryActionMenuAction[]) {
    if (!mounted || !openActionMenu || openActionMenu.kind !== kind || openActionMenu.id !== id || actions.length === 0) return null;

    return createPortal(
      <div
        ref={actionMenuRef}
        role="menu"
        aria-label={label}
        onKeyDown={handleActionMenuKeyDown}
        className="fixed z-[120] rounded-[10px] border border-slate-200/90 bg-white p-[6px] text-left shadow-[0_16px_42px_rgba(21,40,66,0.16)]"
        style={{
          left: openActionMenu.position.left,
          top: openActionMenu.position.top,
          width: openActionMenu.position.width,
        }}
      >
        {actions.map((action) => (
          <button
            key={action.key}
            type="button"
            role="menuitem"
            disabled={action.disabled}
            title={action.title}
            onClick={() => {
              if (action.disabled) return;
              action.onSelect();
            }}
            className={`flex h-[40px] w-full items-center rounded-[8px] px-[11px] text-left text-[14px] font-medium transition focus:outline-none focus:ring-2 disabled:cursor-not-allowed ${getMenuItemClasses(
              action.tone
            )}`}
          >
            {action.label}
          </button>
        ))}
      </div>,
      document.body
    );
  }

  function requestConfirmation(dialog: ConfirmationDialogState) {
    const menuTrigger = actionMenuTriggerRef.current;
    confirmationOpenerRef.current =
      menuTrigger && document.contains(menuTrigger) ? menuTrigger : document.activeElement instanceof HTMLElement ? document.activeElement : null;
    setOpenActionMenu(null);
    actionMenuTriggerRef.current = null;
    setConfirmationDialog(dialog);
  }

  async function runConfirmationAction() {
    if (!confirmationDialog) return;
    setConfirmationPending(true);
    try {
      await confirmationDialog.onConfirm();
      setConfirmationDialog(null);
    } catch (error) {
      console.error("[historyPage] confirmed action failed", error);
    } finally {
      setConfirmationPending(false);
    }
  }

  function togglePlanModify(planGroup: PlanExecutionGroup, nextOpen?: boolean) {
    const shouldOpen = nextOpen ?? !(modifyingPlans[planGroup.key] ?? false);
    setExpandedPlans((current) => ({ ...current, [planGroup.key]: shouldOpen || current[planGroup.key] || false }));
    setModifyingPlans((current) => ({ ...current, [planGroup.key]: shouldOpen }));
    if (shouldOpen) {
      const snapshot = getExecutionPlanSnapshot(planGroup.items[0]);
      setPlanModifyDates((current) => ({
        ...current,
        [planGroup.key]: current[planGroup.key] ?? getCurrentEventDateValue(planGroup, snapshot),
      }));
      setPlanModifyTimes((current) => ({
        ...current,
        [planGroup.key]: current[planGroup.key] ?? getCurrentEventTimeValue(planGroup, snapshot),
      }));
      setPlanModifyWeekendRules((current) => ({
        ...current,
        [planGroup.key]: current[planGroup.key] ?? snapshot?.weekendRule ?? "prior_business_day",
      }));
    }
  }

  async function recallHistoryItem(record: ExecutionHistoryRecord): Promise<HistoryRecallResult> {
    const recallState = getExecutionHistoryRecallState(record);
    if (!recallState.canRecall || !recallState.recallImplemented || !record.providerObjectId) {
      throw new Error(recallState.recallReason || "This item cannot be recalled.");
    }

    if (record.provider === "gmail") {
      if (record.providerObjectType === "event") {
        return await deleteGoogleCalendarEvent({
          eventId: record.providerObjectId,
        });
      }
      throw new Error("This Google item cannot be recalled.");
    }

    if (record.providerObjectType === "message") {
      return await deleteOutlookMessage({
        messageId: record.providerObjectId,
        requireDraft: getHistoryAction(record) === "email_scheduled",
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

  function isAlreadySentEmailError(error: unknown) {
    const message = error instanceof Error ? error.message.toLowerCase() : String(error || "").toLowerCase();
    return (
      message.includes("already been sent") ||
      message.includes("can't be recalled") ||
      message.includes("can't be changed")
    );
  }

  async function persistScheduledEmailSentState(record: ExecutionHistoryRecord, error: unknown) {
    await updateExecutionHistoryRecord(record.id, {
      details: {
        scheduledEmailState: "sent",
        recallError: error instanceof Error ? error.message : null,
        planRescheduleError: error instanceof Error ? error.message : null,
      },
    });
  }

  function getRecallStatusMessage(result: HistoryRecallResult) {
    if (result === "already_removed") {
      return { status: "already_removed" as const, text: "This item is no longer available.", tone: "neutral" as const };
    }
    if (result === "already_canceled") {
      return { status: "already_canceled" as const, text: "This item is no longer available.", tone: "neutral" as const };
    }
    return { status: "recalled" as const, text: "Recalled.", tone: "success" as const };
  }

  async function performRecallItem(record: ExecutionHistoryRecord) {
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
        [record.id]: { tone: statusMessage.tone, text: statusMessage.text },
      }));
    } catch (error) {
      if (record.itemType === "email" && getHistoryAction(record) === "email_scheduled" && isAlreadySentEmailError(error)) {
        await persistScheduledEmailSentState(record, error);
        setItemMessages((current) => ({
          ...current,
          [record.id]: { tone: "neutral", text: "This email has already been sent and can't be recalled." },
        }));
        return;
      }
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

  function handleRecallItem(record: ExecutionHistoryRecord) {
    const recallState = getExecutionHistoryRecallState(record);
    if (!recallState.canRecall || !recallState.recallImplemented) return;
    const confirmation = getProviderRecallConfirmation(record);
    requestConfirmation({
      title: confirmation.title,
      body: confirmation.body,
      confirmLabel: confirmation.confirmLabel,
      tone: "provider",
      onConfirm: () => performRecallItem(record),
    });
  }

  async function performRecallPlan(planGroup: PlanExecutionGroup, recallableItems: ExecutionHistoryRecord[]) {
    setPendingPlanRecalls((current) => ({ ...current, [planGroup.key]: true }));
    setPlanMessages((current) => {
      const next = { ...current };
      delete next[planGroup.key];
      return next;
    });

    let successCount = 0;
    let unavailableCount = 0;
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
          unavailableCount += 1;
        }
      } catch (error) {
        if (item.itemType === "email" && getHistoryAction(item) === "email_scheduled" && isAlreadySentEmailError(error)) {
          await persistScheduledEmailSentState(item, error);
          continue;
        }
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
        failedCount === 0 && successCount === 0 && unavailableCount > 0
          ? {
              tone: "warning",
              text: "Nothing to recall.",
              helperText: "Some items were already unavailable.",
            }
          : failedCount === 0
          ? {
              tone: "success",
              text: "Event recalled.",
              helperText: unavailableCount > 0 ? "Some items were already unavailable." : undefined,
            }
          : successCount > 0
            ? {
                tone: "warning",
                text: "Partially recalled.",
                helperText: unavailableCount > 0 ? "Some items were already unavailable." : undefined,
              }
            : { tone: "error", text: "Recall failed." },
    }));
  }

  function handleRecallPlan(planGroup: PlanExecutionGroup) {
    const recallableItems = planGroup.items.filter((item) => {
      const recallState = getExecutionHistoryRecallState(item);
      return recallState.canRecall && recallState.recallImplemented;
    });

    if (recallableItems.length === 0) {
      setPlanMessages((current) => ({
        ...current,
        [planGroup.key]: getPlanRecallUnavailableMessage(planGroup),
      }));
      return;
    }

    const providerLabel = getConsistentGroupProviderLabel(planGroup);
    requestConfirmation({
      title:
        providerLabel === "Google Calendar"
          ? "Remove these events from Google Calendar?"
          : providerLabel === "Outlook"
            ? "Recall these items from Outlook?"
            : "Recall these provider items?",
      body:
        providerLabel === "Google Calendar"
          ? "This attempts to remove the supported calendar events. History records will remain and update with the result."
          : providerLabel === "Outlook"
            ? "This attempts to remove supported items from Outlook. History records will remain and update with the result."
            : "This attempts to change supported connected-provider items. History records will remain and update with the result.",
      confirmLabel: providerLabel === "Google Calendar" ? "Remove events" : "Recall items",
      tone: "provider",
      onConfirm: () => performRecallPlan(planGroup, recallableItems),
    });
  }

  async function performDeleteItem(record: ExecutionHistoryRecord) {
    setPendingItemDeletes((current) => ({ ...current, [record.id]: true }));
    try {
      await deleteExecutionHistoryRecord(record.id);
      setItemMessages((current) => {
        const next = { ...current };
        delete next[record.id];
        return next;
      });
    } finally {
      setPendingItemDeletes((current) => ({ ...current, [record.id]: false }));
    }
  }

  function handleDeleteItem(record: ExecutionHistoryRecord) {
    requestConfirmation({
      title: "Remove from History?",
      body: getItemRemovalConfirmationBody(record),
      confirmLabel: "Remove from History",
      tone: "destructive",
      onConfirm: () => performDeleteItem(record),
    });
  }

  async function performDeletePlanGroup(planGroup: PlanExecutionGroup) {
    setPendingPlanDeletes((current) => ({ ...current, [planGroup.key]: true }));
    try {
      await deleteExecutionHistoryRecords(planGroup.items.map((item) => item.id));
      setPlanMessages((current) => {
        const next = { ...current };
        delete next[planGroup.key];
        return next;
      });
    } finally {
      setPendingPlanDeletes((current) => ({ ...current, [planGroup.key]: false }));
    }
  }

  function handleDeletePlanGroup(planGroup: PlanExecutionGroup) {
    requestConfirmation({
      title: getGroupRemovalConfirmationTitle(planGroup),
      body: getGroupRemovalConfirmationBody(planGroup),
      confirmLabel: "Remove from History",
      tone: "destructive",
      onConfirm: () => performDeletePlanGroup(planGroup),
    });
  }

  async function performClearHistory() {
    setLoading(true);
    try {
      await clearExecutionHistory();
      setRecords([]);
      setHasHistorySnapshot(true);
      setExpandedPlans({});
      setExpandedItems({});
      setPlanMessages({});
      setItemMessages({});
    } finally {
      setLoading(false);
    }
  }

  function handleClearHistory() {
    requestConfirmation({
      title: "Clear all History?",
      body: "This removes all History records from this workspace. It does not recall or delete items from connected providers.",
      confirmLabel: "Clear History",
      tone: "destructive",
      onConfirm: performClearHistory,
    });
  }

  async function handleModifyPlan(planGroup: PlanExecutionGroup) {
    const nextEventDate = planModifyDates[planGroup.key] ?? "";
    const nextEventTime = planModifyTimes[planGroup.key] ?? "";
    const nextWeekendRule = planModifyWeekendRules[planGroup.key] ?? "prior_business_day";
    const { snapshot, items } = getPlanModifyPreview(planGroup, nextEventDate, nextEventTime, nextWeekendRule);
    if (!snapshot || !nextEventDate) return;

    setPendingPlanModifies((current) => ({ ...current, [planGroup.key]: true }));
    setPlanMessages((current) => {
      const next = { ...current };
      delete next[planGroup.key];
      return next;
    });

    let updatedCount = 0;
    let replacedCount = 0;
    let failedCount = 0;

    for (const previewItem of items) {
      const { record, nextItem, action } = previewItem;
      if (!nextItem || action === "Locked" || action === "Unsupported" || action === "Unchanged") {
        continue;
      }

      try {
        const nextTiming = getComputedPlanItemTiming(nextItem);
        const rescheduleDetails = {
          appliedAt: new Date().toISOString(),
          fromEventDate: snapshot.anchorDate,
          toEventDate: nextEventDate,
          toEventTime: nextEventTime || null,
          weekendRule: nextWeekendRule,
          action,
        };

        if (record.itemType === "meeting" || record.itemType === "teams_meeting") {
          if (record.provider === "gmail") {
            await updateGoogleCalendarEvent({
              eventId: record.providerObjectId || "",
              subject: nextItem.customTitle || nextItem.title,
              bodyText: nextItem.body ?? "",
              startISO: nextTiming.scheduledFor,
              endISO: nextTiming.endsAt || addMinutesToIso(nextTiming.scheduledFor, 30),
              timeZone: "America/New_York",
              isAllDay: nextTiming.isAllDay,
              location: nextItem.meetingDraft?.location ?? "",
              attendees: nextItem.meetingDraft?.attendees ?? [],
              addGoogleMeet: Boolean(nextItem.meetingDraft?.addGoogleMeet),
            });
          } else {
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
          }

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
                addGoogleMeet: Boolean(nextItem.meetingDraft?.addGoogleMeet),
                teamsMeeting: Boolean(nextItem.meetingDraft?.teamsMeeting),
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
          const createResult =
            record.provider === "gmail"
              ? await createGoogleCalendarEvent({
                  subject: nextItem.customTitle || nextItem.title,
                  bodyText: nextItem.body ?? "",
                  startISO: nextTiming.scheduledFor,
                  endISO: nextTiming.endsAt || addMinutesToIso(nextTiming.scheduledFor, 30),
                  timeZone: "America/New_York",
                  isAllDay: nextTiming.isAllDay,
                })
              : await createOutlookCalendarEvent({
                  subject: nextItem.customTitle || nextItem.title,
                  bodyText: nextItem.body ?? "",
                  startISO: nextTiming.scheduledFor,
                  endISO: nextTiming.endsAt || addMinutesToIso(nextTiming.scheduledFor, 30),
                  timeZone: "America/New_York",
                  isAllDay: nextTiming.isAllDay,
                });

          if (record.providerObjectId) {
            if (record.provider === "gmail") {
              await deleteGoogleCalendarEvent({
                eventId: record.providerObjectId,
              });
            } else {
              await deleteOutlookCalendarEvent({
                eventId: record.providerObjectId,
              });
            }
          }

          const replacedProviderObjectId = record.providerObjectId;
          await updateExecutionHistoryRecord(record.id, {
            status: "modified",
            title: nextItem.customTitle || nextItem.title,
            subject: nextItem.customTitle || nextItem.title,
            providerObjectId: createResult.id,
            outlookWebLink: record.provider === "outlook" ? createResult.webLink : null,
            teamsJoinLink: record.provider === "outlook" ? createResult.joinUrl || null : null,
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

        continue;
      } catch (error) {
        failedCount += 1;
        if (record.itemType === "email" && getHistoryAction(record) === "email_scheduled" && isAlreadySentEmailError(error)) {
          await persistScheduledEmailSentState(record, error);
        }
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
      [planGroup.key]: getPlanModifyResultMessage(items, {
        updatedCount,
        replacedCount,
        failedCount,
      }),
    }));
  }

  return (
    <div className="mx-auto w-full max-w-[960px] min-w-0 pb-[44px] pt-[28px] text-slate-900">
      <header className="mb-[20px] flex flex-col items-start justify-between gap-[16px] sm:flex-row sm:gap-[24px]">
        <div className="min-w-0">
          <h1 className="text-[30px] font-bold leading-[1.08] text-slate-950 sm:text-[34px]">History</h1>
          <p className="mt-[8px] max-w-[650px] text-[15px] leading-[1.45] text-slate-600">
            Review what your workflows created, updated, recalled, or removed.
          </p>
        </div>
        <button
          type="button"
          onClick={handleClearHistory}
          disabled={records.length === 0 || loading}
          className="inline-flex h-[40px] shrink-0 items-center justify-center rounded-[10px] border border-red-200 bg-white px-[16px] text-[14px] font-semibold text-red-700 shadow-[0_1px_2px_rgba(15,23,42,0.04)] transition hover:border-red-300 hover:bg-red-50 focus:outline-none focus:ring-2 focus:ring-red-500/30 disabled:cursor-not-allowed disabled:border-slate-200 disabled:text-slate-400 disabled:hover:bg-white"
        >
          Clear History
        </button>
      </header>

      {attentionCount > 0 ? (
        <section className="mb-[16px] rounded-[12px] border border-amber-200/80 bg-amber-50/80 px-[14px] py-[12px] text-[14px] leading-[1.4] text-amber-900">
          <h2 className="text-[14px] font-semibold text-amber-950">Needs attention</h2>
          <p className="mt-[3px]">
            {attentionCount} workflow {attentionCount === 1 ? "action reported" : "actions reported"} an error or incomplete result.
          </p>
        </section>
      ) : null}

      {staleRefreshFailed ? (
        <section className="mb-[16px] flex flex-col gap-[10px] rounded-[12px] border border-amber-200/80 bg-amber-50/70 px-[14px] py-[12px] text-[14px] text-amber-900 sm:flex-row sm:items-center sm:justify-between">
          <p>Couldn&apos;t refresh. Showing saved History.</p>
          <button
            type="button"
            onClick={retryHistoryLoad}
            className="inline-flex h-[36px] items-center justify-center rounded-[9px] border border-amber-300 bg-white px-[13px] text-[13px] font-semibold text-amber-900 transition hover:bg-amber-50 focus:outline-none focus:ring-2 focus:ring-amber-500/30"
          >
            Retry
          </button>
        </section>
      ) : null}

      <main>
        {loading && groupedRecords.length === 0 && !hasHistorySnapshot ? (
          <div className="space-y-[10px]" aria-label="Loading History">
            {[0, 1, 2, 3].map((index) => (
              <div
                key={index}
                className="h-[68px] animate-pulse rounded-[14px] border border-slate-200/70 bg-white/80 motion-reduce:animate-none"
              />
            ))}
          </div>
        ) : loadError && groupedRecords.length === 0 ? (
          <section className="rounded-[14px] border border-amber-200/80 bg-white/95 p-[20px] shadow-[0_8px_24px_rgba(30,64,100,0.05)]">
            <h2 className="text-[16px] font-semibold text-slate-950">Unable to load History</h2>
            <p className="mt-[6px] text-[14px] leading-[1.45] text-slate-600">History could not be loaded. Try again.</p>
            <button
              type="button"
              onClick={retryHistoryLoad}
              className="mt-[14px] inline-flex h-[40px] items-center justify-center rounded-[10px] border border-slate-300 bg-white px-[16px] text-[14px] font-semibold text-slate-800 transition hover:bg-slate-50 focus:outline-none focus:ring-2 focus:ring-slate-500/25"
            >
              Retry
            </button>
          </section>
        ) : groupedRecords.length === 0 ? (
          <section className="flex min-h-[190px] items-center justify-center rounded-[16px] border border-slate-200/80 bg-white/95 p-[20px] text-center shadow-[0_8px_24px_rgba(30,64,100,0.05)]">
            <div className="max-w-[360px]">
              <h2 className="text-[18px] font-semibold text-slate-950">No activity yet</h2>
              <p className="mt-[7px] text-[14px] leading-[1.45] text-slate-600">Export a plan to begin building your workflow history.</p>
              <Link
                href="/plans"
                className="mt-[14px] inline-flex h-[40px] items-center justify-center rounded-[10px] bg-slate-900 px-[16px] text-[14px] font-semibold text-white transition hover:bg-slate-800 focus:outline-none focus:ring-2 focus:ring-slate-500/30"
              >
                Go to Plans
              </Link>
            </div>
          </section>
        ) : (
          <div className="space-y-[18px]">
            {groupedRecords.map((dayGroup, dayIndex) => {
              const dayHeadingId = `history-day-${dayGroup.day}`;
              return (
                <section key={dayGroup.day} aria-labelledby={dayHeadingId}>
                  <h2
                    id={dayHeadingId}
                    className={`${dayIndex === 0 ? "" : "mt-[18px]"} mb-[8px] text-[13px] font-semibold leading-[1.35] text-slate-500`}
                  >
                    {formatDayLabel(dayGroup.day)}
                  </h2>
                  <ul className="space-y-[10px]">
                    {dayGroup.plans.map((planGroup, planIndex) => {
                      const isPlanExpanded = expandedPlans[planGroup.key] ?? false;
                      const isPlanModifying = exposeModifyUI && (modifyingPlans[planGroup.key] ?? false);
                      const { snapshot: planSnapshot, canModifyPlan } = getPlanModifyEligibility(planGroup);
                      const planModifyDate = planModifyDates[planGroup.key] ?? getCurrentEventDateValue(planGroup, planSnapshot);
                      const planModifyTime = planModifyTimes[planGroup.key] ?? getCurrentEventTimeValue(planGroup, planSnapshot);
                      const planModifyWeekendRule = planModifyWeekendRules[planGroup.key] ?? planSnapshot?.weekendRule ?? "prior_business_day";
                      const planModifyPreview = planModifyDate
                        ? getPlanModifyPreview(planGroup, planModifyDate, planModifyTime, planModifyWeekendRule)
                        : { snapshot: planSnapshot, items: [] };
                      const planModifyAvailabilityMessage = getPlanModifyAvailabilityMessage(planModifyPreview.items);
                      const recallablePlanItems = planGroup.items.filter((item) => {
                        const recallState = getExecutionHistoryRecallState(item);
                        return recallState.canRecall && recallState.recallImplemented;
                      });
	                      const planGroupUnavailable = isUnavailablePlanGroup(planGroup);
	                      const planMessage = planMessages[planGroup.key] ?? null;
	                      const planDisplayName = getPlanDisplayName(planGroup);
	                      const planMenuActions: HistoryActionMenuAction[] = [
	                        ...(exposeModifyUI && canModifyPlan && !planGroupUnavailable
	                          ? [
	                              {
	                                key: "modify",
	                                label: pendingPlanModifies[planGroup.key] ? "Updating..." : "Modify scheduled item",
	                                disabled: pendingPlanModifies[planGroup.key],
	                                onSelect: () => {
	                                  togglePlanModify(planGroup, true);
	                                  closeActionMenu();
	                                },
	                              },
	                            ]
	                          : []),
	                        ...(!planGroupUnavailable && recallablePlanItems.length > 0
	                          ? [
	                              {
	                                key: "provider-recall",
	                                label: pendingPlanRecalls[planGroup.key] ? "Working..." : getPlanRecallActionLabel(planGroup),
	                                tone: "provider" as const,
	                                disabled: pendingPlanRecalls[planGroup.key],
	                                onSelect: () => handleRecallPlan(planGroup),
	                              },
	                            ]
	                          : []),
	                        {
	                          key: "remove-group",
	                          label: "Remove group from History",
	                          tone: "history",
	                          disabled: pendingPlanDeletes[planGroup.key],
	                          title: "Removes these records from History only.",
	                          onSelect: () => handleDeletePlanGroup(planGroup),
	                        },
	                      ];
	                      const isPlanMenuOpen = openActionMenu?.kind === "plan" && openActionMenu.id === planGroup.key;

	                      return (
                        <li key={planGroup.key} className="relative pl-[18px]">
                          <span className="absolute left-[1px] top-[22px] h-[10px] w-[10px] rounded-full border-2 border-white bg-sky-500 shadow-[0_0_0_1px_rgba(14,116,144,0.18)]" />
                          {planIndex < dayGroup.plans.length - 1 ? (
                            <span className="absolute bottom-[-10px] left-[5px] top-[36px] w-px bg-slate-200/70" aria-hidden="true" />
                          ) : null}
                          <article className="overflow-hidden rounded-[14px] border border-slate-200/80 bg-white/95 shadow-[0_8px_24px_rgba(30,64,100,0.05)]">
	                            <div className="flex flex-col gap-[12px] px-[16px] py-[14px] min-[900px]:flex-row min-[900px]:items-center min-[900px]:justify-between">
	                              <div className="min-w-0">
                                <h3 className="text-[17px] font-semibold leading-[1.2] text-slate-950">{planDisplayName}</h3>
                                <p className="mt-[4px] text-[13px] leading-[1.35] text-slate-500">{getGroupMetadata(planGroup)}</p>
                                {shouldShowPlanMessage(planMessage) ? (
                                  <p
                                    className={`mt-[6px] text-[13px] font-medium ${
                                      planMessage.tone === "success"
                                        ? "text-green-700"
                                        : planMessage.tone === "warning"
                                          ? "text-amber-700"
                                          : "text-red-700"
                                    }`}
                                  >
                                    {planMessage.text}
                                  </p>
                                ) : null}
                                {shouldShowPlanHelperText(planMessage) ? (
                                  <p className="mt-[3px] text-[12px] leading-[1.35] text-slate-500">{planMessage.helperText}</p>
                                ) : null}
                              </div>
	                              <div className="flex w-full shrink-0 items-center justify-end gap-[8px] min-[900px]:w-auto">
	                                <button
	                                  type="button"
	                                  onClick={() => setExpandedPlans((current) => ({ ...current, [planGroup.key]: !isPlanExpanded }))}
                                  className="inline-flex h-[38px] items-center justify-center rounded-[9px] border border-slate-200 bg-white px-[13px] text-[13px] font-semibold text-slate-700 transition hover:bg-slate-50 focus:outline-none focus:ring-2 focus:ring-slate-500/25"
                                  aria-expanded={isPlanExpanded}
	                                >
	                                  {isPlanExpanded ? "Hide details" : "View details"}
	                                </button>
	                                <button
	                                  type="button"
	                                  onClick={(event) => handleActionMenuTriggerClick(event, "plan", planGroup.key, planMenuActions.length)}
	                                  onKeyDown={(event) => handleActionMenuTriggerKeyDown(event, "plan", planGroup.key, planMenuActions.length)}
	                                  aria-label={`More actions for ${planDisplayName}`}
	                                  aria-haspopup="menu"
	                                  aria-expanded={isPlanMenuOpen}
	                                  className="inline-flex h-[32px] w-[32px] items-center justify-center rounded-[8px] border border-slate-200 bg-white text-slate-600 transition hover:bg-slate-50 focus:outline-none focus:ring-2 focus:ring-slate-500/25"
	                                >
	                                  <IconEllipsis />
	                                </button>
	                                {renderActionMenu("plan", planGroup.key, `Actions for ${planDisplayName}`, planMenuActions)}
	                              </div>
                            </div>

                            {isPlanModifying ? (
                              <div className="border-t border-slate-200/70 bg-slate-50/70 px-[16px] py-[15px]">
                                <h4 className="text-[15px] font-semibold text-slate-950">Modify scheduled item</h4>
                                <div className="mt-[12px] grid gap-[12px] sm:grid-cols-2">
                                  <div>
                                    <div className="text-[13px] font-medium text-slate-500">Current event date</div>
                                    <div className="mt-[6px] text-[14px] text-slate-800">
                                      {getCurrentEventDateValue(planGroup, planSnapshot)
                                        ? formatDateOnly(`${getCurrentEventDateValue(planGroup, planSnapshot)}T00:00:00`)
                                        : "Not available"}
                                    </div>
                                  </div>
                                  <div>
                                    <div className="text-[13px] font-medium text-slate-500">Current event time</div>
                                    <div className="mt-[6px] text-[14px] text-slate-800">
                                      {getCurrentEventTimeValue(planGroup, planSnapshot)
                                        ? formatTimeOnly(`2000-01-01T${getCurrentEventTimeValue(planGroup, planSnapshot)}:00`)
                                        : "Not available"}
                                    </div>
                                  </div>
                                </div>
                                <div className="mt-[14px] grid gap-[12px] min-[860px]:grid-cols-[minmax(0,1fr)_minmax(0,1fr)_minmax(0,1.25fr)]">
                                  <label className="block min-w-0 text-[13px] font-medium text-slate-600">
                                    New event date
                                    <input
                                      type="date"
                                      value={planModifyDate}
                                      onChange={(event) => setPlanModifyDates((current) => ({ ...current, [planGroup.key]: event.target.value }))}
                                      className="mt-[6px] h-[40px] w-full rounded-[10px] border border-slate-300 bg-white px-[12px] text-[14px] text-slate-900 focus:outline-none focus:ring-2 focus:ring-slate-500/25"
                                    />
                                  </label>
                                  <label className="block min-w-0 text-[13px] font-medium text-slate-600">
                                    New event time
                                    <input
                                      type="time"
                                      value={planModifyTime}
                                      onChange={(event) => setPlanModifyTimes((current) => ({ ...current, [planGroup.key]: event.target.value }))}
                                      className="mt-[6px] h-[40px] w-full rounded-[10px] border border-slate-300 bg-white px-[12px] text-[14px] text-slate-900 focus:outline-none focus:ring-2 focus:ring-slate-500/25"
                                    />
                                  </label>
                                  <label className="block min-w-0 text-[13px] font-medium text-slate-600">
                                    Weekend handling
                                    <select
                                      value={planModifyWeekendRule}
                                      onChange={(event) =>
                                        setPlanModifyWeekendRules((current) => ({
                                          ...current,
                                          [planGroup.key]: event.target.value as WeekendRule,
                                        }))
                                      }
                                      className="mt-[6px] h-[40px] w-full rounded-[10px] border border-slate-300 bg-white px-[12px] text-[14px] text-slate-900 focus:outline-none focus:ring-2 focus:ring-slate-500/25"
                                    >
                                      <option value="prior_business_day">Adjust to prior business day (Fri)</option>
                                      <option value="none">Allow weekends (no adjustment)</option>
                                    </select>
                                  </label>
                                </div>
                                {planModifyAvailabilityMessage ? (
                                  <div className="mt-[12px] rounded-[12px] border border-amber-200 bg-amber-50/80 px-[12px] py-[10px] text-[13px] leading-[1.4] text-amber-800">
                                    <div>{planModifyAvailabilityMessage.text}</div>
                                    {planModifyAvailabilityMessage.helperText ? <div className="mt-[3px]">{planModifyAvailabilityMessage.helperText}</div> : null}
                                  </div>
                                ) : null}
                                <div className="mt-[14px] flex flex-wrap gap-[8px] sm:justify-end">
                                  <button
                                    type="button"
                                    onClick={() => setModifyingPlans((current) => ({ ...current, [planGroup.key]: false }))}
                                    className="inline-flex h-[40px] items-center justify-center rounded-[10px] border border-slate-300 bg-white px-[14px] text-[14px] font-semibold text-slate-700 transition hover:bg-slate-50 focus:outline-none focus:ring-2 focus:ring-slate-500/25"
                                  >
                                    Cancel
                                  </button>
                                  <button
                                    type="button"
                                    onClick={() => void handleModifyPlan(planGroup)}
                                    disabled={!planModifyDate || pendingPlanModifies[planGroup.key]}
                                    className="inline-flex h-[40px] items-center justify-center rounded-[10px] bg-slate-900 px-[14px] text-[14px] font-semibold text-white transition hover:bg-slate-800 focus:outline-none focus:ring-2 focus:ring-slate-500/30 disabled:cursor-not-allowed disabled:bg-slate-300"
                                  >
                                    Save changes
                                  </button>
                                </div>
                              </div>
                            ) : null}

                            {isPlanExpanded ? (
                              <div className="border-t border-slate-200/70 bg-slate-50/45">
                                <ul className="divide-y divide-slate-200/70">
                                  {planGroup.items.map((item) => {
                                    const outcome = getHistoryOutcome(item);
                                    const itemTitle = getItemTitle(item);
                                    const accessibleName = getItemAccessibleName(item);
                                    const itemTypeDisplayLabel = getItemTypeDisplayLabel(item);
                                    const isItemExpanded = expandedItems[item.id] ?? false;
                                    const reminderBody = getHistoryBody(item);
                                    const emailDraft = getHistoryEmailDraftDetails(item);
                                    const meetingDetails = getHistoryMeetingDetails(item);
	                                    const recallState = getExecutionHistoryRecallState(item);
	                                    const itemMessage = itemMessages[item.id] ?? null;
	                                    const canOpenModify = exposeModifyUI && canModifyPlan && !planGroupUnavailable;
	                                    const itemModifyState = getExecutionHistoryModifyState(item);
	                                    const canModifyItem = canOpenModify && itemModifyState.canModify && itemModifyState.modifyImplemented;
	                                    const recallUnavailableCopy = getRecallUnavailableCopy(item, recallState.recallReason);
	                                    const itemMenuActions: HistoryActionMenuAction[] = [];
	                                    if (hasExpandableHistoryItemDetails(item)) {
	                                      itemMenuActions.push({
	                                        key: "details",
	                                        label: isItemExpanded ? "Hide full details" : "View full details",
	                                        onSelect: () => {
	                                          setExpandedItems((current) => ({ ...current, [item.id]: !isItemExpanded }));
	                                          closeActionMenu();
	                                        },
	                                      });
	                                    }
	                                    if (canModifyItem) {
	                                      itemMenuActions.push({
	                                        key: "modify",
	                                        label: pendingPlanModifies[planGroup.key] ? "Updating..." : "Modify scheduled item",
	                                        disabled: pendingPlanModifies[planGroup.key],
	                                        onSelect: () => {
	                                          togglePlanModify(planGroup, true);
	                                          closeActionMenu();
	                                        },
	                                      });
	                                    }
	                                    if (item.outlookWebLink && !isUnavailableHistoryItem(item)) {
	                                      itemMenuActions.push({
	                                        key: "open-provider",
	                                        label: "Open in Outlook",
	                                        onSelect: () => {
	                                          window.open(item.outlookWebLink || "", "_blank", "noopener,noreferrer");
	                                          closeActionMenu();
	                                        },
	                                      });
	                                    }
	                                    if (recallState.canRecall && recallState.recallImplemented && item.providerObjectId) {
	                                      itemMenuActions.push({
	                                        key: "provider-recall",
	                                        label: pendingItemRecalls[item.id] ? "Working..." : getProviderRecallActionLabel(item),
	                                        tone: "provider",
	                                        disabled: pendingItemRecalls[item.id],
	                                        onSelect: () => handleRecallItem(item),
	                                      });
	                                    }
	                                    const hasItemMenuActions = itemMenuActions.length > 0;
	                                    const isItemMenuOpen = openActionMenu?.kind === "item" && openActionMenu.id === item.id;
	                                    const itemGridClassName = hasItemMenuActions
	                                      ? "grid min-h-[68px] grid-cols-[4px_minmax(0,1fr)_32px_32px] items-center gap-x-[10px] gap-y-[8px] px-[16px] py-[13px] min-[800px]:grid-cols-[4px_minmax(0,1fr)_150px_32px_32px] min-[800px]:gap-x-[12px]"
	                                      : "grid min-h-[68px] grid-cols-[4px_minmax(0,1fr)_32px] items-center gap-x-[10px] gap-y-[8px] px-[16px] py-[13px] min-[800px]:grid-cols-[4px_minmax(0,1fr)_150px_32px_32px] min-[800px]:gap-x-[12px]";

	                                    return (
	                                      <li key={item.id}>
	                                        <article>
	                                          <div className={itemGridClassName}>
                                            <span
                                              className={`h-full min-h-[38px] w-[4px] rounded-full ${
                                                outcome.tone === "error"
                                                  ? "bg-red-300"
                                                  : outcome.tone === "neutral"
                                                    ? "bg-slate-300"
                                                    : "bg-sky-400"
                                              }`}
                                              aria-hidden="true"
                                            />
                                            <div className="min-w-0">
                                              <div className={`text-[15px] font-semibold leading-[1.25] ${getOutcomeTextClasses(outcome.tone)}`}>
                                                {outcome.text}
                                              </div>
                                              {itemTitle ? (
                                                <div className="mt-[3px] line-clamp-2 text-[14px] leading-[1.35] text-slate-600">{itemTitle}</div>
                                              ) : null}
                                              <div className="mt-[4px] text-[12px] leading-[1.35] text-slate-500">
                                                {itemTypeDisplayLabel}
                                                {getActivityMetadata(item) ? ` · ${getActivityMetadata(item)}` : ""}
                                              </div>
                                              {itemMessage ? (
                                                <div
                                                  className={`mt-[5px] text-[13px] font-medium ${
                                                    itemMessage.tone === "success"
                                                      ? "text-green-700"
                                                      : itemMessage.tone === "neutral"
                                                        ? "text-slate-600"
                                                        : "text-red-700"
                                                  }`}
                                                >
                                                  {itemMessage.text}
                                                </div>
                                              ) : null}
	                                              {!recallState.canRecall &&
	                                              recallUnavailableCopy &&
	                                              item.status !== "recalled" &&
	                                              item.status !== "already_removed" &&
	                                              item.status !== "already_canceled" ? (
	                                                <div className="mt-[5px] text-[13px] leading-[1.35] text-slate-500">{recallUnavailableCopy}</div>
	                                              ) : null}
                                            </div>
                                            <div className="col-start-2 text-[12px] font-medium leading-[1.35] text-slate-500 min-[800px]:col-start-3 min-[800px]:row-start-1 min-[800px]:text-right">
                                              {formatDateTime(item.executedAt)}
                                            </div>
                                            <button
                                              type="button"
                                              onClick={() => handleDeleteItem(item)}
                                              disabled={pendingItemDeletes[item.id]}
                                              title="Removes this record from History only."
                                              aria-label={`Remove ${accessibleName} from History`}
                                              className="col-start-3 row-start-1 inline-flex h-[32px] w-[32px] items-center justify-center rounded-[8px] border border-red-200 bg-white text-red-700 transition hover:border-red-300 hover:bg-red-50 focus:outline-none focus:ring-2 focus:ring-red-500/30 disabled:cursor-not-allowed disabled:border-slate-200 disabled:text-slate-300 min-[800px]:col-start-4"
                                            >
                                              <IconTrash />
                                            </button>
	                                            {hasItemMenuActions ? (
	                                              <div className="col-start-4 row-start-1 min-[800px]:col-start-5">
	                                                <button
	                                                  type="button"
	                                                  onClick={(event) => handleActionMenuTriggerClick(event, "item", item.id, itemMenuActions.length)}
	                                                  onKeyDown={(event) => handleActionMenuTriggerKeyDown(event, "item", item.id, itemMenuActions.length)}
	                                                  aria-label={`More actions for ${accessibleName}`}
	                                                  aria-haspopup="menu"
	                                                  aria-expanded={isItemMenuOpen}
	                                                  className="inline-flex h-[32px] w-[32px] items-center justify-center rounded-[8px] border border-slate-200 bg-white text-slate-600 transition hover:bg-slate-50 focus:outline-none focus:ring-2 focus:ring-slate-500/25"
	                                                >
	                                                  <IconEllipsis />
	                                                </button>
	                                                {renderActionMenu("item", item.id, `Actions for ${accessibleName}`, itemMenuActions)}
	                                              </div>
	                                            ) : (
	                                              <span
	                                                className="hidden h-[32px] w-[32px] min-[800px]:col-start-5 min-[800px]:row-start-1 min-[800px]:block"
	                                                aria-hidden="true"
	                                              />
	                                            )}
                                          </div>
                                          {isItemExpanded ? (
                                            <div className="border-t border-slate-200/70 bg-white px-[16px] py-[14px]">
                                              <div className="space-y-[12px] text-[13px] leading-[1.45] text-slate-600">
                                                <div>
                                                  <h4 className="text-[13px] font-semibold text-slate-800">Workflow</h4>
                                                  <p className="mt-[3px]">{planDisplayName}</p>
                                                </div>
                                                {item.itemType === "reminder" && reminderBody ? (
                                                  <div>
                                                    <h4 className="text-[13px] font-semibold text-slate-800">Reminder note</h4>
                                                    <p className="mt-[3px] whitespace-pre-wrap">{reminderBody}</p>
                                                  </div>
                                                ) : null}
                                                {(item.itemType === "meeting" || item.itemType === "teams_meeting") && meetingDetails ? (
                                                  <div className="grid gap-[12px] sm:grid-cols-2">
                                                    {meetingDetails.attendees.length > 0 ? (
                                                      <div className="sm:col-span-2">
                                                        <h4 className="text-[13px] font-semibold text-slate-800">Attendees</h4>
                                                        <p className="mt-[3px] break-words">{meetingDetails.attendees.join(", ")}</p>
                                                      </div>
                                                    ) : null}
                                                    {meetingDetails.location ? (
                                                      <div>
                                                        <h4 className="text-[13px] font-semibold text-slate-800">Location</h4>
                                                        <p className="mt-[3px] break-words">{meetingDetails.location}</p>
                                                      </div>
                                                    ) : null}
                                                    {meetingDetails.body ? (
                                                      <div className="sm:col-span-2">
                                                        <h4 className="text-[13px] font-semibold text-slate-800">Message</h4>
                                                        <p className="mt-[3px] whitespace-pre-wrap">{meetingDetails.body}</p>
                                                      </div>
                                                    ) : null}
                                                  </div>
                                                ) : null}
                                                {item.itemType === "email" && emailDraft ? (
                                                  <div className="grid gap-[12px] sm:grid-cols-2">
                                                    {emailDraft.to.length > 0 ? (
                                                      <div className="sm:col-span-2">
                                                        <h4 className="text-[13px] font-semibold text-slate-800">To</h4>
                                                        <p className="mt-[3px] break-words">{emailDraft.to.join(", ")}</p>
                                                      </div>
                                                    ) : null}
                                                    {emailDraft.cc.length > 0 ? (
                                                      <div className="sm:col-span-2">
                                                        <h4 className="text-[13px] font-semibold text-slate-800">Cc</h4>
                                                        <p className="mt-[3px] break-words">{emailDraft.cc.join(", ")}</p>
                                                      </div>
                                                    ) : null}
                                                    {emailDraft.bcc.length > 0 ? (
                                                      <div className="sm:col-span-2">
                                                        <h4 className="text-[13px] font-semibold text-slate-800">Bcc</h4>
                                                        <p className="mt-[3px] break-words">{emailDraft.bcc.join(", ")}</p>
                                                      </div>
                                                    ) : null}
                                                    {emailDraft.subject ? (
                                                      <div className="sm:col-span-2">
                                                        <h4 className="text-[13px] font-semibold text-slate-800">Subject</h4>
                                                        <p className="mt-[3px] break-words">{emailDraft.subject}</p>
                                                      </div>
                                                    ) : null}
                                                    {emailDraft.body ? (
                                                      <div className="sm:col-span-2">
                                                        <h4 className="text-[13px] font-semibold text-slate-800">Message</h4>
                                                        <p className="mt-[3px] whitespace-pre-wrap">{emailDraft.body}</p>
                                                      </div>
                                                    ) : null}
                                                  </div>
                                                ) : null}
                                              </div>
                                            </div>
                                          ) : null}
                                        </article>
                                      </li>
                                    );
                                  })}
                                </ul>
                              </div>
                            ) : null}
                          </article>
                        </li>
                      );
                    })}
                  </ul>
                </section>
              );
            })}
          </div>
        )}
      </main>

      {mounted && confirmationDialog
        ? createPortal(
            <div
              className="fixed inset-0 z-[260] flex items-end justify-center bg-[rgba(15,23,42,0.14)] p-[16px] sm:items-center"
              onMouseDown={(event) => {
                if (event.target === event.currentTarget && !confirmationPending) {
                  setConfirmationDialog(null);
                }
              }}
            >
              <div
                ref={confirmationDialogRef}
                role="dialog"
                aria-modal="true"
                aria-labelledby="history-confirm-title"
                aria-describedby="history-confirm-description"
                className="max-h-[88dvh] w-full max-w-[500px] overflow-auto rounded-t-[16px] border border-slate-200 bg-white shadow-[0_22px_70px_rgba(15,23,42,0.18)] sm:rounded-[16px]"
              >
                <div className="px-[20px] pb-[16px] pt-[20px]">
                  <h2 id="history-confirm-title" className="text-[18px] font-semibold leading-[1.25] text-slate-950">
                    {confirmationDialog.title}
                  </h2>
                  <p id="history-confirm-description" className="mt-[8px] text-[14px] leading-[1.45] text-slate-600">
                    {confirmationDialog.body}
                  </p>
                </div>
                <div className="flex flex-col-reverse gap-[8px] border-t border-slate-200/70 px-[20px] pb-[calc(16px+env(safe-area-inset-bottom))] pt-[14px] sm:flex-row sm:justify-end sm:pb-[16px]">
                  <button
                    ref={confirmationCancelRef}
                    type="button"
                    onClick={() => {
                      if (!confirmationPending) setConfirmationDialog(null);
                    }}
                    disabled={confirmationPending}
                    className="inline-flex h-[40px] items-center justify-center rounded-[10px] border border-slate-300 bg-white px-[16px] text-[14px] font-semibold text-slate-700 transition hover:bg-slate-50 focus:outline-none focus:ring-2 focus:ring-slate-500/25 disabled:cursor-not-allowed disabled:text-slate-400"
                  >
                    Cancel
                  </button>
                  <button
                    type="button"
                    onClick={() => void runConfirmationAction()}
                    disabled={confirmationPending}
	                    className={`inline-flex h-[40px] items-center justify-center rounded-[10px] px-[16px] text-[14px] font-semibold text-white transition focus:outline-none focus:ring-2 disabled:cursor-not-allowed disabled:opacity-70 ${
	                      confirmationDialog.tone === "destructive"
	                        ? "bg-red-700 hover:bg-red-800 focus:ring-red-500/30"
	                        : "bg-amber-600 hover:bg-amber-700 focus:ring-amber-500/30"
	                    }`}
                  >
                    {confirmationPending ? "Working..." : confirmationDialog.confirmLabel}
                  </button>
                </div>
              </div>
            </div>,
            document.body
          )
        : null}
    </div>
  );
}
