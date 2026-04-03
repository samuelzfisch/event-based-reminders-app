"use client";

import Link from "next/link";
import { useEffect, useRef, useState } from "react";

import {
  APP_SETTINGS_UPDATED_EVENT,
  areAppSettingsEqual,
  DEFAULT_APP_SETTINGS,
  hydrateAppSettingsFromSupabase,
  loadAppSettings,
  type AppSettings,
  type EmailHandlingMode,
} from "../../lib/appSettings";
import { todayYYYYMMDD } from "../../lib/dateUtils";
import { writeExecutionHistory } from "../../lib/executionHistory";
import { buildICSForPlan, downloadICS } from "../../lib/ics";
import {
  createOutlookCalendarEvent,
  createOutlookDraftFromEmailDraft,
  getConnectedOutlookMailboxEmail,
  getOutlookConnectionState,
  OUTLOOK_CONNECTION_UPDATED_EVENT,
  resolveOutlookConnectionState,
  scheduleOutlookEmailFromEmailDraft,
  sendOutlookEmailFromEmailDraft,
  type OutlookConnectionState,
} from "../../lib/outlookClient";
import { createPlan, type TemplateItem } from "../../lib/planEngine";
import {
  loadCachedTemplateState,
  loadTemplateStateFromSupabase,
  saveCachedTemplateState,
  saveTemplateStateToSupabase,
  type PersistedPlanTemplate,
  type PersistedTemplateState,
} from "../../lib/templateStore";
import {
  buildAnchorMap,
  classifyPlanRow,
  normalizeAnchorKey,
  partitionPlanItemsByKind,
  resolvePlanAnchors,
  resolveReminderTimeValue,
} from "../../lib/plansRuntime";
import type { Plan, PlanDateBasis, PlanRowType, PlanType, WeekendRule } from "../../types/plan";

type BuilderEmailDraft = {
  to?: string[];
  cc?: string[];
  bcc?: string[];
  subject?: string;
  body?: string;
};

type BuilderAnchor = {
  id: string;
  key: string;
  value: string;
  locked?: boolean;
};

type BuilderRow = {
  id: string;
  title: string;
  body?: string;
  offsetDays: number | null;
  dateBasis?: PlanDateBasis;
  rowType: PlanRowType;
  reminderTime?: string;
  emailDraft?: BuilderEmailDraft;
  durationDraft?: {
    durationMinutes?: number;
    useCustomEnd?: boolean;
    endDate?: string;
    endTime?: string;
    isAllDay?: boolean;
  };
  meetingDraft?: {
    attendees?: string[];
    location?: string;
    durationMinutes?: number;
    useCustomEnd?: boolean;
    endDate?: string;
    endTime?: string;
    isAllDay?: boolean;
    teamsMeeting?: boolean;
  };
};

type SavedPlanTemplate = {
  id: string;
  name: string;
  baseType: PlanType;
  templateMode?: "template" | "custom";
  noEventDate?: boolean;
  weekendRule: WeekendRule;
  anchors: Array<{ key: string; value: string }>;
  items: BuilderRow[];
};

type EmailFieldVisibility = Record<string, { cc?: boolean; bcc?: boolean }>;
type BuilderMode = "template" | "new" | "guided" | "custom";
type BuilderStateSnapshot = {
  builderMode: BuilderMode;
  selectedTemplateId: string | null;
  planType: PlanType;
  templateName: string;
  eventName: string;
  anchorDate: string;
  noEventDate: boolean;
  weekendRule: WeekendRule;
  rows: BuilderRow[];
  anchors: BuilderAnchor[];
  guidedForm: GuidedFormState;
};

type MeetingValidationErrorState = Record<
  string,
  {
    attendees?: boolean;
    time?: boolean;
    duration?: boolean;
  }
>;

type OutlookExecutionResult = {
  kind: "email" | "reminder" | "meeting";
  action:
    | "draft_created"
    | "email_sent"
    | "email_scheduled"
    | "reminder_created"
    | "meeting_created";
  title: string;
  message: string;
  providerObjectId?: string;
  webLink?: string;
  joinUrl?: string;
};

type ExecutionNotice = {
  tone: "success" | "mixed" | "warning";
  title: string;
  message?: string;
  details?: string[];
};

type GuidedFormState = {
  releaseName: string;
  releaseDate: string;
  releaseTime: string;
  quarter: "" | "Q1" | "Q2" | "Q3" | "Q4";
  year: string;
  fiscalYear: boolean;
  earningsDate: string;
  earningsTime: string;
  conferenceName: string;
  conferenceLocation: string;
  conferenceDate: string;
  conferenceEndDate: string;
};

const PRESS_RELEASE_PRESET_ANCHOR_KEYS = [
  "Press Release Name",
  "Dissemination Date",
  "Dissemination Time",
] as const;
const CONFERENCE_PRESET_ANCHOR_KEYS = [
  "Conference Name",
  "Conference Location",
  "Conference Start Date",
  "Conference End Date",
] as const;
const EARNINGS_PRESET_ANCHOR_KEYS = [
  "Quarter",
  "Year / Fiscal Year",
  "Earnings Call Date",
  "Earnings Call Time",
] as const;
const GENERIC_PRESET_ANCHOR_KEYS = ["Event Name", "Event Date"] as const;
const TEAMS_MEETING_LOCATION = "Microsoft Teams Meeting";
const MEETING_DURATION_OPTIONS = [
  { value: "30", label: "30 minutes" },
  { value: "60", label: "60 minutes" },
  { value: "90", label: "90 minutes" },
  { value: "120", label: "2 hours" },
  { value: "custom", label: "Custom" },
] as const;

function makeId(prefix: string) {
  return `${prefix}_${Math.random().toString(36).slice(2, 10)}`;
}

function createEmptyAnchor(): BuilderAnchor {
  return {
    id: crypto.randomUUID(),
    key: "",
    value: "",
  };
}

function createLockedAnchor(key: string, value = ""): BuilderAnchor {
  return {
    id: crypto.randomUUID(),
    key,
    value,
    locked: true,
  };
}

function createGenericPresetAnchors() {
  return GENERIC_PRESET_ANCHOR_KEYS.map((key) => createLockedAnchor(key));
}

function createEmptyGuidedForm(): GuidedFormState {
  return {
    releaseName: "",
    releaseDate: "",
    releaseTime: "",
    quarter: "",
    year: "",
    fiscalYear: false,
    earningsDate: "",
    earningsTime: "",
    conferenceName: "",
    conferenceLocation: "",
    conferenceDate: "",
    conferenceEndDate: "",
  };
}

function OutlookExecutionNoticeCard({
  notice,
  onDismiss,
}: {
  notice: ExecutionNotice;
  onDismiss: () => void;
}) {
  const toneClasses =
    notice.tone === "success"
      ? "border-green-200 bg-green-50 text-green-950"
      : notice.tone === "mixed"
        ? "border-amber-200 bg-amber-50 text-amber-950"
        : "border-gray-200 bg-gray-50 text-gray-900";

  return (
    <div className={`rounded-xl border p-4 shadow-sm ${toneClasses}`}>
      <div className="flex items-start justify-between gap-3">
        <div className="space-y-2">
          <div className="text-sm font-semibold">{notice.title}</div>
          {notice.message ? <p className="text-sm">{notice.message}</p> : null}
          {notice.details && notice.details.length > 0 ? (
            <div className="space-y-1 text-sm">
              {notice.details.map((detail) => (
                <div key={detail}>{detail}</div>
              ))}
            </div>
          ) : null}
        </div>
        <button
          type="button"
          onClick={onDismiss}
          className="rounded-lg border border-current/20 bg-white px-2 py-1 text-xs hover:bg-white/80"
        >
          Dismiss
        </button>
      </div>
    </div>
  );
}

function ExportDoneBadge() {
  return (
    <div className="inline-flex items-center gap-1.5 rounded-full border border-green-200 bg-green-50 px-2.5 py-1 text-xs font-medium text-green-700">
      <svg viewBox="0 0 16 16" aria-hidden="true" className="h-3.5 w-3.5">
        <path
          d="M13 4.5 6.5 11 3 7.5"
          fill="none"
          stroke="currentColor"
          strokeWidth="1.8"
          strokeLinecap="round"
          strokeLinejoin="round"
        />
      </svg>
      <span>Export Done</span>
    </div>
  );
}

function getSeedTemplateName(type: PlanType) {
  if (type === "press_release") return "Press Release";
  return type.charAt(0).toUpperCase() + type.slice(1);
}

function hasMeaningfulBody(body?: string | null) {
  return Boolean(body?.trim());
}

function isProtectedTemplate(template: SavedPlanTemplate) {
  return (
    template.id === "seed:press_release" ||
    template.id === "seed:conference" ||
    template.id === "seed:earnings"
  );
}

function getProtectedTemplateDefinition(template: Pick<SavedPlanTemplate, "id" | "name" | "baseType">) {
  if (template.id === "seed:press_release" || template.name === "Press Release") {
    return {
      id: "seed:press_release",
      name: "Press Release",
      baseType: "press_release" as const,
    };
  }
  if (template.id === "seed:conference" || template.name === "Conference") {
    return {
      id: "seed:conference",
      name: "Conference",
      baseType: "conference" as const,
    };
  }
  if (template.id === "seed:earnings" || template.name === "Earnings") {
    return {
      id: "seed:earnings",
      name: "Earnings",
      baseType: "earnings" as const,
    };
  }
  return null;
}

function getPresetAnchorKeysForType(type: PlanType) {
  if (type === "press_release") return [...PRESS_RELEASE_PRESET_ANCHOR_KEYS];
  if (type === "conference") return [...CONFERENCE_PRESET_ANCHOR_KEYS];
  return [...EARNINGS_PRESET_ANCHOR_KEYS];
}

function normalizeEmailToken(value: string) {
  return value.trim().replace(/^mailto:/i, "").trim().toLowerCase();
}

function isValidEmailToken(value: string) {
  const normalized = normalizeEmailToken(value);
  return normalized.includes("@");
}

function normalizeEmailList(value: unknown): string[] {
  const rawValues = Array.isArray(value) ? value : typeof value === "string" ? value.split(/[,\n;]/) : [];
  return Array.from(
    new Set(
      rawValues
        .map((entry) => normalizeEmailToken(String(entry ?? "")))
        .filter((entry) => entry && isValidEmailToken(entry))
    )
  );
}

function normalizeEmailDraft(draft?: BuilderEmailDraft | null) {
  return {
    to: normalizeEmailList(draft?.to),
    cc: normalizeEmailList(draft?.cc),
    bcc: normalizeEmailList(draft?.bcc),
    subject: typeof draft?.subject === "string" ? draft.subject : "",
    body: typeof draft?.body === "string" ? draft.body : "",
  };
}

function normalizeMeetingDraft(draft?: BuilderRow["meetingDraft"] | null) {
  if (!draft) return undefined;
  return {
    attendees: normalizeEmailList(draft.attendees),
    location: typeof draft.location === "string" ? draft.location : "",
    durationMinutes:
      typeof draft.durationMinutes === "number" && draft.durationMinutes > 0 ? draft.durationMinutes : 30,
    useCustomEnd: Boolean(draft.useCustomEnd),
    endDate: typeof draft.endDate === "string" ? draft.endDate : "",
    endTime: typeof draft.endTime === "string" ? draft.endTime : "",
    isAllDay: Boolean(draft.isAllDay),
    teamsMeeting: Boolean(draft.teamsMeeting),
  };
}

function normalizeDurationDraft(draft?: BuilderRow["durationDraft"] | null) {
  if (!draft) return undefined;
  return {
    durationMinutes:
      typeof draft.durationMinutes === "number" && draft.durationMinutes > 0 ? draft.durationMinutes : 30,
    useCustomEnd: Boolean(draft.useCustomEnd),
    endDate: typeof draft.endDate === "string" ? draft.endDate : "",
    endTime: typeof draft.endTime === "string" ? draft.endTime : "",
    isAllDay: Boolean(draft.isAllDay),
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
    const normalizedHours =
      meridiem === "AM" ? (hours === 12 ? 0 : hours) : hours === 12 ? 12 : hours + 12;
    return `${String(normalizedHours).padStart(2, "0")}:${String(minutes).padStart(2, "0")}`;
  }

  const twentyFourHourMatch = normalized.match(/^([01]?\d|2[0-3]):([0-5]\d)$/);
  if (twentyFourHourMatch) {
    return `${String(Number(twentyFourHourMatch[1])).padStart(2, "0")}:${twentyFourHourMatch[2]}`;
  }

  return null;
}

function formatTimeForEditor(time: string) {
  const parsedTime = parseTimeInput(time);
  if (!parsedTime) return time.trim();
  const [hours, minutes] = parsedTime.split(":").map(Number);
  const parsed = new Date(2000, 0, 1, hours ?? 0, minutes ?? 0);
  return parsed.toLocaleTimeString("en-US", {
    hour: "numeric",
    minute: "2-digit",
  });
}

const REMINDER_TIME_INPUT_MASK = "--:-- --";

function sanitizeReminderTimeInputForParsing(value: string) {
  return value
    .replace(/-/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function getUsableReminderTime(value: string | undefined, anchorMap?: Map<string, string>) {
  const resolved = resolveReminderTimeValue(value, anchorMap);
  return parseTimeInput(resolved);
}

function addMinutesToLocalDateTime(date: string, time: string, minutesToAdd: number) {
  const [year, month, day] = date.split("-").map(Number);
  const [hours, minutes] = (parseTimeInput(time) || "09:00").split(":").map(Number);
  const next = new Date(year ?? 2000, (month ?? 1) - 1, day ?? 1, hours ?? 9, minutes ?? 0);
  next.setMinutes(next.getMinutes() + minutesToAdd);
  return {
    endDate: `${next.getFullYear()}-${String(next.getMonth() + 1).padStart(2, "0")}-${String(next.getDate()).padStart(2, "0")}`,
    endTime: `${String(next.getHours()).padStart(2, "0")}:${String(next.getMinutes()).padStart(2, "0")}`,
  };
}

function normalizeReminderTimeInput(value: string) {
  const trimmed = value.trim();
  if (!trimmed) return "";
  if (trimmed.startsWith("[")) return trimmed;
  const sanitized = sanitizeReminderTimeInputForParsing(trimmed);
  if (!sanitized) return "";
  const parsed = parseTimeInput(sanitized);
  if (!parsed) return "";
  return formatTimeForEditor(sanitized);
}

function buildReminderTimeMaskedValue(value: string) {
  const trimmed = value.trim();
  if (!trimmed) return REMINDER_TIME_INPUT_MASK;
  if (trimmed.startsWith("[")) return trimmed;

  const formattedLiteralTime = normalizeReminderTimeInput(trimmed);
  const effectiveValue = formattedLiteralTime || trimmed;

  const digits = effectiveValue.replace(/\D/g, "").slice(0, 4);
  const letters = effectiveValue.replace(/[^apm]/gi, "").toUpperCase().slice(0, 2);
  const hours = `${digits[0] ?? "-"}${digits[1] ?? "-"}`;
  const minutes = `${digits[2] ?? "-"}${digits[3] ?? "-"}`;
  const meridiem = `${letters[0] ?? "-"}${letters[1] ?? "-"}`;
  return `${hours}:${minutes} ${meridiem}`;
}

function isReminderTimeAnchorValue(value: string) {
  return value.trim().startsWith("[");
}

function getReminderTimeDisplayValue(value: string) {
  const trimmed = value.trim();
  if (!trimmed) return "";
  if (trimmed.startsWith("[")) return trimmed;
  return normalizeReminderTimeInput(trimmed) || trimmed;
}

function clearReminderTimeDraft(drafts: Record<string, string>, rowId: string) {
  const next = { ...drafts };
  delete next[rowId];
  return next;
}

function countReminderTimeMeaningfulChars(value: string, end: number) {
  return value
    .slice(0, end)
    .replace(/[^0-9APMapm]/g, "")
    .length;
}

function findReminderTimeCursorFromMeaningfulCount(value: string, meaningfulCount: number) {
  if (meaningfulCount <= 0) return 0;
  let seen = 0;
  for (let index = 0; index < value.length; index += 1) {
    if (/[0-9APM]/i.test(value[index] ?? "")) {
      seen += 1;
      if (seen >= meaningfulCount) {
        return index + 1;
      }
    }
  }
  return value.length;
}

function maskReminderTimeDraftInput(value: string) {
  const trimmed = value.trimStart();
  if (!trimmed) return REMINDER_TIME_INPUT_MASK;
  if (trimmed.startsWith("[")) return value;

  const digits = trimmed.replace(/\D/g, "").slice(0, 4);
  const letters = trimmed.replace(/[^apm]/gi, "").toUpperCase().slice(0, 2);
  const hours = `${digits[0] ?? "-"}${digits[1] ?? "-"}`;
  const minutes = `${digits[2] ?? "-"}${digits[3] ?? "-"}`;
  const meridiem = `${letters[0] ?? "-"}${letters[1] ?? "-"}`;
  return `${hours}:${minutes} ${meridiem}`;
}

function getGuidedTemplateDisplayName(type: PlanType, form: GuidedFormState) {
  if (type === "press_release") return form.releaseName.trim();
  if (type === "conference") return form.conferenceName.trim();
  if (type === "earnings") {
    if (!form.quarter || !form.year.trim()) return "";
    return form.fiscalYear
      ? `${form.quarter} Fiscal Year ${form.year.trim()} Earnings`
      : `${form.quarter} ${form.year.trim()} Earnings`;
  }
  return "";
}

function formatAnchorDateDisplayValue(value: string) {
  const trimmed = value.trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) return value;
  const [year, month, day] = trimmed.split("-").map(Number);
  const parsed = new Date(year ?? 2000, (month ?? 1) - 1, day ?? 1);
  return parsed.toLocaleDateString("en-US", {
    month: "long",
    day: "numeric",
    year: "numeric",
  });
}

function getAnchorDisplayValue(key: string, value: string) {
  const normalized = normalizeAnchorKey(key);
  const isDateAnchor =
    normalized === normalizeAnchorKey("Event Date") ||
    normalized === normalizeAnchorKey("Dissemination Date") ||
    normalized === normalizeAnchorKey("Conference Start Date") ||
    normalized === normalizeAnchorKey("Conference End Date") ||
    normalized === normalizeAnchorKey("Earnings Call Date");
  return isDateAnchor ? formatAnchorDateDisplayValue(value) : value;
}

function getDerivedAnchorValue(
  type: PlanType,
  planName: string,
  anchorDate: string,
  key: string,
  guidedForm: GuidedFormState,
  options?: {
    eventNameValue?: string;
    eventDateValue?: string;
  }
) {
  const normalized = normalizeAnchorKey(key);
  if (normalized === normalizeAnchorKey("Event Name")) return options?.eventNameValue ?? planName;
  if (normalized === normalizeAnchorKey("Event Date")) return options?.eventDateValue ?? anchorDate;
  if (type === "press_release") {
    if (normalized === normalizeAnchorKey("Press Release Name")) return guidedForm.releaseName.trim();
    if (normalized === normalizeAnchorKey("Dissemination Date")) return guidedForm.releaseDate;
    if (normalized === normalizeAnchorKey("Dissemination Time")) return guidedForm.releaseTime;
  }
  if (type === "conference") {
    if (normalized === normalizeAnchorKey("Conference Name")) return guidedForm.conferenceName.trim();
    if (normalized === normalizeAnchorKey("Conference Location")) return guidedForm.conferenceLocation.trim();
    if (normalized === normalizeAnchorKey("Conference Start Date")) return guidedForm.conferenceDate;
    if (normalized === normalizeAnchorKey("Conference End Date")) return guidedForm.conferenceEndDate;
  }
  if (type === "earnings") {
    if (normalized === normalizeAnchorKey("Quarter")) return guidedForm.quarter;
    if (normalized === normalizeAnchorKey("Year / Fiscal Year")) {
      return guidedForm.year.trim()
        ? guidedForm.fiscalYear
          ? `Fiscal Year ${guidedForm.year.trim()}`
          : guidedForm.year.trim()
        : "";
    }
    if (normalized === normalizeAnchorKey("Earnings Call Date")) return guidedForm.earningsDate;
    if (normalized === normalizeAnchorKey("Earnings Call Time")) return guidedForm.earningsTime;
  }
  return null;
}

function createEmptyBuilderRow(defaultReminderTime = ""): BuilderRow {
  return {
    id: crypto.randomUUID(),
    title: "",
    body: "",
    offsetDays: 0,
    dateBasis: "event",
    rowType: "reminder",
    reminderTime: normalizeReminderTimeInput(defaultReminderTime),
    emailDraft: { to: [], cc: [], bcc: [], subject: "", body: "" },
  };
}

function buildSeedTemplates(defaultReminderTime = "", defaultPressReleaseTime = ""): SavedPlanTemplate[] {
  const normalizedDefaultReminderTime = normalizeReminderTimeInput(defaultReminderTime) || "10:30";
  const normalizedDefaultPressReleaseTime =
    normalizeReminderTimeInput(defaultPressReleaseTime) || normalizedDefaultReminderTime || "08:30";
  const sharedChecklistBody = [
    "Check Forward Looking Statements",
    "Notified",
    "Nasdaq IssuerEntry",
    "Email Kari Sharp",
    "Constant Contact",
    "Review Proof",
  ].join("\n");

  const baseTemplates: SavedPlanTemplate[] = [
    {
      id: "seed:press_release",
      name: "Press Release",
      baseType: "press_release",
      templateMode: "template",
      weekendRule: "prior_business_day",
      anchors: PRESS_RELEASE_PRESET_ANCHOR_KEYS.map((key) => ({ key, value: "" })),
      items: [
        {
          id: crypto.randomUUID(),
          title: "Prepare Press Release Distribution for Tomorrow",
          body: sharedChecklistBody,
          offsetDays: -1,
          rowType: "reminder",
          reminderTime: normalizedDefaultReminderTime,
          dateBasis: "event",
        },
        {
          id: crypto.randomUUID(),
          title: "Finalize Proof & Schedule Press Release",
          offsetDays: -1,
          rowType: "reminder",
          reminderTime: "16:00",
          dateBasis: "event",
        },
        {
          id: crypto.randomUUID(),
          title: "Press Release Going Out Now: Constant Contact & Social",
          offsetDays: 0,
          rowType: "reminder",
          reminderTime: normalizedDefaultPressReleaseTime,
          dateBasis: "event",
        },
      ],
    },
    {
      id: "seed:conference",
      name: "Conference",
      baseType: "conference",
      templateMode: "template",
      weekendRule: "prior_business_day",
      anchors: CONFERENCE_PRESET_ANCHOR_KEYS.map((key) => ({ key, value: "" })),
      items: [
        { id: crypto.randomUUID(), title: "Draft Conference Press Release", offsetDays: -21, rowType: "reminder", reminderTime: "10:30", dateBasis: "event" },
        { id: crypto.randomUUID(), title: "Register Webcast", body: "Register webcast", offsetDays: -21, rowType: "reminder", reminderTime: "10:30", dateBasis: "event" },
        { id: crypto.randomUUID(), title: "Review Conference Agenda & Presentation Timing", offsetDays: -14, rowType: "reminder", reminderTime: "10:30", dateBasis: "event" },
        { id: crypto.randomUUID(), title: "Confirm Conference Presentations", body: ["Confirm meeting version", "Confirm presentation version"].join("\n"), offsetDays: -14, rowType: "reminder", reminderTime: "10:30", dateBasis: "event" },
        { id: crypto.randomUUID(), title: "Prepare Investor Briefs", offsetDays: -7, rowType: "reminder", reminderTime: "10:30", dateBasis: "event" },
        { id: crypto.randomUUID(), title: "Print Presentation Deck(s)", offsetDays: -7, rowType: "reminder", reminderTime: "10:30", dateBasis: "event" },
        { id: crypto.randomUUID(), title: "Pack Business Cards & Presentations", offsetDays: -1, rowType: "reminder", reminderTime: "22:30", dateBasis: "event" },
        { id: crypto.randomUUID(), title: "[Conference Name]", offsetDays: 0, rowType: "reminder", reminderTime: "10:30", dateBasis: "event", meetingDraft: { location: "[Conference Location]", isAllDay: true } },
        { id: crypto.randomUUID(), title: "Add New Corporate Presentation to Website", offsetDays: 1, rowType: "reminder", reminderTime: "10:30", dateBasis: "event" },
        { id: crypto.randomUUID(), title: "Set Up Flights & Hotels", offsetDays: 1, rowType: "reminder", reminderTime: "10:30", dateBasis: "today" },
        {
          id: crypto.randomUUID(),
          title: "Hotel & Flight Coordination",
          offsetDays: -21,
          rowType: "email",
          reminderTime: "10:30",
          dateBasis: "event",
          emailDraft: {
            to: ["dfitzpatrick@verupharma.com"],
            cc: [],
            bcc: [],
            subject: "[Conference Name] - Hotel & Flight Information",
            body: [
              "Hi Dawn!",
              "",
              "We are attending the [Conference Name] on [Conference Start Date] and it ends on [Conference End Date]. The conference is located in [Conference Location]. Can you see what hotels/flights are available nearby the conference?",
              "",
              "Thank you!",
              "Sam",
            ].join("\n"),
          },
        },
      ],
    },
    {
      id: "seed:earnings",
      name: "Earnings",
      baseType: "earnings",
      templateMode: "template",
      weekendRule: "prior_business_day",
      anchors: EARNINGS_PRESET_ANCHOR_KEYS.map((key) => ({ key, value: "" })),
      items: [
        { id: crypto.randomUUID(), title: "Draft Earnings Curtain Raiser Press Release", offsetDays: -21, rowType: "reminder", reminderTime: "10:30", dateBasis: "event" },
        { id: crypto.randomUUID(), title: "Draft Earnings Report Press Release", offsetDays: -14, rowType: "reminder", reminderTime: "10:30", dateBasis: "event" },
        { id: crypto.randomUUID(), title: "Prepare Press Release Distribution for Tomorrow", body: sharedChecklistBody, offsetDays: -8, rowType: "reminder", reminderTime: "11:00", dateBasis: "event" },
        { id: crypto.randomUUID(), title: "Finalize Curtain Raiser Press Release", offsetDays: -8, rowType: "reminder", reminderTime: "16:00", dateBasis: "event" },
        { id: crypto.randomUUID(), title: "Earnings Curtain Raiser Going Out Now: Constant Contact & Social", offsetDays: -7, rowType: "reminder", reminderTime: "08:30", dateBasis: "event" },
        { id: crypto.randomUUID(), title: "Earnings Call Script Walk-Through", offsetDays: -1, rowType: "reminder", reminderTime: "14:30", dateBasis: "event", meetingDraft: { attendees: ["msteiner@verupharma.com", "hfisch@verupharma.com", "pgreenberg@verupharma.com", "mgreco@verupharma.com", "gbarnette@verupharma.com", "kgilbert@verupharma.com"], durationMinutes: 60 } },
        { id: crypto.randomUUID(), title: "Prepare Press Release Distribution for Tomorrow", body: sharedChecklistBody, offsetDays: -1, rowType: "reminder", reminderTime: "10:30", dateBasis: "event", durationDraft: { durationMinutes: 60 } },
        { id: crypto.randomUUID(), title: "Finalize Earnings Press Release", offsetDays: -1, rowType: "reminder", reminderTime: "16:00", dateBasis: "event" },
        { id: crypto.randomUUID(), title: "Send Chorus Call Intro Script and Authorized Callers", offsetDays: -1, rowType: "reminder", reminderTime: "10:30", dateBasis: "event" },
        { id: crypto.randomUUID(), title: "Set Alarm for Early Morning (5:30 AM)", offsetDays: -1, rowType: "reminder", reminderTime: "10:30", dateBasis: "event" },
        { id: crypto.randomUUID(), title: "Earnings Press Release Going Out Now: Constant Contact & Social Media", offsetDays: 0, rowType: "reminder", reminderTime: "06:30", dateBasis: "event", durationDraft: { durationMinutes: 10 } },
        { id: crypto.randomUUID(), title: "Print Earnings Scripts & PR", offsetDays: 0, rowType: "reminder", reminderTime: "06:50", dateBasis: "event", durationDraft: { durationMinutes: 10 } },
        { id: crypto.randomUUID(), title: "Final Walk-Through Earnings Call Script", offsetDays: 0, rowType: "reminder", reminderTime: "07:00", dateBasis: "event", meetingDraft: { attendees: ["msteiner@verupharma.com", "hfisch@verupharma.com", "pgreenberg@verupharma.com", "mgreco@verupharma.com", "gbarnette@verupharma.com", "kgilbert@verupharma.com"], durationMinutes: 60 } },
        { id: crypto.randomUUID(), title: "Earnings Call", offsetDays: 0, rowType: "reminder", reminderTime: "08:00", dateBasis: "event", durationDraft: { durationMinutes: 60 } },
        { id: crypto.randomUUID(), title: "Email Kari Sharp Earnings Webcast Link", offsetDays: 1, rowType: "reminder", reminderTime: "10:30", dateBasis: "today" },
      ],
    },
  ];
  return baseTemplates;
}

function cloneTemplateRows(items: BuilderRow[]) {
  return items.map((item) => ({
    ...item,
    id: crypto.randomUUID(),
    emailDraft: normalizeEmailDraft(item.emailDraft),
    durationDraft: item.durationDraft ? { ...item.durationDraft } : undefined,
    meetingDraft: item.meetingDraft ? { ...normalizeMeetingDraft(item.meetingDraft) } : undefined,
  }));
}

function cloneAnchors(items: BuilderAnchor[]) {
  return items.map((anchor) => ({
    ...anchor,
    id: crypto.randomUUID(),
  }));
}

function hasPresetAnchorsForBaseType(
  baseType: PlanType,
  anchors: Array<{ key: string; value: string }>
) {
  const normalizedKeys = new Set(anchors.map((anchor) => normalizeAnchorKey(anchor.key)));
  const requiredKeys =
    baseType === "press_release"
      ? PRESS_RELEASE_PRESET_ANCHOR_KEYS
      : baseType === "conference"
        ? CONFERENCE_PRESET_ANCHOR_KEYS
        : EARNINGS_PRESET_ANCHOR_KEYS;

  return requiredKeys.every((key) => normalizedKeys.has(normalizeAnchorKey(key)));
}

function inferTemplateMode(template: {
  id: string;
  baseType: PlanType;
  anchors: Array<{ key: string; value: string }>;
  templateMode?: "template" | "custom";
}) {
  if (template.templateMode) return template.templateMode;
  if (template.id.startsWith("seed:")) return "template";
  return hasPresetAnchorsForBaseType(template.baseType, template.anchors) ? "template" : "custom";
}

function normalizeImportedTemplate(template: {
  id: string;
  name: string;
  baseType: PlanType;
  templateMode?: "template" | "custom";
  noEventDate?: boolean;
  weekendRule: WeekendRule;
  anchors: Array<{ key: string; value: string }>;
  items: BuilderRow[];
}): SavedPlanTemplate {
  const protectedDefinition = getProtectedTemplateDefinition(template);
  return {
    id: protectedDefinition?.id ?? template.id,
    name: protectedDefinition?.name ?? template.name,
    baseType: protectedDefinition?.baseType ?? template.baseType,
    templateMode: protectedDefinition ? "template" : inferTemplateMode(template),
    noEventDate: Boolean(template.noEventDate),
    weekendRule: template.weekendRule,
    anchors: template.anchors
      .map((anchor) => ({ key: anchor.key.trim(), value: anchor.value }))
      .filter((anchor) => anchor.key),
    items: cloneTemplateRows(template.items),
  };
}

function toPersistedTemplate(template: SavedPlanTemplate, sortOrder: number): PersistedPlanTemplate {
  return {
    id: template.id,
    name: template.name,
    baseType: template.baseType,
    templateMode: inferTemplateMode(template),
    noEventDate: Boolean(template.noEventDate),
    weekendRule: template.weekendRule,
    anchors: template.anchors.map((anchor) => ({ ...anchor })),
    items: template.items.map((row) => ({
      ...row,
      offsetDays: row.offsetDays ?? 0,
      body: row.body ?? "",
      emailDraft: row.emailDraft ? normalizeEmailDraft(row.emailDraft) : undefined,
      durationDraft: row.durationDraft ? { ...row.durationDraft } : undefined,
      meetingDraft: row.meetingDraft ? { ...normalizeMeetingDraft(row.meetingDraft) } : undefined,
    })),
    isProtected: isProtectedTemplate(template),
    sortOrder,
  };
}

function fromPersistedTemplate(template: PersistedPlanTemplate): SavedPlanTemplate {
  return normalizeImportedTemplate({
    id: template.id,
    name: template.name,
    baseType: template.baseType,
    templateMode: template.templateMode,
    noEventDate: template.noEventDate,
    weekendRule: template.weekendRule,
    anchors: template.anchors.map((anchor) => ({ ...anchor })),
    items: template.items.map((row) => ({
      id: row.id,
      title: row.title,
      body: row.body ?? "",
      offsetDays: row.offsetDays,
      dateBasis: row.dateBasis ?? "event",
      rowType: row.rowType,
      reminderTime: row.reminderTime ?? "",
      emailDraft: row.emailDraft,
      durationDraft: row.durationDraft,
      meetingDraft: row.meetingDraft,
    })),
  });
}

function buildPersistedTemplateState(
  templates: SavedPlanTemplate[],
  selectedTemplateId: string | null
): PersistedTemplateState {
  return {
    selectedTemplateId,
    templates: templates.map((template, index) => toPersistedTemplate(template, index)),
  };
}

function buildTemplateItemsFromRows(rows: BuilderRow[]): TemplateItem[] {
  return rows.map((row) => ({
    id: row.id,
    title: row.title.trim() || "Untitled row",
    body: row.body?.trim() || undefined,
    offsetDays: row.offsetDays ?? 0,
    dateBasis: row.dateBasis ?? "event",
    rowType: row.rowType,
    reminderTime: row.reminderTime?.trim() || undefined,
    emailDraft: row.rowType === "email" ? normalizeEmailDraft(row.emailDraft) : undefined,
    durationDraft: row.durationDraft,
    meetingDraft: row.meetingDraft ? normalizeMeetingDraft(row.meetingDraft) : undefined,
  }));
}

function buildAnchorStateForType(type: PlanType, previous: BuilderAnchor[]) {
  const previousMap = buildAnchorMap(previous);
  return getPresetAnchorKeysForType(type).map((key) =>
    createLockedAnchor(key, previousMap.get(normalizeAnchorKey(key)) ?? "")
  );
}

function formatOffsetLabel(offsetDays: number | null | undefined, options?: { relativeToToday?: boolean; dateBasis?: PlanDateBasis }) {
  if (offsetDays == null) return "No days specified";
  const absoluteDays = Math.abs(offsetDays);
  const dayLabel = absoluteDays === 1 ? "day" : "days";
  const relativeToToday = options?.dateBasis === "today" || options?.relativeToToday;
  if (relativeToToday) {
    if (offsetDays < 0) return absoluteDays === 1 ? "1 day ago" : `${absoluteDays} days ago`;
    if (offsetDays > 0) return `${absoluteDays} ${dayLabel} from today`;
    return "Today";
  }
  if (offsetDays < 0) return `${absoluteDays} ${dayLabel} before event`;
  if (offsetDays > 0) return `${absoluteDays} ${dayLabel} after event`;
  return "Day of event";
}

function formatPreviewTime(time: string) {
  if (!time) return "";
  const parsedTime = parseTimeInput(time);
  if (!parsedTime) return time;
  const [hours, minutes] = parsedTime.split(":").map(Number);
  const parsed = new Date(2000, 0, 1, hours ?? 0, minutes ?? 0);
  return parsed.toLocaleTimeString("en-US", {
    hour: "numeric",
    minute: "2-digit",
  });
}

function buildFinalEmailBody(body: string, signatureSettings: { enabled: boolean; signature: string }) {
  const normalizedBody = body.trim();
  const normalizedSignature = signatureSettings.signature.trim();
  if (signatureSettings.enabled && normalizedSignature) {
    return normalizedBody ? `${normalizedBody}\n\n${normalizedSignature}` : normalizedSignature;
  }
  return normalizedBody;
}

function getBuilderEmailModeMessage(mode: EmailHandlingMode) {
  if (mode === "schedule") {
    return "Emails will be scheduled to send at the specified date and time.";
  }
  if (mode === "send") {
    return "Email will be sent immediately";
  }
  return "Email will be saved to your Drafts";
}

function getPreviewEmailModeMessage(mode: EmailHandlingMode) {
  if (mode === "schedule") {
    return "*** Emails will be scheduled to send at the specified date and time. ***";
  }
  if (mode === "send") {
    return "*** Emails will be sent immediately ***";
  }
  return "*** Emails will be saved to Drafts ***";
}

function getPreviewEmailActionLabel(mode: EmailHandlingMode) {
  if (mode === "schedule") return "Schedule Email";
  if (mode === "send") return "Send Email";
  return "Save to Drafts";
}

function getEffectivePreviewItemDate(item: Plan["items"][number]) {
  return item.customDueDate ?? item.dueDate;
}

function getBuilderRowTypeMeta(row: BuilderRow) {
  const rowKind = classifyPlanRow(row);
  if (rowKind === "email") {
    return { label: "Email", className: "text-green-500" };
  }
  if (rowKind === "meeting") {
    return { label: "Meeting", className: "text-violet-500" };
  }
  if (row.rowType === "calendar_event") {
    return { label: "Calendar Event", className: "text-violet-500" };
  }
  return { label: "Reminder", className: "text-blue-500" };
}

function EmailTokensInput({
  label,
  values,
  onChange,
  placeholder,
  hasError,
}: {
  label?: string;
  values: string[];
  onChange: (nextValues: string[]) => void;
  placeholder?: string;
  hasError?: boolean;
}) {
  const [draftValue, setDraftValue] = useState("");

  function commitRawValue(raw: string) {
    const nextTokens = raw
      .split(/[,\n;]/)
      .map((entry) => normalizeEmailToken(entry))
      .filter((entry) => entry && isValidEmailToken(entry));
    if (nextTokens.length === 0) return;
    onChange(Array.from(new Set([...values, ...nextTokens])));
    setDraftValue("");
  }

  return (
    <div>
      <label className="mb-1 block text-xs font-medium uppercase tracking-wide text-zinc-500">{label ?? ""}</label>
      <div
        className={`flex min-h-[44px] flex-wrap items-center gap-2 rounded-lg border bg-white px-3 py-2 ${
          hasError ? "border-red-400 ring-1 ring-red-100" : "border-gray-300"
        }`}
      >
        {values.map((value) => (
          <span key={value} className="inline-flex items-center gap-1 rounded-full bg-amber-100 px-2 py-1 text-xs text-amber-950">
            <span>{value}</span>
            <button
              type="button"
              onClick={() => onChange(values.filter((entry) => entry !== value))}
              className="text-amber-700"
              aria-label={`Remove ${value}`}
            >
              ×
            </button>
          </span>
        ))}
        <input
          className="min-w-[140px] flex-1 border-0 bg-transparent p-0 text-sm text-gray-900 placeholder:text-gray-700 outline-none"
          value={draftValue}
          placeholder={placeholder}
          onChange={(e) => setDraftValue(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter" || e.key === ",") {
              e.preventDefault();
              commitRawValue(draftValue);
            }
          }}
          onBlur={() => commitRawValue(draftValue)}
        />
      </div>
    </div>
  );
}

function areOutlookConnectionStatesEqual(left: OutlookConnectionState | null, right: OutlookConnectionState | null) {
  return JSON.stringify(left) === JSON.stringify(right);
}

export default function PlansPage() {
  const initialSettings = DEFAULT_APP_SETTINGS;
  const initialSeedTemplates = buildSeedTemplates(
    initialSettings.defaultReminderTime,
    initialSettings.defaultPressReleaseTime
  );
  const initialTemplates = initialSeedTemplates;
  const initialSelectedTemplateId = initialTemplates[0]?.id ?? null;
  const initialSelectedTemplate =
    initialTemplates.find((template) => template.id === initialSelectedTemplateId) ?? initialTemplates[0] ?? null;
  const [appSettings, setAppSettings] = useState<AppSettings>(() => initialSettings);
  const [outlookConnection, setOutlookConnection] = useState<OutlookConnectionState | null>(() =>
    getOutlookConnectionState(initialSettings.outlookAccountEmail)
  );
  const [builderMode, setBuilderMode] = useState<BuilderMode>("template");
  const [planType, setPlanType] = useState<PlanType>(initialSelectedTemplate?.baseType ?? "press_release");
  const [templateName, setTemplateName] = useState(initialSelectedTemplate?.name ?? "Press Release");
  const [eventName, setEventName] = useState(initialSelectedTemplate?.name ?? "Press Release");
  const [anchorDate, setAnchorDate] = useState(todayYYYYMMDD());
  const [noEventDate, setNoEventDate] = useState(false);
  const [weekendRule, setWeekendRule] = useState<WeekendRule>(initialSelectedTemplate?.weekendRule ?? "prior_business_day");
  const [rows, setRows] = useState<BuilderRow[]>(() =>
    cloneTemplateRows(
      initialSelectedTemplate?.items ?? []
    )
  );
  const [anchors, setAnchors] = useState<BuilderAnchor[]>(() =>
    buildAnchorStateForType(
      initialSelectedTemplate?.baseType ?? "press_release",
      (initialSelectedTemplate?.anchors ?? []).map((anchor) => ({
        id: crypto.randomUUID(),
        key: anchor.key,
        value: anchor.value,
      }))
    )
  );
  const [savedTemplates, setSavedTemplates] = useState<SavedPlanTemplate[]>(() => initialTemplates);
  const [selectedTemplateId, setSelectedTemplateId] = useState<string | null>(initialSelectedTemplateId);
  const [lastTemplateSnapshot, setLastTemplateSnapshot] = useState<BuilderStateSnapshot | null>(null);
  const [guidedForm, setGuidedForm] = useState<GuidedFormState>(() => createEmptyGuidedForm());
  const [isTemplateManageMode, setIsTemplateManageMode] = useState(false);
  const [isTemplateCopyMode, setIsTemplateCopyMode] = useState(false);
  const [areAnchorsHidden, setAreAnchorsHidden] = useState(false);
  const [templateActionMessage, setTemplateActionMessage] = useState("");
  const [openMenuRowId, setOpenMenuRowId] = useState<string | null>(null);
  const [openEmailDraftRowId, setOpenEmailDraftRowId] = useState<string | null>(null);
  const [openDetailIndicatorId, setOpenDetailIndicatorId] = useState<string | null>(null);
  const [pinnedDetailIndicatorId, setPinnedDetailIndicatorId] = useState<string | null>(null);
  const [openMeetingEditorRowId, setOpenMeetingEditorRowId] = useState<string | null>(null);
  const [openDurationEditorRowId, setOpenDurationEditorRowId] = useState<string | null>(null);
  const [forcedOpenMeetingEditorRowIds, setForcedOpenMeetingEditorRowIds] = useState<string[]>([]);
  const [meetingValidationErrors, setMeetingValidationErrors] = useState<MeetingValidationErrorState>({});
  const [openBodyEditorRowId, setOpenBodyEditorRowId] = useState<string | null>(null);
  const [emailFieldVisibility, setEmailFieldVisibility] = useState<EmailFieldVisibility>({});
  const [editingOffsetRowId, setEditingOffsetRowId] = useState<string | null>(null);
  const [offsetDrafts, setOffsetDrafts] = useState<Record<string, string>>({});
  const [focusedTimeInputRowId, setFocusedTimeInputRowId] = useState<string | null>(null);
  const [timeInputDrafts, setTimeInputDrafts] = useState<Record<string, string>>({});
  const [isBuilderPreviewOpen, setIsBuilderPreviewOpen] = useState(false);
  const [expandedPreviewReminderRowIds, setExpandedPreviewReminderRowIds] = useState<string[]>([]);
  const [expandedPreviewEmailRowIds, setExpandedPreviewEmailRowIds] = useState<string[]>([]);
  const [expandedPreviewMeetingRowIds, setExpandedPreviewMeetingRowIds] = useState<string[]>([]);
  const [openPreviewRowMenuId, setOpenPreviewRowMenuId] = useState<string | null>(null);
  const [executionNotice, setExecutionNotice] = useState<ExecutionNotice | null>(null);
  const [hasMounted, setHasMounted] = useState(false);
  const [hasHydratedTemplates, setHasHydratedTemplates] = useState(false);
  const hasLocalTemplateMutationRef = useRef(false);
  const kebabMenuRef = useRef<HTMLDivElement | null>(null);
  const builderTimeInputRefs = useRef<Record<string, HTMLInputElement | HTMLTextAreaElement | null>>({});

  useEffect(() => {
    function handleClickOutside(event: MouseEvent) {
      if (!kebabMenuRef.current) return;
      if (!kebabMenuRef.current.contains(event.target as Node)) {
        setOpenMenuRowId(null);
      }
    }

    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);

  useEffect(() => {
    setHasMounted(true);
  }, []);

  useEffect(() => {
    function handleSettingsRefresh() {
      const nextSettings = loadAppSettings();
      setAppSettings((current) => (areAppSettingsEqual(current, nextSettings) ? current : nextSettings));
    }

    window.addEventListener("focus", handleSettingsRefresh);
    window.addEventListener(APP_SETTINGS_UPDATED_EVENT, handleSettingsRefresh as EventListener);
    return () => {
      window.removeEventListener("focus", handleSettingsRefresh);
      window.removeEventListener(APP_SETTINGS_UPDATED_EVENT, handleSettingsRefresh as EventListener);
    };
  }, []);

  useEffect(() => {
    let active = true;

    async function hydrateSettings() {
      const hydratedSettings = await hydrateAppSettingsFromSupabase();
      if (!active) return;
      setAppSettings((current) => (areAppSettingsEqual(current, hydratedSettings) ? current : hydratedSettings));
    }

    void hydrateSettings();

    return () => {
      active = false;
    };
  }, []);

  // The initial remote template hydration intentionally runs once on mount.
  /* eslint-disable react-hooks/exhaustive-deps */
  useEffect(() => {
    let active = true;

    async function hydrateTemplates() {
      const seedState = buildSeedTemplates(
        loadAppSettings().defaultReminderTime,
        loadAppSettings().defaultPressReleaseTime
      ).map((template, index) => toPersistedTemplate(template, index));
      const remoteState = await loadTemplateStateFromSupabase(seedState);
      const nextState = remoteState ?? loadCachedTemplateState(seedState);

      if (!active) return;
      if (hasLocalTemplateMutationRef.current) {
        setHasHydratedTemplates(true);
        return;
      }

      saveCachedTemplateState(nextState);
      const nextTemplates = nextState.templates.map((template) => fromPersistedTemplate(template));
      setSavedTemplates(nextTemplates);
      setHasHydratedTemplates(true);

      const nextSelectedTemplate =
        nextTemplates.find((template) => template.id === nextState.selectedTemplateId) ?? nextTemplates[0] ?? null;

      if (nextSelectedTemplate) {
        applyTemplateRecord(nextSelectedTemplate);
      }
    }

    void hydrateTemplates();

    return () => {
      active = false;
    };
  }, []);
  /* eslint-enable react-hooks/exhaustive-deps */

  useEffect(() => {
    if (!hasHydratedTemplates) return;
    const nextState = buildPersistedTemplateState(savedTemplates, selectedTemplateId);
    saveCachedTemplateState(nextState);
    void saveTemplateStateToSupabase(nextState);
  }, [hasHydratedTemplates, savedTemplates, selectedTemplateId]);

  useEffect(() => {
    if (selectedTemplateId && !savedTemplates.some((template) => template.id === selectedTemplateId)) {
      const fallbackTemplate = savedTemplates[0] ?? null;
      if (fallbackTemplate) {
        applyTemplateRecord(fallbackTemplate);
      } else {
        setSelectedTemplateId(null);
      }
    }
    // This repair effect intentionally runs from template/selection state only.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [savedTemplates, selectedTemplateId]);

  useEffect(() => {
    let active = true;

    async function refreshConnection() {
      const connection = await resolveOutlookConnectionState(appSettings.outlookAccountEmail);
      if (!active) return;
      setOutlookConnection((current) => (areOutlookConnectionStatesEqual(current, connection) ? current : connection));
    }

    void refreshConnection();

    function handleConnectionRefresh() {
      void refreshConnection();
    }

    window.addEventListener("focus", handleConnectionRefresh);
    window.addEventListener(OUTLOOK_CONNECTION_UPDATED_EVENT, handleConnectionRefresh as EventListener);
    return () => {
      active = false;
      window.removeEventListener("focus", handleConnectionRefresh);
      window.removeEventListener(OUTLOOK_CONNECTION_UPDATED_EVENT, handleConnectionRefresh as EventListener);
    };
  }, [appSettings.outlookAccountEmail]);

  const normalizedDefaultReminderTime = normalizeReminderTimeInput(appSettings.defaultReminderTime);
  const normalizedDefaultPressReleaseTime = normalizeReminderTimeInput(appSettings.defaultPressReleaseTime);
  const currentSelectedTemplate = selectedTemplateId
    ? savedTemplates.find((template) => template.id === selectedTemplateId) ?? null
    : null;
  const effectiveTemplateMode = currentSelectedTemplate ? inferTemplateMode(currentSelectedTemplate) : builderMode;

  const effectivePlanName =
    effectiveTemplateMode === "template" &&
    (planType === "press_release" || planType === "earnings" || planType === "conference")
      ? getGuidedTemplateDisplayName(planType, guidedForm) || eventName || templateName || getSeedTemplateName(planType)
      : eventName || templateName || getSeedTemplateName(planType);
  const previewEffectiveAnchorDate = noEventDate ? todayYYYYMMDD() : anchorDate;
  const previewAnchorDateForComputation = previewEffectiveAnchorDate || todayYYYYMMDD();
  const genericEventAnchorName = eventName.trim();
  const genericEventAnchorDate = noEventDate ? "" : anchorDate;
  const resolvedAnchors = anchors.map((anchor) => ({
    id: anchor.id,
    key: anchor.key,
    value:
      getDerivedAnchorValue(planType, effectivePlanName, anchorDate, anchor.key, guidedForm, {
        eventNameValue: genericEventAnchorName,
        eventDateValue: genericEventAnchorDate,
      }) ?? anchor.value,
    displayValue: getAnchorDisplayValue(
      anchor.key,
      getDerivedAnchorValue(planType, effectivePlanName, anchorDate, anchor.key, guidedForm, {
        eventNameValue: genericEventAnchorName,
        eventDateValue: genericEventAnchorDate,
      }) ?? anchor.value
    ),
    locked: anchor.locked,
  }));
  const anchorMap = buildAnchorMap(resolvedAnchors);
  const previewPlan = resolvePlanAnchors(
    createPlan({
      name: effectivePlanName,
      type: planType,
      anchorDate: previewAnchorDateForComputation,
      weekendRule,
      template: buildTemplateItemsFromRows(rows),
    }),
    anchorMap
  );
  const previewPlanForRender = previewEffectiveAnchorDate ? previewPlan : null;
  const selectedEditableTemplate =
    currentSelectedTemplate && !isProtectedTemplate(currentSelectedTemplate) ? currentSelectedTemplate : null;

  function persistTemplateStateImmediately(nextTemplates: SavedPlanTemplate[], nextSelectedTemplateId: string | null) {
    const nextState = buildPersistedTemplateState(nextTemplates, nextSelectedTemplateId);
    saveCachedTemplateState(nextState);
    if (hasHydratedTemplates) {
      void saveTemplateStateToSupabase(nextState);
    }
  }

  function updateRow(rowId: string, updater: (row: BuilderRow) => BuilderRow) {
    setRows((currentRows) => {
      const nextRows = currentRows.map((row) => (row.id === rowId ? updater(row) : row));
      if (meetingValidationErrors[rowId]) {
        const nextRow = nextRows.find((row) => row.id === rowId);
        const nextErrors = nextRow
          ? getMeetingValidationErrorsForRow(nextRow, previewPlan.items.find((item) => item.id === rowId))
          : null;
        setMeetingValidationErrors((current) => {
          const updated = { ...current };
          if (nextErrors) {
            updated[rowId] = nextErrors;
          } else {
            delete updated[rowId];
          }
          return updated;
        });
        if (!nextErrors) {
          setForcedOpenMeetingEditorRowIds((current) => current.filter((id) => id !== rowId));
        }
      }
      return nextRows;
    });
  }

  function onInsertAnchor(anchorKey: string) {
    if (typeof document === "undefined") return;

    const normalizedKey = normalizeAnchorKey(anchorKey);
    if (!normalizedKey) return;

    const activeElement = document.activeElement;
    if (
      !activeElement ||
      !(activeElement instanceof HTMLInputElement || activeElement instanceof HTMLTextAreaElement) ||
      activeElement.readOnly ||
      activeElement.disabled
    ) {
      return;
    }

    const token = `[${normalizedKey}]`;
    const start = activeElement.selectionStart ?? activeElement.value.length;
    const end = activeElement.selectionEnd ?? start;
    const nextValue = `${activeElement.value.slice(0, start)}${token}${activeElement.value.slice(end)}`;

    const descriptor = Object.getOwnPropertyDescriptor(
      activeElement instanceof HTMLTextAreaElement ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype,
      "value"
    );
    descriptor?.set?.call(activeElement, nextValue);
    activeElement.dispatchEvent(new Event("input", { bubbles: true }));

    const nextCursorPosition = start + token.length;
    requestAnimationFrame(() => {
      activeElement.focus();
      activeElement.setSelectionRange(nextCursorPosition, nextCursorPosition);
    });
  }

  function downloadTextFile(filename: string, content: string) {
    if (typeof window === "undefined") return;
    const blob = new Blob([content], { type: "text/plain;charset=utf-8" });
    const objectUrl = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = objectUrl;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(objectUrl);
  }

  function confirmExport() {
    if (typeof window === "undefined") return true;
    return window.confirm("Are you sure you want to export?");
  }

  function getMeetingValidationMessage(itemIds?: string[]) {
    const scopedRows = itemIds ? rows.filter((row) => itemIds.includes(row.id)) : rows;
    const nextErrors: MeetingValidationErrorState = {};
    const invalidMeetings = scopedRows
      .filter((row) => classifyPlanRow(row) === "meeting")
      .map((row) => {
        const rowErrors = getMeetingValidationErrorsForRow(
          row,
          previewPlan.items.find((item) => item.id === row.id)
        );
        if (!rowErrors) return null;
        nextErrors[row.id] = rowErrors;
        const missing: string[] = [];
        if (rowErrors.attendees) missing.push("attendees");
        if (rowErrors.time) missing.push("time");
        if (rowErrors.duration) missing.push("duration");
        return `${row.title.trim() || "Untitled meeting"} (${missing.join(", ")})`;
      })
      .filter((entry): entry is string => Boolean(entry));

    setMeetingValidationErrors(nextErrors);
    if (invalidMeetings.length === 0) {
      setForcedOpenMeetingEditorRowIds([]);
      return null;
    }
    const affectedIds = Object.keys(nextErrors);
    setForcedOpenMeetingEditorRowIds(affectedIds);
    setOpenMeetingEditorRowId(affectedIds[0] ?? null);
    setOpenDurationEditorRowId(null);
    return `Some meetings are missing required information.\n\n${invalidMeetings.join("\n")}`;
  }

  function getMeetingValidationErrorsForRow(
    row: BuilderRow,
    previewItem?: Plan["items"][number]
  ): MeetingValidationErrorState[string] | null {
    const draft = normalizeMeetingDraft(row.meetingDraft);
    if (!draft) return null;

    const nextErrors: MeetingValidationErrorState[string] = {};
    if (draft.attendees.length === 0) {
      nextErrors.attendees = true;
    }
    if (!draft.isAllDay && !getUsableReminderTime(previewItem?.reminderTime ?? row.reminderTime, anchorMap)) {
      nextErrors.time = true;
    }
    if (draft.useCustomEnd && !draft.isAllDay) {
      if (!draft.endDate.trim() || !draft.endTime.trim()) {
        nextErrors.duration = true;
      }
    } else if (!(draft.durationMinutes > 0)) {
      nextErrors.duration = true;
    }

    return Object.keys(nextErrors).length > 0 ? nextErrors : null;
  }

  function warnIfMissingReminderTimes(options?: { usePopup?: boolean; itemIds?: string[] }) {
    const scopedRows = options?.itemIds ? rows.filter((row) => options.itemIds?.includes(row.id)) : rows;
    const hasMissingTimes = scopedRows.some((row) => {
      const rowKind = classifyPlanRow(row);
      if (rowKind === "email") return false;
      if (row.meetingDraft?.isAllDay || row.durationDraft?.isAllDay) return false;
      return !getUsableReminderTime(
        previewPlan.items.find((item) => item.id === row.id)?.reminderTime ?? row.reminderTime,
        anchorMap
      );
    });
    if (!hasMissingTimes) return false;
    if (options?.usePopup && typeof window !== "undefined") {
      window.alert("One or more reminder rows are missing a reminder time.");
    }
    return true;
  }

  function validateMeetingRowsForExport(itemIds?: string[], options?: { usePopup?: boolean }) {
    const message = getMeetingValidationMessage(itemIds);
    if (!message) return false;
    if (options?.usePopup && typeof window !== "undefined") {
      window.alert(message);
    }
    return true;
  }

  function buildEmailDraftFileContent(item: Plan["items"][number]) {
    const draft = normalizeEmailDraft(item.emailDraft);
    const finalBody = buildFinalEmailBody(draft.body, {
      enabled: appSettings.emailSignatureEnabled,
      signature: appSettings.emailSignatureText,
    });
    return [
      `To: ${draft.to.join(", ")}`,
      `Cc: ${draft.cc.join(", ")}`,
      `Bcc: ${draft.bcc.join(", ")}`,
      `Subject: ${draft.subject.trim() || item.title || "Email draft"}`,
      `X-Event-Based-Reminders-Mode: local-${appSettings.emailHandlingMode}`,
      "",
      finalBody || "(No body)",
    ].join("\n");
  }

  function getResolvedEmailDraftForExecution(item: Plan["items"][number]) {
    const draft = normalizeEmailDraft(item.emailDraft);
    return {
      ...draft,
      body: buildFinalEmailBody(draft.body, {
        enabled: appSettings.emailSignatureEnabled,
        signature: appSettings.emailSignatureText,
      }),
    };
  }

  function buildLocalDateTimeIso(date: string, time: string) {
    return `${date}T${time}:00`;
  }

  function buildNextDayMidnightIso(date: string) {
    const [year, month, day] = date.split("-").map(Number);
    const next = new Date(year ?? 2000, (month ?? 1) - 1, day ?? 1, 0, 0, 0);
    next.setDate(next.getDate() + 1);
    return `${next.getFullYear()}-${String(next.getMonth() + 1).padStart(2, "0")}-${String(next.getDate()).padStart(2, "0")}T00:00:00`;
  }

  function buildPreviewItemEndDateTime(item: Plan["items"][number], startDate: string, startTime: string) {
    const meetingDraft = normalizeMeetingDraft(item.meetingDraft);
    const durationDraft = normalizeDurationDraft(item.durationDraft);

    if (meetingDraft?.useCustomEnd && meetingDraft.endDate && meetingDraft.endTime) {
      const customEndTime = parseTimeInput(meetingDraft.endTime) || meetingDraft.endTime;
      return buildLocalDateTimeIso(meetingDraft.endDate, customEndTime);
    }

    if (durationDraft?.useCustomEnd && durationDraft.endDate && durationDraft.endTime) {
      const customEndTime = parseTimeInput(durationDraft.endTime) || durationDraft.endTime;
      return buildLocalDateTimeIso(durationDraft.endDate, customEndTime);
    }

    const durationMinutes = meetingDraft?.durationMinutes ?? durationDraft?.durationMinutes ?? 30;
    const computedEnd = addMinutesToLocalDateTime(startDate, startTime, durationMinutes);
    return buildLocalDateTimeIso(computedEnd.endDate, computedEnd.endTime);
  }

  function buildPreviewItemGraphTiming(item: Plan["items"][number]) {
    const dueDate = getEffectivePreviewItemDate(item);
    const resolvedTime = getUsableReminderTime(item.reminderTime);
    const isAllDay = Boolean(item.meetingDraft?.isAllDay || item.durationDraft?.isAllDay);

    if (!resolvedTime || isAllDay) {
      return {
        startISO: `${dueDate}T00:00:00`,
        endISO: buildNextDayMidnightIso(dueDate),
        isAllDay: true,
      };
    }

    return {
      startISO: buildLocalDateTimeIso(dueDate, resolvedTime),
      endISO: buildPreviewItemEndDateTime(item, dueDate, resolvedTime),
      isAllDay: false,
    };
  }

  function getEmailScheduledSendISO(item: Plan["items"][number]) {
    const dueDate = getEffectivePreviewItemDate(item);
    const resolvedTime = getUsableReminderTime(item.reminderTime);
    if (!resolvedTime) return "";
    return buildLocalDateTimeIso(dueDate, resolvedTime);
  }

  function getExecutionHistoryItemType(item: Plan["items"][number]) {
    const rowKind = classifyPlanRow(item);
    if (rowKind === "meeting" && item.meetingDraft?.teamsMeeting) return "teams_meeting";
    return rowKind;
  }

  function getExecutionHistoryTitle(item: Plan["items"][number]) {
    if (classifyPlanRow(item) === "email") {
      const draft = getResolvedEmailDraftForExecution(item);
      return draft.subject.trim() || item.customTitle || item.title || "Email draft";
    }
    return item.customTitle || item.title;
  }

  function getExecutionHistoryRecipients(item: Plan["items"][number]) {
    if (classifyPlanRow(item) !== "email") return [];
    const draft = getResolvedEmailDraftForExecution(item);
    return [...draft.to, ...draft.cc, ...draft.bcc].filter(Boolean);
  }

  function getExecutionHistoryAttendees(item: Plan["items"][number]) {
    if (classifyPlanRow(item) !== "meeting") return [];
    return normalizeMeetingDraft(item.meetingDraft)?.attendees ?? [];
  }

  function getExecutionHistoryDetailFields(item: Plan["items"][number]) {
    const rowKind = classifyPlanRow(item);

    if (rowKind === "email") {
      const draft = getResolvedEmailDraftForExecution(item);
      return {
        body: draft.body,
        emailDraft: {
          to: draft.to,
          cc: draft.cc,
          bcc: draft.bcc,
          subject: draft.subject,
          body: draft.body,
        },
      };
    }

    if (rowKind === "meeting") {
      const meetingDraft = normalizeMeetingDraft(item.meetingDraft);
      return {
        body: item.body ?? "",
        meetingDraft: {
          attendees: meetingDraft?.attendees ?? [],
          location: meetingDraft?.location ?? "",
          title: item.customTitle ?? item.title,
          body: item.body ?? "",
        },
      };
    }

    return {
      body: item.body ?? "",
    };
  }

  function getExecutionHistoryTiming(item: Plan["items"][number]) {
    const rowKind = classifyPlanRow(item);

    if (rowKind === "meeting" || rowKind === "reminder") {
      const timing = buildPreviewItemGraphTiming(item);
      return {
        scheduledFor: timing.startISO,
        endsAt: timing.endISO,
        isAllDay: timing.isAllDay,
      };
    }

    if (rowKind === "email" && appSettings.emailHandlingMode === "schedule") {
      return {
        scheduledFor: getEmailScheduledSendISO(item) || null,
        endsAt: null,
        isAllDay: false,
      };
    }

    const dueDate = getEffectivePreviewItemDate(item);
    const reminderTime = getUsableReminderTime(item.reminderTime);
    return {
      scheduledFor: dueDate && reminderTime ? buildLocalDateTimeIso(dueDate, reminderTime) : null,
      endsAt: null,
      isAllDay: false,
    };
  }

  function getExecutionHistoryCapabilities(options: {
    item: Plan["items"][number];
    status: "success" | "fallback" | "failed";
    path: "graph" | "fallback";
    result?: OutlookExecutionResult;
  }) {
    if (options.path !== "graph") {
      return {
        provider: "local_export" as const,
        providerObjectType: "file" as const,
        providerObjectId: null,
        canRecall: false,
        canModify: false,
        recallImplemented: false,
        modifyImplemented: false,
        recallReason: "Local exports do not create provider-backed objects, so they cannot be recalled.",
        modifyReason: "Local export files cannot be modified after download from History.",
      };
    }

    const action = options.result?.action ?? null;
    const rowKind = classifyPlanRow(options.item);
    const providerObjectType = rowKind === "email" ? "message" : "event";
    const providerObjectId = options.result?.providerObjectId ?? null;

    if (options.status !== "success" || !providerObjectId) {
      return {
        provider: "outlook" as const,
        providerObjectType,
        providerObjectId,
        canRecall: false,
        canModify: false,
        recallImplemented: false,
        modifyImplemented: false,
        recallReason: "This record does not include a provider object id, so recall is not available.",
        modifyReason: "This record does not include a provider object id, so modify is not available.",
      };
    }

    if (action === "email_sent") {
      return {
        provider: "outlook" as const,
        providerObjectType,
        providerObjectId,
        canRecall: false,
        canModify: false,
        recallImplemented: false,
        modifyImplemented: false,
        recallReason: "Sent emails are not recallable from this app.",
        modifyReason: "Sent emails cannot be modified after send.",
      };
    }

    return {
      provider: "outlook" as const,
      providerObjectType,
      providerObjectId,
      canRecall: true,
      canModify: true,
      recallImplemented: false,
      modifyImplemented: false,
      recallReason: "Outlook object metadata is stored, but provider recall is not implemented yet.",
      modifyReason: "Outlook object metadata is stored, but provider modify is not implemented yet.",
    };
  }

  async function recordExecutionHistory(options: {
    item: Plan["items"][number];
    status: "success" | "fallback" | "failed";
    path: "graph" | "fallback";
    executionGroupId?: string;
    result?: OutlookExecutionResult;
    fallbackExportKind?: "eml" | "ics";
    reason?: string;
  }) {
    try {
      const subject = getExecutionHistoryTitle(options.item);
      const timing = getExecutionHistoryTiming(options.item);
      const capabilities = getExecutionHistoryCapabilities(options);
      const detailFields = getExecutionHistoryDetailFields(options.item);
      await writeExecutionHistory({
        executionGroupId: options.executionGroupId ?? null,
        planName: previewPlan.name,
        itemType: getExecutionHistoryItemType(options.item),
        title: options.item.customTitle || options.item.title,
        subject,
        status: options.status,
        path: options.path,
        recipients: getExecutionHistoryRecipients(options.item),
        attendees: getExecutionHistoryAttendees(options.item),
        outlookWebLink: options.result?.webLink ?? null,
        teamsJoinLink: options.result?.joinUrl ?? null,
        fallbackExportKind: options.fallbackExportKind ?? null,
        provider: capabilities.provider,
        providerObjectId: capabilities.providerObjectId,
        providerObjectType: capabilities.providerObjectType,
        canRecall: capabilities.canRecall,
        canModify: capabilities.canModify,
        recallImplemented: capabilities.recallImplemented,
        modifyImplemented: capabilities.modifyImplemented,
        recallReason: capabilities.recallReason,
        modifyReason: capabilities.modifyReason,
        scheduledFor: timing.scheduledFor,
        endsAt: timing.endsAt,
        isAllDay: timing.isAllDay,
        details: {
          action: options.result?.action ?? null,
          message: options.result?.message ?? null,
          reason: options.reason ?? null,
          rowType: classifyPlanRow(options.item),
          itemId: options.item.id,
          scheduledSendAt:
            classifyPlanRow(options.item) === "email" && appSettings.emailHandlingMode === "schedule"
              ? getEmailScheduledSendISO(options.item)
              : null,
          reminderTime: options.item.reminderTime ?? null,
          teamsMeeting: Boolean(options.item.meetingDraft?.teamsMeeting),
          emailHandlingMode: classifyPlanRow(options.item) === "email" ? appSettings.emailHandlingMode : null,
          ...detailFields,
        },
      });
    } catch {
      // Preserve existing export UX if history persistence is unavailable.
    }
  }

  async function getGraphExecutionAvailability(requiredScopes: string[]) {
    const connection = await resolveOutlookConnectionState(appSettings.outlookAccountEmail, requiredScopes);
    if (connection.connected && !connection.stale && connection.supportedMailbox) {
      return { canUseGraph: true, connection };
    }

    const reason = connection.stale
      ? "Your Outlook connection no longer matches the selected email. Falling back to local export."
      : !connection.supportedMailbox && connection.identity
        ? `${connection.identity.mailboxEligibilityReason} Falling back to local export.`
        : connection.status === "reconnect_required"
          ? "Reconnect Outlook in Settings to continue. Falling back to local export."
          : "Outlook is not connected. Falling back to local export.";

    return { canUseGraph: false, connection, reason };
  }

  function downloadLocalEmailItem(
    item: Plan["items"][number],
    options?: { showAlert?: boolean; alertMessage?: string }
  ) {
    downloadTextFile(
      `${(normalizeEmailDraft(item.emailDraft).subject.trim() || item.title || "email-draft").toLowerCase().replace(/[^a-z0-9]+/g, "-") || "email-draft"}.eml`,
      buildEmailDraftFileContent(item)
    );
    if (options?.showAlert === false || typeof window === "undefined") return;
    window.alert(
      options?.alertMessage ||
        (appSettings.emailHandlingMode === "schedule"
          ? "Downloaded a local scheduled-email draft file. Outlook schedule-send is not connected in the Event-Based Reminders app yet."
          : appSettings.emailHandlingMode === "send"
            ? "Downloaded a local email draft file. Live send is not connected in the Event-Based Reminders app yet."
            : "Downloaded a local email draft file. Outlook draft creation is not connected in the Event-Based Reminders app yet.")
    );
  }

  function downloadLocalMeetingItem(
    item: Plan["items"][number],
    options?: { showAlert?: boolean; alertMessage?: string }
  ) {
    downloadICS(
      `${(item.title || "meeting").toLowerCase().replace(/[^a-z0-9]+/g, "-") || "meeting"}.ics`,
      buildICSForPlan({
        ...previewPlan,
        items: [item],
      })
    );
    if (options?.showAlert === false || typeof window === "undefined") return;
    window.alert(
      options?.alertMessage ||
        "Downloaded a local calendar file for this meeting. Outlook calendar creation is not connected in the Event-Based Reminders app yet."
    );
  }

  function downloadLocalReminderItem(
    item: Plan["items"][number],
    options?: { showAlert?: boolean; alertMessage?: string }
  ) {
    downloadICS(
      `${(item.title || previewPlan.name).toLowerCase().replace(/[^a-z0-9]+/g, "-") || "plan-item"}.ics`,
      buildICSForPlan({
        ...previewPlan,
        items: [item],
      })
    );
    if (options?.showAlert === false || typeof window === "undefined") return;
    window.alert(
      options?.alertMessage ||
        "Downloaded a local calendar file for this reminder. Outlook calendar creation is not connected in the Event-Based Reminders app yet."
    );
  }

  function downloadLocalPlanCalendarItems(
    items: Plan["items"],
    options?: { showAlert?: boolean; alertMessage?: string }
  ) {
    if (items.length === 0) return;
    downloadICS(
      `${previewPlan.name.toLowerCase().replace(/[^a-z0-9]+/g, "-") || "plan"}.ics`,
      buildICSForPlan({
        ...previewPlan,
        items,
      })
    );
    if (options?.showAlert === false || typeof window === "undefined") return;
    if (options?.alertMessage) {
      window.alert(options.alertMessage);
    }
  }

  function setExecutionNoticeForResult(result: OutlookExecutionResult) {
    const details = [result.title];

    setExecutionNotice({
      tone: "success",
      title: result.message,
      details,
    });
  }

  function setExecutionNoticeForFallback(options: {
    title: string;
    reason: string;
    fallbackMessage: string;
  }) {
    setExecutionNotice({
      tone: "mixed",
      title: options.title,
      message: options.reason,
      details: [options.fallbackMessage],
    });
  }

  function setExecutionNoticeForExportSummary(options: {
    graphUnavailableReason?: string;
    graphResults: OutlookExecutionResult[];
    failedEmailItems: Plan["items"];
    failedCalendarItems: Plan["items"];
    draftCount: number;
    scheduledCount: number;
    sentCount: number;
    reminderCount: number;
    meetingCount: number;
  }) {
    const details: string[] = [];

    if (options.draftCount > 0) {
      details.push(`${options.draftCount} Outlook draft${options.draftCount === 1 ? "" : "s"} created`);
    }
    if (options.scheduledCount > 0) {
      details.push(`${options.scheduledCount} Outlook email${options.scheduledCount === 1 ? "" : "s"} scheduled`);
    }
    if (options.sentCount > 0) {
      details.push(`${options.sentCount} Outlook email${options.sentCount === 1 ? "" : "s"} sent`);
    }
    if (options.reminderCount > 0) {
      details.push(`${options.reminderCount} Outlook reminder${options.reminderCount === 1 ? "" : "s"} created`);
    }
    if (options.meetingCount > 0) {
      details.push(`${options.meetingCount} Outlook meeting${options.meetingCount === 1 ? "" : "s"} created`);
    }
    if (options.failedEmailItems.length > 0) {
      details.push(
        `${options.failedEmailItems.length} email fallback file${options.failedEmailItems.length === 1 ? "" : "s"} downloaded`
      );
      details.push(
        `Email fallback: ${options.failedEmailItems
          .slice(0, 3)
          .map((item) => item.customTitle ?? item.title)
          .join(", ")}${options.failedEmailItems.length > 3 ? ", ..." : ""}`
      );
    }
    if (options.failedCalendarItems.length > 0) {
      details.push(
        `${options.failedCalendarItems.length} calendar item${options.failedCalendarItems.length === 1 ? "" : "s"} exported to ICS`
      );
      details.push(
        `Calendar fallback: ${options.failedCalendarItems
          .slice(0, 3)
          .map((item) => item.customTitle ?? item.title)
          .join(", ")}${options.failedCalendarItems.length > 3 ? ", ..." : ""}`
      );
    }
    const hasFallbacks = options.failedEmailItems.length > 0 || options.failedCalendarItems.length > 0;
    const hasGraphSuccesses = options.graphResults.length > 0;

    setExecutionNotice({
      tone: !hasGraphSuccesses && hasFallbacks ? "warning" : hasFallbacks ? "mixed" : "success",
      title: hasFallbacks
        ? hasGraphSuccesses
          ? "Some Outlook actions fell back to local export"
          : "Exported locally"
        : "Outlook export completed",
      message: options.graphUnavailableReason,
      details,
    });
  }

  async function executePreviewEmailViaGraph(item: Plan["items"][number]) {
    const draft = getResolvedEmailDraftForExecution(item);
    const fallbackSubject = draft.subject.trim() || item.title || "Email draft";
    const title = fallbackSubject;

    if (appSettings.emailHandlingMode === "send") {
      await sendOutlookEmailFromEmailDraft({
        draft,
        fallbackSubject,
        expectedEmail: appSettings.outlookAccountEmail,
      });
      return {
        kind: "email",
        action: "email_sent",
        title,
        message: "Outlook email sent.",
      } satisfies OutlookExecutionResult;
    }

    if (appSettings.emailHandlingMode === "schedule") {
      const scheduledSendISO = getEmailScheduledSendISO(item);
      if (!scheduledSendISO) {
        throw new Error("Add a time before scheduling this email.");
      }
      const result = await scheduleOutlookEmailFromEmailDraft({
        draft,
        fallbackSubject,
        scheduledSendISO,
        expectedEmail: appSettings.outlookAccountEmail,
      });
      return {
        kind: "email",
        action: "email_scheduled",
        title,
        message: "Outlook email scheduled.",
        providerObjectId: result.id,
        webLink: result.webLink,
      } satisfies OutlookExecutionResult;
    }

    const result = await createOutlookDraftFromEmailDraft({
      draft,
      fallbackSubject,
      expectedEmail: appSettings.outlookAccountEmail,
    });
    return {
      kind: "email",
      action: "draft_created",
      title,
      message: "Outlook draft created.",
      providerObjectId: result.id,
      webLink: result.webLink,
    } satisfies OutlookExecutionResult;
  }

  async function executePreviewCalendarViaGraph(item: Plan["items"][number]) {
    const timing = buildPreviewItemGraphTiming(item);
    const rowKind = classifyPlanRow(item);
    const result = await createOutlookCalendarEvent({
      subject: item.customTitle ?? item.title,
      bodyText: item.body?.trim() || "",
      startISO: timing.startISO,
      endISO: timing.endISO,
      timeZone: "America/New_York",
      isAllDay: timing.isAllDay,
      location: item.meetingDraft?.location,
      attendees: item.meetingDraft?.attendees ?? [],
      teamsMeeting: item.meetingDraft?.teamsMeeting,
      expectedEmail: appSettings.outlookAccountEmail,
    });
    return {
      kind: rowKind === "meeting" ? "meeting" : "reminder",
      action: rowKind === "meeting" ? "meeting_created" : "reminder_created",
      title: item.customTitle ?? item.title,
      message: rowKind === "meeting" ? "Outlook meeting created." : "Outlook calendar reminder created.",
      providerObjectId: result.id,
      webLink: result.webLink,
      joinUrl: result.joinUrl,
    } satisfies OutlookExecutionResult;
  }

  async function exportPreviewEmailItem(itemId: string) {
    const item = previewPlan.items.find((entry) => entry.id === itemId);
    if (!item) return;
    if (warnIfMissingReminderTimes({ usePopup: true, itemIds: [itemId] })) return;
    if (!confirmExport()) return;
    const executionGroupId = crypto.randomUUID();

    const graphAvailability = await getGraphExecutionAvailability(["Mail.ReadWrite", "Mail.Send"]);
    if (graphAvailability.canUseGraph) {
      try {
        const result = await executePreviewEmailViaGraph(item);
        await recordExecutionHistory({
          item,
          status: "success",
          path: "graph",
          executionGroupId,
          result,
        });
        setExecutionNoticeForResult(result);
        return;
      } catch (error) {
        const reason = error instanceof Error ? error.message : "Outlook email action failed.";
        downloadLocalEmailItem(item, {
          showAlert: false,
        });
        await recordExecutionHistory({
          item,
          status: "fallback",
          path: "fallback",
          executionGroupId,
          fallbackExportKind: "eml",
          reason,
        });
        setExecutionNoticeForFallback({
          title: "Email action fell back to local export",
          reason,
          fallbackMessage: "Downloaded a local email draft file instead.",
        });
        return;
      }
    }

    downloadLocalEmailItem(item, {
      showAlert: false,
    });
    await recordExecutionHistory({
      item,
      status: "fallback",
      path: "fallback",
      executionGroupId,
      fallbackExportKind: "eml",
      reason: graphAvailability.reason ?? "Outlook is not connected. Falling back to local export.",
    });
    setExecutionNoticeForFallback({
      title: "Email action fell back to local export",
      reason: graphAvailability.reason ?? "Outlook is not connected. Falling back to local export.",
      fallbackMessage: "Downloaded a local email draft file instead.",
    });
  }

  async function exportPreviewMeetingItem(itemId: string) {
    const item = previewPlan.items.find((entry) => entry.id === itemId);
    if (!item) return;
    if (validateMeetingRowsForExport([itemId], { usePopup: true })) return;
    if (warnIfMissingReminderTimes({ usePopup: true, itemIds: [itemId] })) return;
    if (!confirmExport()) return;
    const executionGroupId = crypto.randomUUID();

    const graphAvailability = await getGraphExecutionAvailability(["Calendars.ReadWrite"]);
    if (graphAvailability.canUseGraph) {
      try {
        const result = await executePreviewCalendarViaGraph(item);
        await recordExecutionHistory({
          item,
          status: "success",
          path: "graph",
          executionGroupId,
          result,
        });
        setExecutionNoticeForResult(result);
        return;
      } catch (error) {
        const reason = error instanceof Error ? error.message : "Outlook calendar event creation failed.";
        downloadLocalMeetingItem(item, {
          showAlert: false,
        });
        await recordExecutionHistory({
          item,
          status: "fallback",
          path: "fallback",
          executionGroupId,
          fallbackExportKind: "ics",
          reason,
        });
        setExecutionNoticeForFallback({
          title: "Meeting action fell back to local export",
          reason,
          fallbackMessage: "Downloaded a local calendar file instead.",
        });
        return;
      }
    }

    downloadLocalMeetingItem(item, {
      showAlert: false,
    });
    await recordExecutionHistory({
      item,
      status: "fallback",
      path: "fallback",
      executionGroupId,
      fallbackExportKind: "ics",
      reason: graphAvailability.reason ?? "Outlook is not connected. Falling back to local export.",
    });
    setExecutionNoticeForFallback({
      title: "Meeting action fell back to local export",
      reason: graphAvailability.reason ?? "Outlook is not connected. Falling back to local export.",
      fallbackMessage: "Downloaded a local calendar file instead.",
    });
  }

  function clearTransientEditingState() {
    setTemplateActionMessage("");
    setOpenMenuRowId(null);
    setOpenEmailDraftRowId(null);
    setOpenDetailIndicatorId(null);
    setPinnedDetailIndicatorId(null);
    setOpenMeetingEditorRowId(null);
    setOpenDurationEditorRowId(null);
    setForcedOpenMeetingEditorRowIds([]);
    setMeetingValidationErrors({});
    setOpenBodyEditorRowId(null);
    setEditingOffsetRowId(null);
    setOffsetDrafts({});
    setFocusedTimeInputRowId(null);
    setTimeInputDrafts({});
    setEmailFieldVisibility({});
    setIsBuilderPreviewOpen(false);
    setExpandedPreviewReminderRowIds([]);
    setExpandedPreviewEmailRowIds([]);
    setExpandedPreviewMeetingRowIds([]);
    setOpenPreviewRowMenuId(null);
  }

  function buildTemplateSnapshot(template: SavedPlanTemplate): BuilderStateSnapshot {
    const templateMode = inferTemplateMode(template);
    const nextGuidedForm = createEmptyGuidedForm();
    if (templateMode === "template" && template.baseType === "press_release") {
      nextGuidedForm.releaseName = "";
      nextGuidedForm.releaseDate = "";
      nextGuidedForm.releaseTime =
        template.anchors.find((anchor) => normalizeAnchorKey(anchor.key) === normalizeAnchorKey("Dissemination Time"))?.value ??
        normalizedDefaultPressReleaseTime;
    } else if (templateMode === "template" && template.baseType === "earnings") {
      nextGuidedForm.quarter =
        (template.anchors.find((anchor) => normalizeAnchorKey(anchor.key) === normalizeAnchorKey("Quarter"))?.value as GuidedFormState["quarter"]) ?? "";
      const yearValue =
        template.anchors.find((anchor) => normalizeAnchorKey(anchor.key) === normalizeAnchorKey("Year / Fiscal Year"))?.value ?? "";
      nextGuidedForm.fiscalYear = yearValue.toLowerCase().startsWith("fiscal year ");
      nextGuidedForm.year = nextGuidedForm.fiscalYear ? yearValue.replace(/^Fiscal Year\s+/i, "") : yearValue;
      nextGuidedForm.earningsDate =
        template.anchors.find((anchor) => normalizeAnchorKey(anchor.key) === normalizeAnchorKey("Earnings Call Date"))?.value ?? "";
      nextGuidedForm.earningsTime =
        template.anchors.find((anchor) => normalizeAnchorKey(anchor.key) === normalizeAnchorKey("Earnings Call Time"))?.value ?? "";
    } else if (templateMode === "template" && template.baseType === "conference") {
      nextGuidedForm.conferenceName = "";
      nextGuidedForm.conferenceLocation =
        template.anchors.find((anchor) => normalizeAnchorKey(anchor.key) === normalizeAnchorKey("Conference Location"))?.value ?? "";
      nextGuidedForm.conferenceDate =
        template.anchors.find((anchor) => normalizeAnchorKey(anchor.key) === normalizeAnchorKey("Conference Start Date"))?.value ?? "";
      nextGuidedForm.conferenceEndDate =
        template.anchors.find((anchor) => normalizeAnchorKey(anchor.key) === normalizeAnchorKey("Conference End Date"))?.value ?? "";
    }
    return {
      builderMode: templateMode,
      selectedTemplateId: template.id,
      planType: template.baseType,
      templateName: template.name,
      eventName: templateMode === "template" ? template.name : "",
      anchorDate:
        templateMode === "template" && template.baseType === "press_release"
          ? nextGuidedForm.releaseDate || todayYYYYMMDD()
          : templateMode === "template" && template.baseType === "earnings"
            ? nextGuidedForm.earningsDate || todayYYYYMMDD()
            : templateMode === "template" && template.baseType === "conference"
              ? nextGuidedForm.conferenceDate || todayYYYYMMDD()
              : todayYYYYMMDD(),
      noEventDate: Boolean(template.noEventDate),
      weekendRule: template.weekendRule,
      rows: cloneTemplateRows(template.items),
      anchors:
        templateMode === "template"
          ? buildAnchorStateForType(
              template.baseType,
              template.anchors.map((anchor) => ({ id: crypto.randomUUID(), key: anchor.key, value: anchor.value }))
            )
          : template.anchors.map((anchor) => ({
              id: crypto.randomUUID(),
              key: anchor.key,
              value: anchor.value,
            })),
      guidedForm: nextGuidedForm,
    };
  }

  function buildCurrentTemplateSnapshot(): BuilderStateSnapshot {
    return {
      builderMode,
      selectedTemplateId,
      planType,
      templateName,
      eventName,
      anchorDate,
      noEventDate,
      weekendRule,
      rows: cloneTemplateRows(rows),
      anchors: cloneAnchors(anchors),
      guidedForm: { ...guidedForm },
    };
  }

  function applyBuilderSnapshot(snapshot: BuilderStateSnapshot) {
    setBuilderMode(snapshot.builderMode);
    setSelectedTemplateId(snapshot.selectedTemplateId);
    setPlanType(snapshot.planType);
    setTemplateName(snapshot.templateName);
    setEventName(snapshot.eventName);
    setAnchorDate(snapshot.anchorDate);
    setNoEventDate(snapshot.noEventDate);
    setWeekendRule(snapshot.weekendRule);
    setRows(cloneTemplateRows(snapshot.rows));
    setAnchors(cloneAnchors(snapshot.anchors));
    setAreAnchorsHidden(false);
    setGuidedForm({ ...snapshot.guidedForm });
    clearTransientEditingState();
  }

  function applyTemplateRecord(template: SavedPlanTemplate) {
    setTemplateActionMessage("");
    const snapshot = buildTemplateSnapshot(template);
    applyBuilderSnapshot(snapshot);
    setLastTemplateSnapshot(snapshot);
  }

  function applyTemplate(templateId: string) {
    const template = savedTemplates.find((entry) => entry.id === templateId);
    if (!template) return;
    applyTemplateRecord(template);
  }

  function hasDuplicateTemplateName(name: string, options?: { excludeTemplateId?: string }) {
    const normalizedName = name.trim().toLowerCase();
    return savedTemplates.some(
      (template) =>
        template.id !== options?.excludeTemplateId && template.name.trim().toLowerCase() === normalizedName
    );
  }

  function buildTemplateToSave(templateId: string, nameOverride?: string): SavedPlanTemplate {
    const protectedName =
      currentSelectedTemplate && isProtectedTemplate(currentSelectedTemplate) ? currentSelectedTemplate.name : null;
    const nextTemplateName = protectedName || nameOverride?.trim() || templateName.trim() || "Untitled Template";
    const nextTemplateMode =
      currentSelectedTemplate && isProtectedTemplate(currentSelectedTemplate)
        ? "template"
        : builderMode === "template"
          ? "template"
          : "custom";
    return {
      id: templateId,
      name: nextTemplateName,
      baseType: planType,
      templateMode: nextTemplateMode,
      noEventDate,
      weekendRule,
      anchors: resolvedAnchors
        .map((anchor) => ({ key: anchor.key.trim(), value: anchor.value }))
        .filter((anchor) => anchor.key),
      items: cloneTemplateRows(rows),
    };
  }

  function saveTemplateEntry(nextTemplate: SavedPlanTemplate) {
    const nextTemplateMode = inferTemplateMode(nextTemplate);
    hasLocalTemplateMutationRef.current = true;
    const existingIndex = savedTemplates.findIndex((template) => template.id === nextTemplate.id);
    const nextTemplates =
      existingIndex >= 0
        ? savedTemplates.map((template, index) => (index === existingIndex ? nextTemplate : template))
        : [...savedTemplates, nextTemplate];
    setSavedTemplates(nextTemplates);
    const snapshot: BuilderStateSnapshot = {
      builderMode: nextTemplateMode,
      selectedTemplateId: nextTemplate.id,
      planType,
      templateName: nextTemplate.name,
      eventName,
      anchorDate,
      noEventDate: Boolean(nextTemplate.noEventDate),
      weekendRule,
      rows: cloneTemplateRows(rows),
      anchors:
        nextTemplateMode === "template"
          ? cloneAnchors(
              buildAnchorStateForType(
                nextTemplate.baseType,
                nextTemplate.anchors.map((anchor) => ({ id: crypto.randomUUID(), key: anchor.key, value: anchor.value }))
              )
            )
          : cloneAnchors(
              nextTemplate.anchors.map((anchor) => ({
                id: crypto.randomUUID(),
                key: anchor.key,
                value: anchor.value,
              }))
            ),
      guidedForm: { ...guidedForm },
    };
    persistTemplateStateImmediately(nextTemplates, nextTemplate.id);
    applyBuilderSnapshot(snapshot);
    setLastTemplateSnapshot(snapshot);
  }

  function promptForNewTemplateName(defaultName: string) {
    const nextName = window.prompt("Save As", defaultName);
    if (!nextName) return null;
    const trimmedName = nextName.trim();
    if (!trimmedName) {
      setTemplateActionMessage("Template name is required.");
      return null;
    }
    if (hasDuplicateTemplateName(trimmedName)) {
      setTemplateActionMessage("That name is already taken. Please choose another name.");
      return null;
    }
    return trimmedName;
  }

  function onSelectSavedTemplate(templateId: string) {
    const template = savedTemplates.find((entry) => entry.id === templateId);
    if (!template) return;

    if (isTemplateCopyMode && typeof window !== "undefined") {
      const confirmed = window.confirm(`Are you sure you want to copy "${template.name}"?`);
      if (!confirmed) {
        setIsTemplateCopyMode(false);
        return;
      }
      const nextName = promptForNewTemplateName(`${template.name} Copy`);
      if (!nextName) {
        setIsTemplateCopyMode(false);
        return;
      }
      const duplicatedTemplate: SavedPlanTemplate = {
        ...template,
        id: makeId("template"),
        name: nextName,
        anchors: template.anchors.map((anchor) => ({ ...anchor })),
        items: cloneTemplateRows(template.items),
      };
      setTemplateActionMessage("");
      setSavedTemplates((current) => [...current, duplicatedTemplate]);
      setIsTemplateCopyMode(false);
      applyTemplate(duplicatedTemplate.id);
      return;
    }

    setIsTemplateManageMode(false);
    setIsTemplateCopyMode(false);
    applyTemplate(templateId);
  }

  function saveCurrentTemplate() {
    setTemplateActionMessage("");
    if (typeof window === "undefined") return;

    const promptDefaultName =
      templateName.trim() || selectedEditableTemplate?.name || getSeedTemplateName(planType) || "Untitled Template";

    const promptForNewTemplate = () => {
      const nextName = promptForNewTemplateName(promptDefaultName);
      if (!nextName) return;
      const nextTemplate = buildTemplateToSave(makeId("template"), nextName);
      saveTemplateEntry(nextTemplate);
    };

    const chooseSave = window.confirm(
      "Choose how you want to save this reminder template.\n\nChoose OK for Save.\nChoose Cancel for Save As."
    );

    if (chooseSave) {
      if (!selectedEditableTemplate) {
        promptForNewTemplate();
        return;
      }

      const nextName = templateName.trim() || selectedEditableTemplate.name;
      if (hasDuplicateTemplateName(nextName, { excludeTemplateId: selectedEditableTemplate.id })) {
        setTemplateActionMessage("That name is already taken. Please choose another name.");
        return;
      }

      const confirmedOverwrite = window.confirm("Are you sure you want to overwrite the existing template?");
      if (!confirmedOverwrite) return;

      const nextTemplate = buildTemplateToSave(selectedEditableTemplate.id, nextName);
      saveTemplateEntry(nextTemplate);
      return;
    }

    promptForNewTemplate();
  }

  function startNewPlan() {
    if (builderMode !== "new") {
      setLastTemplateSnapshot(buildCurrentTemplateSnapshot());
    }
    setBuilderMode("new");
    setSelectedTemplateId(null);
    setTemplateName("");
    setEventName("");
    setAnchorDate(todayYYYYMMDD());
    setNoEventDate(false);
    setWeekendRule("prior_business_day");
    setRows([createEmptyBuilderRow(normalizedDefaultReminderTime)]);
    setAnchors(createGenericPresetAnchors());
    setAreAnchorsHidden(false);
    setGuidedForm(createEmptyGuidedForm());
    clearTransientEditingState();
  }

  function cancelEditing() {
    if (builderMode === "new" && lastTemplateSnapshot) {
      applyBuilderSnapshot(lastTemplateSnapshot);
      return;
    }
    if (selectedTemplateId) {
      applyTemplate(selectedTemplateId);
      return;
    }
    clearTransientEditingState();
  }

  async function exportPreviewReminderItem(itemId: string) {
    const item = previewPlan.items.find((entry) => entry.id === itemId);
    if (!item) return;
    if (warnIfMissingReminderTimes({ usePopup: true, itemIds: [itemId] })) return;
    if (!confirmExport()) return;
    const executionGroupId = crypto.randomUUID();

    const graphAvailability = await getGraphExecutionAvailability(["Calendars.ReadWrite"]);
    if (graphAvailability.canUseGraph) {
      try {
        const result = await executePreviewCalendarViaGraph(item);
        await recordExecutionHistory({
          item,
          status: "success",
          path: "graph",
          executionGroupId,
          result,
        });
        setExecutionNoticeForResult(result);
        return;
      } catch (error) {
        const reason = error instanceof Error ? error.message : "Outlook calendar event creation failed.";
        downloadLocalReminderItem(item, {
          showAlert: false,
        });
        await recordExecutionHistory({
          item,
          status: "fallback",
          path: "fallback",
          executionGroupId,
          fallbackExportKind: "ics",
          reason,
        });
        setExecutionNoticeForFallback({
          title: "Reminder action fell back to local export",
          reason,
          fallbackMessage: "Downloaded a local calendar file instead.",
        });
        return;
      }
    }

    downloadLocalReminderItem(item, {
      showAlert: false,
    });
    await recordExecutionHistory({
      item,
      status: "fallback",
      path: "fallback",
      executionGroupId,
      fallbackExportKind: "ics",
      reason: graphAvailability.reason ?? "Outlook is not connected. Falling back to local export.",
    });
    setExecutionNoticeForFallback({
      title: "Reminder action fell back to local export",
      reason: graphAvailability.reason ?? "Outlook is not connected. Falling back to local export.",
      fallbackMessage: "Downloaded a local calendar file instead.",
    });
  }

  async function exportCurrentPlan(options?: { skipValidation?: boolean; skipConfirm?: boolean }) {
    if (!options?.skipValidation && validateExportBeforeRun()) return;
    if (!options?.skipConfirm && !confirmExport()) return;

    const { emailItems, calendarItems } = partitionPlanItemsByKind(previewPlan.items);
    const graphAvailability = await getGraphExecutionAvailability([
      "Mail.ReadWrite",
      "Mail.Send",
      "Calendars.ReadWrite",
    ]);

    const failedEmailItems: Plan["items"] = [];
    const failedCalendarItems: Plan["items"] = [];
    const graphResults: OutlookExecutionResult[] = [];
    const executionGroupId = crypto.randomUUID();
    let draftCount = 0;
    let scheduledCount = 0;
    let sentCount = 0;
    let reminderCount = 0;
    let meetingCount = 0;

    if (graphAvailability.canUseGraph) {
      for (const item of previewPlan.items) {
        const rowKind = classifyPlanRow(item);
        try {
          if (rowKind === "email") {
            const result = await executePreviewEmailViaGraph(item);
            graphResults.push(result);
            await recordExecutionHistory({
              item,
              status: "success",
              path: "graph",
              executionGroupId,
              result,
            });
            if (result.action === "email_sent") {
              sentCount += 1;
            } else if (result.action === "email_scheduled") {
              scheduledCount += 1;
            } else {
              draftCount += 1;
            }
          } else {
            const result = await executePreviewCalendarViaGraph(item);
            graphResults.push(result);
            await recordExecutionHistory({
              item,
              status: "success",
              path: "graph",
              executionGroupId,
              result,
            });
            if (rowKind === "meeting") {
              meetingCount += 1;
            } else {
              reminderCount += 1;
            }
          }
        } catch {
          if (rowKind === "email") {
            failedEmailItems.push(item);
          } else {
            failedCalendarItems.push(item);
          }
        }
      }
    } else {
      failedEmailItems.push(...emailItems);
      failedCalendarItems.push(...calendarItems);
    }

    for (const item of failedEmailItems) {
      downloadLocalEmailItem(item, { showAlert: false });
      await recordExecutionHistory({
        item,
        status: "fallback",
        path: "fallback",
        executionGroupId,
        fallbackExportKind: "eml",
        reason: graphAvailability.canUseGraph ? "Outlook email action fell back to local export." : graphAvailability.reason,
      });
    }
    if (failedCalendarItems.length > 0) {
      downloadLocalPlanCalendarItems(failedCalendarItems, { showAlert: false });
      for (const item of failedCalendarItems) {
        await recordExecutionHistory({
          item,
          status: "fallback",
          path: "fallback",
          executionGroupId,
          fallbackExportKind: "ics",
          reason: graphAvailability.canUseGraph
            ? "Outlook calendar action fell back to local export."
            : graphAvailability.reason,
        });
      }
    }

    setExecutionNoticeForExportSummary({
      graphUnavailableReason: graphAvailability.canUseGraph ? undefined : graphAvailability.reason,
      graphResults,
      failedEmailItems,
      failedCalendarItems,
      draftCount,
      scheduledCount,
      sentCount,
      reminderCount,
      meetingCount,
    });
  }

  function getMissingPreviewRequirements() {
    const missing: string[] = [];

    if (effectiveTemplateMode === "template") {
      if (planType === "press_release") {
        if (!guidedForm.releaseName.trim()) missing.push("Press Release Name");
        if (!guidedForm.releaseDate) missing.push("Dissemination Date");
        if (!guidedForm.releaseTime) missing.push("Dissemination Time");
      } else if (planType === "earnings") {
        if (!guidedForm.quarter) missing.push("Quarter");
        if (!guidedForm.year.trim()) missing.push("Year");
        if (!guidedForm.earningsDate) missing.push("Earnings Call Date");
        if (!guidedForm.earningsTime) missing.push("Earnings Call Time");
      } else if (planType === "conference") {
        if (!guidedForm.conferenceName.trim()) missing.push("Conference Name");
        if (!guidedForm.conferenceLocation.trim()) missing.push("Conference Location");
        if (!guidedForm.conferenceDate) missing.push("Conference Start Date");
        if (!guidedForm.conferenceEndDate) missing.push("Conference End Date");
      }
    } else {
      if (!templateName.trim()) missing.push("Template Name");
      if (!eventName.trim()) missing.push("Event Name");
      if (!noEventDate && !anchorDate) missing.push("Event Date");
    }

    return missing;
  }

  function validatePreviewBeforeOpen() {
    const missingRequirements = getMissingPreviewRequirements();
    if (missingRequirements.length > 0) {
      window.alert(`Preview is missing:\n\n${missingRequirements.join("\n")}`);
      return true;
    }
    if (validateMeetingRowsForExport(undefined, { usePopup: true })) return true;
    if (warnIfMissingReminderTimes({ usePopup: true })) return true;
    return false;
  }

  function validateExportBeforeRun() {
    const missingRequirements = getMissingPreviewRequirements();
    if (missingRequirements.length > 0) {
      window.alert(`Export is missing:\n\n${missingRequirements.join("\n")}`);
      return true;
    }
    if (validateMeetingRowsForExport(undefined, { usePopup: true })) return true;
    if (warnIfMissingReminderTimes({ usePopup: true })) return true;
    return false;
  }

  function deleteSavedTemplate(templateId: string) {
    const template = savedTemplates.find((entry) => entry.id === templateId);
    if (!template || isProtectedTemplate(template)) return;
    if (typeof window !== "undefined" && !window.confirm(`Delete "${template.name}"? This cannot be undone.`)) {
      return;
    }
    const nextTemplates = savedTemplates.filter((entry) => entry.id !== templateId);
    const fallbackTemplate =
      nextTemplates.find((entry) => entry.baseType === template.baseType && isProtectedTemplate(entry)) ??
      nextTemplates.find((entry) => isProtectedTemplate(entry)) ??
      nextTemplates[0] ??
      null;
    const nextSelectedTemplateId = selectedTemplateId === templateId ? fallbackTemplate?.id ?? null : selectedTemplateId;
    hasLocalTemplateMutationRef.current = true;
    setTemplateActionMessage("");
    setSavedTemplates(nextTemplates);
    setSelectedTemplateId(nextSelectedTemplateId);
    persistTemplateStateImmediately(nextTemplates, nextSelectedTemplateId);
    if (selectedTemplateId === templateId) {
      if (fallbackTemplate) {
        applyTemplateRecord(fallbackTemplate);
      } else {
        startNewPlan();
      }
    }
  }

  function addReminderRow() {
    setRows((current) => [...current, createEmptyBuilderRow(normalizedDefaultReminderTime)]);
  }

  function addEmailRow() {
    setRows((current) => [
      ...current,
      {
        ...createEmptyBuilderRow(),
        reminderTime: normalizedDefaultReminderTime,
        rowType: "email",
        emailDraft: { to: [], cc: [], bcc: [], subject: "", body: "" },
      },
    ]);
  }

  function addMeetingRow() {
    setRows((current) => [
      ...current,
      {
        ...createEmptyBuilderRow(),
        reminderTime: normalizedDefaultReminderTime,
        rowType: "calendar_event",
        meetingDraft: {
          attendees: [],
          location: "",
          durationMinutes: 30,
        },
      },
    ]);
  }

  function collapseEmptyEmailFields(rowId: string, draft?: BuilderEmailDraft | null) {
    const normalizedDraft = normalizeEmailDraft(draft);
    setEmailFieldVisibility((prev) => ({
      ...prev,
      [rowId]: {
        cc: normalizedDraft.cc.length > 0,
        bcc: normalizedDraft.bcc.length > 0,
      },
    }));
  }

  const accountConnectionStatus = hasMounted
    ? outlookConnection?.status ?? appSettings.outlookConnectionStatus
    : "not_connected";
  const connectedMailboxEmail = hasMounted
    ? getConnectedOutlookMailboxEmail(outlookConnection?.identity) || appSettings.outlookAccountEmail
    : "";
  const accountStatusLabel =
    accountConnectionStatus === "connected"
      ? "Connected"
      : accountConnectionStatus === "reconnect_required"
        ? "Reconnect required"
        : "Not connected";
  const accountStatusClass =
    accountConnectionStatus === "connected"
      ? "text-green-600"
      : accountConnectionStatus === "reconnect_required"
        ? "text-amber-600"
        : "text-red-600";
  const accountPrimaryText =
    accountConnectionStatus === "connected"
      ? connectedMailboxEmail || "Connected email account"
      : accountConnectionStatus === "reconnect_required"
        ? connectedMailboxEmail || "Reconnect required"
        : "No connected email account";
  const accountButtonLabel =
    accountConnectionStatus === "connected" ? "Manage in Settings" : "Reconnect in Settings";

  return (
    <div className="space-y-8 text-gray-900">
      <section className="space-y-2">
        <div className="flex flex-col gap-4 md:flex-row md:items-start md:justify-between">
          <div className="space-y-2">
            <h1 className="text-3xl font-bold text-gray-900">Plans</h1>
            <p className="text-sm text-gray-600">
              Build a plan from an event and export it as a calendar file for review.
            </p>
          </div>
          <div className="rounded-xl border bg-white px-4 py-3 text-sm shadow-sm md:min-w-[260px] md:max-w-[280px]">
            <div className="text-xs font-semibold uppercase tracking-wide text-gray-500">Connected Email Account</div>
            <div className={`mt-1 font-medium ${accountStatusClass}`}>{accountPrimaryText}</div>
            <div className={`mt-1 text-xs font-semibold uppercase tracking-wide ${accountStatusClass}`}>
              {accountStatusLabel}
            </div>
            <Link
              href="/settings"
              className="mt-2 inline-flex rounded-lg border border-gray-200 px-3 py-1.5 text-xs text-gray-700 hover:bg-gray-50"
            >
              {accountButtonLabel}
            </Link>
          </div>
        </div>
      </section>

      {executionNotice && !isBuilderPreviewOpen ? (
        <OutlookExecutionNoticeCard notice={executionNotice} onDismiss={() => setExecutionNotice(null)} />
      ) : null}

      <section className="rounded-2xl border bg-white shadow-sm">
        <div className="space-y-5 p-6">
          <div className="flex items-center justify-between gap-3">
            <h2 className="text-lg font-semibold text-gray-900">Templates</h2>
            <div className="flex items-center gap-2">
              {isTemplateManageMode ? (
                <button
                  type="button"
                  onClick={() => {
                    if (typeof window !== "undefined" && !isTemplateCopyMode) {
                      window.alert("Click the current template you’d like to make a copy of.");
                    }
                    setTemplateActionMessage("");
                    setIsTemplateCopyMode((prev) => !prev);
                  }}
                  className={`rounded-lg border px-3 py-2 text-sm hover:bg-gray-50 ${
                    isTemplateCopyMode ? "border-blue-300 bg-blue-50 text-blue-700" : ""
                  }`}
                >
                  Copy
                </button>
              ) : null}
              <button
                  type="button"
                  onClick={() => {
                    setIsTemplateManageMode((prev) => !prev);
                    setIsTemplateCopyMode(false);
                    setTemplateActionMessage("");
                  }}
                  className="rounded-lg border px-3 py-2 text-sm hover:bg-gray-50"
                >
                  {isTemplateManageMode ? "Done" : "Edit"}
                </button>
              </div>
          </div>
          {templateActionMessage ? <p className="text-xs text-gray-500">{templateActionMessage}</p> : null}

          <div className="grid justify-center gap-x-3 gap-y-5 md:grid-cols-[repeat(3,260px)]">
            {savedTemplates.map((template) => (
              <div key={template.id} className="relative w-full max-w-[260px]">
                {isTemplateManageMode && !isProtectedTemplate(template) ? (
                  <button
                    type="button"
                    onClick={() => deleteSavedTemplate(template.id)}
                    className="absolute right-2 top-2 z-10 rounded-md border bg-white px-2 py-1 text-xs text-red-700 hover:bg-red-50"
                  >
                    Delete
                  </button>
                ) : null}
                <button
                  type="button"
                  onClick={() => onSelectSavedTemplate(template.id)}
                  className={`flex min-h-[72px] w-full items-center justify-center rounded-xl border px-4 py-4 text-center hover:bg-gray-50 ${
                    selectedTemplateId === template.id ? "border-blue-300 bg-blue-50" : ""
                  }`}
                >
                  <div className="font-semibold text-gray-900">{template.name}</div>
                </button>
              </div>
            ))}
          </div>

          <div className="flex justify-end">
            <button
              type="button"
              onClick={startNewPlan}
              className="rounded-lg border px-4 py-2 text-sm hover:bg-gray-50"
            >
              New Plan
            </button>
          </div>
        </div>
      </section>

      <section className="rounded-2xl border bg-white shadow-sm">
        <div className="space-y-5 p-6">
              <div className="flex flex-col gap-4 md:flex-row md:items-start md:justify-between">
                <div>
                  <div className="text-sm font-semibold uppercase tracking-wide text-gray-500">Plan Builder</div>
                  <h2 className="text-2xl font-semibold text-gray-900">{eventName || templateName || getSeedTemplateName(planType)}</h2>
                </div>
                <div className="w-full md:max-w-sm">
                  <label className="mb-1 block text-sm font-medium text-gray-700">Template Name</label>
                  <input
                    className="w-full rounded-lg border px-3 py-2"
                    value={templateName}
                    onChange={(e) => setTemplateName(e.target.value)}
                    readOnly={Boolean(currentSelectedTemplate && isProtectedTemplate(currentSelectedTemplate))}
                  />
                  {currentSelectedTemplate && isProtectedTemplate(currentSelectedTemplate) ? (
                    <p className="mt-1 text-xs text-gray-500">Protected templates keep their original template name.</p>
                  ) : null}
                </div>
              </div>

              <div className="space-y-4 border-b pb-6">
                {effectiveTemplateMode === "template" && planType === "press_release" ? (
                  <div className="grid grid-cols-1 gap-4 md:grid-cols-4">
                    <div className="md:col-span-2">
                      <label className="mb-1 block text-sm font-medium text-gray-700">
                        Press Release Name (Keep Name Short For System)
                      </label>
                      <input
                        className="w-full rounded-lg border px-3 py-2"
                        value={guidedForm.releaseName}
                        onChange={(e) =>
                          setGuidedForm((current) => ({
                            ...current,
                            releaseName: e.target.value,
                          }))
                        }
                      />
                    </div>
                    <div>
                      <label className="mb-1 block text-sm font-medium text-gray-700">Dissemination Date</label>
                      <input
                        type="date"
                        className="w-full rounded-lg border px-3 py-2"
                        value={guidedForm.releaseDate}
                        onChange={(e) => {
                          setGuidedForm((current) => ({
                            ...current,
                            releaseDate: e.target.value,
                          }));
                          setAnchorDate(e.target.value);
                        }}
                      />
                    </div>
                    <div>
                      <label className="mb-1 block text-sm font-medium text-gray-700">Dissemination Time</label>
                      <input
                        type="time"
                        className="w-full rounded-lg border px-3 py-2"
                        value={guidedForm.releaseTime}
                        onChange={(e) =>
                          setGuidedForm((current) => ({
                            ...current,
                            releaseTime: e.target.value,
                          }))
                        }
                      />
                    </div>
                  </div>
                ) : effectiveTemplateMode === "template" && planType === "earnings" ? (
                  <div className="grid grid-cols-1 gap-4 md:grid-cols-4">
                    <div>
                      <label className="mb-1 block text-sm font-medium text-gray-700">Quarter</label>
                      <select
                        className="w-full rounded-lg border px-3 py-2"
                        value={guidedForm.quarter}
                        onChange={(e) =>
                          setGuidedForm((current) => ({
                            ...current,
                            quarter: e.target.value as GuidedFormState["quarter"],
                          }))
                        }
                      >
                        <option value="">Select quarter</option>
                        <option value="Q1">Q1</option>
                        <option value="Q2">Q2</option>
                        <option value="Q3">Q3</option>
                        <option value="Q4">Q4</option>
                      </select>
                    </div>
                    <div>
                      <label className="mb-1 block text-sm font-medium text-gray-700">Year</label>
                      <input
                        className="w-full rounded-lg border px-3 py-2"
                        value={guidedForm.year}
                        onChange={(e) =>
                          setGuidedForm((current) => ({
                            ...current,
                            year: e.target.value,
                          }))
                        }
                      />
                      <label className="mt-2 flex items-center gap-2 text-sm text-gray-700">
                        <input
                          type="checkbox"
                          checked={guidedForm.fiscalYear}
                          onChange={(e) =>
                            setGuidedForm((current) => ({
                              ...current,
                              fiscalYear: e.target.checked,
                            }))
                          }
                        />
                        Use Fiscal Year
                      </label>
                    </div>
                    <div>
                      <label className="mb-1 block text-sm font-medium text-gray-700">Earnings Call Date</label>
                      <input
                        type="date"
                        className="w-full rounded-lg border px-3 py-2"
                        value={guidedForm.earningsDate}
                        onChange={(e) => {
                          setGuidedForm((current) => ({
                            ...current,
                            earningsDate: e.target.value,
                          }));
                          setAnchorDate(e.target.value);
                        }}
                      />
                    </div>
                    <div>
                      <label className="mb-1 block text-sm font-medium text-gray-700">Earnings Call Time</label>
                      <input
                        type="time"
                        className="w-full rounded-lg border px-3 py-2"
                        value={guidedForm.earningsTime}
                        onChange={(e) =>
                          setGuidedForm((current) => ({
                            ...current,
                            earningsTime: e.target.value,
                          }))
                        }
                      />
                    </div>
                  </div>
                ) : effectiveTemplateMode === "template" && planType === "conference" ? (
                  <div className="grid grid-cols-1 gap-4 md:grid-cols-4">
                    <div className="md:col-span-2">
                      <label className="mb-1 block text-sm font-medium text-gray-700">Conference Name</label>
                      <input
                        className="w-full rounded-lg border px-3 py-2"
                        value={guidedForm.conferenceName}
                        onChange={(e) =>
                          setGuidedForm((current) => ({
                            ...current,
                            conferenceName: e.target.value,
                          }))
                        }
                      />
                    </div>
                    <div className="md:col-span-1">
                      <label className="mb-1 block text-sm font-medium text-gray-700">Conference Location</label>
                      <input
                        className="w-full rounded-lg border px-3 py-2"
                        value={guidedForm.conferenceLocation}
                        onChange={(e) =>
                          setGuidedForm((current) => ({
                            ...current,
                            conferenceLocation: e.target.value,
                          }))
                        }
                      />
                    </div>
                    <div className="space-y-3 md:col-span-1">
                      <div>
                        <label className="mb-1 block text-sm font-medium text-gray-700">Conference Start Date</label>
                        <input
                          type="date"
                          className="w-full rounded-lg border px-3 py-2"
                          value={guidedForm.conferenceDate}
                          onChange={(e) => {
                            setGuidedForm((current) => ({
                              ...current,
                              conferenceDate: e.target.value,
                            }));
                            setAnchorDate(e.target.value);
                          }}
                        />
                      </div>
                      <div>
                        <label className="mb-1 block text-sm font-medium text-gray-700">Conference End Date</label>
                        <input
                          type="date"
                          className="w-full rounded-lg border px-3 py-2"
                          value={guidedForm.conferenceEndDate}
                          onChange={(e) =>
                            setGuidedForm((current) => ({
                              ...current,
                              conferenceEndDate: e.target.value,
                            }))
                          }
                        />
                      </div>
                    </div>
                  </div>
                ) : (
                  <div className="grid grid-cols-1 gap-4 md:grid-cols-4">
                    <div className="md:col-span-2">
                      <label className="mb-1 block text-sm font-medium text-gray-700">Event Name</label>
                      <input
                        className="w-full rounded-lg border px-3 py-2"
                        value={eventName}
                        onChange={(e) => setEventName(e.target.value)}
                      />
                    </div>
                    <div className="md:col-span-2">
                      <div className="flex flex-col gap-3 md:flex-row md:items-end md:justify-between">
                        <div className="w-full md:max-w-xs">
                          {noEventDate ? (
                            <div className="invisible" aria-hidden="true">
                              <label className="mb-1 block text-sm font-medium text-gray-700">Event Date</label>
                              <input
                                type="date"
                                className="w-full rounded-lg border px-3 py-2"
                                value=""
                                readOnly
                              />
                            </div>
                          ) : (
                            <>
                              <label className="mb-1 block text-sm font-medium text-gray-700">Event Date</label>
                              <input
                                type="date"
                                className="w-full rounded-lg border px-3 py-2"
                                value={anchorDate}
                                onChange={(e) => setAnchorDate(e.target.value)}
                              />
                            </>
                          )}
                        </div>
                        <label className="flex items-center gap-2 text-sm font-medium text-gray-700">
                          <input
                            type="checkbox"
                            checked={noEventDate}
                            onChange={(e) => setNoEventDate(e.target.checked)}
                          />
                          No Event Date
                        </label>
                      </div>
                    </div>
                  </div>
                )}
              </div>

              <div className="space-y-2 border-b pb-6" ref={kebabMenuRef}>
                <div className="space-y-2">
                  <div className="grid grid-cols-1 gap-3 pr-2 text-[10px] font-semibold uppercase tracking-wide text-gray-600 md:grid-cols-[minmax(0,1.8fr)_120px_140px_88px]">
                    <div className="flex min-h-[24px] items-center justify-center text-center" />
                    <div className="flex min-h-[24px] items-center justify-center text-center">Days from event/today</div>
                    <div className="flex min-h-[24px] items-center justify-center text-center">Time</div>
                    <div className="flex min-h-[24px] items-center justify-center text-center">Actions</div>
                  </div>

                  {rows.map((row, index) => {
                    const rowMeta = getBuilderRowTypeMeta(row);
                    const emailUsesTimingFields = row.rowType === "email" && appSettings.emailHandlingMode === "schedule";
                    const hasAdditionalDetails = hasMeaningfulBody(row.body);
                    const additionalDetailsText = hasAdditionalDetails
                      ? row.rowType === "calendar_event"
                        ? "This meeting includes a body"
                        : row.rowType === "email"
                          ? "This email includes a body"
                          : "This reminder includes a body"
                      : "";
                    const isDetailIndicatorOpen = openDetailIndicatorId === row.id;
                    const isDetailIndicatorPinned = pinnedDetailIndicatorId === row.id;
                    const meetingErrors = meetingValidationErrors[row.id];
                    return (
                      <div key={row.id} className="space-y-3 py-2 pr-2">
                        <div className="grid grid-cols-1 gap-3 md:grid-cols-[minmax(0,1.8fr)_120px_140px_88px] md:items-start">
                          <div className="relative flex min-h-[72px] items-center" data-detail-indicator-root="true">
                            {hasAdditionalDetails ? (
                              <>
                                <button
                                  type="button"
                                  className="absolute -left-5 top-1/2 flex h-4 w-4 -translate-y-1/2 items-center justify-center rounded-full border border-gray-400 bg-white text-[10px] font-semibold text-gray-700 hover:border-gray-500 hover:text-gray-800"
                                  aria-label={additionalDetailsText}
                                  onMouseEnter={() => setOpenDetailIndicatorId(row.id)}
                                  onMouseLeave={() => {
                                    if (!isDetailIndicatorPinned) {
                                      setOpenDetailIndicatorId((current) => (current === row.id ? null : current));
                                    }
                                  }}
                                  onFocus={() => setOpenDetailIndicatorId(row.id)}
                                  onBlur={() => {
                                    if (!isDetailIndicatorPinned) {
                                      setOpenDetailIndicatorId((current) => (current === row.id ? null : current));
                                    }
                                  }}
                                  onClick={() => {
                                    if (isDetailIndicatorPinned) {
                                      setPinnedDetailIndicatorId(null);
                                      setOpenDetailIndicatorId(null);
                                    } else {
                                      setPinnedDetailIndicatorId(row.id);
                                      setOpenDetailIndicatorId(row.id);
                                    }
                                  }}
                                >
                                  i
                                </button>
                                {isDetailIndicatorOpen ? (
                                  <div
                                    className="absolute left-0 top-1/2 z-20 -translate-x-[calc(100%+0.5rem)] -translate-y-1/2 whitespace-nowrap rounded-lg border bg-white px-2 py-1 text-[11px] text-gray-700 shadow-sm"
                                    onMouseEnter={() => setOpenDetailIndicatorId(row.id)}
                                    onMouseLeave={() => {
                                      if (!isDetailIndicatorPinned) {
                                        setOpenDetailIndicatorId((current) => (current === row.id ? null : current));
                                      }
                                    }}
                                  >
                                    {additionalDetailsText}
                                  </div>
                                ) : null}
                              </>
                            ) : null}
                            <div className={`pointer-events-none absolute -top-4 left-1/2 -translate-x-1/2 text-xs font-medium ${rowMeta.className}`}>
                              {rowMeta.label}
                            </div>
                            <textarea
                              rows={2}
                              className="min-h-[72px] w-full rounded-lg border px-3 pb-3 pt-6 text-[13px] leading-5 [overflow-wrap:anywhere]"
                              style={{ fieldSizing: "content" }}
                              value={row.title}
                              placeholder={
                                row.rowType === "email"
                                  ? "Email row name"
                                  : row.rowType === "calendar_event"
                                    ? "Calendar event title"
                                    : rowMeta.label === "Meeting"
                                      ? "Meeting title"
                                      : "Reminder name"
                              }
                              onChange={(e) => updateRow(row.id, (current) => ({ ...current, title: e.target.value }))}
                            />
                          </div>

                          {!emailUsesTimingFields && row.rowType === "email" ? (
                            <div className="md:col-span-2">
                              <div className="flex min-h-[72px] w-full items-center justify-center rounded-lg border bg-gray-50 px-4 py-3 text-center text-[13px] font-medium text-gray-600">
                                {getBuilderEmailModeMessage(appSettings.emailHandlingMode)}
                              </div>
                            </div>
                          ) : (
                            <>
                              <div className="flex flex-col items-center">
                                <div className="flex min-h-[72px] w-full items-center rounded-lg border px-3 py-2">
                                  {editingOffsetRowId === row.id ? (
                                    <input
                                      type="text"
                                      inputMode="numeric"
                                      className="w-full border-0 bg-transparent p-0 text-center text-[13px] leading-5 focus:outline-none focus:ring-0"
                                      value={offsetDrafts[row.id] ?? (row.offsetDays == null ? "" : String(row.offsetDays))}
                                      onChange={(e) =>
                                        setOffsetDrafts((current) => ({
                                          ...current,
                                          [row.id]: e.target.value,
                                        }))
                                      }
                                      onBlur={() => {
                                        updateRow(row.id, (current) => ({
                                          ...current,
                                          offsetDays: Number.isNaN(Number(offsetDrafts[row.id])) ? current.offsetDays ?? 0 : Number(offsetDrafts[row.id]),
                                        }));
                                        setEditingOffsetRowId((current) => (current === row.id ? null : current));
                                      }}
                                      autoFocus
                                    />
                                  ) : (
                                    <button
                                      type="button"
                                      className="flex min-h-[56px] w-full items-center justify-center bg-transparent p-0 text-center text-[13px] leading-5 text-gray-900"
                                      onClick={() => {
                                        setEditingOffsetRowId(row.id);
                                        setOffsetDrafts((current) => ({
                                          ...current,
                                          [row.id]: row.offsetDays == null ? "" : String(row.offsetDays),
                                        }));
                                      }}
                                    >
                                      {formatOffsetLabel(row.offsetDays, {
                                        relativeToToday: noEventDate,
                                        dateBasis: row.dateBasis,
                                      })}
                                    </button>
                                  )}
                                </div>
                              </div>

                              <div className="flex flex-col items-center justify-center">
                                {meetingErrors?.time ? (
                                  <div className="mb-1 text-xs font-medium text-red-500">Time *</div>
                                ) : null}
                                {(() => {
                              const shouldWrapTimeField =
                                Boolean(row.reminderTime?.includes("[")) || Boolean(row.reminderTime?.includes("\n"));
                              const isAnchorTimeValue = isReminderTimeAnchorValue(row.reminderTime ?? "");
                              const literalTimeEditorValue =
                                focusedTimeInputRowId === row.id
                                  ? (timeInputDrafts[row.id] ?? buildReminderTimeMaskedValue(row.reminderTime ?? ""))
                                  : getReminderTimeDisplayValue(row.reminderTime ?? "");

                              if (row.meetingDraft?.isAllDay || row.durationDraft?.isAllDay) {
                                return (
                                  <div className="flex min-h-[72px] w-full items-center justify-center rounded-lg border bg-gray-50 px-3 py-3 text-center text-[13px] font-medium text-gray-600">
                                    All day
                                  </div>
                                );
                              }

                              if (!shouldWrapTimeField && !isAnchorTimeValue) {
                                return (
                                  <input
                                    ref={(node) => {
                                      builderTimeInputRefs.current[row.id] = node;
                                    }}
                                    type="text"
                                    inputMode="text"
                                    className="h-[72px] w-full rounded-lg border px-3 text-center text-[13px] leading-5"
                                    value={literalTimeEditorValue}
                                    placeholder={REMINDER_TIME_INPUT_MASK}
                                    onFocus={(e) => {
                                      const input = e.currentTarget;
                                      setFocusedTimeInputRowId(row.id);
                                      setTimeInputDrafts((current) => ({
                                        ...current,
                                        [row.id]: buildReminderTimeMaskedValue(row.reminderTime ?? ""),
                                      }));
                                      requestAnimationFrame(() => {
                                        input.setSelectionRange(0, 0);
                                      });
                                    }}
                                    onChange={(e) => {
                                      const rawValue = e.target.value;
                                      const selectionEnd = e.target.selectionEnd ?? rawValue.length;
                                      const meaningfulCount = countReminderTimeMeaningfulChars(rawValue, selectionEnd);
                                      const nextMaskedValue = maskReminderTimeDraftInput(rawValue);
                                      const nextCursor = findReminderTimeCursorFromMeaningfulCount(
                                        nextMaskedValue,
                                        meaningfulCount
                                      );

                                      setTimeInputDrafts((current) => ({
                                        ...current,
                                        [row.id]: nextMaskedValue,
                                      }));

                                      requestAnimationFrame(() => {
                                        const input = builderTimeInputRefs.current[row.id];
                                        if (
                                          input &&
                                          input instanceof HTMLInputElement &&
                                          document.activeElement === input
                                        ) {
                                          input.setSelectionRange(nextCursor, nextCursor);
                                        }
                                      });
                                    }}
                                    onBlur={(e) => {
                                      const nextValue = normalizeReminderTimeInput(e.target.value);
                                      updateRow(row.id, (current) => ({
                                        ...current,
                                        reminderTime: nextValue,
                                      }));
                                      setFocusedTimeInputRowId((current) => (current === row.id ? null : current));
                                      setTimeInputDrafts((current) => clearReminderTimeDraft(current, row.id));
                                    }}
                                  />
                                );
                              }

                              return (
                                <textarea
                                  ref={(node) => {
                                    builderTimeInputRefs.current[row.id] = node;
                                  }}
                                  rows={shouldWrapTimeField ? 2 : 1}
                                  className={`h-[72px] w-full resize-none rounded-lg border px-3 text-center text-[13px] leading-5 [overflow-wrap:anywhere] ${
                                    shouldWrapTimeField ? "py-[14px]" : "py-[24px]"
                                  }`}
                                  value={row.reminderTime ?? ""}
                                  placeholder={REMINDER_TIME_INPUT_MASK}
                                  onChange={(e) =>
                                    updateRow(row.id, (current) => ({
                                      ...current,
                                      reminderTime: maskReminderTimeDraftInput(e.target.value),
                                    }))
                                  }
                                  onBlur={(e) =>
                                    updateRow(row.id, (current) => ({
                                      ...current,
                                      reminderTime: normalizeReminderTimeInput(e.target.value),
                                    }))
                                  }
                                />
                              );
                                })()}
                              </div>
                            </>
                          )}

                          <div className="relative flex min-h-[72px] items-start justify-center pt-5 md:col-start-4 md:row-start-1 md:text-center">
                            <button
                              type="button"
                              onClick={() => setOpenMenuRowId((current) => (current === row.id ? null : row.id))}
                              title="Actions"
                              aria-label={`Actions for row ${index + 1}`}
                              className="flex h-8 w-10 items-center justify-center rounded-lg border border-gray-200 bg-white text-gray-500 hover:border-gray-300 hover:bg-gray-100 hover:text-gray-700"
                            >
                              <span className="text-base leading-none">•••</span>
                            </button>
                            {openMenuRowId === row.id ? (
                              <div className="absolute right-0 top-[calc(100%+0.5rem)] z-20 w-48 rounded-xl border bg-white p-2 text-left shadow-lg">
                                {row.rowType === "email" ? (
                                  <>
                                    <button
                                      type="button"
                                      onClick={() => {
                                        setOpenEmailDraftRowId(row.id);
                                        setOpenMenuRowId(null);
                                      }}
                                      className="w-full rounded-lg px-3 py-2 text-left text-[12px] hover:bg-gray-50"
                                    >
                                      Edit Email
                                    </button>
                                    <div className="my-1 border-t-2 border-double border-gray-200" />
                                  </>
                                ) : null}
                                {row.rowType === "calendar_event" ? (
                                  <>
                                    <button
                                      type="button"
                                      onClick={() => {
                                        setOpenMeetingEditorRowId(row.id);
                                        setOpenMenuRowId(null);
                                      }}
                                      className="w-full rounded-lg px-3 py-2 text-left text-[12px] hover:bg-gray-50"
                                    >
                                      Edit Meeting
                                    </button>
                                    <div className="my-1 border-t-2 border-double border-gray-200" />
                                  </>
                                ) : null}
                                {row.rowType !== "email" && !normalizeMeetingDraft(row.meetingDraft)?.teamsMeeting && (row.body ?? "").trim() ? (
                                  <>
                                    <button
                                      type="button"
                                      onClick={() => {
                                        setOpenBodyEditorRowId((prev) => (prev === row.id ? null : row.id));
                                        setOpenMenuRowId(null);
                                      }}
                                      className="w-full rounded-lg px-3 py-2 text-left text-[12px] hover:bg-gray-50"
                                    >
                                      Edit Body
                                    </button>
                                    <div className="my-1 border-t-2 border-double border-gray-200" />
                                  </>
                                ) : null}
                                {row.rowType !== "email" && !normalizeMeetingDraft(row.meetingDraft)?.teamsMeeting && !(row.body ?? "").trim() ? (
                                  <>
                                    <button
                                      type="button"
                                      onClick={() => {
                                        setOpenBodyEditorRowId((prev) => (prev === row.id ? null : row.id));
                                        setOpenMenuRowId(null);
                                      }}
                                      className="w-full rounded-lg px-3 py-2 text-left text-[12px] hover:bg-gray-50"
                                    >
                                      Add Body
                                    </button>
                                    <div className="my-1 border-t-2 border-double border-gray-200" />
                                  </>
                                ) : null}
                                {row.rowType === "calendar_event" ? (
                                  <>
                                    <button
                                      type="button"
                                      onClick={() => {
                                        setOpenMeetingEditorRowId(row.id);
                                        setOpenMenuRowId(null);
                                      }}
                                      className="w-full rounded-lg px-3 py-2 text-left text-[12px] hover:bg-gray-50"
                                    >
                                      Edit Duration
                                    </button>
                                  </>
                                ) : null}
                                {row.rowType !== "email" && row.rowType !== "calendar_event" ? (
                                  <>
                                    <button
                                      type="button"
                                      onClick={() => {
                                        setOpenDurationEditorRowId(row.id);
                                        setOpenMenuRowId(null);
                                      }}
                                      className="w-full rounded-lg px-3 py-2 text-left text-[12px] hover:bg-gray-50"
                                    >
                                      Edit Duration
                                    </button>
                                  </>
                                ) : null}
                                <div className="my-1 border-t border-gray-200" />
                                {row.dateBasis === "today" ? (
                                  <>
                                    <button
                                      type="button"
                                      onClick={() => {
                                        updateRow(row.id, (current) => ({ ...current, dateBasis: "event" }));
                                        setOpenMenuRowId(null);
                                      }}
                                      className="w-full rounded-lg px-3 py-2 text-left text-[12px] hover:bg-gray-50"
                                    >
                                      Base on Event Date
                                    </button>
                                    <div className="my-1 border-t border-gray-200" />
                                  </>
                                ) : (
                                  <>
                                    <button
                                      type="button"
                                      onClick={() => {
                                        updateRow(row.id, (current) => ({ ...current, dateBasis: "today" }));
                                        setOpenMenuRowId(null);
                                      }}
                                      className="w-full rounded-lg px-3 py-2 text-left text-[12px] hover:bg-gray-50"
                                    >
                                      Base on Today&apos;s Date
                                    </button>
                                    <div className="my-1 border-t border-gray-200" />
                                  </>
                                )}
                                {anchors
                                  .map((anchor) => anchor.key)
                                  .filter((anchorKey) => normalizeAnchorKey(anchorKey).includes("time"))
                                  .map((anchorKey) => (
                                    <button
                                      key={`${row.id}:menu:${anchorKey}`}
                                      type="button"
                                      onClick={() => {
                                        updateRow(row.id, (current) => ({
                                          ...current,
                                          reminderTime: `[${anchorKey}]`,
                                        }));
                                        setOpenMenuRowId(null);
                                        requestAnimationFrame(() => {
                                          const input = builderTimeInputRefs.current[row.id];
                                          input?.focus();
                                          input?.select();
                                        });
                                      }}
                                      className="w-full rounded-lg px-3 py-2 text-left text-[12px] hover:bg-gray-50"
                                    >
                                      Use [{anchorKey}]
                                    </button>
                                  ))}
                                {anchors.some((anchor) => normalizeAnchorKey(anchor.key).includes("time")) ? (
                                  <div className="my-1 border-t border-gray-200" />
                                ) : null}
                                {row.rowType !== "email" ? (
                                  <button
                                    type="button"
                                    onClick={() => {
                                      updateRow(row.id, (current) => ({
                                        ...current,
                                        rowType: "email",
                                        emailDraft: normalizeEmailDraft(current.emailDraft),
                                      }));
                                      setOpenMenuRowId(null);
                                    }}
                                    className="w-full rounded-lg px-3 py-2 text-left text-[12px] hover:bg-gray-50"
                                  >
                                    Make this an Email
                                  </button>
                                ) : null}
                                {row.rowType !== "calendar_event" ? (
                                  <button
                                    type="button"
                                    onClick={() => {
                                      updateRow(row.id, (current) => ({
                                        ...current,
                                        rowType: "calendar_event",
                                        meetingDraft: normalizeMeetingDraft(current.meetingDraft) ?? {
                                          attendees: [],
                                          location: "",
                                          durationMinutes: 30,
                                        },
                                      }));
                                      setOpenMenuRowId(null);
                                    }}
                                    className="w-full rounded-lg px-3 py-2 text-left text-[12px] hover:bg-gray-50"
                                  >
                                    Make this a Meeting
                                  </button>
                                ) : null}
                              <button
                                type="button"
                                onClick={() => {
                                  setRows((current) => current.filter((entry) => entry.id !== row.id));
                                  setOpenMenuRowId(null);
                                  }}
                                  className="w-full rounded-lg px-3 py-2 text-left text-[12px] text-red-700 hover:bg-red-50"
                                >
                                  Delete
                                </button>
                              </div>
                            ) : null}
                          </div>
                        </div>
                        {openEmailDraftRowId === row.id ? (
                          <div className="rounded-xl border border-amber-200 bg-amber-50 p-4">
                            {(() => {
                              const normalizedDraft = normalizeEmailDraft(row.emailDraft);
                              const showCc = Boolean(emailFieldVisibility[row.id]?.cc || normalizedDraft.cc.length);
                              const showBcc = Boolean(emailFieldVisibility[row.id]?.bcc || normalizedDraft.bcc.length);

                              return (
                                <div className="grid grid-cols-1 gap-3 md:grid-cols-2">
                                  <div className="md:col-span-2">
                                    <div className="mb-1 flex items-center justify-between gap-3">
                                      <label className="block text-sm font-medium text-gray-700">To</label>
                                      <div className="flex items-center gap-3 text-xs font-medium text-amber-900">
                                        {!showCc ? (
                                          <button
                                            type="button"
                                            onClick={() =>
                                              setEmailFieldVisibility((prev) => ({
                                                ...prev,
                                                [row.id]: { ...prev[row.id], cc: true },
                                              }))
                                            }
                                            className="hover:underline"
                                          >
                                            +Cc
                                          </button>
                                        ) : null}
                                        {!showBcc ? (
                                          <button
                                            type="button"
                                            onClick={() =>
                                              setEmailFieldVisibility((prev) => ({
                                                ...prev,
                                                [row.id]: { ...prev[row.id], bcc: true },
                                              }))
                                            }
                                            className="hover:underline"
                                          >
                                            +Bcc
                                          </button>
                                        ) : null}
                                      </div>
                                    </div>
                                    <EmailTokensInput
                                      label=""
                                      values={normalizedDraft.to}
                                      placeholder="Type an email and press Enter or comma"
                                      onChange={(nextValues) =>
                                        updateRow(row.id, (current) => ({
                                          ...current,
                                          emailDraft: { ...normalizeEmailDraft(current.emailDraft), to: nextValues },
                                        }))
                                      }
                                    />
                                  </div>
                                  {showCc ? (
                                    <div className="md:col-span-2">
                                      <EmailTokensInput
                                        label="CC"
                                        values={normalizedDraft.cc}
                                        placeholder="Add CC recipients"
                                        onChange={(nextValues) =>
                                          updateRow(row.id, (current) => ({
                                            ...current,
                                            emailDraft: { ...normalizeEmailDraft(current.emailDraft), cc: nextValues },
                                          }))
                                        }
                                      />
                                    </div>
                                  ) : null}
                                  {showBcc ? (
                                    <div className="md:col-span-2">
                                      <EmailTokensInput
                                        label="BCC"
                                        values={normalizedDraft.bcc}
                                        placeholder="Add BCC recipients"
                                        onChange={(nextValues) =>
                                          updateRow(row.id, (current) => ({
                                            ...current,
                                            emailDraft: { ...normalizeEmailDraft(current.emailDraft), bcc: nextValues },
                                          }))
                                        }
                                      />
                                    </div>
                                  ) : null}
                                  <div className="md:col-span-2">
                                    <label className="mb-1 block text-sm font-medium text-gray-700">Subject</label>
                                    <input
                                      className="w-full rounded-lg border bg-white px-3 py-2"
                                      value={normalizedDraft.subject}
                                      onChange={(e) =>
                                        updateRow(row.id, (current) => ({
                                          ...current,
                                          emailDraft: { ...normalizeEmailDraft(current.emailDraft), subject: e.target.value },
                                        }))
                                      }
                                    />
                                  </div>
                                  <div className="md:col-span-2">
                                    <label className="mb-1 block text-sm font-medium text-gray-700">Message</label>
                                    <textarea
                                      className="min-h-28 w-full rounded-lg border bg-white px-3 py-2"
                                      value={normalizedDraft.body}
                                      onChange={(e) =>
                                        updateRow(row.id, (current) => ({
                                          ...current,
                                          emailDraft: { ...normalizeEmailDraft(current.emailDraft), body: e.target.value },
                                        }))
                                      }
                                    />
                                  </div>
                                </div>
                              );
                            })()}
                            <div className="mt-4 flex justify-end">
                              <button
                                type="button"
                                onClick={() => {
                                  collapseEmptyEmailFields(row.id, row.emailDraft);
                                  setOpenEmailDraftRowId(null);
                                }}
                                className="rounded-lg border px-3 py-2 text-sm hover:bg-white"
                              >
                                Done
                              </button>
                            </div>
                          </div>
                        ) : null}

                        {openDurationEditorRowId === row.id && row.rowType !== "email" && row.rowType !== "calendar_event" ? (
                          <div className="rounded-xl border border-blue-200 bg-blue-50 p-4">
                            {(() => {
                              const previewDurationItem = previewPlan.items.find((preview) => preview.id === row.id);
                              const computedStartDate = (previewDurationItem ? getEffectivePreviewItemDate(previewDurationItem) : null) || todayYYYYMMDD();
                              const computedStartTime =
                                getUsableReminderTime(previewDurationItem?.reminderTime, anchorMap) ||
                                getUsableReminderTime(row.reminderTime, anchorMap) ||
                                "09:00";
                              const normalizedDuration = normalizeDurationDraft(row.durationDraft);
                              const derivedEnd = addMinutesToLocalDateTime(
                                computedStartDate,
                                computedStartTime,
                                normalizedDuration?.durationMinutes ?? 30
                              );

                              return (
                                <div className="grid grid-cols-1 gap-3">
                                  <div>
                                    <label className="mb-1 block text-sm font-medium text-gray-700">Duration</label>
                                    <select
                                      className="w-full rounded-lg border bg-white px-3 py-2"
                                      value={normalizedDuration?.useCustomEnd ? "custom" : String(normalizedDuration?.durationMinutes ?? 30)}
                                      onChange={(e) =>
                                        e.target.value === "custom"
                                          ? updateRow(row.id, (current) => ({
                                              ...current,
                                              durationDraft: {
                                                useCustomEnd: true,
                                                endDate: normalizedDuration?.endDate || computedStartDate,
                                                endTime: normalizedDuration?.endTime || derivedEnd.endTime,
                                              },
                                            }))
                                          : updateRow(row.id, (current) => ({
                                              ...current,
                                              durationDraft: {
                                                durationMinutes: Number(e.target.value),
                                                useCustomEnd: false,
                                                endDate: computedStartDate,
                                                endTime: addMinutesToLocalDateTime(
                                                  computedStartDate,
                                                  computedStartTime,
                                                  Number(e.target.value)
                                                ).endTime,
                                              },
                                            }))
                                      }
                                    >
                                      {MEETING_DURATION_OPTIONS.map((option) => (
                                        <option key={option.value} value={option.value}>
                                          {option.label}
                                        </option>
                                      ))}
                                    </select>
                                  </div>
                                  {normalizedDuration?.useCustomEnd ? (
                                    <>
                                      <div>
                                        <label className="mb-1 block text-sm font-medium text-gray-700">End Date</label>
                                        <input
                                          type="date"
                                          className="w-full rounded-lg border bg-white px-3 py-2"
                                          value={normalizedDuration.endDate || computedStartDate}
                                          onChange={(e) =>
                                            updateRow(row.id, (current) => ({
                                              ...current,
                                              durationDraft: {
                                                ...normalizeDurationDraft(current.durationDraft),
                                                endDate: e.target.value,
                                                useCustomEnd: true,
                                              },
                                            }))
                                          }
                                        />
                                      </div>
                                      <div>
                                        <label className="mb-1 block text-sm font-medium text-gray-700">End Time</label>
                                        <input
                                          type="time"
                                          className="w-full rounded-lg border bg-white px-3 py-2"
                                          value={normalizedDuration.endTime || derivedEnd.endTime}
                                          onChange={(e) =>
                                            updateRow(row.id, (current) => ({
                                              ...current,
                                              durationDraft: {
                                                ...normalizeDurationDraft(current.durationDraft),
                                                endTime: e.target.value,
                                                useCustomEnd: true,
                                              },
                                            }))
                                          }
                                        />
                                      </div>
                                    </>
                                  ) : null}
                                  <div className="flex items-end">
                                    <label className="flex items-center gap-2 text-sm font-medium text-gray-700">
                                      <input
                                        type="checkbox"
                                        checked={Boolean(normalizedDuration?.isAllDay)}
                                        onChange={(e) =>
                                          updateRow(row.id, (current) => ({
                                            ...current,
                                            durationDraft: {
                                              ...normalizeDurationDraft(current.durationDraft),
                                              isAllDay: e.target.checked,
                                            },
                                          }))
                                        }
                                      />
                                      All day
                                    </label>
                                  </div>
                                </div>
                              );
                            })()}
                            <div className="mt-4 flex justify-end">
                              <button
                                type="button"
                                onClick={() => setOpenDurationEditorRowId(null)}
                                className="rounded-lg border px-3 py-2 text-sm hover:bg-white"
                              >
                                Done
                              </button>
                            </div>
                          </div>
                        ) : null}

                        {openMeetingEditorRowId === row.id || forcedOpenMeetingEditorRowIds.includes(row.id) ? (
                          <div className="rounded-xl border border-violet-200 bg-violet-50 p-4">
                            <div className="grid grid-cols-1 gap-3 md:grid-cols-2">
                              <div className="md:col-span-2">
                                <label className="mb-1 block text-sm font-medium text-gray-700">Meeting Title</label>
                                <input
                                  className="w-full rounded-lg border bg-white px-3 py-2 text-gray-700"
                                  value={row.title}
                                  onChange={(e) => updateRow(row.id, (current) => ({ ...current, title: e.target.value }))}
                                />
                              </div>
                              <div className="md:col-span-2">
                                <label className="mb-1 block text-sm font-medium text-gray-700">
                                  Attendees
                                  {meetingErrors?.attendees ? <span className="ml-1 text-red-500">*</span> : null}
                                </label>
                                <EmailTokensInput
                                  label=""
                                  values={normalizeMeetingDraft(row.meetingDraft)?.attendees ?? []}
                                  onChange={(nextValues) =>
                                    updateRow(row.id, (current) => ({
                                      ...current,
                                      meetingDraft: { ...normalizeMeetingDraft(current.meetingDraft), attendees: nextValues },
                                    }))
                                  }
                                  placeholder="Add attendee emails"
                                  hasError={Boolean(meetingErrors?.attendees)}
                                />
                              </div>
                              <div>
                                <label className="mb-1 block text-sm font-medium text-gray-700">Location</label>
                                <input
                                  className={`w-full rounded-lg border px-3 py-2 ${
                                    normalizeMeetingDraft(row.meetingDraft)?.teamsMeeting ? "bg-gray-100 text-gray-600" : "bg-white"
                                  }`}
                                  value={
                                    normalizeMeetingDraft(row.meetingDraft)?.teamsMeeting
                                      ? TEAMS_MEETING_LOCATION
                                      : normalizeMeetingDraft(row.meetingDraft)?.location ?? ""
                                  }
                                  readOnly={Boolean(normalizeMeetingDraft(row.meetingDraft)?.teamsMeeting)}
                                  onChange={(e) =>
                                    updateRow(row.id, (current) => ({
                                      ...current,
                                      meetingDraft: { ...normalizeMeetingDraft(current.meetingDraft), location: e.target.value },
                                    }))
                                  }
                                />
                              </div>
                              <div>
                                <label className="mb-1 block text-sm font-medium text-gray-700">
                                  Meeting Duration
                                  {meetingErrors?.duration ? <span className="ml-1 text-red-500">*</span> : null}
                                </label>
                                <select
                                  className={`w-full rounded-lg border bg-white px-3 py-2 ${
                                    meetingErrors?.duration ? "border-red-400 ring-1 ring-red-100" : ""
                                  }`}
                                  value={
                                    normalizeMeetingDraft(row.meetingDraft)?.useCustomEnd
                                      ? "custom"
                                      : String(normalizeMeetingDraft(row.meetingDraft)?.durationMinutes ?? 30)
                                  }
                                  onChange={(e) =>
                                    e.target.value === "custom"
                                      ? (() => {
                                          const defaultCustomEnd = addMinutesToLocalDateTime(
                                            anchorDate,
                                            getUsableReminderTime(row.reminderTime, anchorMap) || "09:00",
                                            normalizeMeetingDraft(row.meetingDraft)?.durationMinutes ?? 30
                                          );
                                          updateRow(row.id, (current) => ({
                                            ...current,
                                            meetingDraft: {
                                              ...normalizeMeetingDraft(current.meetingDraft),
                                              useCustomEnd: true,
                                              endDate: normalizeMeetingDraft(current.meetingDraft)?.endDate || defaultCustomEnd.endDate,
                                              endTime: normalizeMeetingDraft(current.meetingDraft)?.endTime || defaultCustomEnd.endTime,
                                            },
                                          }));
                                        })()
                                      : updateRow(row.id, (current) => ({
                                          ...current,
                                          meetingDraft: {
                                            ...normalizeMeetingDraft(current.meetingDraft),
                                            useCustomEnd: false,
                                            durationMinutes: Number(e.target.value),
                                          },
                                        }))
                                  }
                                >
                                  {MEETING_DURATION_OPTIONS.map((option) => (
                                    <option key={option.value} value={option.value}>
                                      {option.label}
                                    </option>
                                  ))}
                                </select>
                              </div>
                              <div className="md:col-span-2 flex items-center gap-2 text-sm text-gray-700">
                                <input
                                  type="checkbox"
                                  checked={Boolean(normalizeMeetingDraft(row.meetingDraft)?.teamsMeeting)}
                                  onChange={(e) =>
                                    updateRow(row.id, (current) => ({
                                      ...current,
                                      meetingDraft: {
                                        ...normalizeMeetingDraft(current.meetingDraft),
                                        teamsMeeting: e.target.checked,
                                        location: e.target.checked
                                          ? TEAMS_MEETING_LOCATION
                                          : normalizeMeetingDraft(current.meetingDraft)?.location ?? "",
                                      },
                                    }))
                                  }
                                />
                                <span>Microsoft Teams Meeting</span>
                              </div>
                              {normalizeMeetingDraft(row.meetingDraft)?.teamsMeeting ? (
                                <div className="md:col-span-2 rounded-xl border border-violet-200 bg-white p-4">
                                  <div className="text-sm font-semibold text-violet-950">Teams Meeting Details</div>
                                  <div className="mt-3 space-y-3 text-sm text-gray-700">
                                    <div>
                                      <div className="mb-1 font-medium text-violet-950">Join link</div>
                                      <div className="text-gray-500">
                                        Teams join info will appear here after the Outlook event is created.
                                      </div>
                                    </div>
                                  </div>
                                </div>
                              ) : null}
                              {normalizeMeetingDraft(row.meetingDraft)?.useCustomEnd ? (
                                <>
                                  {!normalizeMeetingDraft(row.meetingDraft)?.isAllDay ? (
                                    <>
                                      <div className="md:col-span-2">
                                        <label className="mb-1 block text-sm font-medium text-gray-700">
                                          End Date
                                          {meetingErrors?.duration ? <span className="ml-1 text-red-500">*</span> : null}
                                        </label>
                                        <input
                                          type="date"
                                          className={`w-full rounded-lg border bg-white px-3 py-2 ${
                                            meetingErrors?.duration ? "border-red-400 ring-1 ring-red-100" : ""
                                          }`}
                                          value={normalizeMeetingDraft(row.meetingDraft)?.endDate || anchorDate}
                                          onChange={(e) =>
                                            updateRow(row.id, (current) => ({
                                              ...current,
                                              meetingDraft: {
                                                ...normalizeMeetingDraft(current.meetingDraft),
                                                endDate: e.target.value,
                                                useCustomEnd: true,
                                              },
                                            }))
                                          }
                                        />
                                      </div>
                                      <div className="md:col-span-2">
                                        <label className="mb-1 block text-sm font-medium text-gray-700">
                                          End Time
                                          {meetingErrors?.duration ? <span className="ml-1 text-red-500">*</span> : null}
                                        </label>
                                        <input
                                          type="time"
                                          className={`w-full rounded-lg border bg-white px-3 py-2 ${
                                            meetingErrors?.duration ? "border-red-400 ring-1 ring-red-100" : ""
                                          }`}
                                          value={normalizeMeetingDraft(row.meetingDraft)?.endTime || "09:30"}
                                          onChange={(e) =>
                                            updateRow(row.id, (current) => ({
                                              ...current,
                                              meetingDraft: {
                                                ...normalizeMeetingDraft(current.meetingDraft),
                                                endTime: e.target.value,
                                                useCustomEnd: true,
                                              },
                                            }))
                                          }
                                        />
                                      </div>
                                    </>
                                  ) : null}
                                  <div className="md:col-span-2 flex items-center gap-2 text-sm text-gray-700">
                                    <input
                                      type="checkbox"
                                      checked={Boolean(normalizeMeetingDraft(row.meetingDraft)?.isAllDay)}
                                      onChange={(e) =>
                                        updateRow(row.id, (current) => ({
                                          ...current,
                                          meetingDraft: {
                                            ...normalizeMeetingDraft(current.meetingDraft),
                                            isAllDay: e.target.checked,
                                          },
                                        }))
                                      }
                                    />
                                    <span>All day meeting</span>
                                  </div>
                                </>
                              ) : null}
                            </div>
                            <div className="mt-4 flex justify-end">
                              <button
                                type="button"
                                onClick={() => {
                                  setOpenMeetingEditorRowId(null);
                                  setForcedOpenMeetingEditorRowIds((current) => current.filter((id) => id !== row.id));
                                }}
                                className="rounded-lg border px-3 py-2 text-sm hover:bg-white"
                              >
                                Done
                              </button>
                            </div>
                          </div>
                        ) : null}

                        {openBodyEditorRowId === row.id ? (
                          <div className="rounded-xl border bg-gray-50 p-4">
                            {normalizeMeetingDraft(row.meetingDraft)?.teamsMeeting ? (
                              <div className="space-y-3">
                                <label className="block text-sm font-medium text-gray-700">Teams Meeting Details</label>
                                <div className="rounded-lg border bg-white px-3 py-3 text-sm text-gray-500">
                                  Teams join info will appear here after the Outlook event is created.
                                </div>
                              </div>
                            ) : (
                              <>
                                <label className="mb-2 block text-sm font-medium text-gray-700">Body</label>
                                <textarea
                                  className="min-h-32 w-full rounded-lg border px-3 py-2 text-sm"
                                  value={row.body ?? ""}
                                  onChange={(e) => updateRow(row.id, (current) => ({ ...current, body: e.target.value }))}
                                />
                              </>
                            )}
                            <div className="mt-4 flex justify-end">
                              <button
                                type="button"
                                onClick={() => setOpenBodyEditorRowId(null)}
                                className="rounded-lg border px-3 py-2 text-sm hover:bg-white"
                              >
                                Done
                              </button>
                            </div>
                          </div>
                        ) : null}
                      </div>
                    );
                  })}
                </div>
              </div>

              <div className="pt-2">
                <div className="space-y-4">
                  <div className="flex flex-wrap items-center gap-3">
                    <button
                      type="button"
                      onClick={addReminderRow}
                      className="rounded-lg border border-blue-200 bg-blue-50 px-3 py-2 text-sm text-blue-700 hover:bg-blue-100"
                    >
                      + Add a Reminder
                    </button>
                    <button
                      type="button"
                      onClick={addEmailRow}
                      className="rounded-lg border border-green-200 bg-green-50 px-3 py-2 text-sm text-green-700 hover:bg-green-100"
                    >
                      + Add an Email
                    </button>
                    <button
                      type="button"
                      onClick={addMeetingRow}
                      className="rounded-lg border border-violet-200 bg-violet-50 px-3 py-2 text-sm text-violet-700 hover:bg-violet-100"
                    >
                      + Add a Meeting
                    </button>
                  </div>

                  <div className="space-y-4">
                    <div className="text-lg font-semibold text-gray-900">Anchors</div>
                    <div className="flex flex-wrap items-center gap-2">
                      <button
                        type="button"
                        onClick={() => {
                          setAnchors((current) => [...current, createEmptyAnchor()]);
                          setAreAnchorsHidden(false);
                        }}
                        className="rounded-lg border px-3 py-2 text-sm hover:bg-gray-50"
                      >
                        + Add Anchor
                      </button>
                      <button
                        type="button"
                        onClick={() => {
                          if (anchors.length === 0) return;
                          setAreAnchorsHidden((prev) => !prev);
                        }}
                        disabled={anchors.length === 0}
                        className="rounded-lg border px-3 py-2 text-sm hover:bg-gray-50 disabled:cursor-not-allowed disabled:opacity-50"
                      >
                        {areAnchorsHidden ? "Show Anchors" : "Hide Anchors"}
                      </button>
                    </div>
                    {!areAnchorsHidden ? (
                    <div className="space-y-3">
                      {resolvedAnchors.map((anchor) => {
                        return (
                          <div key={anchor.id} className="grid grid-cols-1 gap-3 md:grid-cols-[240px_minmax(0,320px)_auto_auto]">
                            <div className="flex items-center rounded-lg border bg-white px-3 py-2 font-mono text-sm">
                              <span className="text-gray-500">[</span>
                              <input
                                className="min-w-0 flex-1 border-0 bg-transparent px-1 text-center uppercase focus:outline-none focus:ring-0"
                                placeholder="KEY"
                                value={anchor.key}
                                readOnly={Boolean(anchor.locked)}
                                onChange={(e) =>
                                  setAnchors((current) =>
                                    current.map((entry) =>
                                      entry.id === anchor.id ? { ...entry, key: e.target.value.toUpperCase() } : entry
                                    )
                                  )
                                }
                              />
                              <span className="text-gray-500">]</span>
                            </div>
                              <input
                                className="w-full rounded-lg border px-3 py-2 text-center text-sm"
                                placeholder="Value"
                                value={anchor.locked ? anchor.displayValue : anchor.value}
                                readOnly={Boolean(anchor.locked)}
                                onChange={(e) =>
                                  setAnchors((current) =>
                                  current.map((entry) =>
                                    entry.id === anchor.id ? { ...entry, value: e.target.value } : entry
                                  )
                                )
                              }
                            />
                            <button
                              type="button"
                              onMouseDown={(e) => e.preventDefault()}
                              className="rounded-lg border px-2 py-2 text-sm hover:bg-gray-50"
                              onClick={() => onInsertAnchor(anchor.key)}
                            >
                              Insert
                            </button>
                            {anchor.locked ? (
                              <div />
                            ) : (
                              <button
                                type="button"
                                onClick={() => setAnchors((current) => current.filter((entry) => entry.id !== anchor.id))}
                                className="rounded-lg border px-2 py-2 text-sm text-red-700 hover:bg-red-50"
                              >
                                Delete
                              </button>
                            )}
                          </div>
                        );
                      })}
                    </div>
                    ) : null}
                  </div>

                  <div className="flex flex-col gap-4 md:flex-row md:items-end md:justify-between">
                    <div className="max-w-md">
                      <label className="mb-1 block text-sm font-medium text-gray-700">Weekend handling</label>
                      <select
                        className="w-full rounded-lg border px-3 py-2"
                        value={weekendRule}
                        onChange={(e) => setWeekendRule(e.target.value as WeekendRule)}
                      >
                        <option value="prior_business_day">Adjust to prior business day (Fri)</option>
                        <option value="none">Allow weekends (no adjustment)</option>
                      </select>
                      <p className="mt-1 text-xs text-gray-500">
                        If a computed date lands on Sat/Sun, we either move it to Friday or leave it as-is.
                      </p>
                    </div>

                    <div className="flex justify-end">
                      <div className="flex w-full max-w-64 flex-col gap-2">
                        <button
                          type="button"
                          onClick={() => {
                            if (validatePreviewBeforeOpen()) return;
                            setExpandedPreviewReminderRowIds([]);
                            setExpandedPreviewEmailRowIds([]);
                            setExpandedPreviewMeetingRowIds([]);
                            setOpenPreviewRowMenuId(null);
                            setIsBuilderPreviewOpen(true);
                          }}
                          className="rounded-lg border px-4 py-2 text-sm hover:bg-gray-50"
                        >
                          Preview
                        </button>
                        <button
                          type="button"
                          onClick={saveCurrentTemplate}
                          className="rounded-lg border px-4 py-2 text-sm hover:bg-gray-50"
                        >
                          Save
                        </button>
                        <button
                          type="button"
                          onClick={cancelEditing}
                          className="rounded-lg border px-4 py-2 text-sm hover:bg-gray-50"
                        >
                          Cancel
                        </button>
                        <button
                          type="button"
                          onClick={() => {
                            void exportCurrentPlan();
                          }}
                          className="rounded-lg bg-blue-600 px-4 py-2 text-white hover:bg-blue-700"
                        >
                          Export
                        </button>
                        {executionNotice?.tone === "success" ? (
                          <div className="flex justify-end">
                            <ExportDoneBadge />
                          </div>
                        ) : null}
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            </div>
      </section>

      {isBuilderPreviewOpen ? (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/35 p-4">
          <div className="max-h-[90vh] w-full max-w-6xl overflow-auto rounded-2xl bg-white shadow-2xl">
            <div className="flex items-center justify-between border-b px-6 py-4">
              <div>
                <h2 className="text-xl font-semibold text-gray-900">Plan Preview</h2>
                <p className="mt-1 text-sm text-gray-600">
                  Review the current builder plan schedule before exporting it to Outlook or as a local file for review.
                </p>
              </div>
              <button
                onClick={() => {
                  setIsBuilderPreviewOpen(false);
                  setExpandedPreviewReminderRowIds([]);
                  setExpandedPreviewEmailRowIds([]);
                  setExpandedPreviewMeetingRowIds([]);
                  setOpenPreviewRowMenuId(null);
                }}
                className="rounded-lg border px-3 py-2 text-sm hover:bg-gray-50"
              >
                Close
              </button>
            </div>

            <div className="space-y-6 p-6">
              {executionNotice ? (
                <OutlookExecutionNoticeCard notice={executionNotice} onDismiss={() => setExecutionNotice(null)} />
              ) : null}
              {previewPlanForRender ? (
                <>
                  <div className="rounded-xl border">
                    <div className="border-b bg-gray-50 px-4 py-3">
                      <div className="font-semibold text-gray-900">{previewPlanForRender.name}</div>
                      <div className="mt-1 text-sm text-gray-600">
                        {previewPlanForRender.items.filter((item) => classifyPlanRow(item) !== "email").length} plan events
                      </div>
                    </div>
                    <div className="grid items-center gap-x-6 border-b bg-gray-50 px-4 py-2 text-xs font-semibold uppercase tracking-wide text-gray-600 md:grid-cols-[120px_minmax(0,5.5fr)_150px_130px_170px_56px]">
                      <div className="flex min-h-[28px] items-center">Plan type</div>
                      <div className="flex min-h-[28px] items-center">Plan name</div>
                      <div className="flex min-h-[28px] items-center justify-center text-center">
                        {noEventDate ? "Days from today" : "Days from event"}
                      </div>
                      <div className="flex min-h-[28px] items-center justify-center text-center">Time</div>
                      <div className="flex min-h-[28px] items-center justify-center text-center">Review</div>
                      <div className="flex min-h-[28px] items-center justify-end text-right">Actions</div>
                    </div>
                    <div className="divide-y">
                      {previewPlanForRender.items.map((item) => {
                    const rowKind = classifyPlanRow(item);
                    const builderItem = rows.find((entry) => entry.id === item.id);
                    const previewReminderBody = item.body ?? builderItem?.body ?? "";
                    const hasReminderPreview = rowKind === "reminder";
                    const hasEmailPreview = rowKind === "email";
                    const hasMeetingPreview = rowKind === "meeting";
                    const isReminderExpanded = expandedPreviewReminderRowIds.includes(item.id);
                    const isEmailExpanded = expandedPreviewEmailRowIds.includes(item.id);
                    const isMeetingExpanded = expandedPreviewMeetingRowIds.includes(item.id);
                    const previewEmailDraft = normalizeEmailDraft(item.emailDraft);
                    const previewEmailSubject = previewEmailDraft.subject.trim() || item.title || "Email draft";
                    const previewEmailFinalBody = buildFinalEmailBody(previewEmailDraft.body, {
                      enabled: appSettings.emailSignatureEnabled,
                      signature: appSettings.emailSignatureText,
                    });
                    const previewMeetingDraft = normalizeMeetingDraft(item.meetingDraft);
                    const isPreviewTeamsMeetingEnabled = Boolean(previewMeetingDraft?.teamsMeeting);
                    const rowTypeLabel = rowKind === "email" ? "Email" : rowKind === "meeting" ? "Meeting" : "Reminder";

                    return (
                      <div key={item.id}>
                        <div className="grid items-center gap-x-6 px-4 py-3 md:grid-cols-[120px_minmax(0,5.5fr)_150px_130px_170px_56px]">
                          <div className="flex min-h-[52px] items-center text-sm font-medium text-gray-700">
                            {rowTypeLabel}
                          </div>
                          <div className="flex min-h-[52px] items-center">
                            <div>
                              <div className="text-sm leading-6 text-gray-900 [overflow-wrap:anywhere]">
                                {item.customTitle ?? item.title}
                              </div>
                              {hasEmailPreview ? (
                                <div className="mt-1 text-xs font-medium text-gray-600">
                                  {getPreviewEmailModeMessage(appSettings.emailHandlingMode)}
                                </div>
                              ) : null}
                            </div>
                          </div>
                          <div className="flex min-h-[52px] items-center justify-center text-center text-sm text-gray-700">
                            {formatOffsetLabel(item.offsetDays, { relativeToToday: noEventDate, dateBasis: item.dateBasis })}
                          </div>
                          <div className="flex min-h-[52px] items-center justify-center text-center text-sm text-gray-700">
                            {item.meetingDraft?.isAllDay || item.durationDraft?.isAllDay
                              ? "All day"
                              : item.reminderTime
                                ? formatPreviewTime(item.reminderTime)
                                : "All day"}
                          </div>
                          <div className="flex min-h-[52px] items-center justify-center">
                            <div className="flex flex-col items-center">
                              <button
                                type="button"
                                onClick={() => {
                                  if (hasEmailPreview) {
                                    exportPreviewEmailItem(item.id);
                                    return;
                                  }
                                  if (hasMeetingPreview) {
                                    exportPreviewMeetingItem(item.id);
                                    return;
                                  }
                                  exportPreviewReminderItem(item.id);
                                }}
                                title={
                                  hasReminderPreview
                                    ? "Export this reminder"
                                    : hasEmailPreview
                                      ? "Execute this email action"
                                      : "Create this meeting event"
                                }
                                className="flex h-14 w-44 items-center justify-center rounded-lg border px-3 py-2 text-center text-sm hover:bg-gray-50"
                              >
                                {hasEmailPreview
                                  ? getPreviewEmailActionLabel(appSettings.emailHandlingMode)
                                  : hasMeetingPreview
                                    ? "Add Meeting to Calendar"
                                    : "Export This Reminder"}
                              </button>
                            </div>
                          </div>
                          <div className="relative flex min-h-[52px] items-center justify-end">
                            {hasReminderPreview || hasEmailPreview || hasMeetingPreview ? (
                              <>
                                <button
                                  type="button"
                                  onClick={() => setOpenPreviewRowMenuId((prev) => (prev === item.id ? null : item.id))}
                                  title="Actions"
                                  aria-label="Actions"
                                  className="flex h-8 w-10 items-center justify-center rounded-lg border border-gray-200 bg-white text-gray-500 hover:border-gray-300 hover:bg-gray-100 hover:text-gray-700"
                                >
                                  <span className="text-base leading-none">•••</span>
                                </button>
                                {openPreviewRowMenuId === item.id ? (
                                  <div className="absolute right-0 top-[calc(100%+0.5rem)] z-20 w-40 rounded-xl border bg-white p-2 text-left shadow-lg">
                                    {hasReminderPreview ? (
                                      <button
                                        type="button"
                                        onClick={() => {
                                          setExpandedPreviewReminderRowIds((prev) =>
                                            prev.includes(item.id) ? prev.filter((id) => id !== item.id) : [...prev, item.id]
                                          );
                                          setOpenPreviewRowMenuId(null);
                                        }}
                                        className="w-full rounded-lg px-3 py-2 text-left text-[12px] hover:bg-gray-50"
                                      >
                                        {isReminderExpanded ? "Hide Reminder" : "View Reminder"}
                                      </button>
                                    ) : null}
                                    {hasEmailPreview ? (
                                      <button
                                        type="button"
                                        onClick={() => {
                                          setExpandedPreviewEmailRowIds((prev) =>
                                            prev.includes(item.id) ? prev.filter((id) => id !== item.id) : [...prev, item.id]
                                          );
                                          setOpenPreviewRowMenuId(null);
                                        }}
                                        className="w-full rounded-lg px-3 py-2 text-left text-[12px] hover:bg-gray-50"
                                      >
                                        {isEmailExpanded ? "Hide Email" : "View Email"}
                                      </button>
                                    ) : null}
                                    {hasMeetingPreview ? (
                                      <button
                                        type="button"
                                        onClick={() => {
                                          setExpandedPreviewMeetingRowIds((prev) =>
                                            prev.includes(item.id) ? prev.filter((id) => id !== item.id) : [...prev, item.id]
                                          );
                                          setOpenPreviewRowMenuId(null);
                                        }}
                                        className="w-full rounded-lg px-3 py-2 text-left text-[12px] hover:bg-gray-50"
                                      >
                                        {isMeetingExpanded ? "Hide Meeting" : "View Meeting"}
                                      </button>
                                    ) : null}
                                  </div>
                                ) : null}
                              </>
                            ) : null}
                          </div>
                        </div>

                        {rowKind === "reminder" && isReminderExpanded ? (
                          <div className="bg-blue-50 px-4 py-3">
                            <div className="grid grid-cols-1 gap-3">
                              <div>
                                <label className="mb-1 block text-sm font-medium text-blue-950">Reminder Body</label>
                                <textarea
                                  rows={5}
                                  className="w-full rounded-lg border border-blue-200 bg-white px-3 py-2 text-sm text-gray-900"
                                  value={previewReminderBody}
                                  onChange={(e) =>
                                    builderItem ? updateRow(builderItem.id, (current) => ({ ...current, body: e.target.value })) : undefined
                                  }
                                />
                              </div>
                            </div>
                          </div>
                        ) : null}

                        {rowKind === "meeting" && isMeetingExpanded ? (
                          <div className="bg-violet-50 px-4 py-3">
                            <div className="mb-3 text-sm font-medium text-violet-900">
                              This meeting will be created in Outlook when connected, or downloaded as a local calendar file otherwise.
                            </div>
                            <div className="grid grid-cols-1 gap-3 md:grid-cols-2">
                              <div className="md:col-span-2">
                                <label className="mb-1 block text-sm font-medium text-violet-950">Meeting Title</label>
                                <input
                                  className="w-full rounded-lg border border-violet-200 bg-white px-3 py-2 text-sm text-gray-900"
                                  value={builderItem?.title ?? item.title}
                                  onChange={(e) =>
                                    builderItem ? updateRow(builderItem.id, (current) => ({ ...current, title: e.target.value })) : undefined
                                  }
                                />
                              </div>
                              <div className="md:col-span-2">
                                <label className="mb-1 block text-sm font-medium text-violet-950">Attendees</label>
                                <EmailTokensInput
                                  label=""
                                  values={previewMeetingDraft?.attendees ?? []}
                                  onChange={(nextValues) =>
                                    builderItem
                                      ? updateRow(builderItem.id, (current) => ({
                                          ...current,
                                          meetingDraft: { ...normalizeMeetingDraft(current.meetingDraft), attendees: nextValues },
                                        }))
                                      : undefined
                                  }
                                  placeholder="Add attendee emails"
                                />
                              </div>
                              <div>
                                <label className="mb-1 block text-sm font-medium text-violet-950">Location</label>
                                <input
                                  className={`w-full rounded-lg border border-violet-200 px-3 py-2 text-sm ${
                                    isPreviewTeamsMeetingEnabled ? "bg-gray-100 text-gray-600" : "bg-white text-gray-900"
                                  }`}
                                  value={
                                    isPreviewTeamsMeetingEnabled
                                      ? TEAMS_MEETING_LOCATION
                                      : previewMeetingDraft?.location ?? ""
                                  }
                                  readOnly={isPreviewTeamsMeetingEnabled}
                                  onChange={(e) =>
                                    builderItem
                                      ? updateRow(builderItem.id, (current) => ({
                                          ...current,
                                          meetingDraft: { ...normalizeMeetingDraft(current.meetingDraft), location: e.target.value },
                                        }))
                                      : undefined
                                  }
                                />
                              </div>
                              <div>
                                <label className="mb-1 block text-sm font-medium text-violet-950">Meeting Duration</label>
                                <select
                                  className="w-full rounded-lg border border-violet-200 bg-white px-3 py-2 text-sm text-gray-900"
                                  value={String(previewMeetingDraft?.durationMinutes ?? 30)}
                                  onChange={(e) =>
                                    builderItem
                                      ? updateRow(builderItem.id, (current) => ({
                                          ...current,
                                          meetingDraft: { ...normalizeMeetingDraft(current.meetingDraft), durationMinutes: Number(e.target.value) },
                                        }))
                                      : undefined
                                  }
                                >
                                  {MEETING_DURATION_OPTIONS.filter((option) => option.value !== "custom").map((option) => (
                                    <option key={option.value} value={option.value}>
                                      {option.label}
                                    </option>
                                  ))}
                                </select>
                              </div>
                              {previewMeetingDraft?.useCustomEnd ? (
                                <>
                                  {!previewMeetingDraft?.isAllDay ? (
                                    <>
                                      <div className="md:col-span-2">
                                        <label className="mb-1 block text-sm font-medium text-violet-950">End Date</label>
                                        <input
                                          type="date"
                                          className="w-full rounded-lg border border-violet-200 bg-white px-3 py-2 text-sm text-gray-900"
                                          value={previewMeetingDraft?.endDate ?? ""}
                                          onChange={(e) =>
                                            builderItem
                                              ? updateRow(builderItem.id, (current) => ({
                                                  ...current,
                                                  meetingDraft: { ...normalizeMeetingDraft(current.meetingDraft), endDate: e.target.value },
                                                }))
                                              : undefined
                                          }
                                        />
                                      </div>
                                      <div className="md:col-span-2">
                                        <label className="mb-1 block text-sm font-medium text-violet-950">End Time</label>
                                        <input
                                          type="time"
                                          className="w-full rounded-lg border border-violet-200 bg-white px-3 py-2 text-sm text-gray-900"
                                          value={previewMeetingDraft?.endTime ?? ""}
                                          onChange={(e) =>
                                            builderItem
                                              ? updateRow(builderItem.id, (current) => ({
                                                  ...current,
                                                  meetingDraft: { ...normalizeMeetingDraft(current.meetingDraft), endTime: e.target.value },
                                                }))
                                              : undefined
                                          }
                                        />
                                      </div>
                                    </>
                                  ) : null}
                                  <div className="md:col-span-2 flex items-center gap-2 text-sm text-violet-950">
                                    <input
                                      type="checkbox"
                                      checked={Boolean(previewMeetingDraft?.isAllDay)}
                                      onChange={(e) =>
                                        builderItem
                                          ? updateRow(builderItem.id, (current) => ({
                                              ...current,
                                              meetingDraft: { ...normalizeMeetingDraft(current.meetingDraft), isAllDay: e.target.checked },
                                            }))
                                          : undefined
                                      }
                                    />
                                    <span>All day meeting</span>
                                  </div>
                                </>
                              ) : null}
                              <div className="md:col-span-2 flex items-center gap-2 text-sm text-violet-950">
                                <input
                                  type="checkbox"
                                  checked={Boolean(previewMeetingDraft?.teamsMeeting)}
                                  onChange={(e) =>
                                    builderItem
                                      ? updateRow(builderItem.id, (current) => ({
                                          ...current,
                                          meetingDraft: {
                                            ...normalizeMeetingDraft(current.meetingDraft),
                                            teamsMeeting: e.target.checked,
                                            location: e.target.checked
                                              ? TEAMS_MEETING_LOCATION
                                              : normalizeMeetingDraft(current.meetingDraft)?.location ?? "",
                                          },
                                        }))
                                      : undefined
                                  }
                                />
                                <span>Microsoft Teams Meeting</span>
                              </div>
                              {isPreviewTeamsMeetingEnabled ? (
                                <div className="md:col-span-2 rounded-xl border border-violet-200 bg-white p-4">
                                  <div className="text-sm font-semibold text-violet-950">Teams Meeting Details</div>
                                  <div className="mt-3 space-y-3 text-sm text-gray-700">
                                    <div>
                                      <div className="mb-1 font-medium text-violet-950">Join link</div>
                                      <div className="text-gray-500">
                                        Teams join info will appear here after the Outlook event is created.
                                      </div>
                                    </div>
                                  </div>
                                </div>
                              ) : null}
                            </div>
                          </div>
                        ) : null}

                        {rowKind === "email" && isEmailExpanded ? (
                          <div className="bg-amber-50 px-4 py-3">
                            <div className="font-medium text-amber-950">Subject: {previewEmailSubject}</div>
                            <div className="mt-2 text-sm font-medium text-amber-900">
                              {getPreviewEmailModeMessage(appSettings.emailHandlingMode)}
                            </div>
                            <div className="mt-2 grid grid-cols-1 gap-3">
                              <div>
                                <label className="mb-1 block text-sm font-medium text-amber-950">To</label>
                                <EmailTokensInput
                                  label=""
                                  values={previewEmailDraft.to}
                                  onChange={(nextValues) =>
                                    builderItem
                                      ? updateRow(builderItem.id, (current) => ({
                                          ...current,
                                          emailDraft: { ...normalizeEmailDraft(current.emailDraft), to: nextValues },
                                        }))
                                      : undefined
                                  }
                                  placeholder="Type an email and press Enter or comma"
                                />
                              </div>
                              <div>
                                <label className="mb-1 block text-sm font-medium text-amber-950">CC</label>
                                <EmailTokensInput
                                  label=""
                                  values={previewEmailDraft.cc}
                                  onChange={(nextValues) =>
                                    builderItem
                                      ? updateRow(builderItem.id, (current) => ({
                                          ...current,
                                          emailDraft: { ...normalizeEmailDraft(current.emailDraft), cc: nextValues },
                                        }))
                                      : undefined
                                  }
                                  placeholder="Add CC emails"
                                />
                              </div>
                              <div>
                                <label className="mb-1 block text-sm font-medium text-amber-950">BCC</label>
                                <EmailTokensInput
                                  label=""
                                  values={previewEmailDraft.bcc}
                                  onChange={(nextValues) =>
                                    builderItem
                                      ? updateRow(builderItem.id, (current) => ({
                                          ...current,
                                          emailDraft: { ...normalizeEmailDraft(current.emailDraft), bcc: nextValues },
                                        }))
                                      : undefined
                                  }
                                  placeholder="Add BCC emails"
                                />
                              </div>
                              <div>
                                <label className="mb-1 block text-sm font-medium text-amber-950">Subject</label>
                                <input
                                  className="w-full rounded-lg border border-amber-200 bg-white px-3 py-2 text-sm text-gray-900"
                                  value={previewEmailDraft.subject}
                                  onChange={(e) =>
                                    builderItem
                                      ? updateRow(builderItem.id, (current) => ({
                                          ...current,
                                          emailDraft: { ...normalizeEmailDraft(current.emailDraft), subject: e.target.value },
                                        }))
                                      : undefined
                                  }
                                />
                              </div>
                              <div>
                                <label className="mb-1 block text-sm font-medium text-amber-950">Message Body</label>
                                <textarea
                                  rows={5}
                                  className="w-full rounded-lg border border-amber-200 bg-white px-3 py-2 text-sm text-gray-900"
                                  value={previewEmailDraft.body}
                                  onChange={(e) =>
                                    builderItem
                                      ? updateRow(builderItem.id, (current) => ({
                                          ...current,
                                          emailDraft: { ...normalizeEmailDraft(current.emailDraft), body: e.target.value },
                                        }))
                                      : undefined
                                  }
                                />
                              </div>
                              {previewEmailFinalBody ? (
                                <div>
                                  <label className="mb-1 block text-sm font-medium text-amber-950">Final Email Body</label>
                                  <pre className="whitespace-pre-wrap rounded-lg border border-amber-200 bg-white px-3 py-2 text-sm text-gray-900">
                                    {previewEmailFinalBody}
                                  </pre>
                                </div>
                              ) : null}
                            </div>
                          </div>
                        ) : null}
                      </div>
                    );
                  })}
                    </div>
                  </div>

                  <div className="flex flex-col gap-4">
                    <div className="max-w-md">
                      <label className="mb-1 block text-sm font-medium text-gray-700">Weekend handling</label>
                      <select
                        className="w-full rounded-lg border px-3 py-2"
                        value={weekendRule}
                        onChange={(e) => setWeekendRule(e.target.value as WeekendRule)}
                      >
                        <option value="prior_business_day">Adjust to prior business day (Fri)</option>
                        <option value="none">Allow weekends (no adjustment)</option>
                      </select>
                      <p className="mt-1 text-xs text-gray-500">
                        If a computed date lands on Sat/Sun, we either move it to Friday or leave it as-is.
                      </p>
                    </div>

                    <div className="flex flex-wrap items-center gap-3">
                      <button
                        onClick={() => {
                          void exportCurrentPlan({ skipValidation: true, skipConfirm: true });
                        }}
                        className="rounded-lg bg-blue-600 px-4 py-2 text-white hover:bg-blue-700"
                      >
                        Export
                      </button>
                      {executionNotice?.tone === "success" ? <ExportDoneBadge /> : null}
                    </div>
                  </div>
                </>
              ) : (
                <div className="rounded-xl border border-dashed px-4 py-8 text-sm text-gray-500">
                  {noEventDate
                    ? "Preview is not ready for the current plan yet."
                    : "Add an event date to preview the current plan schedule."}
                </div>
              )}
            </div>
          </div>
        </div>
      ) : null}
    </div>
  );
}
