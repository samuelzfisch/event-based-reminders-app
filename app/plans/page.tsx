"use client";

import Link from "next/link";
import { useEffect, useMemo, useRef, useState } from "react";

import {
  APP_SETTINGS_UPDATED_EVENT,
  areAppSettingsEqual,
  hydrateAppSettingsFromSupabase,
  loadAppSettings,
  type AppSettings,
  type EmailHandlingMode,
} from "../../lib/appSettings";
import { todayYYYYMMDD } from "../../lib/dateUtils";
import { writeExecutionHistory } from "../../lib/executionHistory";
import type { ExecutionHistoryProviderObjectType } from "../../lib/executionHistory";
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
import {
  createGmailDraftFromEmailDraft,
  createGoogleCalendarEvent,
  GMAIL_COMPOSE_SCOPE,
  GMAIL_CONNECTION_UPDATED_EVENT,
  GOOGLE_CALENDAR_EVENTS_SCOPE,
  getConnectedGmailMailboxEmail,
  getGmailConnectionState,
  sendGmailEmailFromEmailDraft,
  type GmailConnectionState,
  resolveGmailConnectionState,
} from "../../lib/gmailClient";
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
import type {
  AIChatMessage,
  AIPlanBuilderContext,
  AIPlanChatRequest,
  AIPlanChatTurnResult,
  AIPlanDraft,
} from "../../lib/aiPlanGeneration";
import type { Plan, PlanDateBasis, PlanRowType, PlanType, WeekendRule } from "../../types/plan";
import { useAuthContext } from "../components/auth-provider";

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
    addGoogleMeet?: boolean;
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

type ProviderExecutionResult = {
  provider: "outlook" | "gmail";
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

type ProviderExecutionAvailability = {
  provider: "outlook" | "gmail";
  canExecute: boolean;
  reason?: string;
  outlookAvailable: boolean;
  gmailAvailable: boolean;
};

type ExecutionSnapshotRowDefinition = {
  id: string;
  title: string;
  body: string;
  offsetDays: number;
  dateBasis: PlanDateBasis;
  rowType: PlanRowType;
  reminderTime: string;
  emailDraft: ReturnType<typeof normalizeEmailDraft> | null;
  durationDraft: BuilderRow["durationDraft"] | null;
  meetingDraft: ReturnType<typeof normalizeMeetingDraft> | null;
};

type ExecutionNotice = {
  tone: "pending" | "success" | "mixed" | "warning";
  title: string;
  message?: string;
  details?: string[];
};

type ExecutionNoticeEntry = {
  id: string;
  notice: ExecutionNotice;
};

type AIConversationMessage = {
  id: string;
  role: "user" | "assistant";
  text: string;
  summary?: string;
  status?: "needs_more_info" | "ready_to_apply";
  followUpQuestions?: string[];
  changeSummary?: string[];
  confidenceNote?: string;
  suggestedNextActions?: string[];
  starterPrompts?: string[];
  modeOptions?: Array<{ id: "refine_current" | "start_new"; label: string }>;
};

type AIPlanningSessionBackup = {
  messages: AIConversationMessage[];
  summary: string;
  draft: AIPlanDraft | null;
  status: "needs_more_info" | "ready_to_apply";
  changeSummary: string[];
  confidenceNote: string;
  suggestedNextActions: string[];
  builderContextMode: "refine_current" | "start_new" | null;
  baseline: AIDraftBaseline | null;
  sessionSource: AISessionSource | null;
};

type AIDraftBaseline = {
  sourceLabel: string;
  planType: PlanType;
  noEventDate: boolean;
  anchorDate: string;
  eventTime: string;
  weekendRule: WeekendRule;
  totalRows: number;
  reminderCount: number;
  emailCount: number;
  meetingCount: number;
};

type AISessionSource =
  | { type: "new" }
  | { type: "current_builder" }
  | { type: "saved_template"; name: string }
  | { type: "branched_draft" };

type BuilderSourceProvenance = {
  sourceType: "manual" | "saved_template" | "ai_draft";
  sourceLabel: string;
  loadedAt: string;
  sourceSignature: string;
  hadMissingDetails?: boolean;
};

type BuilderSourceSeed = {
  sourceType: BuilderSourceProvenance["sourceType"];
  sourceLabel: string;
  hadMissingDetails?: boolean;
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
const GOOGLE_MEET_LOCATION = "Google Meet";
const MEETING_DURATION_OPTIONS = [
  { value: "30", label: "30 minutes" },
  { value: "60", label: "60 minutes" },
  { value: "90", label: "90 minutes" },
  { value: "120", label: "2 hours" },
  { value: "custom", label: "Custom" },
] as const;
const AI_STARTER_PROMPTS = [
  "I need a plan for an earnings call with prep reminders and follow-up emails.",
  "Help me make a conference follow-up plan.",
  "I want a press release timeline with internal review reminders.",
  "Build a workflow for a board meeting with prep tasks, a draft email, and day-of reminders.",
] as const;
const AI_ENABLED = false;

function makeId(prefix: string) {
  return `${prefix}_${Math.random().toString(36).slice(2, 10)}`;
}

function getAiReadinessLabel(status: "needs_more_info" | "ready_to_apply") {
  return status === "ready_to_apply" ? "Draft ready to apply" : "Needs a bit more information";
}

function getAiBuilderContextRowLabel(rowType: PlanRowType) {
  if (rowType === "email") return "Email";
  if (rowType === "calendar_event") return "Meeting";
  return "Reminder";
}

function getBaselineDeltaLabel(delta: number) {
  if (delta === 0) return "unchanged";
  return delta > 0 ? `+${delta}` : `${delta}`;
}

function getAiSessionSourceDetails(source: AISessionSource | null) {
  if (!source || source.type === "new") {
    return {
      label: "New AI draft",
      note: "Nothing changes in the builder unless you click Apply to Builder or save a new template.",
      classes: "border-gray-200 bg-gray-50 text-gray-900",
    };
  }
  if (source.type === "current_builder") {
    return {
      label: "Current builder plan",
      note: "Your builder stays unchanged until you review the draft and click Apply to Builder.",
      classes: "border-blue-200 bg-blue-50 text-blue-900",
    };
  }
  if (source.type === "saved_template") {
    return {
      label: `Saved template: ${source.name}`,
      note: "The saved template stays unchanged unless you save a new template later.",
      classes: "border-purple-200 bg-purple-50 text-purple-900",
    };
  }
  return {
    label: "Branched AI draft",
    note: "This branch will not affect your earlier draft unless you restore, apply, or save it explicitly.",
    classes: "border-amber-200 bg-amber-50 text-amber-900",
  };
}

function getAiDraftIdentityDetails(source: AISessionSource | null) {
  if (!source || source.type === "new") {
    return {
      source: "New AI draft",
      currentDraft: "AI working draft",
      applyDestination: "Current builder",
      saveDestination: "New custom template",
      note: "Applying updates the builder only. Saving creates a new reusable template.",
    };
  }
  if (source.type === "current_builder") {
    return {
      source: "Current builder plan",
      currentDraft: "AI working draft",
      applyDestination: "Current builder",
      saveDestination: "New custom template",
      note: "Your builder stays unchanged until Apply to Builder. Saving creates a separate template.",
    };
  }
  if (source.type === "saved_template") {
    return {
      source: `Saved template: ${source.name}`,
      currentDraft: "AI working draft",
      applyDestination: "Current builder",
      saveDestination: "New custom template",
      note: "The original template stays unchanged. Saving creates a new template from this draft.",
    };
  }
  return {
    source: "Branched AI draft",
    currentDraft: "AI working draft",
    applyDestination: "Current builder",
    saveDestination: "New custom template",
    note: "This branch is separate from the earlier draft unless you explicitly apply or save it.",
  };
}

function getAiDraftStageDetails(options: {
  hasDraft: boolean;
  readiness: "needs_more_info" | "ready_to_apply";
  hasFollowUpQuestions: boolean;
  wasSavedAsTemplate: boolean;
  source: AISessionSource | null;
  rowCount: number;
}) {
  if (!options.hasDraft) {
    return {
      stage: "Exploring",
      nextStep: "Tell the assistant what you want to plan so it can shape a first draft.",
    };
  }

  if (options.hasFollowUpQuestions || options.readiness === "needs_more_info") {
    return {
      stage: "Needs clarification",
      nextStep: "Answer the remaining questions so the draft can tighten up.",
    };
  }

  if (!options.wasSavedAsTemplate && options.rowCount >= 2 && options.source?.type !== "saved_template") {
    return {
      stage: "Good template candidate",
      nextStep: "Save this as a template if you expect to reuse this workflow.",
    };
  }

  return {
    stage: "Ready to apply",
    nextStep: options.wasSavedAsTemplate
      ? "Review the rows, then apply to builder if you want to use this version now."
      : "Review the rows, then apply to builder when you’re ready.",
  };
}

function getAiDraftMissingDetails(options: {
  draft: AIPlanDraft | null;
  confidenceNote: string;
  hasFollowUpQuestions: boolean;
  source: AISessionSource | null;
}) {
  if (!options.draft) return [];

  const details = new Set<string>();
  const draft = options.draft;

  if (!draft.noEventDate && !(draft.anchorDate ?? "").trim()) {
    details.add("Event date not specified.");
  }

  if (draft.rows.some((row) => (row.rowType === "reminder" || row.rowType === "calendar_event") && !row.reminderTime?.trim())) {
    details.add("Reminder or meeting timing is still general.");
  }

  if (
    draft.rows.some(
      (row) =>
        row.rowType === "email" &&
        (!row.emailDraft || (row.emailDraft.to.length === 0 && row.emailDraft.cc.length === 0 && row.emailDraft.bcc.length === 0))
    )
  ) {
    details.add("Email recipients still need confirmation.");
  }

  if (
    draft.rows.some(
      (row) =>
        row.rowType === "calendar_event" &&
        (!row.meetingDraft || row.meetingDraft.attendees.length === 0)
    )
  ) {
    details.add("Meeting attendees may still need confirmation.");
  }

  if (draft.weekendRule === "prior_business_day" && options.source?.type !== "saved_template") {
    details.add("Weekend handling may need review.");
  }

  if (options.hasFollowUpQuestions) {
    details.add("A few details still need clarification from you.");
  }

  const normalizedConfidence = options.confidenceNote.toLowerCase();
  if (normalizedConfidence.includes("assumed") || normalizedConfidence.includes("needs confirmation")) {
    details.add(options.confidenceNote);
  }

  return Array.from(details).slice(0, 5);
}

function buildBuilderContentSignature(input: {
  planType: PlanType;
  templateName: string;
  eventName: string;
  anchorDate: string;
  noEventDate: boolean;
  weekendRule: WeekendRule;
  anchors: BuilderAnchor[];
  rows: BuilderRow[];
}) {
  return JSON.stringify({
    planType: input.planType,
    templateName: input.templateName.trim(),
    eventName: input.eventName.trim(),
    anchorDate: input.anchorDate,
    noEventDate: input.noEventDate,
    weekendRule: input.weekendRule,
    anchors: input.anchors.map((anchor) => ({
      key: anchor.key.trim(),
      value: anchor.value,
    })),
    rows: input.rows.map((row) => {
      const emailDraft = normalizeEmailDraft(row.emailDraft);
      const meetingDraft = normalizeMeetingDraft(row.meetingDraft);
      return {
        title: row.title.trim(),
        body: row.body ?? "",
        offsetDays: row.offsetDays ?? 0,
        dateBasis: row.dateBasis ?? "event",
        rowType: row.rowType,
        reminderTime: row.reminderTime ?? "",
        emailDraft,
        durationDraft: row.durationDraft ?? null,
        meetingDraft: meetingDraft ?? null,
      };
    }),
  });
}

function getBuilderSourceTimestamp() {
  return new Date().toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
}

function buildBuilderSourceProvenance(
  snapshot: Pick<BuilderStateSnapshot, "planType" | "templateName" | "eventName" | "anchorDate" | "noEventDate" | "weekendRule" | "anchors" | "rows">,
  source: BuilderSourceSeed
): BuilderSourceProvenance {
  return {
    sourceType: source.sourceType,
    sourceLabel: source.sourceLabel,
    loadedAt: getBuilderSourceTimestamp(),
    sourceSignature: buildBuilderContentSignature({
      planType: snapshot.planType,
      templateName: snapshot.templateName,
      eventName: snapshot.eventName,
      anchorDate: snapshot.anchorDate,
      noEventDate: snapshot.noEventDate,
      weekendRule: snapshot.weekendRule,
      anchors: snapshot.anchors,
      rows: snapshot.rows,
    }),
    hadMissingDetails: source.hadMissingDetails,
  };
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
    notice.tone === "pending"
      ? "border-blue-200 bg-blue-50 text-blue-950"
      : notice.tone === "success"
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
    addGoogleMeet: Boolean(draft.addGoogleMeet),
  };
}

function getNextMeetingLocationOnToggle(params: {
  currentLocation?: string;
  checked: boolean;
  enabledLocation: string;
  disabledLocation: string;
}) {
  const currentLocation = typeof params.currentLocation === "string" ? params.currentLocation.trim() : "";
  if (params.checked) return params.enabledLocation;
  return currentLocation === params.enabledLocation ? params.disabledLocation : params.currentLocation ?? "";
}

function getMeetingLocationValue(
  draft: BuilderRow["meetingDraft"] | null | undefined,
  activeProvider: "outlook" | "gmail" | null
) {
  const normalizedDraft = normalizeMeetingDraft(draft);
  if (!normalizedDraft) return "";
  if (normalizedDraft.teamsMeeting) return TEAMS_MEETING_LOCATION;
  if (activeProvider === "gmail" && normalizedDraft.addGoogleMeet) return GOOGLE_MEET_LOCATION;

  const trimmedLocation = normalizedDraft.location.trim();
  if (trimmedLocation === TEAMS_MEETING_LOCATION || trimmedLocation === GOOGLE_MEET_LOCATION) {
    return "";
  }

  return normalizedDraft.location ?? "";
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

function hasAnchorToken(value: unknown): boolean {
  if (typeof value === "string") {
    return /\[[^\]]+\]/.test(value);
  }
  if (Array.isArray(value)) {
    return value.some((entry) => hasAnchorToken(entry));
  }
  if (value && typeof value === "object") {
    return Object.values(value).some((entry) => hasAnchorToken(entry));
  }
  return false;
}

function normalizeExecutionSnapshotRow(row: BuilderRow): ExecutionSnapshotRowDefinition {
  return {
    id: row.id,
    title: row.title,
    body: row.body ?? "",
    offsetDays: row.offsetDays ?? 0,
    dateBasis: row.dateBasis ?? "event",
    rowType: row.rowType ?? "reminder",
    reminderTime: row.reminderTime ?? "",
    emailDraft: row.rowType === "email" ? normalizeEmailDraft(row.emailDraft) : null,
    durationDraft: row.durationDraft ?? null,
    meetingDraft: row.meetingDraft ? normalizeMeetingDraft(row.meetingDraft) : null,
  };
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
    if (offsetDays < 0) return `-${absoluteDays} ${dayLabel} from today`;
    if (offsetDays > 0) return `+${absoluteDays} ${dayLabel} from today`;
    return "today";
  }
  if (offsetDays < 0) return `-${absoluteDays} ${dayLabel} before event`;
  if (offsetDays > 0) return `+${absoluteDays} ${dayLabel} after event`;
  return "day of event";
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

function buildFinalEmailBody(body: string, signatureSettings: { signature: string }) {
  const normalizedBody = body.trim();
  const normalizedSignature = signatureSettings.signature.trim();
  if (normalizedSignature) {
    return normalizedBody ? `${normalizedBody}\n\n${normalizedSignature}` : normalizedSignature;
  }
  return normalizedBody;
}

function getBuilderEmailModeMessage(mode: EmailHandlingMode) {
  if (mode === "schedule") {
    return "Outlook emails will be scheduled to send at the specified date and time. Gmail will save a draft instead.";
  }
  if (mode === "send") {
    return "Email will be sent immediately";
  }
  return "Email will be saved to your Drafts";
}

function getPreviewEmailModeMessage(mode: EmailHandlingMode) {
  if (mode === "schedule") {
    return "Outlook emails will be scheduled to send at the specified date and time. Gmail will save a draft instead.";
  }
  if (mode === "send") {
    return "Emails will be sent immediately.";
  }
  return "Emails will be saved to Drafts.";
}

function getPreviewEmailActionLabel(mode: EmailHandlingMode) {
  if (mode === "schedule") return "Schedule Email (Outlook only)";
  if (mode === "send") return "Send Email";
  return "Save to Drafts";
}

function getEffectivePreviewItemDate(item: Plan["items"][number]) {
  return item.customDueDate ?? item.dueDate;
}

function getBuilderRowTypeMeta(row: BuilderRow) {
  const rowKind = classifyPlanRow(row);
  if (rowKind === "email") {
    return {
      label: "Email",
      className: "text-green-600",
      badgeClass: "border-green-200 bg-green-50 text-green-700",
      borderClass: "border-l-green-400",
      timingPanelClass: "border-green-100 bg-green-50/60",
      timelineDotClass: "bg-green-500",
    };
  }
  if (rowKind === "meeting") {
    return {
      label: "Meeting",
      className: "text-violet-600",
      badgeClass: "border-violet-200 bg-violet-50 text-violet-700",
      borderClass: "border-l-violet-400",
      timingPanelClass: "border-violet-100 bg-violet-50/60",
      timelineDotClass: "bg-violet-500",
    };
  }
  if (row.rowType === "calendar_event") {
    return {
      label: "Calendar Event",
      className: "text-violet-600",
      badgeClass: "border-violet-200 bg-violet-50 text-violet-700",
      borderClass: "border-l-violet-400",
      timingPanelClass: "border-violet-100 bg-violet-50/60",
      timelineDotClass: "bg-violet-500",
    };
  }
  return {
    label: "Reminder",
    className: "text-blue-600",
    badgeClass: "border-blue-200 bg-blue-50 text-blue-700",
    borderClass: "border-l-blue-400",
    timingPanelClass: "border-blue-100 bg-blue-50/60",
    timelineDotClass: "bg-blue-500",
  };
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

function areGmailConnectionStatesEqual(left: GmailConnectionState | null, right: GmailConnectionState | null) {
  return JSON.stringify(left) === JSON.stringify(right);
}

export default function PlansPage() {
  const { authEnabled, currentUser, currentOrgId, loading, refreshAuthContext } = useAuthContext();
  const initialSettings = loadAppSettings();
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
  const [gmailConnection, setGmailConnection] = useState<GmailConnectionState | null>(() => getGmailConnectionState());
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
  const [lastBuilderSourceProvenance, setLastBuilderSourceProvenance] = useState<BuilderSourceProvenance | null>(null);
  const [guidedForm, setGuidedForm] = useState<GuidedFormState>(() => createEmptyGuidedForm());
  const [isTemplateManageMode, setIsTemplateManageMode] = useState(false);
  const [isTemplateCopyMode, setIsTemplateCopyMode] = useState(false);
  const [areAnchorsHidden, setAreAnchorsHidden] = useState(true);
  const [templateActionMessage, setTemplateActionMessage] = useState("");
  const [openMenuRowId, setOpenMenuRowId] = useState<string | null>(null);
  const [openEmailDraftRowId, setOpenEmailDraftRowId] = useState<string | null>(null);
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
  const [executionNotices, setExecutionNotices] = useState<ExecutionNoticeEntry[]>([]);
  const [executionState, setExecutionState] = useState<"pending" | "success" | "failure" | null>(null);
  const exportQueueRef = useRef(Promise.resolve());
  const activeExportCountRef = useRef(0);
  const [providerLoading, setProviderLoading] = useState({ outlook: true, gmail: true });
  const [previewLoading, setPreviewLoading] = useState(true);
  const [isAiPanelOpen, setIsAiPanelOpen] = useState(false);
  const [aiComposer, setAiComposer] = useState("");
  const [aiGenerating, setAiGenerating] = useState(false);
  const [aiChatError, setAiChatError] = useState<string | null>(null);
  const [aiChatMessages, setAiChatMessages] = useState<AIConversationMessage[]>([]);
  const [aiChatSummary, setAiChatSummary] = useState("");
  const [aiChatDraft, setAiChatDraft] = useState<AIPlanDraft | null>(null);
  const [aiChatStatus, setAiChatStatus] = useState<"needs_more_info" | "ready_to_apply">("needs_more_info");
  const [aiChatChangeSummary, setAiChatChangeSummary] = useState<string[]>([]);
  const [aiChatConfidenceNote, setAiChatConfidenceNote] = useState("");
  const [aiChatSuggestedNextActions, setAiChatSuggestedNextActions] = useState<string[]>([]);
  const [aiBuilderContextMode, setAiBuilderContextMode] = useState<"refine_current" | "start_new" | null>(null);
  const [aiSessionBackup, setAiSessionBackup] = useState<AIPlanningSessionBackup | null>(null);
  const [aiDraftBaseline, setAiDraftBaseline] = useState<AIDraftBaseline | null>(null);
  const [aiSessionSource, setAiSessionSource] = useState<AISessionSource | null>(null);
  const [showAiTemplateSaveDialog, setShowAiTemplateSaveDialog] = useState(false);
  const [aiTemplateNameDraft, setAiTemplateNameDraft] = useState("");
  const [aiTemplateSaveMessage, setAiTemplateSaveMessage] = useState<string | null>(null);
  const [aiSavedTemplateInfo, setAiSavedTemplateInfo] = useState<{ id: string; name: string } | null>(null);
  const [showBuilderTemplateSaveDialog, setShowBuilderTemplateSaveDialog] = useState(false);
  const [builderTemplateNameDraft, setBuilderTemplateNameDraft] = useState("");
  const [builderTemplateSaveMessage, setBuilderTemplateSaveMessage] = useState<string | null>(null);
  const [highlightedTemplateId, setHighlightedTemplateId] = useState<string | null>(null);
  const [showAiApplyConfirm, setShowAiApplyConfirm] = useState(false);
  const [aiApplySuccessMessage, setAiApplySuccessMessage] = useState<string | null>(null);
  const [builderSourceProvenance, setBuilderSourceProvenance] = useState<BuilderSourceProvenance | null>(null);
  const [hasMounted, setHasMounted] = useState(false);
  const [hasHydratedTemplates, setHasHydratedTemplates] = useState(false);
  const hasLocalTemplateMutationRef = useRef(false);
  const kebabMenuRef = useRef<HTMLDivElement | null>(null);
  const builderTimeInputRefs = useRef<Record<string, HTMLInputElement | HTMLTextAreaElement | null>>({});
  const aiConversationRef = useRef<HTMLDivElement | null>(null);
  const aiComposerRef = useRef<HTMLTextAreaElement | null>(null);
  const builderSectionRef = useRef<HTMLElement | null>(null);
  const rowsRef = useRef<BuilderRow[]>(rows);

  function getLatestPreviewPlan() {
    return resolvePlanAnchors(
      createPlan({
        name: effectivePlanName,
        type: planType,
        anchorDate: previewAnchorDateForComputation,
        weekendRule,
        template: buildTemplateItemsFromRows(rowsRef.current),
      }),
      anchorMap
    );
  }

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
    rowsRef.current = rows;
  }, [rows]);

  useEffect(() => {
    if (!isAiPanelOpen) return;
    aiConversationRef.current?.scrollTo({
      top: aiConversationRef.current.scrollHeight,
      behavior: "smooth",
    });
  }, [aiChatMessages, isAiPanelOpen]);

  useEffect(() => {
    if (!aiApplySuccessMessage) return;
    const timeoutId = window.setTimeout(() => {
      setAiApplySuccessMessage(null);
    }, 5000);
    return () => window.clearTimeout(timeoutId);
  }, [aiApplySuccessMessage]);

  useEffect(() => {
    if (!highlightedTemplateId) return;
    const timeoutId = window.setTimeout(() => {
      setHighlightedTemplateId(null);
    }, 5000);
    return () => window.clearTimeout(timeoutId);
  }, [highlightedTemplateId]);

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
      setPreviewLoading(true);
      const seedState = buildSeedTemplates(
        loadAppSettings().defaultReminderTime,
        loadAppSettings().defaultPressReleaseTime
      ).map((template, index) => toPersistedTemplate(template, index));
      const remoteState = await loadTemplateStateFromSupabase(seedState);
      const nextState = remoteState ?? loadCachedTemplateState(seedState);

      if (!active) return;
      if (hasLocalTemplateMutationRef.current) {
        setHasHydratedTemplates(true);
        setPreviewLoading(false);
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
      setPreviewLoading(false);
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
      setProviderLoading((current) => ({ ...current, outlook: true }));
      const outlookConnectionState = await resolveOutlookConnectionState(appSettings.outlookAccountEmail);
      if (!active) return;
      setOutlookConnection((current) =>
        areOutlookConnectionStatesEqual(current, outlookConnectionState) ? current : outlookConnectionState
      );
      setProviderLoading((current) => ({ ...current, outlook: false }));
    }

    void refreshConnection();

    function handleConnectionRefresh() {
      void refreshConnection();
    }

    window.addEventListener(OUTLOOK_CONNECTION_UPDATED_EVENT, handleConnectionRefresh as EventListener);
    return () => {
      active = false;
      window.removeEventListener(OUTLOOK_CONNECTION_UPDATED_EVENT, handleConnectionRefresh as EventListener);
    };
  }, [appSettings.outlookAccountEmail]);

  useEffect(() => {
    let active = true;

    async function refreshConnection() {
      setProviderLoading((current) => ({ ...current, gmail: true }));
      const gmailConnectionState = await resolveGmailConnectionState();
      if (!active) return;
      setGmailConnection((current) => (areGmailConnectionStatesEqual(current, gmailConnectionState) ? current : gmailConnectionState));
      setProviderLoading((current) => ({ ...current, gmail: false }));
    }

    void refreshConnection();

    function handleConnectionRefresh() {
      void refreshConnection();
    }

    window.addEventListener(GMAIL_CONNECTION_UPDATED_EVENT, handleConnectionRefresh as EventListener);
    return () => {
      active = false;
      window.removeEventListener(GMAIL_CONNECTION_UPDATED_EVENT, handleConnectionRefresh as EventListener);
    };
  }, []);

  const normalizedDefaultReminderTime = normalizeReminderTimeInput(appSettings.defaultReminderTime);
  const normalizedDefaultPressReleaseTime = normalizeReminderTimeInput(appSettings.defaultPressReleaseTime);
  const currentSelectedTemplate = useMemo(
    () => (selectedTemplateId ? savedTemplates.find((template) => template.id === selectedTemplateId) ?? null : null),
    [savedTemplates, selectedTemplateId]
  );
  const effectiveTemplateMode = useMemo(
    () => (currentSelectedTemplate ? inferTemplateMode(currentSelectedTemplate) : builderMode),
    [builderMode, currentSelectedTemplate]
  );

  const effectivePlanName = useMemo(
    () =>
      effectiveTemplateMode === "template" &&
      (planType === "press_release" || planType === "earnings" || planType === "conference")
        ? getGuidedTemplateDisplayName(planType, guidedForm) || eventName || templateName || getSeedTemplateName(planType)
        : eventName || templateName || getSeedTemplateName(planType),
    [effectiveTemplateMode, eventName, guidedForm, planType, templateName]
  );
  const previewEffectiveAnchorDate = noEventDate ? todayYYYYMMDD() : anchorDate;
  const previewAnchorDateForComputation = previewEffectiveAnchorDate || todayYYYYMMDD();
  const genericEventAnchorName = eventName.trim();
  const genericEventAnchorDate = noEventDate ? "" : anchorDate;
  const resolvedAnchors = useMemo(
    () =>
      anchors.map((anchor) => {
        const resolvedValue =
          getDerivedAnchorValue(planType, effectivePlanName, anchorDate, anchor.key, guidedForm, {
            eventNameValue: genericEventAnchorName,
            eventDateValue: genericEventAnchorDate,
          }) ?? anchor.value;
        return {
          id: anchor.id,
          key: anchor.key,
          value: resolvedValue,
          displayValue: getAnchorDisplayValue(anchor.key, resolvedValue),
          locked: anchor.locked,
        };
      }),
    [anchorDate, anchors, effectivePlanName, genericEventAnchorDate, genericEventAnchorName, guidedForm, planType]
  );
  const anchorMap = useMemo(() => buildAnchorMap(resolvedAnchors), [resolvedAnchors]);
  const previewTemplate = useMemo(() => buildTemplateItemsFromRows(rows), [rows]);
  const previewPlan = useMemo(
    () =>
      resolvePlanAnchors(
        createPlan({
          name: effectivePlanName,
          type: planType,
          anchorDate: previewAnchorDateForComputation,
          weekendRule,
          template: previewTemplate,
        }),
        anchorMap
      ),
    [anchorMap, effectivePlanName, planType, previewAnchorDateForComputation, previewTemplate, weekendRule]
  );
  const previewPlanForRender = useMemo(
    () => (previewEffectiveAnchorDate ? previewPlan : null),
    [previewEffectiveAnchorDate, previewPlan]
  );
  const selectedEditableTemplate =
    currentSelectedTemplate && !isProtectedTemplate(currentSelectedTemplate) ? currentSelectedTemplate : null;
  const shouldRenderSimpleEventHeaderFields = !(
    effectiveTemplateMode === "template" &&
    (planType === "press_release" || planType === "earnings" || planType === "conference")
  );

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
      rowsRef.current = nextRows;
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

  function getPastScheduledItemsMessage(itemIds?: string[]) {
    const now = Date.now();
    const scopedItems = itemIds ? previewPlan.items.filter((item) => itemIds.includes(item.id)) : previewPlan.items;

    const pastItems = scopedItems
      .map((item) => {
        const rowKind = classifyPlanRow(item);

        if (rowKind === "reminder") {
          const dueDate = getEffectivePreviewItemDate(item);
          const reminderTime = getUsableReminderTime(item.reminderTime, anchorMap);
          const scheduledAtISO = dueDate && reminderTime ? buildLocalDateTimeIso(dueDate, reminderTime) : null;
          const scheduledAt = scheduledAtISO ? new Date(scheduledAtISO) : null;
          if (scheduledAt && !Number.isNaN(scheduledAt.getTime()) && scheduledAt.getTime() < now) {
            return `${item.customTitle ?? item.title ?? "Untitled reminder"} (reminder)`;
          }
        }

        if (rowKind === "meeting") {
          const timing = buildPreviewItemGraphTiming(item);
          const scheduledAt = new Date(timing.startISO);
          if (!Number.isNaN(scheduledAt.getTime()) && scheduledAt.getTime() < now) {
            return `${item.customTitle ?? item.title ?? "Untitled meeting"} (meeting)`;
          }
        }

        if (rowKind === "email" && appSettings.emailHandlingMode === "schedule") {
          const scheduledSendISO = getEmailScheduledSendISO(item);
          if (scheduledSendISO) {
            const scheduledAt = new Date(scheduledSendISO);
            if (!Number.isNaN(scheduledAt.getTime()) && scheduledAt.getTime() < now) {
              return `${(normalizeEmailDraft(item.emailDraft).subject || item.customTitle || item.title || "Untitled email").trim()} (email)`;
            }
          }
        }

        return null;
      })
      .filter((entry): entry is string => Boolean(entry));

    if (pastItems.length === 0) return null;
    return `This item is scheduled in the past.\n\nPlease update the date/time before continuing.\n\n${pastItems.join("\n")}`;
  }

  function warnIfPastScheduledItems(options?: { usePopup?: boolean; itemIds?: string[] }) {
    const message = getPastScheduledItemsMessage(options?.itemIds);
    if (!message) return false;
    if (options?.usePopup && typeof window !== "undefined") {
      window.alert(message);
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

  function getEmailValidationMessage(itemIds?: string[]) {
    const scopedItems = itemIds ? previewPlan.items.filter((item) => itemIds.includes(item.id)) : previewPlan.items;
    const missingTimeRows: string[] = [];
    const missingRecipientRows: string[] = [];

    for (const item of scopedItems) {
      if (classifyPlanRow(item) !== "email") continue;

      const draft = normalizeEmailDraft(item.emailDraft);
      const rowName = draft.subject.trim() || item.customTitle || item.title || "Untitled email";

      if ((appSettings.emailHandlingMode === "send" || appSettings.emailHandlingMode === "schedule") && draft.to.length === 0) {
        missingRecipientRows.push(rowName);
      }

      if (appSettings.emailHandlingMode === "schedule" && !getEmailScheduledSendISO(item)) {
        missingTimeRows.push(rowName);
      }
    }

    const sections: string[] = [];

    if (missingTimeRows.length > 0) {
      sections.push("Scheduled emails require a send time.");
      sections.push("Please add a time before continuing.");
      sections.push("");
      sections.push(...missingTimeRows.map((entry) => `- ${entry}`));
    }

    if (missingRecipientRows.length > 0) {
      if (sections.length > 0) sections.push("");
      sections.push("Please add at least one recipient before continuing.");
      sections.push("");
      sections.push(...missingRecipientRows.map((entry) => `- ${entry}`));
    }

    return sections.length > 0 ? sections.join("\n") : null;
  }

  function validateEmailRowsForExport(itemIds?: string[], options?: { usePopup?: boolean }) {
    const message = getEmailValidationMessage(itemIds);
    if (!message) return false;
    if (options?.usePopup && typeof window !== "undefined") {
      window.alert(message);
    }
    return true;
  }

  function getUserFixableEmailExecutionMessage(error: unknown) {
    const message = error instanceof Error ? error.message : String(error || "");
    const normalized = message.toLowerCase();

    if (normalized.includes("scheduled emails require a send time") || normalized.includes("add a time before scheduling this email")) {
      return "Scheduled emails require a send time.\n\nPlease add a time before continuing.";
    }

    if (
      normalized.includes("please add at least one recipient") ||
      normalized.includes("recipient") ||
      normalized.includes("recipients") ||
      normalized.includes("torecipients") ||
      normalized.includes("email address") ||
      normalized.includes("invalid address") ||
      normalized.includes("malformed")
    ) {
      return "Please add at least one recipient before continuing.";
    }

    return null;
  }

  function getResolvedEmailDraftForExecution(item: Plan["items"][number]) {
    const draft = normalizeEmailDraft(item.emailDraft);
    return {
      ...draft,
      body: buildFinalEmailBody(draft.body, {
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

  function getExecutionHistoryRowLinkage(item: Plan["items"][number]) {
    const sourceRow = rows.find((row) => row.id === item.id);
    const normalizedSourceRow = sourceRow ? normalizeExecutionSnapshotRow(sourceRow) : null;
    const rowKind = classifyPlanRow(item);
    return {
      sourceRowId: item.id,
      rowType: sourceRow?.rowType ?? item.rowType ?? "reminder",
      rowKind,
      dateBasis: sourceRow?.dateBasis ?? item.dateBasis ?? "event",
      offsetDays: sourceRow?.offsetDays ?? item.offsetDays ?? 0,
      reminderTimeMode: sourceRow?.reminderTime?.trim().startsWith("[") ? "anchor" : "fixed",
      anchorDerivedContent: {
        title: hasAnchorToken(sourceRow?.title),
        body: hasAnchorToken(sourceRow?.body),
        reminderTime: hasAnchorToken(sourceRow?.reminderTime),
        emailRecipients: hasAnchorToken(sourceRow?.emailDraft?.to) || hasAnchorToken(sourceRow?.emailDraft?.cc) || hasAnchorToken(sourceRow?.emailDraft?.bcc),
        emailSubject: hasAnchorToken(sourceRow?.emailDraft?.subject),
        emailBody: hasAnchorToken(sourceRow?.emailDraft?.body),
        meetingAttendees: hasAnchorToken(sourceRow?.meetingDraft?.attendees),
        meetingLocation: hasAnchorToken(sourceRow?.meetingDraft?.location),
      },
      originalRowDefinition: normalizedSourceRow,
      resolvedRowAtExecution: {
        title: item.title,
        customTitle: item.customTitle ?? "",
        body: item.body ?? "",
        dueDate: item.dueDate,
        rawDueDate: item.rawDueDate,
        customDueDate: item.customDueDate ?? null,
        reminderTime: item.reminderTime ?? "",
        emailDraft: item.emailDraft ?? null,
        durationDraft: item.durationDraft ?? null,
        meetingDraft: item.meetingDraft ?? null,
      },
    };
  }

  function getExecutionPlanSnapshot() {
    return {
      executionGroupPlanName: previewPlan.name,
      templateBaseType: planType,
      templateMode: effectiveTemplateMode,
      templateId: selectedTemplateId,
      templateName,
      eventName: effectivePlanName,
      anchorDate: previewPlan.anchorDate,
      noEventDate,
      weekendRule,
      anchorValues: resolvedAnchors.map((anchor) => ({
        id: anchor.id,
        key: anchor.key,
        value: anchor.value,
        displayValue: anchor.displayValue,
        locked: Boolean(anchor.locked),
      })),
      originalRowDefinitions: rows.map((row) => normalizeExecutionSnapshotRow(row)),
      resolvedItemsAtExecution: previewPlan.items.map((previewItem) => ({
        sourceRowId: previewItem.id,
        rowType: previewItem.rowType ?? "reminder",
        rowKind: classifyPlanRow(previewItem),
        offsetDays: previewItem.offsetDays,
        dateBasis: previewItem.dateBasis ?? "event",
        dueDate: previewItem.dueDate,
        rawDueDate: previewItem.rawDueDate,
        customDueDate: previewItem.customDueDate ?? null,
        reminderTime: previewItem.reminderTime ?? "",
        title: previewItem.title,
        customTitle: previewItem.customTitle ?? "",
        body: previewItem.body ?? "",
        emailDraft: previewItem.emailDraft ?? null,
        durationDraft: previewItem.durationDraft ?? null,
        meetingDraft: previewItem.meetingDraft ?? null,
        wasAdjusted: previewItem.wasAdjusted,
      })),
    };
  }

  function getExecutionHistoryCapabilities(options: {
    item: Plan["items"][number];
    status: "success" | "fallback" | "failed";
    path: "graph" | "fallback";
    result?: ProviderExecutionResult;
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
    const provider = options.result?.provider ?? "outlook";
    const rowKind = classifyPlanRow(options.item);
    const providerObjectType: ExecutionHistoryProviderObjectType = rowKind === "email" ? "message" : "event";
    const providerObjectId = options.result?.providerObjectId ?? null;

    if (options.status !== "success" || !providerObjectId) {
      return {
        provider: provider as "outlook" | "gmail",
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

    if (provider === "gmail") {
      if (providerObjectType === "event") {
        return {
          provider: "gmail" as const,
          providerObjectType,
          providerObjectId,
          canRecall: true,
          canModify: true,
          recallImplemented: true,
          modifyImplemented: true,
          recallReason: null,
          modifyReason: null,
        };
      }
      return {
        provider: "gmail" as const,
        providerObjectType,
        providerObjectId,
        canRecall: false,
        canModify: false,
        recallImplemented: false,
        modifyImplemented: false,
        recallReason: "This Google item cannot be recalled from History.",
        modifyReason: "This Google item cannot be modified from History.",
      };
    }

    if (action === "draft_created") {
      return {
        provider: "outlook" as const,
        providerObjectType,
        providerObjectId,
        canRecall: true,
        canModify: true,
        recallImplemented: true,
        modifyImplemented: true,
        recallReason: null,
        modifyReason: null,
      };
    }

    if (providerObjectType === "event") {
      return {
        provider: "outlook" as const,
        providerObjectType,
        providerObjectId,
        canRecall: true,
        canModify: true,
        recallImplemented: true,
        modifyImplemented: true,
        recallReason: null,
        modifyReason: null,
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

    if (action === "email_scheduled") {
      return {
        provider: "outlook" as const,
        providerObjectType,
        providerObjectId,
        canRecall: true,
        canModify: true,
        recallImplemented: true,
        modifyImplemented: true,
        recallReason: null,
        modifyReason: null,
      };
    }

    return {
      provider: "outlook" as const,
      providerObjectType,
      providerObjectId,
      canRecall: false,
      canModify: false,
      recallImplemented: false,
      modifyImplemented: false,
      recallReason: "This item is not recallable from History.",
      modifyReason: "This item cannot be modified from History.",
    };
  }

  async function recordExecutionHistory(options: {
    item: Plan["items"][number];
    status: "success" | "fallback" | "failed";
    path: "graph" | "fallback";
    executionGroupId?: string;
    result?: ProviderExecutionResult;
    fallbackExportKind?: "eml" | "ics";
    reason?: string;
  }) {
    try {
      const subject = getExecutionHistoryTitle(options.item);
      const timing = getExecutionHistoryTiming(options.item);
      const capabilities = getExecutionHistoryCapabilities(options);
      const detailFields = getExecutionHistoryDetailFields(options.item);
      const rowLinkage = getExecutionHistoryRowLinkage(options.item);
      const executionPlanSnapshot = getExecutionPlanSnapshot();
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
        outlookWebLink: options.result?.provider === "outlook" ? options.result?.webLink ?? null : null,
        teamsJoinLink: options.result?.provider === "outlook" ? options.result?.joinUrl ?? null : null,
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
          sourceRowId: rowLinkage.sourceRowId,
          scheduledSendAt:
            classifyPlanRow(options.item) === "email" && appSettings.emailHandlingMode === "schedule"
              ? getEmailScheduledSendISO(options.item)
              : null,
          scheduledEmailState:
            options.result?.action === "email_scheduled" ? "scheduled" : options.result?.action === "email_sent" ? "sent" : null,
          reminderTime: options.item.reminderTime ?? null,
          teamsMeeting: Boolean(options.item.meetingDraft?.teamsMeeting),
          emailHandlingMode: classifyPlanRow(options.item) === "email" ? appSettings.emailHandlingMode : null,
          executionPlanSnapshot,
          rowLinkage,
          overrideTracking: {
            isOverridden: false,
            overriddenAt: null,
            overrideSource: null,
            changedFields: [],
          },
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
      ? "Your Outlook connection no longer matches the selected email."
      : !connection.supportedMailbox && connection.identity
        ? `${connection.identity.mailboxEligibilityReason}`
        : connection.status === "reconnect_required"
          ? "Reconnect Outlook in Settings to continue."
          : "Connect Outlook to continue.";

    return { canUseGraph: false, connection, reason };
  }

  async function getEmailExecutionAvailability() {
    const requestedEmailAction =
      appSettings.emailHandlingMode === "send"
        ? "send"
        : appSettings.emailHandlingMode === "schedule"
          ? "schedule"
          : "draft";
    const outlookConnected = Boolean(outlookConnection?.connected);

    if (authEnabled && currentUser && !currentOrgId) {
      await refreshAuthContext();
    }

    if (appSettings.emailHandlingMode === "schedule") {
      const outlookAvailability = await getGraphExecutionAvailability(["Mail.ReadWrite", "Mail.Send"]);
      const gmailAvailability = await resolveGmailConnectionState(undefined, [GMAIL_COMPOSE_SCOPE]);
      const gmailConnected = gmailAvailability.connected && !gmailAvailability.stale;
      const chosenProvider: "outlook" | "gmail" = outlookAvailability.canUseGraph
        ? "outlook"
        : gmailConnected
          ? "gmail"
          : "outlook";
      return {
        provider: chosenProvider,
        canExecute: outlookAvailability.canUseGraph || gmailConnected,
        reason: outlookAvailability.canUseGraph || gmailConnected ? undefined : "Connect Outlook or Gmail to continue.",
        outlookAvailable: outlookAvailability.canUseGraph,
        gmailAvailable: gmailConnected,
      };
    }

    const outlookAvailability = await getGraphExecutionAvailability(["Mail.ReadWrite", "Mail.Send"]);
    if (outlookAvailability.canUseGraph) {
      return {
        provider: "outlook" as const,
        canExecute: true,
        outlookAvailable: true,
        gmailAvailable: false,
      };
    }

    const gmailAvailability = await resolveGmailConnectionState(undefined, [GMAIL_COMPOSE_SCOPE]);
    if (gmailAvailability.connected && !gmailAvailability.stale) {
      return {
        provider: "gmail" as const,
        canExecute: true,
        outlookAvailable: false,
        gmailAvailable: true,
      };
    }

    const reason =
      gmailAvailability.status === "reconnect_required"
        ? "Reconnect Gmail in Settings to continue."
        : outlookAvailability.connection?.status === "reconnect_required"
          ? "Reconnect Outlook or Gmail in Settings to continue."
          : "Connect Outlook or Gmail to continue.";

    return {
      provider: "outlook" as const,
      canExecute: false,
      reason,
      outlookAvailable: false,
      gmailAvailable: false,
    };
  }

  async function getCalendarExecutionAvailability(): Promise<ProviderExecutionAvailability> {
    const outlookAvailability = await getGraphExecutionAvailability(["Calendars.ReadWrite"]);
    if (outlookAvailability.canUseGraph) {
      return {
        provider: "outlook" as const,
        canExecute: true,
        reason: undefined,
        outlookAvailable: true,
        gmailAvailable: false,
      };
    }

    if (authEnabled && currentUser && !currentOrgId) {
      await refreshAuthContext();
    }

    const gmailAvailability = await resolveGmailConnectionState(undefined, [GOOGLE_CALENDAR_EVENTS_SCOPE]);
    const gmailConnected = gmailAvailability.connected && !gmailAvailability.stale;
    const gmailNeedsReconnect = gmailAvailability.status === "reconnect_required";
    const reason = gmailConnected
      ? undefined
      : gmailNeedsReconnect
        ? "Reconnect Google in Settings to enable Google Calendar, or connect Outlook to continue."
        : "Connect Outlook to continue.";

    return {
      provider: gmailConnected ? ("gmail" as const) : ("outlook" as const),
      canExecute: gmailConnected,
      reason,
      outlookAvailable: false,
      gmailAvailable: gmailConnected,
    };
  }

  function setExecutionNoticeForResult(result: ProviderExecutionResult) {
    const details = [result.title];

    setExecutionState("success");
    setExecutionNotices((current) => [
      {
        id: crypto.randomUUID(),
        notice: {
          tone: "success",
          title: result.message,
          details,
        },
      },
      ...current,
    ]);
  }

  function setExecutionNoticeForUnavailable(options: {
    title: string;
    reason: string;
    detail?: string;
  }) {
    setExecutionState("failure");
    setExecutionNotices((current) => [
      {
        id: crypto.randomUUID(),
        notice: {
          tone: "warning",
          title: options.title,
          message: options.reason,
          details: options.detail ? [options.detail] : undefined,
        },
      },
      ...current,
    ]);
  }

  function setExecutionNoticeForExportSummary(options: {
    graphUnavailableReason?: string;
    graphResults: ProviderExecutionResult[];
    failedEmailItems: Plan["items"];
    failedCalendarItems: Plan["items"];
    gmailScheduledDraftCount: number;
  }) {
    const details: string[] = [];
    const countResults = (filter: (result: ProviderExecutionResult) => boolean) =>
      options.graphResults.filter(filter).length;
    const draftCount = countResults((result) => result.kind === "email" && result.action === "draft_created");
    const gmailDraftCount = countResults(
      (result) =>
        result.provider === "gmail" &&
        result.kind === "email" &&
        result.action === "draft_created"
    );
    const outlookDraftCount = draftCount - gmailDraftCount;
    const gmailSentCount = countResults(
      (result) => result.provider === "gmail" && result.kind === "email" && result.action === "email_sent"
    );
    const outlookSentCount = countResults(
      (result) => result.provider === "outlook" && result.kind === "email" && result.action === "email_sent"
    );
    const gmailScheduledCount = countResults(
      (result) => result.provider === "gmail" && result.kind === "email" && result.action === "email_scheduled"
    );
    const outlookScheduledCount = countResults(
      (result) => result.provider === "outlook" && result.kind === "email" && result.action === "email_scheduled"
    );
    const gmailReminderCount = countResults(
      (result) => result.provider === "gmail" && result.kind === "reminder" && result.action === "reminder_created"
    );
    const outlookReminderCount = countResults(
      (result) => result.provider === "outlook" && result.kind === "reminder" && result.action === "reminder_created"
    );
    const gmailMeetingCount = countResults(
      (result) => result.provider === "gmail" && result.kind === "meeting" && result.action === "meeting_created"
    );
    const outlookMeetingCount = countResults(
      (result) => result.provider === "outlook" && result.kind === "meeting" && result.action === "meeting_created"
    );

    if (gmailDraftCount > 0) {
      details.push(`${gmailDraftCount} Gmail email draft${gmailDraftCount === 1 ? "" : "s"} created`);
    }
    if (outlookDraftCount > 0) {
      details.push(`${outlookDraftCount} Outlook email draft${outlookDraftCount === 1 ? "" : "s"} created`);
    }
    if (gmailScheduledCount > 0) {
      details.push(`${gmailScheduledCount} Gmail email${gmailScheduledCount === 1 ? "" : "s"} scheduled`);
    }
    if (outlookScheduledCount > 0) {
      details.push(`${outlookScheduledCount} Outlook email${outlookScheduledCount === 1 ? "" : "s"} scheduled`);
    }
    if (gmailSentCount > 0) {
      details.push(`${gmailSentCount} Gmail email${gmailSentCount === 1 ? "" : "s"} sent`);
    }
    if (outlookSentCount > 0) {
      details.push(`${outlookSentCount} Outlook email${outlookSentCount === 1 ? "" : "s"} sent`);
    }
    if (gmailReminderCount > 0) {
      details.push(`${gmailReminderCount} Google reminder${gmailReminderCount === 1 ? "" : "s"} created`);
    }
    if (outlookReminderCount > 0) {
      details.push(`${outlookReminderCount} Outlook reminder${outlookReminderCount === 1 ? "" : "s"} created`);
    }
    if (gmailMeetingCount > 0) {
      details.push(`${gmailMeetingCount} Google meeting${gmailMeetingCount === 1 ? "" : "s"} created`);
    }
    if (outlookMeetingCount > 0) {
      details.push(`${outlookMeetingCount} Outlook meeting${outlookMeetingCount === 1 ? "" : "s"} created`);
    }
    if (options.gmailScheduledDraftCount > 0) {
      details.push(
        `${options.gmailScheduledDraftCount} Gmail email${options.gmailScheduledDraftCount === 1 ? "" : "s"} saved as draft because scheduled send is not supported yet`
      );
    }
    if (options.failedEmailItems.length > 0) {
      details.push(
        `${options.failedEmailItems.length} email action${options.failedEmailItems.length === 1 ? "" : "s"} could not be completed`
      );
      details.push(
        `Email issue: ${options.failedEmailItems
          .slice(0, 3)
          .map((item) => item.customTitle ?? item.title)
          .join(", ")}${options.failedEmailItems.length > 3 ? ", ..." : ""}`
      );
    }
    if (options.failedCalendarItems.length > 0) {
      details.push(
        `${options.failedCalendarItems.length} calendar action${options.failedCalendarItems.length === 1 ? "" : "s"} could not be completed`
      );
      details.push(
        `Calendar issue: ${options.failedCalendarItems
          .slice(0, 3)
          .map((item) => item.customTitle ?? item.title)
          .join(", ")}${options.failedCalendarItems.length > 3 ? ", ..." : ""}`
      );
    }
    const hasFallbacks = options.failedEmailItems.length > 0 || options.failedCalendarItems.length > 0;
    const hasGraphSuccesses = options.graphResults.length > 0;

    setExecutionState(hasFallbacks ? (hasGraphSuccesses ? "success" : "failure") : "success");
    setExecutionNotices((current) => [
      {
        id: crypto.randomUUID(),
        notice: {
          tone: !hasGraphSuccesses && hasFallbacks ? "warning" : hasFallbacks ? "mixed" : "success",
          title: hasFallbacks
            ? hasGraphSuccesses
              ? "Some actions could not be completed"
              : "Export could not be completed"
            : "Export completed",
          message: options.graphUnavailableReason,
          details,
        },
      },
      ...current,
    ]);
  }

  function setExecutionNoticePending(options: {
    title: string;
    message: string;
    details?: string[];
  }) {
    const noticeId = crypto.randomUUID();
    setExecutionState("pending");
    setExecutionNotices((current) => [
      {
        id: noticeId,
        notice: {
          tone: "pending",
          title: options.title,
          message: options.message,
          details: options.details,
        },
      },
      ...current,
    ]);
    return noticeId;
  }

  function dismissExecutionNotice(id: string) {
    setExecutionNotices((current) => current.filter((entry) => entry.id !== id));
  }

  async function runQueuedExport(task: () => Promise<void>) {
    const runTask = async () => {
      activeExportCountRef.current += 1;
      setExecutionState("pending");
      try {
        await task();
      } finally {
        activeExportCountRef.current = Math.max(0, activeExportCountRef.current - 1);
        if (activeExportCountRef.current === 0) {
          setExecutionState(null);
        }
      }
    };

    const nextTask = exportQueueRef.current.then(runTask, runTask);
    exportQueueRef.current = nextTask.catch(() => undefined);
    await nextTask;
  }

  async function executePreviewEmailViaProvider(item: Plan["items"][number], provider: "outlook" | "gmail") {
    const draft = getResolvedEmailDraftForExecution(item);
    const fallbackSubject = draft.subject.trim() || item.title || "Email draft";
    const title = fallbackSubject;

    if (provider === "gmail") {
      try {
        if (appSettings.emailHandlingMode === "send") {
          const result = await sendGmailEmailFromEmailDraft({
            draft,
            fallbackSubject,
          });
          return {
            provider: "gmail",
            kind: "email",
            action: "email_sent",
            title,
            message: "Gmail email sent.",
            providerObjectId: result.id,
          } satisfies ProviderExecutionResult;
        }
        const result = await createGmailDraftFromEmailDraft({
          draft,
          fallbackSubject,
        });
        return {
          provider: "gmail",
          kind: "email",
          action: "draft_created",
          title,
          message:
            appSettings.emailHandlingMode === "schedule"
              ? "Gmail scheduled send is not supported yet. Draft created instead."
              : "Gmail draft created.",
          providerObjectId: result.id,
          webLink: result.webLink,
        } satisfies ProviderExecutionResult;
      } catch (error) {
        console.error("[plans] gmail email execution failed", {
          itemId: item.id,
          provider,
          action: appSettings.emailHandlingMode,
          error: error instanceof Error ? error.message : String(error),
        });
        throw error;
      }
    }

    if (appSettings.emailHandlingMode === "send") {
      await sendOutlookEmailFromEmailDraft({
        draft,
        fallbackSubject,
        expectedEmail: appSettings.outlookAccountEmail,
      });
      return {
        provider: "outlook",
        kind: "email",
        action: "email_sent",
        title,
        message: "Outlook email sent.",
      } satisfies ProviderExecutionResult;
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
        provider: "outlook",
        kind: "email",
        action: "email_scheduled",
        title,
        message: "Outlook email scheduled.",
        providerObjectId: result.id,
        webLink: result.webLink,
      } satisfies ProviderExecutionResult;
    }

    const result = await createOutlookDraftFromEmailDraft({
      draft,
      fallbackSubject,
      expectedEmail: appSettings.outlookAccountEmail,
    });
    return {
      provider: "outlook",
      kind: "email",
      action: "draft_created",
      title,
      message: "Outlook draft created.",
      providerObjectId: result.id,
      webLink: result.webLink,
    } satisfies ProviderExecutionResult;
  }

  async function executePreviewCalendarViaProvider(item: Plan["items"][number], provider: "outlook" | "gmail") {
    const timing = buildPreviewItemGraphTiming(item);
    const rowKind = classifyPlanRow(item);
    if (provider === "gmail") {
      const result = await createGoogleCalendarEvent({
        subject: item.customTitle ?? item.title,
        bodyText: item.body?.trim() || "",
        startISO: timing.startISO,
        endISO: timing.endISO,
        timeZone: "America/New_York",
        isAllDay: timing.isAllDay,
        location: item.meetingDraft?.location,
        attendees: item.meetingDraft?.attendees ?? [],
        teamsMeeting: item.meetingDraft?.teamsMeeting,
        addGoogleMeet: item.meetingDraft?.addGoogleMeet,
      });
      return {
        provider: "gmail",
        kind: rowKind === "meeting" ? "meeting" : "reminder",
        action: rowKind === "meeting" ? "meeting_created" : "reminder_created",
        title: item.customTitle ?? item.title,
        message: rowKind === "meeting" ? "Google Calendar meeting created." : "Google Calendar reminder created.",
        providerObjectId: result.id,
        webLink: result.webLink,
        joinUrl: result.joinUrl,
      } satisfies ProviderExecutionResult;
    }
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
      provider: "outlook",
      kind: rowKind === "meeting" ? "meeting" : "reminder",
      action: rowKind === "meeting" ? "meeting_created" : "reminder_created",
      title: item.customTitle ?? item.title,
      message: rowKind === "meeting" ? "Outlook meeting created." : "Outlook calendar reminder created.",
      providerObjectId: result.id,
      webLink: result.webLink,
      joinUrl: result.joinUrl,
    } satisfies ProviderExecutionResult;
  }

  async function exportPreviewEmailItem(itemId: string) {
    await runQueuedExport(async () => {
      const item = getLatestPreviewPlan().items.find((entry) => entry.id === itemId);
      if (!item) return;
      if (validateEmailRowsForExport([itemId], { usePopup: true })) return;
      if (warnIfMissingReminderTimes({ usePopup: true, itemIds: [itemId] })) return;
      if (warnIfPastScheduledItems({ usePopup: true, itemIds: [itemId] })) return;
      if (!confirmExport()) return;
      const executionGroupId = crypto.randomUUID();
      const pendingNoticeId = setExecutionNoticePending({
        title: "Starting email export",
        message:
          appSettings.emailHandlingMode === "send"
            ? "Sending email..."
            : "Creating email draft...",
        details: [item.customTitle ?? item.title],
      });

      const emailAvailability = await getEmailExecutionAvailability();
      if (emailAvailability.canExecute) {
        try {
          const result = await executePreviewEmailViaProvider(item, emailAvailability.provider);
          if (result.provider === "outlook") {
            await recordExecutionHistory({
              item,
              status: "success",
              path: "graph",
              executionGroupId,
              result,
            });
          }
          dismissExecutionNotice(pendingNoticeId);
          setExecutionNoticeForResult(result);
          return;
        } catch (error) {
          console.error("[plans] single email export unavailable", {
            requestedEmailAction: appSettings.emailHandlingMode,
            provider: emailAvailability.provider,
            outlookAvailable: emailAvailability.outlookAvailable,
            gmailAvailable: emailAvailability.gmailAvailable,
            itemId,
            reason: error instanceof Error ? error.message : String(error),
          });
          const userFixableMessage = getUserFixableEmailExecutionMessage(error);
          if (userFixableMessage) {
            if (typeof window !== "undefined") {
              window.alert(userFixableMessage);
            }
            return;
          }
          const reason = error instanceof Error ? error.message : "Email action failed.";
          await recordExecutionHistory({
            item,
            status: "failed",
            path: "graph",
            executionGroupId,
            reason,
          });
          dismissExecutionNotice(pendingNoticeId);
          setExecutionNoticeForUnavailable({
            title: "Email action could not be completed",
            reason,
          });
          return;
        }
      }

      await recordExecutionHistory({
        item,
        status: "failed",
        path: "graph",
        executionGroupId,
        reason: emailAvailability.reason ?? "Connect Outlook or Gmail to continue.",
      });
      dismissExecutionNotice(pendingNoticeId);
      setExecutionNoticeForUnavailable({
        title: "Email action is not available",
        reason: emailAvailability.reason ?? "Connect Outlook or Gmail to continue.",
      });
    });
  }

  async function exportPreviewMeetingItem(itemId: string) {
    await runQueuedExport(async () => {
      const item = getLatestPreviewPlan().items.find((entry) => entry.id === itemId);
      if (!item) return;
      if (validateMeetingRowsForExport([itemId], { usePopup: true })) return;
      if (warnIfMissingReminderTimes({ usePopup: true, itemIds: [itemId] })) return;
      if (warnIfPastScheduledItems({ usePopup: true, itemIds: [itemId] })) return;
      if (!confirmExport()) return;
      const executionGroupId = crypto.randomUUID();
      const pendingNoticeId = setExecutionNoticePending({
        title: "Starting calendar export",
        message: "Creating calendar event...",
        details: [item.customTitle ?? item.title],
      });

      const calendarAvailability = await getCalendarExecutionAvailability();
      if (calendarAvailability.canExecute) {
        try {
          const result = await executePreviewCalendarViaProvider(item, calendarAvailability.provider);
          await recordExecutionHistory({
            item,
            status: "success",
            path: "graph",
            executionGroupId,
            result,
          });
          dismissExecutionNotice(pendingNoticeId);
          setExecutionNoticeForResult(result);
          return;
        } catch (error) {
          const reason = error instanceof Error ? error.message : "Outlook calendar event creation failed.";
          await recordExecutionHistory({
            item,
            status: "failed",
            path: "graph",
            executionGroupId,
            reason,
          });
          dismissExecutionNotice(pendingNoticeId);
          setExecutionNoticeForUnavailable({
            title: "Meeting action could not be completed",
            reason,
          });
          return;
        }
      }

      await recordExecutionHistory({
        item,
        status: "failed",
        path: "graph",
        executionGroupId,
        reason: calendarAvailability.reason ?? "Connect Outlook to continue.",
      });
      dismissExecutionNotice(pendingNoticeId);
      setExecutionNoticeForUnavailable({
        title: "Meeting action is not available",
        reason: calendarAvailability.reason ?? "Connect Outlook to continue.",
      });
    });
  }

  function clearTransientEditingState() {
    setTemplateActionMessage("");
    setOpenMenuRowId(null);
    setOpenEmailDraftRowId(null);
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
    setAreAnchorsHidden(true);
    setGuidedForm({ ...snapshot.guidedForm });
    clearTransientEditingState();
  }

  function applyTemplateRecord(template: SavedPlanTemplate) {
    setTemplateActionMessage("");
    const snapshot = buildTemplateSnapshot(template);
    setBuilderSourceProvenance(
      buildBuilderSourceProvenance(snapshot, {
        sourceType: "saved_template",
        sourceLabel: template.name,
      })
    );
    applyBuilderSnapshot(snapshot);
    setLastTemplateSnapshot(snapshot);
  }

  function buildAiAnchorState(draft: AIPlanDraft) {
    const genericAnchors = createGenericPresetAnchors().map((anchor) => {
      const normalizedKey = normalizeAnchorKey(anchor.key);
      if (normalizedKey === normalizeAnchorKey("Event Name")) {
        return { ...anchor, value: draft.eventName || draft.templateName };
      }
      if (normalizedKey === normalizeAnchorKey("Event Date")) {
        return { ...anchor, value: draft.noEventDate ? "" : draft.anchorDate || todayYYYYMMDD() };
      }
      return anchor;
    });

    const genericKeys = new Set(genericAnchors.map((anchor) => normalizeAnchorKey(anchor.key)));
    const extraAnchors = draft.anchors
      .filter((anchor) => !genericKeys.has(normalizeAnchorKey(anchor.key)))
      .map((anchor) => ({
        id: crypto.randomUUID(),
        key: anchor.key,
        value: anchor.value,
      }));

    return [...genericAnchors, ...extraAnchors];
  }

  function buildBuilderSnapshotFromAiDraft(draft: AIPlanDraft): BuilderStateSnapshot {
    return {
      builderMode: "new",
      selectedTemplateId: null,
      planType: draft.baseType,
      templateName: draft.templateName || draft.eventName || "AI Generated Plan",
      eventName: draft.eventName || draft.templateName || "AI Generated Plan",
      anchorDate: draft.anchorDate || todayYYYYMMDD(),
      noEventDate: draft.noEventDate,
      weekendRule: draft.weekendRule,
      rows: draft.rows.map((row) => ({
        id: crypto.randomUUID(),
        title: row.title,
        body: row.body ?? "",
        offsetDays: row.offsetDays,
        dateBasis: row.dateBasis,
        rowType: row.rowType,
        reminderTime: normalizeReminderTimeInput(row.reminderTime ?? normalizedDefaultReminderTime),
        emailDraft:
          row.rowType === "email"
            ? normalizeEmailDraft(row.emailDraft)
            : { to: [], cc: [], bcc: [], subject: "", body: "" },
        durationDraft: row.durationDraft ? { ...row.durationDraft } : undefined,
        meetingDraft: row.meetingDraft ? { ...normalizeMeetingDraft(row.meetingDraft) } : undefined,
      })),
      anchors: buildAiAnchorState(draft),
      guidedForm: createEmptyGuidedForm(),
    };
  }

  function buildSavedTemplateFromAiDraft(draft: AIPlanDraft, name: string): SavedPlanTemplate {
    return normalizeImportedTemplate({
      id: makeId("template"),
      name: name.trim() || draft.templateName || draft.eventName || "AI Template",
      baseType: draft.baseType,
      templateMode: "custom",
      noEventDate: draft.noEventDate,
      weekendRule: draft.weekendRule,
      anchors: buildAiAnchorState(draft).map((anchor) => ({
        key: anchor.key,
        value: anchor.value,
      })),
      items: draft.rows.map((row) => ({
        id: crypto.randomUUID(),
        title: row.title,
        body: row.body ?? "",
        offsetDays: row.offsetDays,
        dateBasis: row.dateBasis,
        rowType: row.rowType,
        reminderTime: normalizeReminderTimeInput(row.reminderTime ?? normalizedDefaultReminderTime),
        emailDraft:
          row.rowType === "email"
            ? normalizeEmailDraft(row.emailDraft)
            : { to: [], cc: [], bcc: [], subject: "", body: "" },
        durationDraft: row.durationDraft ? { ...row.durationDraft } : undefined,
        meetingDraft: row.meetingDraft ? { ...normalizeMeetingDraft(row.meetingDraft) } : undefined,
      })),
    });
  }

  function buildCurrentBuilderContext(): AIPlanBuilderContext {
    return {
      title: eventName.trim() || templateName.trim() || "Current plan",
      planType,
      noEventDate,
      anchorDate,
      weekendRule,
      anchors: anchors
        .filter((anchor) => anchor.key.trim() || anchor.value.trim())
        .map((anchor) => ({
          key: anchor.key.trim(),
          value: anchor.value.trim(),
        })),
      rows: rows.map((row) => {
        const normalizedEmailDraft = normalizeEmailDraft(row.emailDraft);
        const normalizedMeetingDraft = normalizeMeetingDraft(row.meetingDraft);
        return {
          rowType: row.rowType,
          title: row.title.trim() || row.body?.trim() || getAiBuilderContextRowLabel(row.rowType),
          offsetDays: row.offsetDays ?? 0,
          dateBasis: row.dateBasis ?? "event",
          reminderTime: row.reminderTime?.trim() || undefined,
          emailSubject: normalizedEmailDraft.subject.trim() || undefined,
          recipientCount:
            row.rowType === "email"
              ? normalizedEmailDraft.to.length + normalizedEmailDraft.cc.length + normalizedEmailDraft.bcc.length
              : undefined,
          attendeeCount: normalizedMeetingDraft?.attendees?.length || undefined,
        };
      }),
    };
  }

  function buildBaselineFromBuilder(): AIDraftBaseline {
    const reminderCount = rows.filter((row) => row.rowType === "reminder").length;
    const emailCount = rows.filter((row) => row.rowType === "email").length;
    const meetingCount = rows.filter((row) => row.rowType === "calendar_event").length;

    return {
      sourceLabel: "starting builder plan",
      planType,
      noEventDate,
      anchorDate,
      eventTime: "",
      weekendRule,
      totalRows: rows.length,
      reminderCount,
      emailCount,
      meetingCount,
    };
  }

  function buildBaselineFromDraft(draft: AIPlanDraft, sourceLabel: string): AIDraftBaseline {
    const reminderCount = draft.rows.filter((row) => row.rowType === "reminder").length;
    const emailCount = draft.rows.filter((row) => row.rowType === "email").length;
    const meetingCount = draft.rows.filter((row) => row.rowType === "calendar_event").length;

    return {
      sourceLabel,
      planType: draft.baseType,
      noEventDate: draft.noEventDate,
      anchorDate: draft.anchorDate || "",
      eventTime: draft.eventTime || "",
      weekendRule: draft.weekendRule,
      totalRows: draft.rows.length,
      reminderCount,
      emailCount,
      meetingCount,
    };
  }

  function getAiDraftComparisonSummary(draft: AIPlanDraft, baseline: AIDraftBaseline) {
    const current = buildBaselineFromDraft(draft, baseline.sourceLabel);
    const totalDelta = current.totalRows - baseline.totalRows;
    const reminderDelta = current.reminderCount - baseline.reminderCount;
    const emailDelta = current.emailCount - baseline.emailCount;
    const meetingDelta = current.meetingCount - baseline.meetingCount;
    const timingChanged =
      baseline.noEventDate !== current.noEventDate ||
      baseline.anchorDate !== current.anchorDate ||
      baseline.eventTime !== current.eventTime ||
      baseline.weekendRule !== current.weekendRule;

    let qualitativeLabel = "";
    if (totalDelta > 0) {
      qualitativeLabel = "Timeline expanded";
    } else if (totalDelta < 0) {
      qualitativeLabel = "Timeline simplified";
    } else if (reminderDelta !== 0 || emailDelta !== 0 || meetingDelta !== 0) {
      qualitativeLabel = "Mix of actions changed";
    }

    return {
      current,
      totalDelta,
      reminderDelta,
      emailDelta,
      meetingDelta,
      timingChanged,
      qualitativeLabel,
    };
  }

  function buildInitialAiMessages(): AIConversationMessage[] {
    if (hasMeaningfulBuilderContent()) {
      return [
        {
          id: crypto.randomUUID(),
          role: "assistant",
          text: "I can help refine your current plan or create a new one. Choose how you’d like to start.",
          status: "needs_more_info",
          modeOptions: [
            { id: "refine_current", label: "Refine current plan" },
            { id: "start_new", label: "Start a new plan" },
          ],
        },
      ];
    }

    return [
      {
        id: crypto.randomUUID(),
        role: "assistant",
        text: "Tell me what kind of event or workflow you’re planning for. You can describe your job, the event, and any reminders, emails, or meetings you want help creating.",
        status: "needs_more_info",
        starterPrompts: [...AI_STARTER_PROMPTS],
      },
    ];
  }

  function buildFreshNewPlanAiMessages(): AIConversationMessage[] {
    return [
      {
        id: crypto.randomUUID(),
        role: "assistant",
        text: "Tell me what kind of event or workflow you want to plan, and I’ll draft a timeline for you.",
        status: "needs_more_info",
        starterPrompts: [...AI_STARTER_PROMPTS],
      },
    ];
  }

  function buildSeededExplorationMessages(): AIConversationMessage[] {
    return [
      {
        id: crypto.randomUUID(),
        role: "assistant",
        text: "Starting a fresh exploration from your last draft. Tell me how you want this version to change.",
        status: aiChatStatus,
        starterPrompts: [
          "Make the timeline more aggressive.",
          "Move reminders earlier.",
          "Add an internal prep meeting.",
          "Remove the follow-up email.",
        ],
      },
    ];
  }

  function buildCurrentBuilderSeededAiMessages(): AIConversationMessage[] {
    return [
      {
        id: crypto.randomUUID(),
        role: "assistant",
        text: "Continuing from your current builder plan. Tell me how you want to change it.",
        status: "needs_more_info",
        starterPrompts: [
          "Move the reminders earlier.",
          "Add an internal prep meeting.",
          "Remove the follow-up email.",
          "Make the timeline more aggressive.",
        ],
      },
    ];
  }

  function hasMeaningfulAiSession() {
    return Boolean(
      aiChatMessages.length > 1 ||
        aiChatDraft ||
        aiChatSummary ||
        aiChatChangeSummary.length ||
        aiChatConfidenceNote ||
        aiChatSuggestedNextActions.length ||
        aiBuilderContextMode === "refine_current"
    );
  }

  function restoreAiInitialState() {
    setAiComposer("");
    setAiGenerating(false);
    setAiChatError(null);
    setAiChatMessages(buildInitialAiMessages());
    setAiChatSummary("");
    setAiChatDraft(null);
    setAiChatStatus("needs_more_info");
    setAiChatChangeSummary([]);
    setAiChatConfidenceNote("");
    setAiChatSuggestedNextActions([]);
    setAiBuilderContextMode(hasMeaningfulBuilderContent() ? null : "start_new");
    setAiDraftBaseline(null);
    setAiSessionSource({ type: "new" });
    setShowAiApplyConfirm(false);
  }

  function openAiPanel() {
    if (!AI_ENABLED) return;
    setAiChatError(null);
    setShowAiApplyConfirm(false);
    setIsAiPanelOpen(true);
    setAiComposer("");
    setAiBuilderContextMode(hasMeaningfulBuilderContent() ? null : "start_new");
    if (aiChatMessages.length > 0) return;
    setAiSessionSource({ type: "new" });
    setAiChatMessages(buildInitialAiMessages());
  }

  function resetAiChatState() {
    setAiComposer("");
    setAiGenerating(false);
    setAiChatError(null);
    setAiChatMessages([]);
    setAiChatSummary("");
    setAiChatDraft(null);
    setAiChatStatus("needs_more_info");
    setAiChatChangeSummary([]);
    setAiChatConfidenceNote("");
    setAiChatSuggestedNextActions([]);
    setAiBuilderContextMode(null);
    setAiSessionBackup(null);
    setAiSavedTemplateInfo(null);
    setAiTemplateSaveMessage(null);
    setAiDraftBaseline(null);
    setAiSessionSource(null);
    setShowAiApplyConfirm(false);
  }

  function onStartOverAiSession() {
    if (hasMeaningfulAiSession()) {
      setAiSessionBackup({
        messages: aiChatMessages,
        summary: aiChatSummary,
        draft: aiChatDraft,
        status: aiChatStatus,
        changeSummary: aiChatChangeSummary,
        confidenceNote: aiChatConfidenceNote,
        suggestedNextActions: aiChatSuggestedNextActions,
        builderContextMode: aiBuilderContextMode,
        baseline: aiDraftBaseline,
        sessionSource: aiSessionSource,
      });
    }
    restoreAiInitialState();
  }

  function onRestoreAiSession() {
    if (!aiSessionBackup) return;
    setAiComposer("");
    setAiGenerating(false);
    setAiChatError(null);
    setAiChatMessages(aiSessionBackup.messages);
    setAiChatSummary(aiSessionBackup.summary);
    setAiChatDraft(aiSessionBackup.draft);
    setAiChatStatus(aiSessionBackup.status);
    setAiChatChangeSummary(aiSessionBackup.changeSummary);
    setAiChatConfidenceNote(aiSessionBackup.confidenceNote);
    setAiChatSuggestedNextActions(aiSessionBackup.suggestedNextActions);
    setAiBuilderContextMode(aiSessionBackup.builderContextMode);
    setAiDraftBaseline(aiSessionBackup.baseline);
    setAiSessionSource(aiSessionBackup.sessionSource);
    setShowAiApplyConfirm(false);
  }

  function onDuplicateAiDraftIntoNewExploration() {
    if (!hasMeaningfulAiSession() || !aiChatDraft) return;
    setAiSessionBackup({
      messages: aiChatMessages,
      summary: aiChatSummary,
      draft: aiChatDraft,
      status: aiChatStatus,
      changeSummary: aiChatChangeSummary,
      confidenceNote: aiChatConfidenceNote,
      suggestedNextActions: aiChatSuggestedNextActions,
      builderContextMode: aiBuilderContextMode,
      baseline: aiDraftBaseline,
      sessionSource: aiSessionSource,
    });
    setAiComposer("");
    setAiGenerating(false);
    setAiChatError(null);
    setAiChatMessages(buildSeededExplorationMessages());
    setAiBuilderContextMode(aiBuilderContextMode ?? "start_new");
    setAiDraftBaseline(buildBaselineFromDraft(aiChatDraft, "seed draft"));
    setAiSessionSource({ type: "branched_draft" });
    setShowAiApplyConfirm(false);
  }

  function openAiPanelFromCurrentBuilder() {
    if (!AI_ENABLED) return;
    setAiChatError(null);
    setShowAiApplyConfirm(false);
    setShowAiTemplateSaveDialog(false);
    setAiSavedTemplateInfo(null);
    setAiTemplateSaveMessage(null);
    setIsAiPanelOpen(true);
    setAiComposer("");
    setAiGenerating(false);
    setAiChatMessages(buildCurrentBuilderSeededAiMessages());
    setAiChatSummary(`Continuing from current builder plan "${eventName || templateName || "Current plan"}".`);
    setAiChatDraft(null);
    setAiChatStatus("needs_more_info");
    setAiChatChangeSummary([]);
    setAiChatConfidenceNote("");
    setAiChatSuggestedNextActions([]);
    setAiBuilderContextMode("refine_current");
    setAiSessionBackup(null);
    setAiDraftBaseline(buildBaselineFromBuilder());
    setAiSessionSource({ type: "current_builder" });
  }

  function focusAiComposer() {
    aiComposerRef.current?.scrollIntoView({ behavior: "smooth", block: "nearest" });
    aiComposerRef.current?.focus();
  }

  function onSelectAiBuilderMode(mode: "refine_current" | "start_new") {
    setAiBuilderContextMode(mode);
    setAiDraftBaseline(mode === "refine_current" ? buildBaselineFromBuilder() : null);
    setAiSessionSource(mode === "refine_current" ? { type: "current_builder" } : { type: "new" });
    if (mode === "start_new") {
      setAiComposer("");
      setAiGenerating(false);
      setAiChatError(null);
      setAiChatMessages(buildFreshNewPlanAiMessages());
      setAiChatSummary("");
      setAiChatDraft(null);
      setAiChatStatus("needs_more_info");
      setAiChatChangeSummary([]);
      setAiChatConfidenceNote("");
      setAiChatSuggestedNextActions([]);
      return;
    }
    void onSendAiMessage(
      mode === "refine_current" ? "Help me refine my current plan." : "Let’s start a new plan.",
      mode
    );
  }

  async function onSendAiMessage(rawInput?: string, explicitBuilderContextMode?: "refine_current" | "start_new") {
    const trimmedPrompt = (rawInput ?? aiComposer).trim();
    if (!trimmedPrompt) {
      setAiChatError("Please enter a message.");
      return;
    }
    const builderContextMode = explicitBuilderContextMode ?? aiBuilderContextMode ?? "start_new";

    const userMessage: AIConversationMessage = {
      id: crypto.randomUUID(),
      role: "user",
      text: trimmedPrompt,
    };
    const requestMessages: AIChatMessage[] = [...aiChatMessages, userMessage].map((message) => ({
      role: message.role,
      text: message.text,
    }));

    setAiGenerating(true);
    setAiChatError(null);
    setAiChatMessages((current) => [...current, userMessage]);
    if (!rawInput) {
      setAiComposer("");
    }

    try {
      const requestBody: AIPlanChatRequest = {
        messages: requestMessages,
        currentSummary: aiChatSummary,
        currentDraft: aiChatDraft,
        builderContextMode,
        currentBuilderContext:
          builderContextMode === "refine_current" ? buildCurrentBuilderContext() : null,
      };

      const response = await fetch("/api/ai/generate-plan", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(requestBody),
      });

      const payload = (await response.json()) as AIPlanChatTurnResult & { error?: string };
      if (!response.ok) {
        throw new Error(payload.error ?? "AI plan generation failed.");
      }
      if (!payload?.draft?.rows?.length || !payload.assistantMessage) {
        throw new Error("AI did not return a usable plan draft.");
      }
      setAiChatSummary(payload.summary);
      setAiChatDraft(payload.draft);
      setAiChatStatus(payload.status);
      setAiChatChangeSummary(payload.changeSummary);
      setAiChatConfidenceNote(payload.confidenceNote);
      setAiChatSuggestedNextActions(payload.suggestedNextActions);
      setAiChatMessages((current) => [
        ...current,
        {
          id: crypto.randomUUID(),
          role: "assistant",
          text: payload.assistantMessage,
          summary: payload.summary,
          status: payload.status,
          followUpQuestions: payload.followUpQuestions,
          changeSummary: payload.changeSummary,
          confidenceNote: payload.confidenceNote,
          suggestedNextActions: payload.suggestedNextActions,
        },
      ]);
    } catch (error) {
      setAiChatError(error instanceof Error ? error.message : "AI plan generation failed.");
    } finally {
      setAiGenerating(false);
    }
  }

  function hasMeaningfulBuilderContent() {
    const hasMeaningfulName = Boolean(templateName.trim() || eventName.trim());
    const hasMeaningfulAnchorValue = anchors.some((anchor) => anchor.value.trim());
    const hasMeaningfulRowContent = rows.some((row) => {
      const normalizedEmailDraft = normalizeEmailDraft(row.emailDraft);
      const normalizedMeetingDraft = normalizeMeetingDraft(row.meetingDraft);
      return Boolean(
        row.title.trim() ||
          row.body?.trim() ||
          normalizedEmailDraft.subject.trim() ||
          normalizedEmailDraft.body.trim() ||
          normalizedEmailDraft.to.length ||
          normalizedEmailDraft.cc.length ||
          normalizedEmailDraft.bcc.length ||
          normalizedMeetingDraft?.attendees?.length ||
          normalizedMeetingDraft?.location?.trim()
      );
    });

    const hasMultipleRows = rows.length > 1;

    return hasMeaningfulName || hasMeaningfulAnchorValue || hasMeaningfulRowContent || hasMultipleRows;
  }

  function applyAiDraftToBuilder() {
    if (!aiChatDraft) return;
    const sourceDetails = getAiSessionSourceDetails(aiSessionSource);
    const nextSnapshot = buildBuilderSnapshotFromAiDraft(aiChatDraft);
    setLastTemplateSnapshot(buildCurrentTemplateSnapshot());
    setLastBuilderSourceProvenance(builderSourceProvenance);
    applyBuilderSnapshot(nextSnapshot);
    setBuilderSourceProvenance(
      buildBuilderSourceProvenance(nextSnapshot, {
        sourceType: "ai_draft",
        sourceLabel: sourceDetails.label,
        hadMissingDetails: aiDraftMissingDetails.length > 0,
      })
    );
    setTemplateActionMessage("AI draft applied. Review and adjust it before exporting.");
    setAiApplySuccessMessage("AI draft loaded into your plan. You can now review, edit, preview, and export it.");
    resetAiChatState();
    setIsAiPanelOpen(false);
    window.setTimeout(() => {
      builderSectionRef.current?.scrollIntoView({ behavior: "smooth", block: "start" });
    }, 0);
  }

  function onApplyAiDraft() {
    if (!aiChatDraft) return;
    if (hasMeaningfulBuilderContent()) {
      setShowAiApplyConfirm(true);
      return;
    }
    applyAiDraftToBuilder();
  }

  function updateAiDraftRow(
    rowIndex: number,
    updater: (row: AIPlanDraft["rows"][number]) => AIPlanDraft["rows"][number]
  ) {
    setAiChatDraft((current) => {
      if (!current) return current;
      return {
        ...current,
        rows: current.rows.map((row, index) => (index === rowIndex ? updater(row) : row)),
      };
    });
  }

  function removeAiDraftRow(rowIndex: number) {
    setAiChatDraft((current) => {
      if (!current) return current;
      return {
        ...current,
        rows: current.rows.filter((_, index) => index !== rowIndex),
      };
    });
  }

  function openAiTemplateSaveDialog() {
    if (!aiChatDraft) return;
    setAiTemplateSaveMessage(null);
    setAiSavedTemplateInfo(null);
    setAiTemplateNameDraft(aiChatDraft.templateName || aiChatDraft.eventName || "AI Template");
    setShowAiTemplateSaveDialog(true);
  }

  function saveAiDraftAsTemplate() {
    if (!aiChatDraft) return;
    const trimmedName = aiTemplateNameDraft.trim();
    if (!trimmedName) {
      setAiTemplateSaveMessage("Template name is required.");
      return;
    }
    if (hasDuplicateTemplateName(trimmedName)) {
      setAiTemplateSaveMessage("That name is already taken. Please choose another name.");
      return;
    }

    const nextTemplate = buildSavedTemplateFromAiDraft(aiChatDraft, trimmedName);
    hasLocalTemplateMutationRef.current = true;
    const nextTemplates = [...savedTemplates, nextTemplate];
    setSavedTemplates(nextTemplates);
    persistTemplateStateImmediately(nextTemplates, selectedTemplateId);
    setAiSavedTemplateInfo({ id: nextTemplate.id, name: nextTemplate.name });
    setHighlightedTemplateId(nextTemplate.id);
    setAiTemplateSaveMessage(null);
    setShowAiTemplateSaveDialog(false);
  }

  function openBuilderTemplateSaveDialog() {
    if (!hasMeaningfulBuilderContent()) return;
    setBuilderTemplateSaveMessage(null);
    setBuilderTemplateNameDraft(templateName.trim() || eventName.trim() || "Current Plan Template");
    setShowBuilderTemplateSaveDialog(true);
  }

  function saveCurrentBuilderAsTemplate() {
    const trimmedName = builderTemplateNameDraft.trim();
    if (!trimmedName) {
      setBuilderTemplateSaveMessage("Template name is required.");
      return;
    }
    if (hasDuplicateTemplateName(trimmedName)) {
      setBuilderTemplateSaveMessage("That name is already taken. Please choose another name.");
      return;
    }

    const nextTemplate = normalizeImportedTemplate({
      id: makeId("template"),
      name: trimmedName,
      baseType: planType,
      templateMode: "custom",
      noEventDate,
      weekendRule,
      anchors: resolvedAnchors
        .map((anchor) => ({ key: anchor.key.trim(), value: anchor.value }))
        .filter((anchor) => anchor.key),
      items: cloneTemplateRows(rows),
    });
    hasLocalTemplateMutationRef.current = true;
    const nextTemplates = [...savedTemplates, nextTemplate];
    setSavedTemplates(nextTemplates);
    persistTemplateStateImmediately(nextTemplates, selectedTemplateId);
    setHighlightedTemplateId(nextTemplate.id);
    setTemplateActionMessage(`Saved "${nextTemplate.name}" as a custom template.`);
    setBuilderTemplateSaveMessage(null);
    setShowBuilderTemplateSaveDialog(false);
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
      setLastBuilderSourceProvenance(builderSourceProvenance);
    }
    setBuilderSourceProvenance(null);
    setBuilderMode("new");
    setSelectedTemplateId(null);
    setTemplateName("");
    setEventName("");
    setAnchorDate(todayYYYYMMDD());
    setNoEventDate(false);
    setWeekendRule("prior_business_day");
    setRows([createEmptyBuilderRow(normalizedDefaultReminderTime)]);
    setAnchors(createGenericPresetAnchors());
    setAreAnchorsHidden(true);
    setGuidedForm(createEmptyGuidedForm());
    clearTransientEditingState();
  }

  function cancelEditing() {
    if (builderMode === "new" && lastTemplateSnapshot) {
      applyBuilderSnapshot(lastTemplateSnapshot);
      setBuilderSourceProvenance(lastBuilderSourceProvenance);
      return;
    }
    if (selectedTemplateId) {
      applyTemplate(selectedTemplateId);
      return;
    }
    clearTransientEditingState();
  }

  async function exportPreviewReminderItem(itemId: string) {
    await runQueuedExport(async () => {
      const item = getLatestPreviewPlan().items.find((entry) => entry.id === itemId);
      if (!item) return;
      if (warnIfMissingReminderTimes({ usePopup: true, itemIds: [itemId] })) return;
      if (warnIfPastScheduledItems({ usePopup: true, itemIds: [itemId] })) return;
      if (!confirmExport()) return;
      const executionGroupId = crypto.randomUUID();
      const pendingNoticeId = setExecutionNoticePending({
        title: "Starting calendar export",
        message: "Creating calendar event...",
        details: [item.customTitle ?? item.title],
      });

      const calendarAvailability = await getCalendarExecutionAvailability();
      if (calendarAvailability.canExecute) {
        try {
          const result = await executePreviewCalendarViaProvider(item, calendarAvailability.provider);
          await recordExecutionHistory({
            item,
            status: "success",
            path: "graph",
            executionGroupId,
            result,
          });
          dismissExecutionNotice(pendingNoticeId);
          setExecutionNoticeForResult(result);
          return;
        } catch (error) {
          const reason = error instanceof Error ? error.message : "Outlook calendar event creation failed.";
          await recordExecutionHistory({
            item,
            status: "failed",
            path: "graph",
            executionGroupId,
            reason,
          });
          dismissExecutionNotice(pendingNoticeId);
          setExecutionNoticeForUnavailable({
            title: "Reminder action could not be completed",
            reason,
          });
          return;
        }
      }

      await recordExecutionHistory({
        item,
        status: "failed",
        path: "graph",
        executionGroupId,
        reason: calendarAvailability.reason ?? "Connect Outlook to continue.",
      });
      dismissExecutionNotice(pendingNoticeId);
      setExecutionNoticeForUnavailable({
        title: "Reminder action is not available",
        reason: calendarAvailability.reason ?? "Connect Outlook to continue.",
      });
    });
  }

  async function exportCurrentPlan(options?: { skipValidation?: boolean; skipConfirm?: boolean }) {
    if (!options?.skipValidation && validateExportBeforeRun()) return;
    if (warnIfPastScheduledItems({ usePopup: true })) return;
    if (!options?.skipConfirm && !confirmExport()) return;

    const latestPreviewPlan = getLatestPreviewPlan();
    const { emailItems, calendarItems } = partitionPlanItemsByKind(latestPreviewPlan.items);
    const pendingNoticeId = setExecutionNoticePending({
      title: "Starting export",
      message:
        emailItems.length > 0 && calendarItems.length === 0
          ? appSettings.emailHandlingMode === "send"
            ? "Sending email..."
            : "Creating email draft..."
          : calendarItems.length > 0 && emailItems.length === 0
            ? "Creating calendar event..."
            : "Sending provider actions...",
    });
    const calendarAvailability: ProviderExecutionAvailability = calendarItems.length > 0
      ? await getCalendarExecutionAvailability()
      : {
          provider: "outlook",
          canExecute: false,
          reason: undefined,
          outlookAvailable: false,
          gmailAvailable: false,
        };

    const failedEmailItems: Plan["items"] = [];
    const failedCalendarItems: Plan["items"] = [];
    const failedCalendarReasons = new Map<string, string>();
    const graphResults: ProviderExecutionResult[] = [];
    const executionGroupId = crypto.randomUUID();
    const emailAvailability: ProviderExecutionAvailability = emailItems.length > 0
      ? await getEmailExecutionAvailability()
      : {
          provider: "outlook",
          canExecute: false,
          reason: undefined,
          outlookAvailable: false,
          gmailAvailable: false,
        };
    let gmailScheduledDraftCount = 0;
    const executableItems = latestPreviewPlan.items.filter((item) => {
      const rowKind = classifyPlanRow(item);
      return rowKind === "email" ? emailAvailability.canExecute : calendarAvailability.canExecute;
    });

    for (const item of executableItems) {
      const rowKind = classifyPlanRow(item);
      try {
        const result =
          rowKind === "email"
            ? await executePreviewEmailViaProvider(item, emailAvailability.provider)
            : await executePreviewCalendarViaProvider(item, calendarAvailability.provider);

        if (result.provider === "outlook" || rowKind !== "email") {
          await recordExecutionHistory({
            item,
            status: "success",
            path: "graph",
            executionGroupId,
            result,
          });
        }

        graphResults.push(result);
        if (rowKind === "email" && result.action === "draft_created" && appSettings.emailHandlingMode === "schedule" && result.provider === "gmail") {
          gmailScheduledDraftCount += 1;
        }
      } catch (error) {
        if (rowKind === "email") {
          console.error("[plans] bulk email export unavailable", {
            requestedEmailAction: appSettings.emailHandlingMode,
            provider: emailAvailability.provider,
            outlookAvailable: emailAvailability.outlookAvailable,
            gmailAvailable: emailAvailability.gmailAvailable,
            itemId: item.id,
            reason: error instanceof Error ? error.message : String(error),
          });
          failedEmailItems.push(item);
        } else {
          const failureReason = error instanceof Error ? error.message : String(error);
          console.error("[plans] bulk calendar export unavailable", {
            provider: calendarAvailability.provider,
            outlookAvailable: calendarAvailability.outlookAvailable,
            gmailAvailable: calendarAvailability.gmailAvailable,
            itemId: item.id,
            reason: failureReason,
          });
          failedCalendarItems.push(item);
          failedCalendarReasons.set(item.id, failureReason);
        }
      }
    }

    if (!emailAvailability.canExecute) {
      failedEmailItems.push(...emailItems);
    }
    if (!calendarAvailability.canExecute) {
      failedCalendarItems.push(...calendarItems);
    }

    for (const item of failedEmailItems) {
      await recordExecutionHistory({
        item,
        status: "failed",
        path: "graph",
        executionGroupId,
        reason: emailAvailability.canExecute
          ? `${emailAvailability.provider === "gmail" ? "Gmail" : "Outlook"} email action could not be completed.`
          : emailAvailability.reason,
      });
    }
    if (failedCalendarItems.length > 0) {
      for (const item of failedCalendarItems) {
        const failureReason = failedCalendarReasons.get(item.id);
        await recordExecutionHistory({
          item,
          status: "failed",
          path: "graph",
          executionGroupId,
          reason: calendarAvailability.canExecute
            ? failureReason ||
              (calendarAvailability.provider === "gmail"
                ? "Reconnect Google to enable calendar access."
                : "Outlook calendar action could not be completed.")
            : calendarAvailability.reason,
        });
      }
    }

    dismissExecutionNotice(pendingNoticeId);
    setExecutionNoticeForExportSummary({
      graphUnavailableReason: calendarAvailability.canExecute ? undefined : calendarAvailability.reason,
      graphResults,
      failedEmailItems,
      failedCalendarItems,
      gmailScheduledDraftCount,
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
    if (validateEmailRowsForExport(undefined, { usePopup: true })) return true;
    if (validateMeetingRowsForExport(undefined, { usePopup: true })) return true;
    if (warnIfMissingReminderTimes({ usePopup: true })) return true;
    if (warnIfPastScheduledItems({ usePopup: true })) return true;
    return false;
  }

  function validateExportBeforeRun() {
    const missingRequirements = getMissingPreviewRequirements();
    if (missingRequirements.length > 0) {
      window.alert(`Export is missing:\n\n${missingRequirements.join("\n")}`);
      return true;
    }
    if (validateEmailRowsForExport(undefined, { usePopup: true })) return true;
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

  function renameSavedTemplate(templateId: string) {
    const template = savedTemplates.find((entry) => entry.id === templateId);
    if (!template || isProtectedTemplate(template) || typeof window === "undefined") return;

    const nextName = window.prompt("Rename template", template.name);
    if (!nextName) return;

    const trimmedName = nextName.trim();
    if (!trimmedName) {
      setTemplateActionMessage("Template name is required.");
      return;
    }

    if (hasDuplicateTemplateName(trimmedName, { excludeTemplateId: template.id })) {
      setTemplateActionMessage("That name is already taken. Please choose another name.");
      return;
    }

    const nextTemplates = savedTemplates.map((entry) => (entry.id === template.id ? { ...entry, name: trimmedName } : entry));
    hasLocalTemplateMutationRef.current = true;
    setTemplateActionMessage(`Renamed "${template.name}" to "${trimmedName}".`);
    setSavedTemplates(nextTemplates);
    persistTemplateStateImmediately(nextTemplates, selectedTemplateId);

    if (selectedTemplateId === template.id) {
      setTemplateName(trimmedName);
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

  const activeAccountProvider = hasMounted
    ? outlookConnection?.status === "connected"
      ? "outlook"
      : gmailConnection?.status === "connected"
        ? "gmail"
        : outlookConnection?.status === "reconnect_required"
          ? "outlook"
          : gmailConnection?.status === "reconnect_required"
            ? "gmail"
            : null
    : null;
  const accountConnectionStatus = hasMounted
    ? activeAccountProvider === "gmail"
      ? gmailConnection?.status ?? "not_connected"
      : outlookConnection?.status ?? appSettings.outlookConnectionStatus
    : "not_connected";
  const connectedMailboxEmail = hasMounted
    ? activeAccountProvider === "gmail"
      ? getConnectedGmailMailboxEmail(gmailConnection?.identity) || ""
      : getConnectedOutlookMailboxEmail(outlookConnection?.identity) || appSettings.outlookAccountEmail
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
      ? `${activeAccountProvider === "gmail" ? "Google" : "Outlook"}: ${connectedMailboxEmail || "Connected email account"}`
      : accountConnectionStatus === "reconnect_required"
        ? `${activeAccountProvider === "gmail" ? "Google" : "Outlook"}: ${connectedMailboxEmail || "Reconnect required"}`
        : "No connected email account";
  const accountButtonLabel =
    accountConnectionStatus === "connected" ? "Manage in Settings" : "Reconnect in Settings";

  useEffect(() => {
    setRows((currentRows) => {
      let changed = false;
      const nextRows = currentRows.map((row) => {
        const normalizedMeetingDraft = normalizeMeetingDraft(row.meetingDraft);
        if (!normalizedMeetingDraft) return row;

        const nextLocation = getMeetingLocationValue(normalizedMeetingDraft, activeAccountProvider);
        if ((normalizedMeetingDraft.location ?? "") === nextLocation) return row;

        changed = true;
        return {
          ...row,
          meetingDraft: {
            ...normalizedMeetingDraft,
            location: nextLocation,
          },
        };
      });

      if (!changed) return currentRows;
      rowsRef.current = nextRows;
      return nextRows;
    });
  }, [activeAccountProvider]);

  const aiSessionSourceDetails = getAiSessionSourceDetails(aiSessionSource);
  const aiDraftIdentityDetails = getAiDraftIdentityDetails(aiSessionSource);
  const latestAssistantMessage = [...aiChatMessages].reverse().find((message) => message.role === "assistant") ?? null;
  const aiDraftStageDetails = getAiDraftStageDetails({
    hasDraft: Boolean(aiChatDraft),
    readiness: aiChatStatus,
    hasFollowUpQuestions: Boolean(latestAssistantMessage?.followUpQuestions?.length),
    wasSavedAsTemplate: Boolean(aiSavedTemplateInfo),
    source: aiSessionSource,
    rowCount: aiChatDraft?.rows.length ?? 0,
  });
  const aiDraftMissingDetails = getAiDraftMissingDetails({
    draft: aiChatDraft,
    confidenceNote: aiChatConfidenceNote,
    hasFollowUpQuestions: Boolean(latestAssistantMessage?.followUpQuestions?.length),
    source: aiSessionSource,
  });
  void openAiPanelFromCurrentBuilder;
  const builderHasMeaningfulContent = hasMeaningfulBuilderContent();
  const currentBuilderSignature = buildBuilderContentSignature({
    planType,
    templateName,
    eventName,
    anchorDate,
    noEventDate,
    weekendRule,
    anchors,
    rows,
  });
  const shouldRenderBuilderSourceBanner = hasMounted && builderHasMeaningfulContent;

  useEffect(() => {
    if (!builderHasMeaningfulContent) {
      if (builderMode === "new" && builderSourceProvenance?.sourceType === "manual") {
        setBuilderSourceProvenance(null);
      }
      return;
    }
    if (builderSourceProvenance) return;
    if (selectedTemplateId || builderMode === "template") return;
    setBuilderSourceProvenance({
      sourceType: "manual",
      sourceLabel: "Manual plan",
      loadedAt: getBuilderSourceTimestamp(),
      sourceSignature: currentBuilderSignature,
    });
  }, [builderHasMeaningfulContent, builderMode, builderSourceProvenance, currentBuilderSignature, selectedTemplateId]);

  useEffect(() => {
    if (AI_ENABLED) return;
    if (!isAiPanelOpen) return;
    setIsAiPanelOpen(false);
  }, [isAiPanelOpen]);

  return (
    <div className="space-y-8 text-gray-900">
      <section className="space-y-2">
        <div className="flex flex-col gap-4 md:flex-row md:items-start md:justify-between">
          <div className="space-y-2">
            <h1 className="text-3xl font-bold text-gray-900">Plans</h1>
            <p className="text-sm text-gray-600">
              Build an event timeline and export it to Outlook when you&apos;re ready.
            </p>
            {AI_ENABLED ? (
              <div>
                <button
                  type="button"
                  onClick={openAiPanel}
                  className="rounded-lg border border-gray-300 bg-white px-4 py-2 text-sm text-gray-900 hover:bg-gray-50"
                >
                  Generate with AI
                </button>
              </div>
            ) : null}
          </div>
          <div className="rounded-xl border bg-white px-4 py-3 text-sm shadow-sm md:min-w-[260px] md:max-w-[280px]">
            <div className="text-xs font-semibold uppercase tracking-wide text-gray-500">Connected Email Account</div>
            <div className={`mt-1 font-medium ${accountStatusClass}`}>{accountPrimaryText}</div>
            <div className={`mt-1 text-xs font-semibold uppercase tracking-wide ${accountStatusClass}`}>
              {accountStatusLabel}
            </div>
            {providerLoading.outlook || providerLoading.gmail ? (
              <div className="mt-1 text-xs text-gray-500">Refreshing provider status…</div>
            ) : null}
            <Link
              href="/settings"
              className="mt-2 inline-flex rounded-lg border border-gray-200 px-3 py-1.5 text-xs text-gray-700 hover:bg-gray-50"
            >
              {accountButtonLabel}
            </Link>
          </div>
        </div>
      </section>

      {executionNotices.length > 0 && !isBuilderPreviewOpen ? (
        <div className="space-y-3">
          {executionNotices.map((entry) => (
            <OutlookExecutionNoticeCard key={entry.id} notice={entry.notice} onDismiss={() => dismissExecutionNotice(entry.id)} />
          ))}
        </div>
      ) : null}

      <section className="rounded-2xl border bg-white shadow-sm">
        <div className="space-y-6 p-6">
          <div className="flex flex-wrap items-center justify-between gap-3">
            <div>
              <h2 className="text-lg font-semibold text-gray-900">Templates</h2>
              <p className="mt-1 text-sm text-gray-600">Choose a saved template or start a fresh plan.</p>
            </div>
            <div className="flex items-center gap-2">
              <button
                type="button"
                onClick={startNewPlan}
                className="rounded-xl border border-slate-300 bg-white px-4 py-2 text-sm font-medium text-slate-700 shadow-sm hover:-translate-y-0.5 hover:border-slate-400 hover:bg-slate-50 hover:shadow-md"
              >
                + New Plan
              </button>
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
                  className={`rounded-xl border px-3 py-2 text-sm shadow-sm hover:-translate-y-0.5 hover:bg-gray-50 hover:shadow-md ${
                    isTemplateCopyMode ? "border-blue-300 bg-blue-50 text-blue-700" : "border-slate-300 bg-white text-slate-700"
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
                  className="rounded-xl border border-slate-300 bg-white px-3 py-2 text-sm text-slate-700 shadow-sm hover:-translate-y-0.5 hover:bg-gray-50 hover:shadow-md"
                >
                  {isTemplateManageMode ? "Done" : "Edit"}
                </button>
              </div>
          </div>
          {templateActionMessage ? <p className="text-xs text-gray-500">{templateActionMessage}</p> : null}

          <div className="grid grid-cols-1 gap-5 sm:grid-cols-2 xl:grid-cols-3">
            {savedTemplates.map((template) => (
              <div key={template.id} className="relative min-w-0">
                {isTemplateManageMode && !isProtectedTemplate(template) ? (
                  <div className="absolute right-3 top-3 z-10 flex gap-2">
                    <button
                      type="button"
                      onClick={() => renameSavedTemplate(template.id)}
                      className="rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs font-medium text-gray-700 shadow-sm hover:bg-gray-50"
                    >
                      Rename
                    </button>
                    <button
                      type="button"
                      onClick={() => deleteSavedTemplate(template.id)}
                      className="rounded-lg border border-red-200 bg-white px-2.5 py-1.5 text-xs font-medium text-red-700 shadow-sm hover:bg-red-50"
                    >
                      Delete
                    </button>
                  </div>
                ) : null}
                <button
                  type="button"
                  onClick={() => onSelectSavedTemplate(template.id)}
                  className={`group flex min-h-[120px] w-full flex-col items-start justify-between rounded-[24px] border bg-white px-5 py-5 text-left shadow-sm transition duration-150 hover:-translate-y-1 hover:shadow-lg ${
                    selectedTemplateId === template.id
                      ? "border-blue-300 bg-blue-50/80 shadow-[0_18px_38px_-24px_rgba(59,130,246,0.65)]"
                      : highlightedTemplateId === template.id
                        ? "border-green-300 bg-green-50/80 shadow-[0_18px_38px_-24px_rgba(34,197,94,0.55)]"
                        : "border-slate-200 hover:border-slate-300"
                  }`}
                >
                  <div className="flex w-full items-start justify-between gap-4">
                    <div className="min-w-0">
                      <div className="max-w-full break-words text-lg font-semibold leading-6 text-gray-900">{template.name}</div>
                    </div>
                    {selectedTemplateId === template.id ? (
                      <span className="inline-flex h-7 w-7 flex-none items-center justify-center rounded-full bg-blue-600 text-sm font-semibold text-white">
                        ✓
                      </span>
                    ) : highlightedTemplateId === template.id ? (
                      <span className="inline-flex h-7 w-7 flex-none items-center justify-center rounded-full bg-green-600 text-sm font-semibold text-white">
                        +
                      </span>
                    ) : (
                      <span className="inline-flex h-7 w-7 flex-none items-center justify-center rounded-full border border-slate-200 text-xs font-semibold text-slate-400 transition group-hover:border-slate-300 group-hover:text-slate-500">
                        →
                      </span>
                    )}
                  </div>
                  <div className="text-xs font-medium uppercase tracking-[0.22em] text-slate-500">
                    {selectedTemplateId === template.id ? "Selected Template" : isTemplateManageMode ? "Manage Template" : "Open Template"}
                  </div>
                </button>
              </div>
            ))}
          </div>
        </div>
      </section>

      {AI_ENABLED && aiApplySuccessMessage ? (
        <div className="rounded-2xl border border-green-200 bg-green-50 px-4 py-3 text-sm text-green-900 shadow-sm">
          {aiApplySuccessMessage}
        </div>
      ) : null}

      <section ref={builderSectionRef} className="rounded-2xl border bg-white shadow-sm">
        <div className="space-y-5 p-6">
              <div>
                <div className="text-sm font-semibold uppercase tracking-wide text-gray-500">Plan Builder</div>
                <h2 className="text-2xl font-semibold text-gray-900">{eventName || "Untitled Plan"}</h2>
              </div>

              <div className="space-y-4 rounded-[28px] border border-slate-200 bg-[linear-gradient(180deg,#ffffff_0%,#f8fafc_100%)] p-5 shadow-[0_18px_45px_-35px_rgba(15,23,42,0.65)]">
                <div className="rounded-[24px] border border-slate-200 bg-[radial-gradient(circle_at_top_left,_rgba(191,219,254,0.45),_transparent_38%),linear-gradient(180deg,#ffffff_0%,#f8fafc_100%)] p-5 shadow-sm">
                  <div>
                    {shouldRenderSimpleEventHeaderFields ? (
                      <div className="grid grid-cols-1 gap-4 md:grid-cols-2">
                        <div className="md:col-span-2">
                          <input
                            className="w-full rounded-2xl border border-slate-200 bg-white px-4 py-3 text-slate-900 shadow-sm"
                            placeholder="Event Name"
                            value={eventName}
                            onChange={(e) => setEventName(e.target.value)}
                          />
                        </div>
                        <div>
                          <label className="mb-1 block text-sm font-medium text-gray-700">Event Date</label>
                          {noEventDate ? (
                            <>
                              <input
                                type="text"
                                className="w-full rounded-2xl border border-slate-200 bg-slate-100 px-4 py-3 text-slate-500 shadow-sm"
                                value="Today"
                                readOnly
                              />
                              <div className="mt-4">
                                <label className="mb-1 flex items-center gap-2 text-sm font-medium text-gray-700">
                                  <span>Weekend Handling</span>
                                  <span className="group relative inline-flex h-4 w-4 items-center justify-center rounded-full border border-slate-300 bg-white text-[10px] font-semibold text-slate-500">
                                    i
                                    <span className="pointer-events-none absolute bottom-full left-1/2 z-10 mb-2 hidden w-56 -translate-x-1/2 rounded-lg border border-slate-200 bg-white px-2 py-1 text-[11px] font-normal leading-4 text-slate-600 shadow-lg group-hover:block">
                                      If a computed date lands on Sat/Sun, we either move it to Friday or leave it as-is.
                                    </span>
                                  </span>
                                </label>
                                <select
                                  className="w-full rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm text-slate-900 shadow-sm"
                                  value={weekendRule}
                                  onChange={(e) => setWeekendRule(e.target.value as WeekendRule)}
                                >
                                  <option value="prior_business_day">Adjust to prior business day (Fri)</option>
                                  <option value="none">Allow weekends (no adjustment)</option>
                                </select>
                              </div>
                            </>
                          ) : (
                            <>
                              <input
                                type="date"
                                className="w-full rounded-2xl border border-slate-200 bg-white px-4 py-3 text-slate-900 shadow-sm"
                                value={anchorDate}
                                onChange={(e) => setAnchorDate(e.target.value)}
                              />
                              <div className="mt-4">
                                <label className="mb-1 flex items-center gap-2 text-sm font-medium text-gray-700">
                                  <span>Weekend Handling</span>
                                  <span className="group relative inline-flex h-4 w-4 items-center justify-center rounded-full border border-slate-300 bg-white text-[10px] font-semibold text-slate-500">
                                    i
                                    <span className="pointer-events-none absolute bottom-full left-1/2 z-10 mb-2 hidden w-56 -translate-x-1/2 rounded-lg border border-slate-200 bg-white px-2 py-1 text-[11px] font-normal leading-4 text-slate-600 shadow-lg group-hover:block">
                                      If a computed date lands on Sat/Sun, we either move it to Friday or leave it as-is.
                                    </span>
                                  </span>
                                </label>
                                <select
                                  className="w-full rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm text-slate-900 shadow-sm"
                                  value={weekendRule}
                                  onChange={(e) => setWeekendRule(e.target.value as WeekendRule)}
                                >
                                  <option value="prior_business_day">Adjust to prior business day (Fri)</option>
                                  <option value="none">Allow weekends (no adjustment)</option>
                                </select>
                              </div>
                            </>
                          )}
                        </div>
                        <div>
                          <div className="mb-1 block text-sm font-medium text-transparent select-none" aria-hidden="true">
                            Event Date
                          </div>
                          <label className="inline-flex w-full items-center gap-3 rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm font-medium text-slate-700 shadow-sm">
                            <input
                              type="checkbox"
                              checked={noEventDate}
                              onChange={(e) => setNoEventDate(e.target.checked)}
                            />
                            No Event Date
                          </label>
                        </div>
                      </div>
                    ) : null}
                  </div>
                </div>
                <div className="space-y-4 border-b border-slate-200 pb-6">
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
                ) : null}
                </div>
              </div>

              <div className="space-y-3 border-b border-slate-200 pb-8" ref={kebabMenuRef}>
                <div className="space-y-3">
                  <div className="hidden grid-cols-1 gap-3 px-6 text-[10px] font-semibold uppercase tracking-[0.24em] text-slate-500 md:grid md:grid-cols-[minmax(0,1.8fr)_120px_140px_88px]">
                    <div className="flex min-h-[24px] items-center text-left">Item</div>
                    <div className="flex min-h-[24px] items-center justify-center text-center">
                      {noEventDate ? "Days from today" : "Days from event"}
                    </div>
                    <div className="flex min-h-[24px] items-center justify-center text-center">Time</div>
                    <div className="flex min-h-[24px] items-center justify-center text-center">Actions</div>
                  </div>

                  {rows.map((row, index) => {
                    const rowMeta = getBuilderRowTypeMeta(row);
                    const emailUsesTimingFields = row.rowType === "email" && appSettings.emailHandlingMode === "schedule";
                    const meetingErrors = meetingValidationErrors[row.id];
                    return (
                      <div
                        key={row.id}
                        className={`cursor-pointer space-y-3 rounded-[24px] border border-slate-200 border-l-4 ${rowMeta.borderClass} bg-white p-4 shadow-[0_16px_35px_-30px_rgba(15,23,42,0.55)] transition duration-150 hover:-translate-y-0.5 hover:shadow-[0_22px_45px_-28px_rgba(15,23,42,0.62)] md:p-5`}
                      >
                        <div className="grid grid-cols-1 gap-3 md:grid-cols-[minmax(0,1.8fr)_120px_140px_88px] md:items-start">
                          <div className="relative flex min-h-[72px] items-center">
                            <div className="absolute left-4 top-3 flex items-center gap-2">
                              <span className={`h-2.5 w-2.5 rounded-full ${rowMeta.timelineDotClass}`} />
                              <span className={`inline-flex items-center rounded-full border px-2.5 py-1 text-[11px] font-semibold uppercase tracking-[0.18em] ${rowMeta.badgeClass}`}>
                                {rowMeta.label}
                              </span>
                            </div>
                            <textarea
                              rows={2}
                              className="min-h-[88px] w-full rounded-[20px] border border-slate-200 bg-white px-4 pb-4 pt-11 text-[13px] leading-5 text-slate-900 shadow-sm [overflow-wrap:anywhere]"
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
                              <div className={`flex min-h-[88px] w-full items-center justify-center rounded-[20px] border px-4 py-3 text-center text-[13px] font-medium text-slate-600 ${rowMeta.timingPanelClass}`}>
                                {getBuilderEmailModeMessage(appSettings.emailHandlingMode)}
                              </div>
                            </div>
                          ) : (
                            <>
                              <div className="flex flex-col items-center">
                                <div className={`flex min-h-[88px] w-full items-center rounded-[20px] border px-3 py-2 shadow-sm ${rowMeta.timingPanelClass}`}>
                                  {editingOffsetRowId === row.id ? (
                                    <input
                                      type="text"
                                      inputMode="numeric"
                                      className="w-full border-0 bg-transparent p-0 text-center text-[13px] leading-5 text-slate-900 focus:outline-none focus:ring-0"
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
                                      className="flex min-h-[64px] w-full items-center justify-center bg-transparent p-0 text-center text-[13px] leading-5 text-slate-900"
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
                                  <div className={`flex min-h-[88px] w-full items-center justify-center rounded-[20px] border px-3 py-3 text-center text-[13px] font-medium text-slate-600 ${rowMeta.timingPanelClass}`}>
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
                                    className={`h-[88px] w-full rounded-[20px] border px-4 text-center text-[13px] leading-5 text-slate-900 shadow-sm ${rowMeta.timingPanelClass}`}
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
                                  className={`h-[88px] w-full resize-none rounded-[20px] border px-4 text-center text-[13px] leading-5 text-slate-900 shadow-sm [overflow-wrap:anywhere] ${rowMeta.timingPanelClass} ${
                                    shouldWrapTimeField ? "py-[18px]" : "py-[32px]"
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

                          <div className="relative flex min-h-[88px] items-start justify-center pt-4 md:col-start-4 md:row-start-1 md:text-center">
                            <button
                              type="button"
                              onClick={() => setOpenMenuRowId((current) => (current === row.id ? null : row.id))}
                              title="Actions"
                              aria-label={`Actions for row ${index + 1}`}
                              className="flex h-10 w-10 items-center justify-center rounded-xl border border-slate-200 bg-slate-50 text-slate-500 hover:border-slate-300 hover:bg-white hover:text-slate-700"
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
                                        setOpenBodyEditorRowId(row.id);
                                        setOpenMenuRowId(null);
                                      }}
                                      className="w-full rounded-lg px-3 py-2 text-left text-[12px] hover:bg-gray-50"
                                    >
                                      Edit Meeting
                                    </button>
                                    <div className="my-1 border-t-2 border-double border-gray-200" />
                                  </>
                                ) : null}
                                {row.rowType !== "email" && row.rowType !== "calendar_event" ? (
                                  <>
                                    <button
                                      type="button"
                                      onClick={() => {
                                        setOpenDurationEditorRowId(row.id);
                                        setOpenBodyEditorRowId(row.id);
                                        setOpenMenuRowId(null);
                                      }}
                                      className="w-full rounded-lg px-3 py-2 text-left text-[12px] hover:bg-gray-50"
                                    >
                                      Edit Reminder
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
                                      className="w-full max-w-[220px] rounded-lg border bg-white px-3 py-2 text-sm"
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
                                {(() => {
                                  const normalizedMeetingDraft = normalizeMeetingDraft(row.meetingDraft);
                                  const isGoogleProviderActive = activeAccountProvider === "gmail";
                                  const isProviderManagedMeeting =
                                    normalizedMeetingDraft?.teamsMeeting || (isGoogleProviderActive && normalizedMeetingDraft?.addGoogleMeet);
                                  const locationValue = getMeetingLocationValue(normalizedMeetingDraft, activeAccountProvider);
                                  return (
                                <input
                                  className={`w-full rounded-lg border px-3 py-2 ${isProviderManagedMeeting ? "bg-gray-100 text-gray-600" : "bg-white"}`}
                                  value={locationValue}
                                  readOnly={Boolean(isProviderManagedMeeting)}
                                  onChange={(e) =>
                                    updateRow(row.id, (current) => ({
                                      ...current,
                                      meetingDraft: { ...normalizeMeetingDraft(current.meetingDraft), location: e.target.value },
                                    }))
                                  }
                                />
                                  );
                                })()}
                              </div>
                              <div>
                                <label className="mb-1 block text-sm font-medium text-gray-700">
                                  Meeting Duration
                                  {meetingErrors?.duration ? <span className="ml-1 text-red-500">*</span> : null}
                                </label>
                                <select
                                  className={`w-full max-w-[220px] rounded-lg border bg-white px-3 py-2 text-sm ${
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
                                          const previewMeetingItem = previewPlan.items.find((preview) => preview.id === row.id);
                                          const effectiveStartDate = (previewMeetingItem ? getEffectivePreviewItemDate(previewMeetingItem) : null) || todayYYYYMMDD();
                                          const effectiveStartTime =
                                            getUsableReminderTime(previewMeetingItem?.reminderTime, anchorMap) ||
                                            getUsableReminderTime(row.reminderTime, anchorMap) ||
                                            "09:00";
                                          const defaultCustomEnd = addMinutesToLocalDateTime(
                                            effectiveStartDate,
                                            effectiveStartTime,
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
                              {activeAccountProvider === "gmail" ? (
                                <div className="md:col-span-2 flex items-center gap-2 text-sm text-gray-700">
                                  <input
                                    type="checkbox"
                                    checked={Boolean(normalizeMeetingDraft(row.meetingDraft)?.addGoogleMeet)}
                                    onChange={(e) =>
                                          updateRow(row.id, (current) => ({
                                            ...current,
                                            meetingDraft: {
                                              ...normalizeMeetingDraft(current.meetingDraft),
                                              addGoogleMeet: e.target.checked,
                                              teamsMeeting: false,
                                              location: getNextMeetingLocationOnToggle({
                                                currentLocation: normalizeMeetingDraft(current.meetingDraft)?.location,
                                                checked: e.target.checked,
                                                enabledLocation: GOOGLE_MEET_LOCATION,
                                                disabledLocation: "",
                                              }),
                                            },
                                          }))
                                    }
                                  />
                                  <span>Add Google Meet link</span>
                                </div>
                              ) : (
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
                                          addGoogleMeet: false,
                                          location: getNextMeetingLocationOnToggle({
                                            currentLocation: normalizeMeetingDraft(current.meetingDraft)?.location,
                                            checked: e.target.checked,
                                            enabledLocation: TEAMS_MEETING_LOCATION,
                                            disabledLocation: "",
                                          }),
                                        },
                                      }))
                                    }
                                  />
                                  <span>Microsoft Teams Meeting</span>
                                </div>
                              )}
                              {activeAccountProvider !== "gmail" && normalizeMeetingDraft(row.meetingDraft)?.teamsMeeting ? (
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
                              {activeAccountProvider === "gmail" && normalizeMeetingDraft(row.meetingDraft)?.addGoogleMeet ? (
                                <div className="md:col-span-2 rounded-xl border border-violet-200 bg-white p-4">
                                  <div className="text-sm font-semibold text-violet-950">Google Meet Details</div>
                                  <div className="mt-3 space-y-3 text-sm text-gray-700">
                                    <div>
                                      <div className="mb-1 font-medium text-violet-950">Join link</div>
                                      <div className="text-gray-500">Google Meet link will be generated after export.</div>
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

                      </div>
                    );
                  })}
                </div>
              </div>

              <div className="pt-2">
                <div className="space-y-4">
                  <div className="flex flex-wrap items-center justify-center gap-3 rounded-[24px] border border-slate-200 bg-[linear-gradient(180deg,rgba(255,255,255,0.96)_0%,rgba(248,250,252,0.96)_100%)] p-5 shadow-[0_16px_38px_-28px_rgba(15,23,42,0.5)]">
                    <button
                      type="button"
                      onClick={addReminderRow}
                      className="rounded-2xl border border-blue-200 bg-white px-5 py-3 text-sm font-medium text-blue-700 shadow-md transition duration-150 hover:-translate-y-1 hover:bg-blue-50 hover:shadow-lg"
                    >
                      + Add a Reminder
                    </button>
                    <button
                      type="button"
                      onClick={addEmailRow}
                      className="rounded-2xl border border-green-200 bg-white px-5 py-3 text-sm font-medium text-green-700 shadow-md transition duration-150 hover:-translate-y-1 hover:bg-green-50 hover:shadow-lg"
                    >
                      + Add an Email
                    </button>
                    <button
                      type="button"
                      onClick={addMeetingRow}
                      className="rounded-2xl border border-violet-200 bg-white px-5 py-3 text-sm font-medium text-violet-700 shadow-md transition duration-150 hover:-translate-y-1 hover:bg-violet-50 hover:shadow-lg"
                    >
                      + Add a Meeting
                    </button>
                  </div>

                  <div className="space-y-4">
                    <div className="text-lg font-semibold text-gray-900">Dynamic Fields</div>
                    <div className="flex flex-wrap items-center gap-2">
                      <button
                        type="button"
                        onClick={() => {
                          setAnchors((current) => [...current, createEmptyAnchor()]);
                          setAreAnchorsHidden(false);
                        }}
                        className="rounded-lg border px-3 py-2 text-sm hover:bg-gray-50"
                      >
                        + Add Dynamic Field
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
                        {areAnchorsHidden ? "Show Dynamic Fields" : "Hide Dynamic Fields"}
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

                  <div className="space-y-5 border-t border-slate-200 pt-6">
                    {!shouldRenderSimpleEventHeaderFields ? (
                      <div className="max-w-md">
                        <label className="mb-1 flex items-center gap-2 text-sm font-medium text-gray-700">
                          <span>Weekend Handling</span>
                          <span className="group relative inline-flex h-4 w-4 items-center justify-center rounded-full border border-slate-300 bg-white text-[10px] font-semibold text-slate-500">
                            i
                            <span className="pointer-events-none absolute bottom-full left-1/2 z-10 mb-2 hidden w-56 -translate-x-1/2 rounded-lg border border-slate-200 bg-white px-2 py-1 text-[11px] font-normal leading-4 text-slate-600 shadow-lg group-hover:block">
                              If a computed date lands on Sat/Sun, we either move it to Friday or leave it as-is.
                            </span>
                          </span>
                        </label>
                        <select
                          className="w-full rounded-lg border px-3 py-2"
                          value={weekendRule}
                          onChange={(e) => setWeekendRule(e.target.value as WeekendRule)}
                        >
                          <option value="prior_business_day">Adjust to prior business day (Fri)</option>
                          <option value="none">Allow weekends (no adjustment)</option>
                        </select>
                      </div>
                    ) : null}

                    <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
                      <div className="flex flex-wrap items-center gap-3">
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
                          onClick={cancelEditing}
                          className="rounded-lg border px-4 py-2 text-sm hover:bg-gray-50"
                        >
                          Cancel
                        </button>
                      </div>
                      <div className="flex flex-wrap items-center gap-3">
                        <button
                          type="button"
                          onClick={saveCurrentTemplate}
                          className="rounded-lg border px-4 py-2 text-sm hover:bg-gray-50"
                        >
                          Save
                        </button>
                        {shouldRenderBuilderSourceBanner ? (
                          <button
                            type="button"
                            onClick={openBuilderTemplateSaveDialog}
                            className="rounded-lg border px-4 py-2 text-sm hover:bg-gray-50"
                          >
                            Save As
                          </button>
                        ) : null}
                        <button
                          type="button"
                          onClick={() => {
                            void exportCurrentPlan();
                          }}
                          className="rounded-lg bg-blue-600 px-4 py-2 text-white hover:bg-blue-700"
                        >
                          Export
                        </button>
                        {executionNotices.some((entry) => entry.notice.tone === "success") ? (
                          <div className="flex items-center">
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
              {executionNotices.length > 0 ? (
                <div className="space-y-3">
                  {executionNotices.map((entry) => (
                    <OutlookExecutionNoticeCard
                      key={entry.id}
                      notice={entry.notice}
                      onDismiss={() => dismissExecutionNotice(entry.id)}
                    />
                  ))}
                </div>
              ) : null}
              {previewPlanForRender ? (
                <>
                  <div className="rounded-xl border">
                    <div className="border-b bg-gray-50 px-4 py-3">
                      <div className="font-semibold text-gray-900">{previewPlanForRender.name}</div>
                      {previewLoading ? <div className="mt-1 text-xs text-gray-500">Refreshing preview…</div> : null}
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
                      signature: appSettings.emailSignatureText,
                    });
                    const previewMeetingDraft = normalizeMeetingDraft(item.meetingDraft);
                    const isPreviewTeamsMeetingEnabled = Boolean(previewMeetingDraft?.teamsMeeting);
                    const isPreviewGoogleMeetEnabled = Boolean(previewMeetingDraft?.addGoogleMeet);
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
                                    void exportPreviewEmailItem(item.id);
                                    return;
                                  }
                                  if (hasMeetingPreview) {
                                    void exportPreviewMeetingItem(item.id);
                                    return;
                                  }
                                  void exportPreviewReminderItem(item.id);
                                }}
                                disabled={executionState === "pending"}
                                title={
                                  hasReminderPreview
                                    ? "Export this reminder"
                                    : hasEmailPreview
                                      ? "Execute this email action"
                                      : "Create this meeting event"
                                }
                                className="flex h-14 w-44 items-center justify-center rounded-lg border px-3 py-2 text-center text-sm hover:bg-gray-50 disabled:opacity-60"
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
                              {activeAccountProvider === "gmail"
                                ? "This meeting will be created in Google Calendar when Google is connected."
                                : "This meeting will be created in Outlook when connected."}
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
                                {(() => {
                                  const isGoogleProviderActive = activeAccountProvider === "gmail";
                                  const isProviderManagedMeeting =
                                    isPreviewTeamsMeetingEnabled || (isGoogleProviderActive && isPreviewGoogleMeetEnabled);
                                  const locationValue = getMeetingLocationValue(previewMeetingDraft, activeAccountProvider);
                                  return (
                                <input
                                  className={`w-full rounded-lg border border-violet-200 px-3 py-2 text-sm ${
                                    isProviderManagedMeeting ? "bg-gray-100 text-gray-600" : "bg-white text-gray-900"
                                  }`}
                                  value={locationValue}
                                  readOnly={isProviderManagedMeeting}
                                  onChange={(e) =>
                                    builderItem
                                      ? updateRow(builderItem.id, (current) => ({
                                          ...current,
                                          meetingDraft: { ...normalizeMeetingDraft(current.meetingDraft), location: e.target.value },
                                        }))
                                      : undefined
                                  }
                                />
                                  );
                                })()}
                              </div>
                              <div>
                                <label className="mb-1 block text-sm font-medium text-violet-950">Meeting Duration</label>
                                <select
                                  className="w-full max-w-[220px] rounded-lg border border-violet-200 bg-white px-3 py-2 text-sm text-gray-900"
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
                              {activeAccountProvider === "gmail" ? (
                                <div className="md:col-span-2 flex items-center gap-2 text-sm text-violet-950">
                                  <input
                                    type="checkbox"
                                    checked={Boolean(previewMeetingDraft?.addGoogleMeet)}
                                    onChange={(e) =>
                                      builderItem
                                        ? updateRow(builderItem.id, (current) => ({
                                            ...current,
                                            meetingDraft: {
                                              ...normalizeMeetingDraft(current.meetingDraft),
                                              addGoogleMeet: e.target.checked,
                                              teamsMeeting: false,
                                              location: getNextMeetingLocationOnToggle({
                                                currentLocation: normalizeMeetingDraft(current.meetingDraft)?.location,
                                                checked: e.target.checked,
                                                enabledLocation: GOOGLE_MEET_LOCATION,
                                                disabledLocation: "",
                                              }),
                                            },
                                          }))
                                        : undefined
                                    }
                                  />
                                  <span>Add Google Meet link</span>
                                </div>
                              ) : (
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
                                              addGoogleMeet: false,
                                              location: getNextMeetingLocationOnToggle({
                                                currentLocation: normalizeMeetingDraft(current.meetingDraft)?.location,
                                                checked: e.target.checked,
                                                enabledLocation: TEAMS_MEETING_LOCATION,
                                                disabledLocation: "",
                                              }),
                                            },
                                          }))
                                        : undefined
                                    }
                                  />
                                  <span>Microsoft Teams Meeting</span>
                                </div>
                              )}
                              {activeAccountProvider !== "gmail" && isPreviewTeamsMeetingEnabled ? (
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
                              {activeAccountProvider === "gmail" && isPreviewGoogleMeetEnabled ? (
                                <div className="md:col-span-2 rounded-xl border border-violet-200 bg-white p-4">
                                  <div className="text-sm font-semibold text-violet-950">Google Meet Details</div>
                                  <div className="mt-3 space-y-3 text-sm text-gray-700">
                                    <div>
                                      <div className="mb-1 font-medium text-violet-950">Join link</div>
                                      <div className="text-gray-500">Google Meet link will be generated after export.</div>
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
                          void exportCurrentPlan({ skipConfirm: true });
                        }}
                        disabled={executionState === "pending"}
                        className="rounded-lg bg-blue-600 px-4 py-2 text-white hover:bg-blue-700 disabled:opacity-60"
                      >
                        {executionState === "pending" ? "Exporting..." : "Export"}
                      </button>
                      {executionState === "success" ? <ExportDoneBadge /> : null}
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

      {AI_ENABLED && isAiPanelOpen ? (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/35 px-4 py-8">
          <div className="max-h-[90vh] w-full max-w-6xl overflow-hidden rounded-2xl border bg-white shadow-xl">
            <div className="flex items-center justify-between border-b px-6 py-4">
              <div>
                <h2 className="text-lg font-semibold text-gray-900">Generate Plan with AI</h2>
                <p className="mt-1 text-sm text-gray-600">
                  {aiSessionSource?.type === "current_builder"
                    ? "Refine your current plan through chat, then review the revised draft before applying it."
                    : aiSessionSource?.type === "saved_template"
                      ? "Refine a saved template through chat, then review the revised draft before applying or saving."
                      : aiSessionSource?.type === "branched_draft"
                        ? "Explore a different version of your draft without affecting the earlier one."
                        : "Chat through the workflow you need, then apply the generated plan into the builder."}
                </p>
              </div>
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={onDuplicateAiDraftIntoNewExploration}
                  disabled={aiGenerating || !aiChatDraft}
                  className="rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-700 hover:bg-gray-50 disabled:opacity-50"
                >
                  Try a different version
                </button>
                <button
                  type="button"
                  onClick={onStartOverAiSession}
                  disabled={aiGenerating}
                  className="rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-700 hover:bg-gray-50 disabled:opacity-60"
                >
                  Start over
                </button>
                <button
                  type="button"
                  onClick={() => setIsAiPanelOpen(false)}
                  className="rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900 hover:bg-gray-50"
                >
                  Close
                </button>
              </div>
            </div>

            <div className="grid h-[calc(90vh-81px)] grid-cols-1 divide-y md:grid-cols-[minmax(0,1.4fr)_minmax(320px,0.9fr)] md:divide-x md:divide-y-0">
              <div className="flex min-h-0 flex-col">
                <div ref={aiConversationRef} className="flex-1 space-y-4 overflow-y-auto p-6">
                  <div className={`rounded-xl border px-4 py-3 text-sm ${aiSessionSourceDetails.classes}`}>
                    <div className="text-xs font-semibold uppercase tracking-wide">Working from</div>
                    <div className="mt-1 font-medium">{aiSessionSourceDetails.label}</div>
                    <div className="mt-1 text-xs opacity-80">{aiSessionSourceDetails.note}</div>
                  </div>
                  {aiSessionBackup && !hasMeaningfulAiSession() ? (
                    <div className="rounded-xl border border-gray-200 bg-white px-4 py-3 text-sm text-gray-700">
                      <div className="flex flex-wrap items-center justify-between gap-3">
                        <div>
                          <div className="font-medium text-gray-900">Last AI draft cleared</div>
                          <div className="mt-1 text-xs text-gray-500">
                            You can restore the last conversation and draft from this modal session.
                          </div>
                        </div>
                        <button
                          type="button"
                          onClick={onRestoreAiSession}
                          className="rounded-lg border border-gray-300 bg-white px-3 py-2 text-xs font-medium text-gray-900 hover:bg-gray-50"
                        >
                          Restore last draft
                        </button>
                      </div>
                    </div>
                  ) : null}
                  {aiChatMessages.map((message) => (
                    <div
                      key={message.id}
                      className={`flex ${message.role === "user" ? "justify-end" : "justify-start"}`}
                    >
                      <div
                        className={`max-w-[85%] rounded-2xl px-4 py-3 text-sm shadow-sm ${
                          message.role === "user"
                            ? "bg-blue-600 text-white"
                            : "border border-gray-200 bg-gray-50 text-gray-900"
                        }`}
                      >
                        <div className="whitespace-pre-wrap">{message.text}</div>
                        {message.status ? (
                          <div className="mt-3 text-xs font-medium uppercase tracking-wide text-gray-500">
                            {getAiReadinessLabel(message.status)}
                          </div>
                        ) : null}
                        {message.followUpQuestions?.length ? (
                          <div className="mt-3 space-y-1 rounded-xl border border-gray-200 bg-white/70 px-3 py-2 text-xs text-gray-700">
                            <div className="font-semibold uppercase tracking-wide text-gray-500">Still helpful to know</div>
                            {message.followUpQuestions.map((question) => (
                              <div key={question}>- {question}</div>
                            ))}
                          </div>
                        ) : null}
                        {message.changeSummary?.length ? (
                          <div className="mt-3 space-y-1 rounded-xl border border-gray-200 bg-white/70 px-3 py-2 text-xs text-gray-700">
                            <div className="font-semibold uppercase tracking-wide text-gray-500">What changed</div>
                            {message.changeSummary.map((item) => (
                              <div key={item}>- {item}</div>
                            ))}
                          </div>
                        ) : null}
                        {message.confidenceNote ? (
                          <div className="mt-3 text-xs text-gray-500">Confidence: {message.confidenceNote}</div>
                        ) : null}
                        {message.suggestedNextActions?.length ? (
                          <div className="mt-3 flex flex-wrap gap-2">
                            {message.suggestedNextActions.map((action) => (
                              <button
                                key={action}
                                type="button"
                                onClick={() => {
                                  void onSendAiMessage(action);
                                }}
                                disabled={aiGenerating}
                                className="rounded-full border border-blue-200 bg-blue-50 px-3 py-1.5 text-xs text-blue-800 hover:bg-blue-100 disabled:opacity-60"
                              >
                                {action}
                              </button>
                            ))}
                          </div>
                        ) : null}
                        {message.starterPrompts?.length ? (
                          <div className="mt-3 flex flex-wrap gap-2">
                            {message.starterPrompts.map((prompt) => (
                              <button
                                key={prompt}
                                type="button"
                                onClick={() => {
                                  void onSendAiMessage(prompt);
                                }}
                                className="rounded-full border border-gray-300 bg-white px-3 py-1.5 text-xs text-gray-700 hover:bg-gray-100"
                              >
                                {prompt}
                              </button>
                            ))}
                            <button
                              type="button"
                              onClick={focusAiComposer}
                              className="rounded-full border border-dashed border-gray-300 bg-white px-3 py-1.5 text-xs text-gray-500 hover:bg-gray-50"
                            >
                              Something else…
                            </button>
                          </div>
                        ) : null}
                        {message.modeOptions?.length ? (
                          <div className="mt-3 flex flex-wrap gap-2">
                            {message.modeOptions.map((option) => (
                              <button
                                key={option.id}
                                type="button"
                                onClick={() => onSelectAiBuilderMode(option.id)}
                                disabled={aiGenerating}
                                className="rounded-full border border-gray-300 bg-white px-3 py-1.5 text-xs text-gray-700 hover:bg-gray-100 disabled:opacity-60"
                              >
                                {option.label}
                              </button>
                            ))}
                          </div>
                        ) : null}
                      </div>
                    </div>
                  ))}
                  {aiGenerating ? (
                    <div className="flex justify-start">
                      <div className="rounded-2xl border border-gray-200 bg-gray-50 px-4 py-3 text-sm text-gray-600 shadow-sm">
                        Thinking…
                      </div>
                    </div>
                  ) : null}
                </div>

                <div className="border-t p-4">
                  {aiChatError ? (
                    <div className="mb-3 rounded-xl border border-red-200 bg-red-50 px-4 py-3 text-sm text-red-800">
                      {aiChatError}
                    </div>
                  ) : null}
                  <div className="flex flex-col gap-3">
                    <textarea
                      ref={aiComposerRef}
                      rows={4}
                      value={aiComposer}
                      onChange={(e) => setAiComposer(e.target.value)}
                      className="w-full rounded-xl border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900"
                      placeholder={
                        aiBuilderContextMode === "refine_current"
                          ? "Describe how you want to change the current plan. For example: move reminders earlier, remove the email, or add a prep meeting."
                          : "Describe the event workflow you need, or answer the assistant’s follow-up question here."
                      }
                    />
                    <div className="flex items-center justify-between gap-3">
                      <div className="text-xs text-gray-500">
                        The builder stays unchanged until you click Apply to Builder.
                      </div>
                      <button
                        type="button"
                        onClick={() => {
                          void onSendAiMessage();
                        }}
                        disabled={aiGenerating}
                        className="rounded-lg bg-blue-600 px-4 py-2 text-sm font-medium text-white hover:bg-blue-700 disabled:opacity-60"
                      >
                        {aiGenerating ? "Sending..." : "Send"}
                      </button>
                    </div>
                  </div>
                </div>
              </div>

              <div className="min-h-0 overflow-y-auto bg-gray-50/60 p-6">
                <div className="space-y-5">
                  {aiSavedTemplateInfo ? (
                    <div className="rounded-xl border border-green-200 bg-green-50 px-4 py-4 text-sm text-green-900">
                      <div className="font-medium">Saved &quot;{aiSavedTemplateInfo.name}&quot; as a custom template.</div>
                      <div className="mt-1 text-xs text-green-800">
                        You can keep refining this draft, apply it to the builder, or use the saved template later from Templates.
                      </div>
                      <div className="mt-3 flex flex-wrap gap-2">
                        <button
                          type="button"
                          onClick={() => setAiSavedTemplateInfo(null)}
                          className="rounded-lg border border-green-300 bg-white px-3 py-2 text-xs font-medium text-green-900 hover:bg-green-100"
                        >
                          Keep editing this draft
                        </button>
                        <button
                          type="button"
                          onClick={onApplyAiDraft}
                          className="rounded-lg bg-blue-600 px-3 py-2 text-xs font-medium text-white hover:bg-blue-700"
                        >
                          Apply this draft to builder
                        </button>
                      </div>
                    </div>
                  ) : null}
                  <div>
                    <div className="text-xs font-semibold uppercase tracking-wide text-gray-500">Current Plan Summary</div>
                    <div className="mt-2 rounded-xl border bg-white p-4 text-sm text-gray-900">
                      {aiChatSummary || "The assistant will keep a running summary here as the plan takes shape."}
                    </div>
                  </div>

                  {aiSessionSource ? (
                    <div className="rounded-xl border bg-white p-4">
                      <div className="text-xs font-semibold uppercase tracking-wide text-gray-500">Working from</div>
                      <div className="mt-2 text-sm font-medium text-gray-900">{aiSessionSourceDetails.label}</div>
                      <div className="mt-2 text-xs text-gray-500">{aiSessionSourceDetails.note}</div>
                    </div>
                  ) : null}

                  {aiChatDraft ? (
                    <div className="rounded-xl border bg-white p-4">
                      <div className="text-xs font-semibold uppercase tracking-wide text-gray-500">Current Draft</div>
                      <div className="mt-3 space-y-1 text-sm text-gray-700">
                        <div>
                          Source: <span className="text-gray-900">{aiDraftIdentityDetails.source}</span>
                        </div>
                        <div>
                          Current draft: <span className="text-gray-900">{aiDraftIdentityDetails.currentDraft}</span>
                        </div>
                        <div>
                          Apply destination: <span className="text-gray-900">{aiDraftIdentityDetails.applyDestination}</span>
                        </div>
                        <div>
                          Save destination: <span className="text-gray-900">{aiDraftIdentityDetails.saveDestination}</span>
                        </div>
                      </div>
                      <div className="mt-3 text-xs text-gray-500">{aiDraftIdentityDetails.note}</div>
                    </div>
                  ) : null}

                  <div className="rounded-xl border bg-white p-4">
                    <div className="text-xs font-semibold uppercase tracking-wide text-gray-500">Draft Status</div>
                    <div className="mt-2 text-sm font-medium text-gray-900">{aiDraftStageDetails.stage}</div>
                    <div className="mt-2 text-xs font-semibold uppercase tracking-wide text-gray-500">Recommended next step</div>
                    <div className="mt-1 text-xs text-gray-500">{aiDraftStageDetails.nextStep}</div>
                  </div>

                  {aiChatDraft ? (
                    <div className="rounded-xl border bg-white p-4">
                      <div className="text-xs font-semibold uppercase tracking-wide text-gray-500">Missing / Assumed Details</div>
                      {aiDraftMissingDetails.length ? (
                        <div className="mt-2 space-y-1 text-sm text-gray-700">
                          {aiDraftMissingDetails.map((detail) => (
                            <div key={detail}>- {detail}</div>
                          ))}
                        </div>
                      ) : (
                        <div className="mt-2 text-sm text-gray-500">No major gaps detected.</div>
                      )}
                    </div>
                  ) : null}

                  {aiBuilderContextMode === "refine_current" && (aiChatChangeSummary.length > 0 || aiChatConfidenceNote) ? (
                    <div className="rounded-xl border bg-white p-4">
                      <div className="text-xs font-semibold uppercase tracking-wide text-gray-500">What Changed</div>
                      {aiChatChangeSummary.length ? (
                        <div className="mt-2 space-y-1 text-sm text-gray-700">
                          {aiChatChangeSummary.map((item) => (
                            <div key={item}>- {item}</div>
                          ))}
                        </div>
                      ) : (
                        <div className="mt-2 text-sm text-gray-500">The assistant is still tightening the revised draft.</div>
                      )}
                      {aiChatConfidenceNote ? (
                        <div className="mt-3 text-xs text-gray-500">Confidence: {aiChatConfidenceNote}</div>
                      ) : null}
                    </div>
                  ) : null}

                  {aiChatDraft ? (
                    <>
                      {aiDraftBaseline ? (() => {
                        const comparison = getAiDraftComparisonSummary(aiChatDraft, aiDraftBaseline);
                        return (
                          <div className="rounded-xl border bg-white p-4">
                            <div className="text-xs font-semibold uppercase tracking-wide text-gray-500">Compared with Starting Point</div>
                            <div className="mt-2 text-xs text-gray-500">
                              Based on your {aiDraftBaseline.sourceLabel}.
                            </div>
                            {comparison.qualitativeLabel ? (
                              <div className="mt-2 text-sm font-medium text-gray-900">{comparison.qualitativeLabel}</div>
                            ) : null}
                            <div className="mt-3 space-y-1 text-sm text-gray-700">
                              <div>
                                Rows: <span className="text-gray-900">{getBaselineDeltaLabel(comparison.totalDelta)}</span>
                              </div>
                              <div>
                                Reminders: <span className="text-gray-900">{getBaselineDeltaLabel(comparison.reminderDelta)}</span>
                              </div>
                              <div>
                                Emails: <span className="text-gray-900">{getBaselineDeltaLabel(comparison.emailDelta)}</span>
                              </div>
                              <div>
                                Meetings: <span className="text-gray-900">{getBaselineDeltaLabel(comparison.meetingDelta)}</span>
                              </div>
                              <div>
                                Event timing: <span className="text-gray-900">{comparison.timingChanged ? "Changed" : "Unchanged"}</span>
                              </div>
                            </div>
                          </div>
                        );
                      })() : null}

                      <div className="rounded-xl border bg-white p-4">
                        <div className="text-xs font-semibold uppercase tracking-wide text-gray-500">Draft Details</div>
                        <div className="mt-2 space-y-1 text-sm text-gray-700">
                          <div>Event name: <span className="text-gray-900">{aiChatDraft.eventName || "—"}</span></div>
                          <div>Template name: <span className="text-gray-900">{aiChatDraft.templateName || "—"}</span></div>
                          <div>Event date: <span className="text-gray-900">{aiChatDraft.noEventDate ? "No event date" : aiChatDraft.anchorDate || "—"}</span></div>
                          <div>Event time: <span className="text-gray-900">{aiChatDraft.eventTime || "—"}</span></div>
                          <div>Plan type: <span className="text-gray-900">{getSeedTemplateName(aiChatDraft.baseType)}</span></div>
                          <div>Weekend handling: <span className="text-gray-900">{aiChatDraft.weekendRule === "none" ? "Allow weekends" : "Prior business day"}</span></div>
                        </div>
                      </div>

                      <div className="space-y-3">
                        <div className="text-xs font-semibold uppercase tracking-wide text-gray-500">Generated Rows</div>
                        {aiChatDraft.rows.map((row, index) => (
                          <div key={`${row.title}-${index}`} className="rounded-xl border bg-white p-4">
                            <div className="flex flex-wrap items-start justify-between gap-3">
                              <div className="min-w-0 flex-1">
                                <div className="text-xs font-medium uppercase tracking-wide text-gray-500">
                                  {row.rowType === "email" ? "Email" : row.rowType === "calendar_event" ? "Meeting" : "Reminder"}
                                </div>
                              </div>
                              <button
                                type="button"
                                onClick={() => removeAiDraftRow(index)}
                                className="rounded-lg border border-red-200 bg-white px-2.5 py-1.5 text-xs font-medium text-red-700 hover:bg-red-50"
                              >
                                Remove
                              </button>
                            </div>

                            <div className="mt-3 grid gap-3 md:grid-cols-2">
                              <div className="md:col-span-2">
                                <label className="mb-1 block text-xs font-semibold uppercase tracking-wide text-gray-500">Title</label>
                                <input
                                  value={row.title}
                                  onChange={(e) =>
                                    updateAiDraftRow(index, (current) => ({
                                      ...current,
                                      title: e.target.value,
                                    }))
                                  }
                                  className="w-full rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900"
                                />
                              </div>

                              <div>
                                <label className="mb-1 block text-xs font-semibold uppercase tracking-wide text-gray-500">Offset</label>
                                <input
                                  type="number"
                                  value={row.offsetDays}
                                  onChange={(e) =>
                                    updateAiDraftRow(index, (current) => ({
                                      ...current,
                                      offsetDays: Number.isFinite(Number(e.target.value)) ? Number(e.target.value) : 0,
                                    }))
                                  }
                                  className="w-full rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900"
                                />
                                <div className="mt-1 text-xs text-gray-500">
                                  {formatOffsetLabel(row.offsetDays, {
                                    relativeToToday: aiChatDraft.noEventDate,
                                    dateBasis: row.dateBasis,
                                  })}
                                </div>
                              </div>

                              <div>
                                <label className="mb-1 block text-xs font-semibold uppercase tracking-wide text-gray-500">Time</label>
                                <input
                                  value={row.reminderTime ?? ""}
                                  onChange={(e) =>
                                    updateAiDraftRow(index, (current) => ({
                                      ...current,
                                      reminderTime: e.target.value,
                                    }))
                                  }
                                  className="w-full rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900"
                                  placeholder="Optional time"
                                />
                              </div>
                            </div>

                            {row.rowType === "email" ? (
                              <div className="mt-3">
                                <label className="mb-1 block text-xs font-semibold uppercase tracking-wide text-gray-500">Email Subject</label>
                                <input
                                  value={row.emailDraft?.subject ?? ""}
                                  onChange={(e) =>
                                    updateAiDraftRow(index, (current) => ({
                                      ...current,
                                      emailDraft: {
                                        to: current.emailDraft?.to ?? [],
                                        cc: current.emailDraft?.cc ?? [],
                                        bcc: current.emailDraft?.bcc ?? [],
                                        subject: e.target.value,
                                        body: current.emailDraft?.body ?? "",
                                      },
                                    }))
                                  }
                                  className="w-full rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900"
                                />
                              </div>
                            ) : null}

                            {row.body ? (
                              <div className="mt-3">
                                <div className="mb-1 text-xs font-semibold uppercase tracking-wide text-gray-500">Notes</div>
                                <div className="whitespace-pre-wrap text-sm text-gray-700">{row.body}</div>
                              </div>
                            ) : null}

                            {row.rationale ? (
                              <div className="mt-2 text-xs text-gray-500">{row.rationale}</div>
                            ) : null}
                          </div>
                        ))}
                      </div>

                      <div className="flex flex-wrap items-center gap-3">
                        <button
                          type="button"
                          onClick={openAiTemplateSaveDialog}
                          className="rounded-lg border border-gray-300 bg-white px-4 py-2 text-sm font-medium text-gray-900 hover:bg-gray-50"
                        >
                          Save as template
                        </button>
                        <button
                          type="button"
                          onClick={onApplyAiDraft}
                          className="rounded-lg bg-blue-600 px-4 py-2 text-sm font-medium text-white hover:bg-blue-700"
                        >
                          Apply to Builder
                        </button>
                      </div>
                    </>
                  ) : (
                    <div className="rounded-xl border border-dashed bg-white px-4 py-8 text-sm text-gray-500">
                      Start the conversation on the left and the current draft will appear here as soon as the assistant can shape one.
                    </div>
                  )}
                </div>
              </div>
            </div>
          </div>
        </div>
      ) : null}

      {AI_ENABLED && showAiApplyConfirm && aiChatDraft ? (
        <div className="fixed inset-0 z-[60] flex items-center justify-center bg-black/45 px-4">
          <div className="w-full max-w-md rounded-2xl border bg-white p-6 shadow-xl">
            <h3 className="text-lg font-semibold text-gray-900">Replace current builder?</h3>
            <p className="mt-2 text-sm text-gray-600">
              Applying this AI draft will replace what is currently in the builder.
            </p>
            <div className="mt-4 rounded-xl border bg-gray-50 p-4 text-sm text-gray-700">
              <div>Current builder rows: <span className="font-medium text-gray-900">{rows.length}</span></div>
              <div className="mt-1">Incoming AI draft rows: <span className="font-medium text-gray-900">{aiChatDraft.rows.length}</span></div>
            </div>
            <div className="mt-5 flex flex-wrap justify-end gap-3">
              <button
                type="button"
                onClick={() => setShowAiApplyConfirm(false)}
                className="rounded-lg border border-gray-300 bg-white px-4 py-2 text-sm text-gray-900 hover:bg-gray-50"
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={applyAiDraftToBuilder}
                className="rounded-lg bg-blue-600 px-4 py-2 text-sm font-medium text-white hover:bg-blue-700"
              >
                Replace current builder
              </button>
            </div>
          </div>
        </div>
      ) : null}

      {AI_ENABLED && showAiTemplateSaveDialog && aiChatDraft ? (
        <div className="fixed inset-0 z-[60] flex items-center justify-center bg-black/45 px-4">
          <div className="w-full max-w-md rounded-2xl border bg-white p-6 shadow-xl">
            <h3 className="text-lg font-semibold text-gray-900">Save AI draft as template</h3>
            <p className="mt-2 text-sm text-gray-600">
              Save this draft as a reusable custom template. Your current builder will stay unchanged.
            </p>
            <div className="mt-4">
              <label className="mb-1 block text-sm font-medium text-gray-700">Template name</label>
              <input
                value={aiTemplateNameDraft}
                onChange={(e) => setAiTemplateNameDraft(e.target.value)}
                className="w-full rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900"
                placeholder="Template name"
              />
            </div>
            {aiTemplateSaveMessage ? (
              <div className="mt-3 rounded-lg border border-gray-200 bg-gray-50 px-3 py-2 text-sm text-gray-700">
                {aiTemplateSaveMessage}
              </div>
            ) : null}
            <div className="mt-5 flex flex-wrap justify-end gap-3">
              <button
                type="button"
                onClick={() => setShowAiTemplateSaveDialog(false)}
                className="rounded-lg border border-gray-300 bg-white px-4 py-2 text-sm text-gray-900 hover:bg-gray-50"
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={saveAiDraftAsTemplate}
                className="rounded-lg bg-blue-600 px-4 py-2 text-sm font-medium text-white hover:bg-blue-700"
              >
                Save template
              </button>
            </div>
          </div>
        </div>
      ) : null}

      {showBuilderTemplateSaveDialog && hasMeaningfulBuilderContent() ? (
        <div className="fixed inset-0 z-[60] flex items-center justify-center bg-black/45 px-4">
          <div className="w-full max-w-md rounded-2xl border bg-white p-6 shadow-xl">
            <h3 className="text-lg font-semibold text-gray-900">Save current plan as template</h3>
            <p className="mt-2 text-sm text-gray-600">
              Create a new reusable custom template from the current builder. The builder will stay unchanged.
            </p>
            <div className="mt-4">
              <label className="mb-1 block text-sm font-medium text-gray-700">Template name</label>
              <input
                value={builderTemplateNameDraft}
                onChange={(e) => setBuilderTemplateNameDraft(e.target.value)}
                className="w-full rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900"
                placeholder="Template name"
              />
            </div>
            {builderTemplateSaveMessage ? (
              <div className="mt-3 rounded-lg border border-gray-200 bg-gray-50 px-3 py-2 text-sm text-gray-700">
                {builderTemplateSaveMessage}
              </div>
            ) : null}
            <div className="mt-5 flex flex-wrap justify-end gap-3">
              <button
                type="button"
                onClick={() => setShowBuilderTemplateSaveDialog(false)}
                className="rounded-lg border border-gray-300 bg-white px-4 py-2 text-sm text-gray-900 hover:bg-gray-50"
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={saveCurrentBuilderAsTemplate}
                className="rounded-lg bg-blue-600 px-4 py-2 text-sm font-medium text-white hover:bg-blue-700"
              >
                Save template
              </button>
            </div>
          </div>
        </div>
      ) : null}
    </div>
  );
}
