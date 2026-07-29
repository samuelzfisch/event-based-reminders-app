"use client";

import Link from "next/link";
import {
  type CSSProperties,
  type HTMLAttributes,
  type KeyboardEvent as ReactKeyboardEvent,
  type MouseEvent as ReactMouseEvent,
  type ReactNode,
  useCallback,
  useEffect,
  useId,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
} from "react";
import { createPortal } from "react-dom";

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
  type PersistedTemplateAnchor,
  type PersistedPlanTemplate,
  type PersistedTemplateState,
} from "../../lib/templateStore";
import { readPersistedValue, removePersistedValue, writePersistedValue } from "../../lib/browserStorage";
import { getScopedStorageKey } from "../../lib/clientPersistence";
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
import { AppShellSidebarContent } from "../components/app-shell";
import { RecipientGroupsModal } from "../components/recipient-groups-modal";
import { useAuthContext } from "../components/auth-provider";
import {
  createEmailRecipientEntry,
  createGroupRecipientEntry,
  deleteRecipientGroup,
  getRecipientGroupFromEntry,
  hydrateRecipientGroupsFromSupabase,
  mergeRecipientEntries,
  normalizeRecipientEntries,
  normalizeRecipientGroupEmail,
  RECIPIENT_GROUPS_UPDATED_EVENT,
  resolveRecipientEntries,
  saveRecipientGroup,
  type RecipientGroup,
  type RecipientEntry,
} from "../../lib/recipientGroups";

type BuilderEmailDraft = {
  to?: Array<RecipientEntry | string>;
  cc?: Array<RecipientEntry | string>;
  bcc?: Array<RecipientEntry | string>;
  subject?: string;
  body?: string;
};

type BuilderAnchor = {
  id: string;
  key: string;
  value: string;
  locked?: boolean;
  isImportant?: boolean;
  lastUpdatedAt?: string | null;
};

type BuilderRow = {
  id: string;
  title: string;
  body?: string;
  offsetDays: number | null;
  dateBasis?: PlanDateBasis;
  rowType: PlanRowType;
  reminderTime?: string;
  timeZone?: string;
  emailDraft?: BuilderEmailDraft;
  durationDraft?: {
    durationMinutes?: number;
    useCustomEnd?: boolean;
    endDate?: string;
    endTime?: string;
    isAllDay?: boolean;
  };
  meetingDraft?: {
    attendees?: Array<RecipientEntry | string>;
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

type RecipientGroupTargetField = "email_to" | "meeting_attendees";

type RecipientGroupsModalTarget = {
  rowId: string;
  field: RecipientGroupTargetField;
};

type SavedPlanTemplate = {
  id: string;
  name: string;
  baseType: PlanType;
  templateMode?: "template" | "custom";
  noEventDate?: boolean;
  weekendRule: WeekendRule;
  anchors: PersistedTemplateAnchor[];
  items: BuilderRow[];
  lastDynamicFieldsExportAt?: string | null;
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
  hasExplicitEventDate: boolean;
  eventTime: string;
  eventTimeZone: string;
  noEventDate: boolean;
  weekendRule: WeekendRule;
  rows: BuilderRow[];
  anchors: BuilderAnchor[];
  guidedForm: GuidedFormState;
  lastDynamicFieldsExportAt: string | null;
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

type RowEditorKind = "email" | "meeting" | "reminder";
type RenderedRowEditorKind = "email" | "meeting" | "reminder" | "reminderBody" | "reminderDuration";
type PlansModalKind = "alert" | "confirm" | "prompt";
type PlansModalSeverity = "information" | "warning" | "validation" | "destructive";
type PlansModalConfig = {
  kind: PlansModalKind;
  title: string;
  message?: string;
  items?: ReactNode[];
  content?: ReactNode;
  secondaryLabel?: string;
  onSecondaryAction?: () => void;
  onConfirmAction?: () => void;
  confirmLabel?: string;
  cancelLabel?: string;
  defaultValue?: string;
  placeholder?: string;
  inputLabel?: string;
  helperText?: ReactNode;
  destructive?: boolean;
  severity?: PlansModalSeverity;
  maxWidthClassName?: string;
};

type NewPlanDraft = {
  eventName: string;
  anchorDate: string;
  eventTime: string;
  noEventDate: boolean;
  weekendRule: WeekendRule;
};

type PlanSetupDialogMode = "new" | "template";

type MissingFieldIssue = {
  message: string;
  issueType?: "required" | "important_anchor" | "anchor_usage";
  severity?: "error" | "warning";
  eventName?: boolean;
  eventDate?: boolean;
  eventTime?: boolean;
  anchorKey?: string;
  fieldTargets?: ValidationFieldTarget[];
  rowId?: string;
  rowIds?: string[];
  isUndefinedAnchor?: boolean;
};
type ValidationFieldName =
  | "title"
  | "body"
  | "reminderTime"
  | "emailTo"
  | "emailCc"
  | "emailBcc"
  | "emailSubject"
  | "emailBody"
  | "meetingAttendees"
  | "meetingLocation";
type ValidationFieldTarget = {
  rowId: string;
  field: ValidationFieldName;
};
type FloatingMenuPosition = {
  top: number;
  left: number;
};
const ROW_EDITOR_OPEN_ANIMATION_MS = 360;
const ROW_EDITOR_CLOSE_ANIMATION_MS = 520;
const INLINE_EDITOR_EXPAND_ANIMATION_MS = 1000;
const INLINE_EDITOR_REVEAL_ANIMATION_MS = 2000;
type InlineEditorFocusPhase = "idle" | "expanding" | "revealing" | "dimmed";

const plansWorkspaceClass =
  "mx-auto w-full max-w-[1120px] min-w-0 space-y-[16px] pb-[44px] pt-[28px] text-gray-900";
const plansSurfaceClass =
  "overflow-hidden rounded-[16px] border border-slate-200/80 bg-white/95 shadow-[0_8px_24px_rgba(30,64,100,0.05)]";
const plansSectionHeaderClass = "border-b border-slate-200/70 bg-white px-[16px] py-[16px] sm:px-[20px] sm:py-[18px]";
const plansCanvasSectionClass = "px-[16px] py-[16px] sm:px-[20px] sm:py-[18px]";
const plansPanelHeadingClass = "text-[20px] font-semibold leading-7 text-slate-950";
const plansPanelHelperClass = "mt-1.5 text-[14px] leading-5 text-slate-600";
const plansFieldLabelClass = "mb-1.5 block text-[13px] font-medium text-slate-600";
const plansInputClass =
  "h-[42px] w-full rounded-[10px] border border-slate-200 bg-white px-3.5 text-[14px] font-medium text-slate-950 shadow-sm placeholder:text-slate-500 focus:border-[#6f9fd1] focus:outline-none focus:ring-2 focus:ring-[#6f9fd1]/20";
const plansSecondaryButtonClass =
  "inline-flex h-[40px] items-center justify-center rounded-[10px] border border-slate-200 bg-white px-4 text-[14px] font-semibold text-slate-700 shadow-sm transition hover:border-slate-300 hover:bg-slate-50 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#6f9fd1]/30";
const plansPrimaryButtonClass =
  "inline-flex h-[40px] items-center justify-center rounded-[10px] border border-[#315f92] bg-[#315f92] px-4 text-[14px] font-semibold text-white shadow-sm transition hover:bg-[#28527d] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#6f9fd1]/35";
const plansEditorInputClass =
  "min-h-[42px] w-full rounded-xl border border-slate-200 bg-white px-3.5 py-2.5 text-[14px] leading-5 text-slate-900 shadow-sm placeholder:text-slate-500 focus:border-[#6f9fd1] focus:outline-none focus:ring-2 focus:ring-[#6f9fd1]/20";
const plansToolbarButtonClass =
  "inline-flex h-[40px] min-w-[112px] items-center justify-center rounded-[10px] border bg-white px-3.5 text-[14px] font-semibold shadow-sm transition focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#6f9fd1]/30";
const plansInspectorSectionClass =
  "rounded-xl border border-slate-200/70 bg-white p-5";
const plansInspectorSectionHeadingClass =
  "mb-3 text-[14px] font-semibold text-slate-700";
const plansInspectorBodyClass =
  "min-h-0 flex-1 space-y-5 overflow-y-auto overscroll-contain bg-slate-50/45 px-4 py-4 pb-6 scroll-pb-6 scroll-pt-4 sm:px-6 sm:py-6";
const plansInspectorFooterClass =
  "flex shrink-0 justify-end border-t border-slate-200/80 bg-white px-4 py-3 sm:px-6 sm:py-4 [padding-bottom:max(12px,env(safe-area-inset-bottom))]";
const plansDialogSecondaryButtonClass =
  "inline-flex h-[40px] items-center justify-center rounded-[10px] border border-slate-200 bg-white px-4 text-[14px] font-semibold text-slate-700 shadow-sm transition hover:border-slate-300 hover:bg-slate-50 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#6f9fd1]/30";
const plansDialogPrimaryButtonClass =
  "inline-flex h-[40px] items-center justify-center rounded-[10px] border border-[#315f92] bg-[#315f92] px-4 text-[14px] font-semibold text-white shadow-sm transition hover:bg-[#28527d] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#6f9fd1]/35";
const plansDialogDangerButtonClass =
  "inline-flex h-[40px] items-center justify-center rounded-[10px] border border-red-600 bg-red-600 px-4 text-[14px] font-semibold text-white shadow-sm transition hover:bg-red-700 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-red-200";
const plansDialogWarningButtonClass =
  "inline-flex h-[40px] items-center justify-center rounded-[10px] border border-amber-500 bg-amber-500 px-4 text-[14px] font-semibold text-white shadow-sm transition hover:bg-amber-600 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-amber-200";
const eventDetailFieldClass = "min-w-0";
const eventDetailLabelClass = "mb-1.5 flex h-[18px] items-center gap-1.5 text-[13px] font-medium text-slate-600";
const eventDetailSupplementClass = "mt-2 flex min-h-[18px] items-center text-[12px] font-medium text-slate-600";
const eventDetailSupplementSpacerClass = "mt-2 hidden h-[18px] min-[900px]:block";

function PlansDialogShell({
  title,
  description,
  children,
  footer,
  animatedIn,
  severity = "information",
  maxWidthClassName = "max-w-[520px]",
  onEscape,
}: {
  title: string;
  description?: ReactNode;
  children?: ReactNode;
  footer: ReactNode;
  animatedIn: boolean;
  severity?: PlansModalSeverity;
  maxWidthClassName?: string;
  onEscape?: () => void;
}) {
  const titleId = useId();
  const descriptionId = useId();
  const severityDotClass =
    severity === "destructive" || severity === "validation"
      ? "bg-red-500"
      : severity === "warning"
        ? "bg-amber-500"
        : "bg-[#4f7fb8]";

  return (
    <div
      className="fixed inset-0 z-[260] flex items-end justify-center bg-slate-950/[0.14] px-4 py-5 sm:items-center sm:py-8"
      onKeyDown={(event) => {
        if (event.key !== "Escape" || !onEscape) return;
        event.preventDefault();
        onEscape();
      }}
    >
      <div
        role="dialog"
        aria-modal="true"
        aria-labelledby={titleId}
        aria-describedby={description ? descriptionId : undefined}
        className={`flex max-h-[calc(100dvh-32px)] w-full ${maxWidthClassName} origin-bottom transform-gpu flex-col overflow-hidden rounded-t-[18px] border border-slate-200 bg-white shadow-[0_28px_80px_rgba(21,40,66,0.24)] transition-all duration-200 ease-[cubic-bezier(0.16,1,0.3,1)] sm:origin-center sm:rounded-[16px] ${
          animatedIn ? "translate-y-0 scale-100 opacity-100" : "translate-y-3 scale-[0.985] opacity-0"
        }`}
      >
        <div className="border-b border-slate-200/80 px-5 py-4 sm:px-6">
          <div className="flex items-start gap-3">
            <span
              aria-hidden="true"
              className={`mt-1 h-2.5 w-2.5 shrink-0 rounded-full ${severityDotClass}`}
            />
            <div className="min-w-0">
              <h3 id={titleId} className="text-[20px] font-semibold leading-6 text-slate-950">
                {title}
              </h3>
	              {description ? (
	                <div id={descriptionId} className="mt-2 whitespace-pre-line text-[14px] leading-5 text-slate-600">
	                  {description}
	                </div>
	              ) : null}
            </div>
          </div>
        </div>
        {children ? <div className="min-h-0 overflow-y-auto px-5 py-4 sm:px-6">{children}</div> : null}
        <div className="flex shrink-0 flex-col gap-2 border-t border-slate-200/80 bg-slate-50/70 px-5 py-4 sm:flex-row sm:justify-end sm:px-6">
          {footer}
        </div>
      </div>
    </div>
  );
}

function AnimatedRowEditor({
  open,
  activeLayer = false,
  animate = true,
  openMs,
  closeMs,
  className = "",
  contentClassName = "pt-2 pb-px",
  innerClassName = "mx-auto w-full max-w-[1160px]",
  children,
}: {
  open: boolean;
  activeLayer?: boolean;
  animate?: boolean;
  openMs?: number;
  closeMs?: number;
  className?: string;
  contentClassName?: string;
  innerClassName?: string;
  children: ReactNode;
}) {
  const contentRef = useRef<HTMLDivElement | null>(null);
  const [shouldRender, setShouldRender] = useState(false);
  const [height, setHeight] = useState("0px");
  const animationDurationMs = open
    ? (openMs ?? ROW_EDITOR_OPEN_ANIMATION_MS)
    : (closeMs ?? ROW_EDITOR_CLOSE_ANIMATION_MS);
  const contentOpacityDurationMs = open ? Math.min(animationDurationMs, 220) : Math.min(animationDurationMs, 420);

  useEffect(() => {
    if (open) {
      setShouldRender(true);
    }
  }, [open]);

  useEffect(() => {
    if (!shouldRender) return;
    const node = contentRef.current;
    if (!node) return;

    let frameOne = 0;
    let frameTwo = 0;

    if (open) {
      const measuredHeight = `${node.scrollHeight}px`;
      setHeight("0px");
      frameOne = window.requestAnimationFrame(() => {
        frameTwo = window.requestAnimationFrame(() => {
          setHeight(measuredHeight);
        });
      });
    } else {
      const measuredHeight = `${node.scrollHeight}px`;
      setHeight(measuredHeight);
      frameOne = window.requestAnimationFrame(() => {
        frameTwo = window.requestAnimationFrame(() => {
          setHeight("0px");
        });
      });
    }

    return () => {
      window.cancelAnimationFrame(frameOne);
      window.cancelAnimationFrame(frameTwo);
    };
  }, [open, shouldRender]);

  if (!animate) {
    if (!open) return null;

    return (
      <div className={`${activeLayer ? "relative z-40" : ""} ${className}`.trim()}>
        <div ref={contentRef} className={contentClassName}>
          <div className={innerClassName}>
            {children}
          </div>
        </div>
      </div>
    );
  }

  if (!shouldRender) return null;

  return (
    <div
      className={`${activeLayer ? "relative z-40" : ""} ${className}`.trim()}
      style={{
        height,
        overflow: open && height === "auto" ? "visible" : "hidden",
        transition: `height ${animationDurationMs}ms cubic-bezier(0.22, 1, 0.36, 1)`,
        willChange: "height",
      }}
      onTransitionEnd={(event) => {
        if (event.target !== event.currentTarget) return;
        if (event.propertyName !== "height") return;
        if (open) {
          setHeight("auto");
          return;
        }
        setShouldRender(false);
      }}
    >
      <div
        ref={contentRef}
        className={contentClassName}
        style={{
          opacity: open ? undefined : 0,
          transform: open ? undefined : "translateY(-8px)",
          transition: `opacity ${contentOpacityDurationMs}ms ease, transform ${animationDurationMs}ms cubic-bezier(0.22, 1, 0.36, 1)`,
        }}
      >
        <div className={innerClassName}>
          {children}
        </div>
      </div>
    </div>
  );
}

function StaggeredInlineEditorItem({
  active,
  delayMs = 0,
  immediate = false,
  animate = true,
  className = "",
  children,
  ...divProps
}: {
  active: boolean;
  delayMs?: number;
  immediate?: boolean;
  animate?: boolean;
  className?: string;
  children: ReactNode;
} & HTMLAttributes<HTMLDivElement>) {
  const [visible, setVisible] = useState(immediate);

  useEffect(() => {
    if (immediate) {
      setVisible(true);
      return;
    }

    if (!active) {
      return;
    }

    const timeoutId = window.setTimeout(() => {
      setVisible(true);
    }, delayMs);

    return () => window.clearTimeout(timeoutId);
  }, [active, delayMs, immediate]);

  if (!animate) {
    return (
      <div className={className} {...divProps}>
        {children}
      </div>
    );
  }

  const isVisible = immediate || visible;

  return (
    <div
      {...divProps}
      className={className}
      style={{
        ...divProps.style,
        opacity: isVisible ? undefined : 0,
        transform: isVisible ? undefined : "translateY(-12px)",
        transition: "opacity 240ms ease, transform 320ms cubic-bezier(0.22, 1, 0.36, 1)",
      }}
    >
      {children}
    </div>
  );
}

type ExecutionSnapshotRowDefinition = {
  id: string;
  title: string;
  body: string;
  offsetDays: number;
  dateBasis: PlanDateBasis;
  rowType: PlanRowType;
  reminderTime: string;
  timeZone: string;
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
const GENERIC_PRESET_ANCHOR_KEYS = ["Event Date", "Event Time", "Event Name"] as const;
const CORE_EVENT_ANCHOR_KEY_SET = new Set(GENERIC_PRESET_ANCHOR_KEYS.map((key) => normalizeAnchorKey(key)));
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
const BUILDER_DRAFT_STORAGE_KEY_PREFIX = "event-based-reminders-app:builder-draft-v1";
const PLANS_SIDEBAR_NEUTRAL_STORAGE_KEY = "event-reminders:plans-sidebar-neutral";
const PLANS_SIDEBAR_NEUTRAL_EVENT = "event-reminders:plans-sidebar-neutral";
const EVENT_TIME_ZONE_MENU_ID = "__event_time_zone__";
const FALLBACK_IANA_TIME_ZONE = "America/New_York";
const FALLBACK_OUTLOOK_TIME_ZONE = "Eastern Standard Time";
const OUTLOOK_TIME_ZONE_OPTIONS = [
  { value: "UTC", label: "(UTC) Coordinated Universal Time", iana: "UTC" },
  { value: "Dateline Standard Time", label: "(UTC-12:00) International Date Line West", iana: "Etc/GMT+12" },
  { value: "UTC-11", label: "(UTC-11:00) Coordinated Universal Time-11", iana: "Etc/GMT+11" },
  { value: "Aleutian Standard Time", label: "(UTC-10:00) Aleutian Islands", iana: "America/Adak" },
  { value: "Hawaiian Standard Time", label: "(UTC-10:00) Hawaii", iana: "Pacific/Honolulu" },
  { value: "Marquesas Standard Time", label: "(UTC-09:30) Marquesas Islands", iana: "Pacific/Marquesas" },
  { value: "Alaskan Standard Time", label: "(UTC-09:00) Alaska", iana: "America/Anchorage" },
  { value: "UTC-09", label: "(UTC-09:00) Coordinated Universal Time-09", iana: "Etc/GMT+9" },
  { value: "Pacific Standard Time (Mexico)", label: "(UTC-08:00) Baja California", iana: "America/Tijuana" },
  { value: "UTC-08", label: "(UTC-08:00) Coordinated Universal Time-08", iana: "Etc/GMT+8" },
  { value: "Pacific Standard Time", label: "(UTC-08:00) Pacific Time (US & Canada)", iana: "America/Los_Angeles" },
  { value: "US Mountain Standard Time", label: "(UTC-07:00) Arizona", iana: "America/Phoenix" },
  { value: "Mountain Standard Time (Mexico)", label: "(UTC-07:00) Chihuahua, La Paz, Mazatlan", iana: "America/Chihuahua" },
  { value: "Mountain Standard Time", label: "(UTC-07:00) Mountain Time (US & Canada)", iana: "America/Denver" },
  { value: "Central America Standard Time", label: "(UTC-06:00) Central America", iana: "America/Guatemala" },
  { value: "Central Standard Time", label: "(UTC-06:00) Central Time (US & Canada)", iana: "America/Chicago" },
  { value: "Easter Island Standard Time", label: "(UTC-06:00) Easter Island", iana: "Pacific/Easter" },
  { value: "Central Standard Time (Mexico)", label: "(UTC-06:00) Guadalajara, Mexico City, Monterrey", iana: "America/Mexico_City" },
  { value: "Canada Central Standard Time", label: "(UTC-06:00) Saskatchewan", iana: "America/Regina" },
  { value: "SA Pacific Standard Time", label: "(UTC-05:00) Bogota, Lima, Quito, Rio Branco", iana: "America/Bogota" },
  { value: "Eastern Standard Time (Mexico)", label: "(UTC-05:00) Chetumal", iana: "America/Cancun" },
  { value: "Eastern Standard Time", label: "(UTC-05:00) Eastern Time (US & Canada)", iana: "America/New_York" },
  { value: "Haiti Standard Time", label: "(UTC-05:00) Haiti", iana: "America/Port-au-Prince" },
  { value: "Cuba Standard Time", label: "(UTC-05:00) Havana", iana: "America/Havana" },
  { value: "US Eastern Standard Time", label: "(UTC-05:00) Indiana (East)", iana: "America/Indiana/Indianapolis" },
  { value: "Turks And Caicos Standard Time", label: "(UTC-05:00) Turks and Caicos", iana: "America/Grand_Turk" },
  { value: "Paraguay Standard Time", label: "(UTC-04:00) Asuncion", iana: "America/Asuncion" },
  { value: "Atlantic Standard Time", label: "(UTC-04:00) Atlantic Time (Canada)", iana: "America/Halifax" },
  { value: "Venezuela Standard Time", label: "(UTC-04:00) Caracas", iana: "America/Caracas" },
  { value: "Central Brazilian Standard Time", label: "(UTC-04:00) Cuiaba", iana: "America/Cuiaba" },
  { value: "SA Western Standard Time", label: "(UTC-04:00) Georgetown, La Paz, Manaus, San Juan", iana: "America/La_Paz" },
  { value: "Pacific SA Standard Time", label: "(UTC-04:00) Santiago", iana: "America/Santiago" },
  { value: "Newfoundland Standard Time", label: "(UTC-03:30) Newfoundland", iana: "America/St_Johns" },
  { value: "Tocantins Standard Time", label: "(UTC-03:00) Araguaina", iana: "America/Araguaina" },
  { value: "E. South America Standard Time", label: "(UTC-03:00) Brasilia", iana: "America/Sao_Paulo" },
  { value: "SA Eastern Standard Time", label: "(UTC-03:00) Cayenne, Fortaleza", iana: "America/Cayenne" },
  { value: "Argentina Standard Time", label: "(UTC-03:00) City of Buenos Aires", iana: "America/Argentina/Buenos_Aires" },
  { value: "Greenland Standard Time", label: "(UTC-03:00) Greenland", iana: "America/Godthab" },
  { value: "Montevideo Standard Time", label: "(UTC-03:00) Montevideo", iana: "America/Montevideo" },
  { value: "Magallanes Standard Time", label: "(UTC-03:00) Punta Arenas", iana: "America/Punta_Arenas" },
  { value: "Saint Pierre Standard Time", label: "(UTC-03:00) Saint Pierre and Miquelon", iana: "America/Miquelon" },
  { value: "Bahia Standard Time", label: "(UTC-03:00) Salvador", iana: "America/Bahia" },
  { value: "UTC-02", label: "(UTC-02:00) Coordinated Universal Time-02", iana: "Etc/GMT+2" },
  { value: "Azores Standard Time", label: "(UTC-01:00) Azores", iana: "Atlantic/Azores" },
  { value: "Cabo Verde Standard Time", label: "(UTC-01:00) Cabo Verde Is.", iana: "Atlantic/Cape_Verde" },
  { value: "GMT Standard Time", label: "(UTC+00:00) Dublin, Edinburgh, Lisbon, London", iana: "Europe/London" },
  { value: "Greenwich Standard Time", label: "(UTC+00:00) Monrovia, Reykjavik", iana: "Atlantic/Reykjavik" },
  { value: "Sao Tome Standard Time", label: "(UTC+00:00) Sao Tome", iana: "Africa/Sao_Tome" },
  { value: "Morocco Standard Time", label: "(UTC+01:00) Casablanca", iana: "Africa/Casablanca" },
  { value: "W. Europe Standard Time", label: "(UTC+01:00) Amsterdam, Berlin, Bern, Rome, Stockholm, Vienna", iana: "Europe/Berlin" },
  { value: "Central Europe Standard Time", label: "(UTC+01:00) Belgrade, Bratislava, Budapest, Ljubljana, Prague", iana: "Europe/Budapest" },
  { value: "Romance Standard Time", label: "(UTC+01:00) Brussels, Copenhagen, Madrid, Paris", iana: "Europe/Paris" },
  { value: "Central European Standard Time", label: "(UTC+01:00) Sarajevo, Skopje, Warsaw, Zagreb", iana: "Europe/Warsaw" },
  { value: "W. Central Africa Standard Time", label: "(UTC+01:00) West Central Africa", iana: "Africa/Lagos" },
  { value: "Jordan Standard Time", label: "(UTC+02:00) Amman", iana: "Asia/Amman" },
  { value: "GTB Standard Time", label: "(UTC+02:00) Athens, Bucharest", iana: "Europe/Athens" },
  { value: "Middle East Standard Time", label: "(UTC+02:00) Beirut", iana: "Asia/Beirut" },
  { value: "Egypt Standard Time", label: "(UTC+02:00) Cairo", iana: "Africa/Cairo" },
  { value: "E. Europe Standard Time", label: "(UTC+02:00) Chisinau", iana: "Europe/Chisinau" },
  { value: "Syria Standard Time", label: "(UTC+02:00) Damascus", iana: "Asia/Damascus" },
  { value: "West Bank Standard Time", label: "(UTC+02:00) Gaza, Hebron", iana: "Asia/Hebron" },
  { value: "South Africa Standard Time", label: "(UTC+02:00) Harare, Pretoria", iana: "Africa/Johannesburg" },
  { value: "FLE Standard Time", label: "(UTC+02:00) Helsinki, Kyiv, Riga, Sofia, Tallinn, Vilnius", iana: "Europe/Helsinki" },
  { value: "Israel Standard Time", label: "(UTC+02:00) Jerusalem", iana: "Asia/Jerusalem" },
  { value: "Kaliningrad Standard Time", label: "(UTC+02:00) Kaliningrad", iana: "Europe/Kaliningrad" },
  { value: "Sudan Standard Time", label: "(UTC+02:00) Khartoum", iana: "Africa/Khartoum" },
  { value: "Libya Standard Time", label: "(UTC+02:00) Tripoli", iana: "Africa/Tripoli" },
  { value: "Arabic Standard Time", label: "(UTC+03:00) Baghdad", iana: "Asia/Baghdad" },
  { value: "Turkey Standard Time", label: "(UTC+03:00) Istanbul", iana: "Europe/Istanbul" },
  { value: "Arab Standard Time", label: "(UTC+03:00) Kuwait, Riyadh", iana: "Asia/Riyadh" },
  { value: "Belarus Standard Time", label: "(UTC+03:00) Minsk", iana: "Europe/Minsk" },
  { value: "Russian Standard Time", label: "(UTC+03:00) Moscow, St. Petersburg", iana: "Europe/Moscow" },
  { value: "E. Africa Standard Time", label: "(UTC+03:00) Nairobi", iana: "Africa/Nairobi" },
  { value: "Iran Standard Time", label: "(UTC+03:30) Tehran", iana: "Asia/Tehran" },
  { value: "Arabian Standard Time", label: "(UTC+04:00) Abu Dhabi, Muscat", iana: "Asia/Dubai" },
  { value: "Astrakhan Standard Time", label: "(UTC+04:00) Astrakhan, Ulyanovsk", iana: "Europe/Astrakhan" },
  { value: "Azerbaijan Standard Time", label: "(UTC+04:00) Baku", iana: "Asia/Baku" },
  { value: "Russia Time Zone 3", label: "(UTC+04:00) Izhevsk, Samara", iana: "Europe/Samara" },
  { value: "Mauritius Standard Time", label: "(UTC+04:00) Port Louis", iana: "Indian/Mauritius" },
  { value: "Georgian Standard Time", label: "(UTC+04:00) Tbilisi", iana: "Asia/Tbilisi" },
  { value: "Caucasus Standard Time", label: "(UTC+04:00) Yerevan", iana: "Asia/Yerevan" },
  { value: "Afghanistan Standard Time", label: "(UTC+04:30) Kabul", iana: "Asia/Kabul" },
  { value: "West Asia Standard Time", label: "(UTC+05:00) Ashgabat, Tashkent", iana: "Asia/Tashkent" },
  { value: "Ekaterinburg Standard Time", label: "(UTC+05:00) Ekaterinburg", iana: "Asia/Yekaterinburg" },
  { value: "Pakistan Standard Time", label: "(UTC+05:00) Islamabad, Karachi", iana: "Asia/Karachi" },
  { value: "India Standard Time", label: "(UTC+05:30) Chennai, Kolkata, Mumbai, New Delhi", iana: "Asia/Kolkata" },
  { value: "Sri Lanka Standard Time", label: "(UTC+05:30) Sri Jayawardenepura", iana: "Asia/Colombo" },
  { value: "Nepal Standard Time", label: "(UTC+05:45) Kathmandu", iana: "Asia/Kathmandu" },
  { value: "Central Asia Standard Time", label: "(UTC+06:00) Astana", iana: "Asia/Almaty" },
  { value: "Bangladesh Standard Time", label: "(UTC+06:00) Dhaka", iana: "Asia/Dhaka" },
  { value: "Omsk Standard Time", label: "(UTC+06:00) Omsk", iana: "Asia/Omsk" },
  { value: "Myanmar Standard Time", label: "(UTC+06:30) Yangon (Rangoon)", iana: "Asia/Yangon" },
  { value: "SE Asia Standard Time", label: "(UTC+07:00) Bangkok, Hanoi, Jakarta", iana: "Asia/Bangkok" },
  { value: "Altai Standard Time", label: "(UTC+07:00) Barnaul, Gorno-Altaysk", iana: "Asia/Barnaul" },
  { value: "W. Mongolia Standard Time", label: "(UTC+07:00) Hovd", iana: "Asia/Hovd" },
  { value: "North Asia Standard Time", label: "(UTC+07:00) Krasnoyarsk", iana: "Asia/Krasnoyarsk" },
  { value: "N. Central Asia Standard Time", label: "(UTC+07:00) Novosibirsk", iana: "Asia/Novosibirsk" },
  { value: "Tomsk Standard Time", label: "(UTC+07:00) Tomsk", iana: "Asia/Tomsk" },
  { value: "China Standard Time", label: "(UTC+08:00) Beijing, Chongqing, Hong Kong, Urumqi", iana: "Asia/Shanghai" },
  { value: "North Asia East Standard Time", label: "(UTC+08:00) Irkutsk", iana: "Asia/Irkutsk" },
  { value: "Singapore Standard Time", label: "(UTC+08:00) Kuala Lumpur, Singapore", iana: "Asia/Singapore" },
  { value: "W. Australia Standard Time", label: "(UTC+08:00) Perth", iana: "Australia/Perth" },
  { value: "Taipei Standard Time", label: "(UTC+08:00) Taipei", iana: "Asia/Taipei" },
  { value: "Ulaanbaatar Standard Time", label: "(UTC+08:00) Ulaanbaatar", iana: "Asia/Ulaanbaatar" },
  { value: "Aus Central W. Standard Time", label: "(UTC+08:45) Eucla", iana: "Australia/Eucla" },
  { value: "Transbaikal Standard Time", label: "(UTC+09:00) Chita", iana: "Asia/Chita" },
  { value: "Tokyo Standard Time", label: "(UTC+09:00) Osaka, Sapporo, Tokyo", iana: "Asia/Tokyo" },
  { value: "North Korea Standard Time", label: "(UTC+09:00) Pyongyang", iana: "Asia/Pyongyang" },
  { value: "Korea Standard Time", label: "(UTC+09:00) Seoul", iana: "Asia/Seoul" },
  { value: "Yakutsk Standard Time", label: "(UTC+09:00) Yakutsk", iana: "Asia/Yakutsk" },
  { value: "Cen. Australia Standard Time", label: "(UTC+09:30) Adelaide", iana: "Australia/Adelaide" },
  { value: "AUS Central Standard Time", label: "(UTC+09:30) Darwin", iana: "Australia/Darwin" },
  { value: "E. Australia Standard Time", label: "(UTC+10:00) Brisbane", iana: "Australia/Brisbane" },
  { value: "AUS Eastern Standard Time", label: "(UTC+10:00) Canberra, Melbourne, Sydney", iana: "Australia/Sydney" },
  { value: "West Pacific Standard Time", label: "(UTC+10:00) Guam, Port Moresby", iana: "Pacific/Port_Moresby" },
  { value: "Tasmania Standard Time", label: "(UTC+10:00) Hobart", iana: "Australia/Hobart" },
  { value: "Vladivostok Standard Time", label: "(UTC+10:00) Vladivostok", iana: "Asia/Vladivostok" },
  { value: "Lord Howe Standard Time", label: "(UTC+10:30) Lord Howe Island", iana: "Australia/Lord_Howe" },
  { value: "Bougainville Standard Time", label: "(UTC+11:00) Bougainville Island", iana: "Pacific/Bougainville" },
  { value: "Russia Time Zone 10", label: "(UTC+11:00) Chokurdakh", iana: "Asia/Srednekolymsk" },
  { value: "Magadan Standard Time", label: "(UTC+11:00) Magadan", iana: "Asia/Magadan" },
  { value: "Norfolk Standard Time", label: "(UTC+11:00) Norfolk Island", iana: "Pacific/Norfolk" },
  { value: "Sakhalin Standard Time", label: "(UTC+11:00) Sakhalin", iana: "Asia/Sakhalin" },
  { value: "Central Pacific Standard Time", label: "(UTC+11:00) Solomon Is., New Caledonia", iana: "Pacific/Guadalcanal" },
  { value: "Russia Time Zone 11", label: "(UTC+12:00) Anadyr, Petropavlovsk-Kamchatsky", iana: "Asia/Kamchatka" },
  { value: "New Zealand Standard Time", label: "(UTC+12:00) Auckland, Wellington", iana: "Pacific/Auckland" },
  { value: "UTC+12", label: "(UTC+12:00) Coordinated Universal Time+12", iana: "Etc/GMT-12" },
  { value: "Fiji Standard Time", label: "(UTC+12:00) Fiji", iana: "Pacific/Fiji" },
  { value: "Chatham Islands Standard Time", label: "(UTC+12:45) Chatham Islands", iana: "Pacific/Chatham" },
  { value: "UTC+13", label: "(UTC+13:00) Coordinated Universal Time+13", iana: "Etc/GMT-13" },
  { value: "Tonga Standard Time", label: "(UTC+13:00) Nuku'alofa", iana: "Pacific/Tongatapu" },
  { value: "Samoa Standard Time", label: "(UTC+13:00) Samoa", iana: "Pacific/Apia" },
  { value: "Line Islands Standard Time", label: "(UTC+14:00) Kiritimati Island", iana: "Pacific/Kiritimati" },
] as const;
const OUTLOOK_TIME_ZONE_VALUES = new Set<string>(OUTLOOK_TIME_ZONE_OPTIONS.map((option) => option.value));
const OUTLOOK_TIME_ZONE_BY_IANA = new Map<string, string>(OUTLOOK_TIME_ZONE_OPTIONS.map((option) => [option.iana, option.value]));
const IANA_TIME_ZONE_BY_OUTLOOK = new Map<string, string>(OUTLOOK_TIME_ZONE_OPTIONS.map((option) => [option.value, option.iana]));

function makeId(prefix: string) {
  return `${prefix}_${Math.random().toString(36).slice(2, 10)}`;
}

function getAiReadinessLabel(status: "needs_more_info" | "ready_to_apply") {
  return status === "ready_to_apply" ? "Draft ready to apply" : "Needs a bit more information";
}

function getBrowserTimeZone() {
  if (typeof Intl === "undefined") return FALLBACK_IANA_TIME_ZONE;
  return Intl.DateTimeFormat().resolvedOptions().timeZone || FALLBACK_IANA_TIME_ZONE;
}

function getDefaultOutlookTimeZone() {
  return OUTLOOK_TIME_ZONE_BY_IANA.get(getBrowserTimeZone()) ?? FALLBACK_OUTLOOK_TIME_ZONE;
}

function normalizeOutlookTimeZone(timeZone?: string | null) {
  const trimmed = timeZone?.trim();
  if (!trimmed) return getDefaultOutlookTimeZone();
  if (OUTLOOK_TIME_ZONE_VALUES.has(trimmed)) return trimmed;
  return OUTLOOK_TIME_ZONE_BY_IANA.get(trimmed) ?? getDefaultOutlookTimeZone();
}

function getIanaTimeZoneForProvider(timeZone?: string | null) {
  const outlookTimeZone = normalizeOutlookTimeZone(timeZone);
  return IANA_TIME_ZONE_BY_OUTLOOK.get(outlookTimeZone) ?? getBrowserTimeZone();
}

function getSupportedTimeZones() {
  return OUTLOOK_TIME_ZONE_OPTIONS;
}

function getOutlookTimeZoneLabel(timeZone: string) {
  return OUTLOOK_TIME_ZONE_OPTIONS.find((option) => option.value === normalizeOutlookTimeZone(timeZone))?.label ?? timeZone;
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
  eventTimeZone?: string;
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
    eventTimeZone: input.eventTimeZone ?? "",
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
        timeZone: row.timeZone ?? "",
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
  snapshot: Pick<BuilderStateSnapshot, "planType" | "templateName" | "eventName" | "anchorDate" | "noEventDate" | "weekendRule" | "eventTimeZone" | "anchors" | "rows">,
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
      eventTimeZone: snapshot.eventTimeZone,
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
    isImportant: false,
    lastUpdatedAt: null,
  };
}

function createLockedAnchor(key: string, value = ""): BuilderAnchor {
  return {
    id: crypto.randomUUID(),
    key,
    value,
    locked: true,
    isImportant: false,
    lastUpdatedAt: value.trim() ? new Date().toISOString() : null,
  };
}

function createGenericPresetAnchors() {
  return GENERIC_PRESET_ANCHOR_KEYS.map((key) => createLockedAnchor(key));
}

function isCoreEventAnchorKey(key: string) {
  return CORE_EVENT_ANCHOR_KEY_SET.has(normalizeAnchorKey(key));
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

function isObject(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function normalizeDraftBuilderMode(value: unknown): BuilderMode {
  if (value === "template" || value === "guided" || value === "custom") return value;
  return "new";
}

function normalizeDraftPlanType(value: unknown): PlanType {
  if (value === "conference" || value === "press_release") return value;
  return "earnings";
}

function normalizeDraftWeekendRule(value: unknown): WeekendRule {
  return value === "prior_business_day" ? "prior_business_day" : "none";
}

function normalizeDraftGuidedForm(value: unknown): GuidedFormState {
  const fallback = createEmptyGuidedForm();
  if (!isObject(value)) return fallback;

  return {
    releaseName: typeof value.releaseName === "string" ? value.releaseName : fallback.releaseName,
    releaseDate: typeof value.releaseDate === "string" ? value.releaseDate : fallback.releaseDate,
    releaseTime: typeof value.releaseTime === "string" ? value.releaseTime : fallback.releaseTime,
    quarter: value.quarter === "Q1" || value.quarter === "Q2" || value.quarter === "Q3" || value.quarter === "Q4" ? value.quarter : "",
    year: typeof value.year === "string" ? value.year : fallback.year,
    fiscalYear: Boolean(value.fiscalYear),
    earningsDate: typeof value.earningsDate === "string" ? value.earningsDate : fallback.earningsDate,
    earningsTime: typeof value.earningsTime === "string" ? value.earningsTime : fallback.earningsTime,
    conferenceName: typeof value.conferenceName === "string" ? value.conferenceName : fallback.conferenceName,
    conferenceLocation: typeof value.conferenceLocation === "string" ? value.conferenceLocation : fallback.conferenceLocation,
    conferenceDate: typeof value.conferenceDate === "string" ? value.conferenceDate : fallback.conferenceDate,
    conferenceEndDate: typeof value.conferenceEndDate === "string" ? value.conferenceEndDate : fallback.conferenceEndDate,
  };
}

function normalizeDraftRow(value: unknown): BuilderRow | null {
  if (!isObject(value) || typeof value.title !== "string") return null;

  return {
    id: typeof value.id === "string" ? value.id : crypto.randomUUID(),
    title: value.title,
    body: typeof value.body === "string" ? value.body : "",
    offsetDays: typeof value.offsetDays === "number" ? value.offsetDays : 0,
    dateBasis: value.dateBasis === "today" ? "today" : "event",
    rowType: value.rowType === "email" || value.rowType === "calendar_event" ? value.rowType : "reminder",
    reminderTime: typeof value.reminderTime === "string" ? value.reminderTime : "",
    timeZone: typeof value.timeZone === "string" ? value.timeZone : "",
    emailDraft: normalizeEmailDraft(isObject(value.emailDraft) ? value.emailDraft : undefined),
    durationDraft: isObject(value.durationDraft)
      ? {
          durationMinutes: typeof value.durationDraft.durationMinutes === "number" ? value.durationDraft.durationMinutes : undefined,
          useCustomEnd: typeof value.durationDraft.useCustomEnd === "boolean" ? value.durationDraft.useCustomEnd : undefined,
          endDate: typeof value.durationDraft.endDate === "string" ? value.durationDraft.endDate : undefined,
          endTime: typeof value.durationDraft.endTime === "string" ? value.durationDraft.endTime : undefined,
          isAllDay: typeof value.durationDraft.isAllDay === "boolean" ? value.durationDraft.isAllDay : undefined,
        }
      : undefined,
    meetingDraft: normalizeMeetingDraft(isObject(value.meetingDraft) ? value.meetingDraft : undefined) ?? undefined,
  };
}

function normalizeDraftAnchor(value: unknown): BuilderAnchor | null {
  if (!isObject(value) || typeof value.key !== "string") return null;

  return {
    id: typeof value.id === "string" ? value.id : crypto.randomUUID(),
    key: value.key,
    value: typeof value.value === "string" ? value.value : "",
    locked: typeof value.locked === "boolean" ? value.locked : undefined,
    isImportant: typeof value.isImportant === "boolean" ? value.isImportant : false,
    lastUpdatedAt:
      typeof value.lastUpdatedAt === "string"
        ? value.lastUpdatedAt
        : value.lastUpdatedAt === null
          ? null
          : null,
  };
}

function normalizeBuilderDraftSnapshot(value: unknown): BuilderStateSnapshot | null {
  if (!isObject(value)) return null;

  const rows = Array.isArray(value.rows)
    ? value.rows.map(normalizeDraftRow).filter((row): row is BuilderRow => Boolean(row))
    : [];
  const anchors = Array.isArray(value.anchors)
    ? value.anchors.map(normalizeDraftAnchor).filter((anchor): anchor is BuilderAnchor => Boolean(anchor))
    : [];

  return {
    builderMode: normalizeDraftBuilderMode(value.builderMode),
    selectedTemplateId: typeof value.selectedTemplateId === "string" ? value.selectedTemplateId : null,
    planType: normalizeDraftPlanType(value.planType),
    templateName: typeof value.templateName === "string" ? value.templateName : "",
    eventName: typeof value.eventName === "string" ? value.eventName : "",
    anchorDate: typeof value.anchorDate === "string" ? value.anchorDate : "",
    hasExplicitEventDate: typeof value.hasExplicitEventDate === "boolean" ? value.hasExplicitEventDate : false,
    eventTime: typeof value.eventTime === "string" ? value.eventTime : "",
    eventTimeZone: normalizeOutlookTimeZone(typeof value.eventTimeZone === "string" ? value.eventTimeZone : ""),
    noEventDate: typeof value.noEventDate === "boolean" ? value.noEventDate : false,
    weekendRule: normalizeDraftWeekendRule(value.weekendRule),
    rows,
    anchors,
    guidedForm: normalizeDraftGuidedForm(value.guidedForm),
    lastDynamicFieldsExportAt:
      typeof value.lastDynamicFieldsExportAt === "string"
        ? value.lastDynamicFieldsExportAt
        : value.lastDynamicFieldsExportAt === null
          ? null
          : null,
  };
}

function getBuilderDraftStorageKey() {
  return getScopedStorageKey(BUILDER_DRAFT_STORAGE_KEY_PREFIX);
}

function consumePlansSidebarNeutralEntry() {
  if (typeof window === "undefined") return false;

  const shouldOpenNeutral = window.sessionStorage.getItem(PLANS_SIDEBAR_NEUTRAL_STORAGE_KEY) === "1";
  if (shouldOpenNeutral) {
    window.sessionStorage.removeItem(PLANS_SIDEBAR_NEUTRAL_STORAGE_KEY);
  }
  return shouldOpenNeutral;
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
      <span>Exported</span>
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
  if (template.id === "seed:press_release") {
    return {
      id: "seed:press_release",
      name: "Press Release",
      baseType: "press_release" as const,
    };
  }
  if (template.id === "seed:conference") {
    return {
      id: "seed:conference",
      name: "Conference",
      baseType: "conference" as const,
    };
  }
  if (template.id === "seed:earnings") {
    return {
      id: "seed:earnings",
      name: "Earnings",
      baseType: "earnings" as const,
    };
  }
  return null;
}

function getPresetAnchorKeysForType(type: PlanType) {
  if (type === "press_release") return ["Event Date", "Event Time", "Event Name", ...PRESS_RELEASE_PRESET_ANCHOR_KEYS];
  if (type === "conference") return ["Event Date", "Event Time", "Event Name", ...CONFERENCE_PRESET_ANCHOR_KEYS];
  return ["Event Date", "Event Time", "Event Name", ...EARNINGS_PRESET_ANCHOR_KEYS];
}

function normalizeEmailDraft(draft?: BuilderEmailDraft | null) {
  return {
    to: normalizeRecipientEntries(draft?.to),
    cc: normalizeRecipientEntries(draft?.cc),
    bcc: normalizeRecipientEntries(draft?.bcc),
    subject: typeof draft?.subject === "string" ? draft.subject : "",
    body: typeof draft?.body === "string" ? draft.body : "",
  };
}

function normalizeMeetingDraft(draft?: BuilderRow["meetingDraft"] | null) {
  if (!draft) return undefined;
  return {
    attendees: normalizeRecipientEntries(draft.attendees),
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

  const compactTwelveHourMatch = normalized.match(/^(\d{1,4})\s*([AP]M)$/);
  if (compactTwelveHourMatch) {
    const rawTime = compactTwelveHourMatch[1];
    const meridiem = compactTwelveHourMatch[2];
    const hours =
      rawTime.length <= 2 ? Number(rawTime) : Number(rawTime.slice(0, rawTime.length - 2));
    const minutes = rawTime.length <= 2 ? 0 : Number(rawTime.slice(-2));
    if (hours < 1 || hours > 12 || minutes < 0 || minutes > 59) return null;
    const normalizedHours =
      meridiem === "AM" ? (hours === 12 ? 0 : hours) : hours === 12 ? 12 : hours + 12;
    return `${String(normalizedHours).padStart(2, "0")}:${String(minutes).padStart(2, "0")}`;
  }

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

  const compactTwentyFourHourMatch = normalized.match(/^\d{1,4}$/);
  if (compactTwentyFourHourMatch) {
    const digits = compactTwentyFourHourMatch[0];
    const hours = digits.length <= 2 ? Number(digits) : Number(digits.slice(0, digits.length - 2));
    const minutes = digits.length <= 2 ? 0 : Number(digits.slice(-2));
    if (hours < 0 || hours > 23 || minutes < 0 || minutes > 59) return null;
    return `${String(hours).padStart(2, "0")}:${String(minutes).padStart(2, "0")}`;
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
  let year: number | undefined;
  let month: number | undefined;
  let day: number | undefined;

  if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) {
    [year, month, day] = trimmed.split("-").map(Number);
  } else if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(trimmed)) {
    [month, day, year] = trimmed.split("/").map(Number);
  } else {
    return value;
  }

  const parsed = new Date(year ?? 2000, (month ?? 1) - 1, day ?? 1);
  return parsed.toLocaleDateString("en-US", {
    month: "long",
    day: "numeric",
    year: "numeric",
  });
}

function formatEventDateForEditor(value: string) {
  const trimmed = value.trim();
  if (!trimmed) return "";

  let year: number | undefined;
  let month: number | undefined;
  let day: number | undefined;

  if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) {
    [year, month, day] = trimmed.split("-").map(Number);
  } else if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(trimmed)) {
    [month, day, year] = trimmed.split("/").map(Number);
  } else {
    return value;
  }

  return `${String(month ?? "").padStart(2, "0")}/${String(day ?? "").padStart(2, "0")}/${year ?? ""}`;
}

function parseEventDateInput(value: string) {
  const trimmed = value.trim();
  if (!trimmed) return "";
  if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) return trimmed;

  const compactMatch = trimmed.match(/^(\d{2})(\d{2})(\d{4})$/);
  if (compactMatch) {
    const month = Number(compactMatch[1]);
    const day = Number(compactMatch[2]);
    const year = Number(compactMatch[3]);
    if (month < 1 || month > 12 || day < 1 || day > 31) return null;

    const candidate = new Date(year, month - 1, day);
    if (
      Number.isNaN(candidate.getTime()) ||
      candidate.getFullYear() !== year ||
      candidate.getMonth() !== month - 1 ||
      candidate.getDate() !== day
    ) {
      return null;
    }

    return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
  }

  const slashMatch = trimmed.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!slashMatch) return null;

  const month = Number(slashMatch[1]);
  const day = Number(slashMatch[2]);
  const year = Number(slashMatch[3]);
  if (month < 1 || month > 12 || day < 1 || day > 31) return null;

  const candidate = new Date(year, month - 1, day);
  if (
    Number.isNaN(candidate.getTime()) ||
    candidate.getFullYear() !== year ||
    candidate.getMonth() !== month - 1 ||
    candidate.getDate() !== day
  ) {
    return null;
  }

  return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
}

function formatTimeDisplayValue(value: string) {
  const trimmed = value.trim();
  if (!/^\d{2}:\d{2}$/.test(trimmed)) return value;
  const [hours, minutes] = trimmed.split(":").map(Number);
  const parsed = new Date(2000, 0, 1, hours ?? 0, minutes ?? 0);
  return parsed.toLocaleTimeString("en-US", {
    hour: "numeric",
    minute: "2-digit",
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
  const isTimeAnchor =
    normalized === normalizeAnchorKey("Event Time") ||
    normalized === normalizeAnchorKey("Dissemination Time") ||
    normalized === normalizeAnchorKey("Earnings Call Time");
  if (isDateAnchor) return formatAnchorDateDisplayValue(value);
  if (isTimeAnchor && value.trim()) return formatTimeDisplayValue(value);
  return value;
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
    eventTimeValue?: string;
  }
) {
  const normalized = normalizeAnchorKey(key);
  if (normalized === normalizeAnchorKey("Event Name")) return options?.eventNameValue ?? planName;
  if (normalized === normalizeAnchorKey("Event Date")) return options?.eventDateValue ?? anchorDate;
  if (normalized === normalizeAnchorKey("Event Time")) return options?.eventTimeValue ?? "";
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
    timeZone: "",
    emailDraft: { to: [], cc: [], bcc: [], subject: "", body: "" },
  };
}

function buildCoreCustomTemplates(defaultReminderTime = "", defaultPressReleaseTime = ""): SavedPlanTemplate[] {
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

  return [
    {
      id: "core:press_release",
      name: "Press Release",
      baseType: "press_release",
      templateMode: "custom",
      weekendRule: "none",
      anchors: PRESS_RELEASE_PRESET_ANCHOR_KEYS.map((key) => ({ key, value: "" })),
      items: [
        { id: crypto.randomUUID(), title: "Prepare Press Release Distribution for Tomorrow", body: sharedChecklistBody, offsetDays: -1, rowType: "reminder", reminderTime: normalizedDefaultReminderTime, dateBasis: "event" },
        { id: crypto.randomUUID(), title: "Finalize Proof & Schedule Press Release", offsetDays: -1, rowType: "reminder", reminderTime: "16:00", dateBasis: "event" },
        { id: crypto.randomUUID(), title: "Press Release Going Out Now: Constant Contact & Social", offsetDays: 0, rowType: "reminder", reminderTime: normalizedDefaultPressReleaseTime, dateBasis: "event" },
      ],
    },
    {
      id: "core:conference",
      name: "Conference",
      baseType: "conference",
      templateMode: "custom",
      weekendRule: "none",
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
      id: "core:earnings",
      name: "Earnings",
      baseType: "earnings",
      templateMode: "custom",
      weekendRule: "none",
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

function sanitizeTemplateAnchorDefinition(anchor: PersistedTemplateAnchor): PersistedTemplateAnchor {
  return {
    key: anchor.key.trim(),
    value: "",
    isImportant: Boolean(anchor.isImportant),
    lastUpdatedAt: null,
  };
}

function buildFreshBuilderAnchorFromTemplate(anchor: PersistedTemplateAnchor): BuilderAnchor {
  return {
    id: crypto.randomUUID(),
    key: anchor.key,
    value: "",
    isImportant: Boolean(anchor.isImportant),
    lastUpdatedAt: null,
  };
}

function normalizeImportedTemplate(template: {
  id: string;
  name: string;
  baseType: PlanType;
  templateMode?: "template" | "custom";
  noEventDate?: boolean;
  weekendRule: WeekendRule;
  anchors: PersistedTemplateAnchor[];
  items: BuilderRow[];
  lastDynamicFieldsExportAt?: string | null;
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
      .map(sanitizeTemplateAnchorDefinition)
      .filter((anchor) => anchor.key),
    items: cloneTemplateRows(template.items),
    lastDynamicFieldsExportAt: template.lastDynamicFieldsExportAt ?? null,
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
    anchors: template.anchors.map(sanitizeTemplateAnchorDefinition),
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
    lastDynamicFieldsExportAt: template.lastDynamicFieldsExportAt ?? null,
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
      timeZone: row.timeZone ?? "",
      emailDraft: row.emailDraft,
      durationDraft: row.durationDraft,
      meetingDraft: row.meetingDraft,
    })),
    lastDynamicFieldsExportAt: template.lastDynamicFieldsExportAt ?? null,
  });
}

function migratePersistedTemplatesToCustom(state: PersistedTemplateState): PersistedTemplateState {
  const templateIdMap = new Map<string, string>();
  const templates = state.templates.map((template) => {
    const shouldConvertId = template.id.startsWith("seed:");
    const nextId = shouldConvertId ? makeId("template") : template.id;
    if (nextId !== template.id) {
      templateIdMap.set(template.id, nextId);
    }
    return {
      ...template,
      id: nextId,
      templateMode: "custom" as const,
      isProtected: false,
    };
  });

  const selectedTemplateId = state.selectedTemplateId ? (templateIdMap.get(state.selectedTemplateId) ?? state.selectedTemplateId) : null;
  return {
    selectedTemplateId,
    templates,
  };
}

function ensureCoreCustomTemplates(
  state: PersistedTemplateState,
  defaultReminderTime: string,
  defaultPressReleaseTime: string
): PersistedTemplateState {
  const existingTemplateNames = new Set(state.templates.map((template) => template.name.trim().toLowerCase()));
  const missingTemplates = buildCoreCustomTemplates(defaultReminderTime, defaultPressReleaseTime)
    .filter((template) => !existingTemplateNames.has(template.name.trim().toLowerCase()))
    .map((template, index) => toPersistedTemplate(template, state.templates.length + index));

  if (missingTemplates.length === 0) return state;

  return {
    selectedTemplateId: state.selectedTemplateId ?? missingTemplates[0]?.id ?? null,
    templates: [...state.templates, ...missingTemplates],
  };
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

function buildTemplateItemsFromRows(rows: BuilderRow[], recipientGroups: RecipientGroup[]): TemplateItem[] {
  return rows.map((row) => ({
    id: row.id,
    title: row.title.trim() || "Untitled row",
    body: row.body?.trim() || undefined,
    offsetDays: row.offsetDays ?? 0,
    dateBasis: row.dateBasis ?? "event",
    rowType: row.rowType,
    reminderTime: row.reminderTime?.trim() || undefined,
    timeZone: row.timeZone?.trim() || undefined,
    emailDraft:
      row.rowType === "email"
        ? {
            ...normalizeEmailDraft(row.emailDraft),
            to: resolveRecipientEntries(normalizeEmailDraft(row.emailDraft).to, recipientGroups),
            cc: resolveRecipientEntries(normalizeEmailDraft(row.emailDraft).cc, recipientGroups),
            bcc: resolveRecipientEntries(normalizeEmailDraft(row.emailDraft).bcc, recipientGroups),
          }
        : undefined,
    durationDraft: row.durationDraft,
    meetingDraft: row.meetingDraft
      ? {
          ...normalizeMeetingDraft(row.meetingDraft),
          attendees: resolveRecipientEntries(normalizeMeetingDraft(row.meetingDraft)?.attendees ?? [], recipientGroups),
        }
      : undefined,
  }));
}

type TemplateLibraryProps = {
  open: boolean;
  mounted: boolean;
  templates: SavedPlanTemplate[];
  selectedTemplateId: string | null;
  highlightedTemplateId: string | null;
  onSelectTemplate: (templateId: string) => void | Promise<void>;
  onDuplicateTemplate: (templateId: string) => void | Promise<void>;
  onRenameTemplate: (templateId: string) => void | Promise<void>;
  onDeleteTemplate: (templateId: string) => void | Promise<void>;
  actionMessage?: string;
  onClose: () => void;
  returnFocus: () => void;
};

function TemplateLibrary({
  open,
  mounted,
  templates,
  selectedTemplateId,
  highlightedTemplateId,
  onSelectTemplate,
  onDuplicateTemplate,
  onRenameTemplate,
  onDeleteTemplate,
  actionMessage,
  onClose,
  returnFocus,
}: TemplateLibraryProps) {
  const titleId = useId();
  const helperId = useId();
  const [isRendered, setIsRendered] = useState(open);
  const [isClosing, setIsClosing] = useState(false);
  const [openActionTemplateId, setOpenActionTemplateId] = useState<string | null>(null);
  const [actionMenuPosition, setActionMenuPosition] = useState<{ top: number; left: number } | null>(null);
  const rootRef = useRef<HTMLDivElement | null>(null);
  const dialogRef = useRef<HTMLDivElement | null>(null);
  const closeButtonRef = useRef<HTMLButtonElement | null>(null);
  const closeTimerRef = useRef<number | null>(null);
  const actionButtonRefs = useRef<Record<string, HTMLButtonElement | null>>({});
  const actionMenuRef = useRef<HTMLDivElement | null>(null);

  const closeActionMenu = useCallback((options?: { restoreFocus?: boolean }) => {
    const currentTemplateId = openActionTemplateId;
    setOpenActionTemplateId(null);
    setActionMenuPosition(null);
    if (options?.restoreFocus && currentTemplateId) {
      window.requestAnimationFrame(() => {
        actionButtonRefs.current[currentTemplateId]?.focus();
      });
    }
  }, [openActionTemplateId]);

  const requestClose = useCallback((options?: { restoreFocus?: boolean }) => {
    closeActionMenu();
    onClose();
    if (options?.restoreFocus) {
      window.requestAnimationFrame(returnFocus);
    }
  }, [closeActionMenu, onClose, returnFocus]);

  useEffect(() => {
    if (open) {
      if (closeTimerRef.current) {
        window.clearTimeout(closeTimerRef.current);
        closeTimerRef.current = null;
      }
      setIsRendered(true);
      setIsClosing(false);
      setOpenActionTemplateId(null);
      setActionMenuPosition(null);
      return;
    }

    if (!isRendered) return;
    setIsClosing(true);
    closeTimerRef.current = window.setTimeout(() => {
      setIsRendered(false);
      setIsClosing(false);
      closeTimerRef.current = null;
    }, 120);
  }, [isRendered, open]);

  useEffect(() => {
    if (!isRendered) return;
    const previousOverflow = document.body.style.overflow;
    document.body.style.overflow = "hidden";
    return () => {
      document.body.style.overflow = previousOverflow;
    };
  }, [isRendered]);

  useEffect(() => {
    if (!open || !isRendered) return;
    window.requestAnimationFrame(() => {
      closeButtonRef.current?.focus();
    });
  }, [isRendered, open]);

  useEffect(() => {
    return () => {
      if (closeTimerRef.current) {
        window.clearTimeout(closeTimerRef.current);
      }
    };
  }, []);

  useEffect(() => {
    if (!openActionTemplateId) return;

    const handlePointerDown = (event: PointerEvent) => {
      const target = event.target as Node;
      if (actionMenuRef.current?.contains(target)) return;
      if (actionButtonRefs.current[openActionTemplateId]?.contains(target)) return;
      closeActionMenu();
    };

    document.addEventListener("pointerdown", handlePointerDown);
    return () => {
      document.removeEventListener("pointerdown", handlePointerDown);
    };
  }, [closeActionMenu, openActionTemplateId]);

  const getFocusableElements = () =>
    Array.from(
      rootRef.current?.querySelectorAll<HTMLElement>(
        "a[href], button:not([disabled]), textarea:not([disabled]), input:not([disabled]), select:not([disabled]), [tabindex]:not([tabindex='-1'])"
      ) ?? []
    ).filter((element) => !element.hasAttribute("disabled") && element.offsetParent !== null);

  const handleRootKeyDown = (event: ReactKeyboardEvent<HTMLDivElement>) => {
    if (event.key === "Escape") {
      event.preventDefault();
      if (openActionTemplateId) {
        closeActionMenu({ restoreFocus: true });
        return;
      }
      requestClose({ restoreFocus: true });
      return;
    }

    if (event.key !== "Tab") return;

    const focusableElements = getFocusableElements();
    if (focusableElements.length === 0) {
      event.preventDefault();
      return;
    }

    const firstElement = focusableElements[0];
    const lastElement = focusableElements[focusableElements.length - 1];
    if (!firstElement || !lastElement) return;
    if (event.shiftKey && document.activeElement === firstElement) {
      event.preventDefault();
      lastElement.focus();
    } else if (!event.shiftKey && document.activeElement === lastElement) {
      event.preventDefault();
      firstElement.focus();
    }
  };

  const handleSelectTemplate = (templateId: string) => {
    requestClose();
    void onSelectTemplate(templateId);
  };

  const openActionMenu = (event: ReactMouseEvent<HTMLButtonElement>, templateId: string) => {
    event.stopPropagation();
    if (openActionTemplateId === templateId) {
      closeActionMenu({ restoreFocus: true });
      return;
    }

    const rect = event.currentTarget.getBoundingClientRect();
    const menuWidth = 168;
    const menuHeight = 142;
    const margin = 12;
    const left = Math.min(window.innerWidth - menuWidth - margin, Math.max(margin, rect.right - menuWidth));
    let top = rect.bottom + 8;
    if (top + menuHeight > window.innerHeight - margin) {
      top = Math.max(margin, rect.top - menuHeight - 8);
    }
    setActionMenuPosition({ top, left });
    setOpenActionTemplateId(templateId);
  };

  const handleTemplateAction = (
    event: ReactMouseEvent<HTMLButtonElement>,
    templateId: string,
    action: (templateId: string) => void | Promise<void>
  ) => {
    event.stopPropagation();
    requestClose();
    void action(templateId);
  };

  if (!mounted || !isRendered) return null;

  const scrimClass = isClosing
    ? "opacity-0 transition-opacity duration-[90ms] ease-in motion-reduce:transition-none"
    : "opacity-100 transition-opacity duration-[120ms] ease-out motion-reduce:transition-none";
  const dialogClass = isClosing
    ? "translate-y-[18px] opacity-[0.98] transition-[opacity,transform] duration-[120ms] ease-in motion-reduce:translate-y-0 motion-reduce:transition-none sm:translate-y-2 sm:duration-[110ms]"
    : "translate-y-0 opacity-100 transition-[opacity,transform] duration-[180ms] ease-[cubic-bezier(0.16,1,0.3,1)] motion-reduce:transition-none sm:duration-[160ms]";

  return createPortal(
    <div
      ref={rootRef}
      className="fixed inset-0 z-[150] flex items-end justify-center sm:items-center"
      onKeyDown={handleRootKeyDown}
      data-template-library-root="true"
    >
      <div
        aria-hidden="true"
        className={`absolute inset-0 bg-slate-950/[0.12] ${scrimClass}`}
        onPointerDown={() => requestClose({ restoreFocus: true })}
      />
      <div
        ref={dialogRef}
        role="dialog"
        aria-modal="true"
        aria-labelledby={titleId}
        aria-describedby={helperId}
        className={`relative z-10 flex max-h-[82dvh] w-full flex-col overflow-hidden rounded-t-[18px] border border-slate-200 bg-white shadow-[0_24px_70px_rgba(21,40,66,0.20)] [padding-bottom:env(safe-area-inset-bottom)] sm:max-h-[72dvh] sm:w-[560px] sm:max-w-[calc(100vw-40px)] sm:rounded-[16px] sm:pb-0 ${dialogClass}`}
      >
        <div className="flex flex-none items-start justify-between gap-4 border-b border-slate-200/80 bg-white px-5 py-[18px]">
          <div className="min-w-0">
            <h2 id={titleId} className="text-[22px] font-semibold leading-7 text-slate-950">
              Choose a template
            </h2>
            <p id={helperId} className="mt-2 max-w-[430px] text-[14px] leading-[1.4] text-slate-600">
              Load a saved workflow for this event. Event details and Anchor Field values remain specific to the current run.
            </p>
          </div>
          <button
            ref={closeButtonRef}
            type="button"
            onClick={() => requestClose({ restoreFocus: true })}
            className="flex h-[34px] w-[34px] shrink-0 items-center justify-center rounded-[9px] border border-slate-200 bg-white text-slate-500 transition hover:border-slate-300 hover:text-slate-700 focus:outline-none focus:ring-2 focus:ring-[#6f9fd1]/25"
            aria-label="Close template library"
          >
            <span aria-hidden="true" className="text-[20px] leading-none">
              ×
            </span>
          </button>
        </div>

        <div className="min-h-0 flex-1 overscroll-contain overflow-y-auto p-3">
          {templates.length > 0 ? (
            <div className="space-y-1.5">
              {templates.map((template) => {
                const isSelected = selectedTemplateId === template.id;
                const isHighlighted = highlightedTemplateId === template.id;
                const actionCount = template.items.length;
                const actionLabel = actionCount === 1 ? "1 action" : `${actionCount} actions`;
                return (
                  <div
                    key={template.id}
                    className={`grid min-h-[64px] grid-cols-[minmax(0,1fr)_auto] items-center gap-3 rounded-[11px] border px-3 py-2.5 ${
                      isSelected
                        ? "border-blue-200 bg-blue-50/80"
                        : isHighlighted
                          ? "border-green-200 bg-green-50/70"
                          : "border-slate-200 bg-white"
                    }`}
                  >
                    <div className="min-w-0">
                      <div className="truncate text-[15px] font-semibold leading-5 text-slate-950" title={template.name}>
                        {template.name}
                      </div>
                      <div className="mt-1 text-[12px] font-medium leading-4 text-slate-500">{actionLabel}</div>
                    </div>
                    <div className="flex shrink-0 items-center gap-2">
                      {isSelected ? (
                        <span className="inline-flex h-[29px] items-center rounded-[8px] bg-blue-100 px-3 text-[12px] font-semibold text-[#315f92]">
                          Current
                        </span>
                      ) : (
                        <button
                          type="button"
                          onClick={() => handleSelectTemplate(template.id)}
                          className="inline-flex h-9 min-w-[68px] items-center justify-center rounded-[9px] border border-blue-200 bg-blue-50/50 px-3 text-[13px] font-semibold text-[#315f92] transition hover:border-blue-300 hover:bg-blue-50 focus:outline-none focus:ring-2 focus:ring-[#6f9fd1]/25"
                          aria-label={`Use ${template.name} template`}
                        >
                          Use
                        </button>
                      )}
                      <button
                        ref={(node) => {
                          actionButtonRefs.current[template.id] = node;
                        }}
                        type="button"
                        onClick={(event) => openActionMenu(event, template.id)}
                        className="flex h-8 w-8 items-center justify-center rounded-[8px] border border-slate-200 bg-white text-slate-500 transition hover:border-slate-300 hover:bg-slate-50 hover:text-slate-700 focus:outline-none focus:ring-2 focus:ring-[#6f9fd1]/25"
                        aria-label={`Manage ${template.name}`}
                        aria-haspopup="menu"
                        aria-expanded={openActionTemplateId === template.id}
                      >
                        <span aria-hidden="true" className="text-[15px] leading-none">
                          •••
                        </span>
                      </button>
                    </div>
                  </div>
                );
              })}
            </div>
          ) : (
            <div className="flex min-h-[165px] flex-col items-center justify-center rounded-[12px] border border-dashed border-slate-200 bg-slate-50/50 px-5 text-center">
              <h3 className="text-[16px] font-semibold text-slate-950">No saved templates yet</h3>
              <p className="mt-2 max-w-[320px] text-[14px] leading-5 text-slate-600">
                Build a plan, then choose Save as Template to reuse it later.
              </p>
            </div>
          )}
        </div>

        <div className="flex flex-none items-center justify-between gap-3 border-t border-slate-200/80 bg-white px-4 py-3 sm:px-5 sm:py-4">
          <div className="hidden text-[13px] font-medium leading-5 text-slate-500 sm:block">
            Templates are managed from this library.
          </div>
          <span className="sr-only" aria-live="polite">
            {actionMessage}
          </span>
          <button
            type="button"
            onClick={() => requestClose({ restoreFocus: true })}
            className={`${plansSecondaryButtonClass} ml-auto h-[38px] rounded-[9px] px-4 text-[13px]`}
          >
            Close
          </button>
        </div>
      </div>

      {openActionTemplateId && actionMenuPosition ? (
        <div
          ref={actionMenuRef}
          role="menu"
          className="plans-menu-enter fixed z-[170] min-w-[168px] rounded-[10px] border border-slate-200 bg-white p-[7px] shadow-[0_18px_46px_rgba(21,40,66,0.18)]"
          style={{ top: actionMenuPosition.top, left: actionMenuPosition.left } as CSSProperties}
          data-template-action-menu="true"
        >
          {(() => {
            const template = templates.find((entry) => entry.id === openActionTemplateId);
            if (!template) return null;
            return (
              <>
                <button
                  type="button"
                  role="menuitem"
                  onClick={(event) => handleTemplateAction(event, template.id, onDuplicateTemplate)}
                  className="block h-10 w-full rounded-[8px] px-3 text-left text-[14px] font-medium text-slate-700 transition hover:bg-slate-50 focus:outline-none focus:ring-2 focus:ring-[#6f9fd1]/20"
                >
                  Copy
                </button>
                <button
                  type="button"
                  role="menuitem"
                  onClick={(event) => handleTemplateAction(event, template.id, onRenameTemplate)}
                  className="block h-10 w-full rounded-[8px] px-3 text-left text-[14px] font-medium text-slate-700 transition hover:bg-slate-50 focus:outline-none focus:ring-2 focus:ring-[#6f9fd1]/20"
                >
                  Rename
                </button>
                <button
                  type="button"
                  role="menuitem"
                  onClick={(event) => handleTemplateAction(event, template.id, onDeleteTemplate)}
                  className="block h-10 w-full rounded-[8px] px-3 text-left text-[14px] font-medium text-red-700 transition hover:bg-red-50 focus:outline-none focus:ring-2 focus:ring-red-200"
                >
                  Delete
                </button>
              </>
            );
          })()}
        </div>
      ) : null}
    </div>,
    document.body
  );
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

function renderTextWithBoldAnchors(
  value: string,
  options?: { anchorClassName?: string; invalidAnchorClassName?: string; knownAnchorKeys?: Set<string> }
) {
  const text = value ?? "";
  const parts = text.split(/(\[[^\]]+\])/g);
  const anchorClassName = options?.anchorClassName ?? "font-semibold text-current";
  const invalidAnchorClassName = options?.invalidAnchorClassName ?? "font-bold text-red-500";
  const knownAnchorKeys = options?.knownAnchorKeys;

  return parts.map((part, index) => {
    if (!/^\[[^\]]+\]$/.test(part)) {
      return <span key={`${part}-${index}`}>{part}</span>;
    }

    if (!knownAnchorKeys) {
      return (
        <span key={`${part}-${index}`} className={anchorClassName}>
          {part}
        </span>
      );
    }

    const isKnownAnchor = knownAnchorKeys.has(normalizeAnchorKey(part.slice(1, -1)));
    return (
      <span
        key={`${part}-${index}`}
        className={isKnownAnchor ? anchorClassName : invalidAnchorClassName}
      >
        {part}
      </span>
    );
  });
}

function formatAnchorTokenDisplay(key: string) {
  return `[${normalizeAnchorKey(key)}]`;
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
    timeZone: row.timeZone ?? "",
    emailDraft: row.rowType === "email" ? normalizeEmailDraft(row.emailDraft) : null,
    durationDraft: row.durationDraft ?? null,
    meetingDraft: row.meetingDraft ? normalizeMeetingDraft(row.meetingDraft) : null,
  };
}

function buildAnchorStateForType(type: PlanType, previous: BuilderAnchor[]) {
  const previousMap = new Map(previous.map((anchor) => [normalizeAnchorKey(anchor.key), anchor]));
  return getPresetAnchorKeysForType(type).map((key) => {
    const previousAnchor = previousMap.get(normalizeAnchorKey(key));
    return {
      ...createLockedAnchor(key, previousAnchor?.value ?? ""),
      isImportant: Boolean(previousAnchor?.isImportant),
      lastUpdatedAt: previousAnchor?.lastUpdatedAt ?? null,
    };
  });
}

function buildAnchorStateWithCoreEventFields(previous: BuilderAnchor[]) {
  const previousMap = new Map(previous.map((anchor) => [normalizeAnchorKey(anchor.key), anchor]));
  const genericAnchors = createGenericPresetAnchors().map((anchor) => {
    const previousAnchor = previousMap.get(normalizeAnchorKey(anchor.key));
    return {
      ...createLockedAnchor(anchor.key, previousAnchor?.value ?? ""),
      isImportant: Boolean(previousAnchor?.isImportant),
      lastUpdatedAt: previousAnchor?.lastUpdatedAt ?? null,
    };
  });
  const genericKeys = new Set(genericAnchors.map((anchor) => normalizeAnchorKey(anchor.key)));
  const extraAnchors = previous
    .filter((anchor) => !genericKeys.has(normalizeAnchorKey(anchor.key)))
    .map((anchor) => ({ ...anchor, id: crypto.randomUUID() }));
  return [...genericAnchors, ...extraAnchors];
}

function formatOffsetLabel(offsetDays: number | null | undefined, options?: { relativeToToday?: boolean; dateBasis?: PlanDateBasis }) {
  if (offsetDays == null) return "No days specified";
  const absoluteDays = Math.abs(offsetDays);
  const dayLabel = absoluteDays === 1 ? "day" : "days";
  const relativeToToday = options?.dateBasis === "today" || options?.relativeToToday;
  if (relativeToToday) {
    if (offsetDays < 0) return `${absoluteDays} ${dayLabel} before today`;
    if (offsetDays > 0) return `${absoluteDays} ${dayLabel} after today`;
    return "today";
  }
  if (offsetDays < 0) return `${absoluteDays} ${dayLabel} before`;
  if (offsetDays > 0) return `${absoluteDays} ${dayLabel} after`;
  return "Day of Event";
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
      timelineDotClass: "border-green-500",
    };
  }
  if (rowKind === "meeting") {
    return {
      label: "Meeting",
      className: "text-violet-600",
      badgeClass: "border-violet-200 bg-violet-50 text-violet-700",
      borderClass: "border-l-violet-400",
      timingPanelClass: "border-violet-100 bg-violet-50/60",
      timelineDotClass: "border-violet-500",
    };
  }
  if (row.rowType === "calendar_event") {
    return {
      label: "Calendar Event",
      className: "text-violet-600",
      badgeClass: "border-violet-200 bg-violet-50 text-violet-700",
      borderClass: "border-l-violet-400",
      timingPanelClass: "border-violet-100 bg-violet-50/60",
      timelineDotClass: "border-violet-500",
    };
  }
  return {
    label: "Reminder",
    className: "text-blue-600",
    badgeClass: "border-blue-200 bg-blue-50 text-blue-700",
    borderClass: "border-l-blue-400",
    timingPanelClass: "border-blue-100 bg-blue-50/60",
    timelineDotClass: "border-blue-500",
  };
}

function moveBuilderRowToIndex(rows: BuilderRow[], rowId: string, toIndex: number) {
  const fromIndex = rows.findIndex((row) => row.id === rowId);
  if (fromIndex === -1) return rows;

  const row = rows[fromIndex];
  if (!row) return rows;

  const remaining = rows.filter((entry) => entry.id !== rowId);
  const clampedIndex = Math.max(0, Math.min(toIndex, remaining.length));
  const nextRows = [...remaining];
  nextRows.splice(clampedIndex, 0, row);
  return nextRows;
}

type BuilderSortMode = "nearest_first" | "latest_first" | "type";

function getBuilderRowSortKind(row: BuilderRow) {
  return classifyPlanRow(row);
}

function sortBuilderRows(rows: BuilderRow[], mode: BuilderSortMode) {
  const rowsWithIndex = rows.map((row, index) => ({ row, index }));

  rowsWithIndex.sort((left, right) => {
    const leftOffset = left.row.offsetDays ?? 0;
    const rightOffset = right.row.offsetDays ?? 0;
    if (mode === "type") {
      const typeOrder = { reminder: 0, meeting: 1, email: 2 } as const;
      const leftType = typeOrder[getBuilderRowSortKind(left.row)];
      const rightType = typeOrder[getBuilderRowSortKind(right.row)];
      if (leftType !== rightType) return leftType - rightType;

      if (leftOffset !== rightOffset) return leftOffset - rightOffset;

      return left.index - right.index;
    }

    if (leftOffset !== rightOffset) {
      return mode === "nearest_first" ? leftOffset - rightOffset : rightOffset - leftOffset;
    }

    return left.index - right.index;
  });

  return rowsWithIndex.map(({ row }) => row);
}

function EmailTokensInput({
  label,
  values,
  onChange,
  placeholder,
  hasError,
  recipientGroups = [],
}: {
  label?: string;
  values: RecipientEntry[];
  onChange: (nextValues: RecipientEntry[]) => void;
  placeholder?: string;
  hasError?: boolean;
  recipientGroups?: RecipientGroup[];
}) {
  const [draftValue, setDraftValue] = useState("");
  const [hoveredGroupInfoKey, setHoveredGroupInfoKey] = useState<string | null>(null);

  function commitRawValue(raw: string) {
    const nextTokens = raw
      .split(/[,\n;]/)
      .map((entry) => createEmailRecipientEntry(entry))
      .filter((entry): entry is RecipientEntry => Boolean(entry));
    if (nextTokens.length === 0) return;
    onChange(mergeRecipientEntries(values, nextTokens));
    setDraftValue("");
  }

  return (
    <div>
      <label className="mb-1 block text-xs font-medium uppercase tracking-wide text-zinc-500">{label ?? ""}</label>
      <div
        className={`plans-token-input-shell flex min-h-[44px] flex-wrap items-center gap-2 rounded-lg border bg-white px-4 py-2 ${
          hasError ? "border-red-400 ring-1 ring-red-100" : "border-gray-300"
        }`}
      >
        {values.map((value) => (
          <span
            key={value.type === "group" ? `group:${value.groupId}` : `email:${normalizeRecipientGroupEmail(value.email)}`}
            className={`inline-flex items-center gap-1 rounded-full px-2 py-1 text-xs ${
              value.type === "group" ? "bg-slate-100 text-slate-950" : "bg-amber-100 text-amber-950"
            }`}
          >
            <span>{value.name}</span>
            {value.type === "group" ? (
              <span
                className="relative inline-flex h-4 w-4 items-center justify-center rounded-full border border-slate-300 bg-white text-[10px] font-semibold text-slate-500"
                onMouseEnter={() => setHoveredGroupInfoKey(value.groupId)}
                onMouseLeave={() =>
                  setHoveredGroupInfoKey((current) => (current === value.groupId ? null : current))
                }
              >
                i
                <span
                  className={`pointer-events-none absolute bottom-full left-1/2 z-10 mb-2 w-64 -translate-x-1/2 rounded-lg border border-slate-200 bg-white px-2 py-1 text-[11px] font-normal leading-4 text-slate-600 shadow-lg ${
                    hoveredGroupInfoKey === value.groupId ? "block" : "hidden"
                  }`}
                >
                  <span className="block font-semibold text-slate-900">{value.name}</span>
                  <span className="mt-1 block">
                    {getRecipientGroupFromEntry(value, recipientGroups)?.emails.join(", ") || "Group recipients unavailable."}
                  </span>
                </span>
              </span>
            ) : null}
            <button
              type="button"
              onClick={() => onChange(values.filter((entry) => entry !== value))}
              className={value.type === "group" ? "text-slate-600" : "text-amber-700"}
              aria-label={`Remove ${value.name}`}
            >
              ×
            </button>
          </span>
        ))}
        <input
          className="min-w-[140px] flex-1 border-0 bg-transparent px-1 py-0 text-sm text-gray-900 placeholder:text-gray-700 outline-none"
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
  const initialTemplates: SavedPlanTemplate[] = [];
  const initialSelectedTemplateId: string | null = null;
  const [appSettings, setAppSettings] = useState<AppSettings>(() => initialSettings);
  const [outlookConnection, setOutlookConnection] = useState<OutlookConnectionState | null>(() =>
    getOutlookConnectionState(initialSettings.outlookAccountEmail)
  );
  const [gmailConnection, setGmailConnection] = useState<GmailConnectionState | null>(() => getGmailConnectionState());
  const [builderMode, setBuilderMode] = useState<BuilderMode>("new");
  const [planType, setPlanType] = useState<PlanType>("press_release");
  const [templateName, setTemplateName] = useState("");
  const [eventName, setEventName] = useState("");
  const [anchorDate, setAnchorDate] = useState("");
  const [hasExplicitEventDate, setHasExplicitEventDate] = useState(false);
  const [eventTime, setEventTime] = useState("");
  const [eventTimeZone, setEventTimeZone] = useState(() => getDefaultOutlookTimeZone());
  const [eventDateInputValue, setEventDateInputValue] = useState("");
  const [eventTimeInputValue, setEventTimeInputValue] = useState("");
  const [noEventDate, setNoEventDate] = useState(false);
  const [weekendRule, setWeekendRule] = useState<WeekendRule>("none");
  const [rows, setRows] = useState<BuilderRow[]>([]);
  const [anchors, setAnchors] = useState<BuilderAnchor[]>(() => createGenericPresetAnchors());
  const [savedTemplates, setSavedTemplates] = useState<SavedPlanTemplate[]>(() => initialTemplates);
  const [selectedTemplateId, setSelectedTemplateId] = useState<string | null>(initialSelectedTemplateId);
  const [isBuilderSectionVisible, setIsBuilderSectionVisible] = useState(false);
  const [hasActivePlanSession, setHasActivePlanSession] = useState(false);
  const [lastTemplateSnapshot, setLastTemplateSnapshot] = useState<BuilderStateSnapshot | null>(null);
  const [lastBuilderSourceProvenance, setLastBuilderSourceProvenance] = useState<BuilderSourceProvenance | null>(null);
  const [guidedForm, setGuidedForm] = useState<GuidedFormState>(() => createEmptyGuidedForm());
  const [areAnchorsHidden, setAreAnchorsHidden] = useState(true);
  const [templateActionMessage, setTemplateActionMessage] = useState("");
  const [draggingRowId, setDraggingRowId] = useState<string | null>(null);
  const [dragInsertionIndex, setDragInsertionIndex] = useState<number | null>(null);
  const [openEmailDraftRowId, setOpenEmailDraftRowId] = useState<string | null>(null);
  const [openMeetingEditorRowId, setOpenMeetingEditorRowId] = useState<string | null>(null);
  const [openDurationEditorRowId, setOpenDurationEditorRowId] = useState<string | null>(null);
  const [focusedTitleInputId, setFocusedTitleInputId] = useState<string | null>(null);
  const [openTimeZoneRowId, setOpenTimeZoneRowId] = useState<string | null>(null);
  const [timeZoneSearch, setTimeZoneSearch] = useState("");
  const [, setForcedOpenMeetingEditorRowIds] = useState<string[]>([]);
  const [meetingValidationErrors, setMeetingValidationErrors] = useState<MeetingValidationErrorState>({});
  const [openBodyEditorRowId, setOpenBodyEditorRowId] = useState<string | null>(null);
  const [closingRowEditor, setClosingRowEditor] = useState<{ rowId: string; kind: RenderedRowEditorKind } | null>(null);
  const [plansModal, setPlansModal] = useState<PlansModalConfig | null>(null);
  const [plansModalInputValue, setPlansModalInputValue] = useState("");
  const [missingFieldHighlights, setMissingFieldHighlights] = useState<{
    eventName: boolean;
    eventDate: boolean;
    eventTime: boolean;
    anchorKeys: string[];
    rowIds: string[];
    fieldTargets: ValidationFieldTarget[];
  }>({ eventName: false, eventDate: false, eventTime: false, anchorKeys: [], rowIds: [], fieldTargets: [] });
  const [todayBasisTooltipRowId, setTodayBasisTooltipRowId] = useState<string | null>(null);
  const scheduledRowEditorOpenRef = useRef<number | null>(null);
  const scheduledRowEditorCloseRef = useRef<number | null>(null);
  const scheduledBuilderCloseCleanupRef = useRef<number | null>(null);
  const plansModalResolverRef = useRef<((value: boolean | string | null) => void) | null>(null);
  const [emailFieldVisibility, setEmailFieldVisibility] = useState<EmailFieldVisibility>({});
  const [editingOffsetRowId, setEditingOffsetRowId] = useState<string | null>(null);
  const [offsetDrafts, setOffsetDrafts] = useState<Record<string, string>>({});
  const [focusedTimeInputRowId, setFocusedTimeInputRowId] = useState<string | null>(null);
  const timeZoneOptions = useMemo(() => getSupportedTimeZones(), []);
  const [timeInputDrafts, setTimeInputDrafts] = useState<Record<string, string>>({});
  const [isSortMenuOpen, setIsSortMenuOpen] = useState(false);
  const [openWorkflowRowMenuId, setOpenWorkflowRowMenuId] = useState<string | null>(null);
  const [workflowRowMenuPosition, setWorkflowRowMenuPosition] = useState<FloatingMenuPosition | null>(null);
  const [openAnchorMenuId, setOpenAnchorMenuId] = useState<string | null>(null);
  const [anchorMenuPosition, setAnchorMenuPosition] = useState<FloatingMenuPosition | null>(null);
  const [instantlyVisibleRowId, setInstantlyVisibleRowId] = useState<string | null>(null);
  const [addRowSettlingRowId, setAddRowSettlingRowId] = useState<string | null>(null);
  const [hiddenAddRowId, setHiddenAddRowId] = useState<string | null>(null);
  const [, setInlineEditorFocusPhase] = useState<InlineEditorFocusPhase>("idle");
  const [isBuilderPreviewOpen, setIsBuilderPreviewOpen] = useState(false);
  const [openPreviewDetail, setOpenPreviewDetail] = useState<{ rowId: string; kind: "reminder" | "email" | "meeting" } | null>(null);
  const [openPreviewRowMenuId, setOpenPreviewRowMenuId] = useState<string | null>(null);
  const [excludedPreviewItemIds, setExcludedPreviewItemIds] = useState<string[]>([]);
  const [showNoPreviewItemsSelectedCallout, setShowNoPreviewItemsSelectedCallout] = useState(false);
  const previewModalScrollRef = useRef<HTMLDivElement | null>(null);
  const previewItemListHeadingRef = useRef<HTMLDivElement | null>(null);
  const previewDetailRowRefs = useRef<Record<string, HTMLDivElement | null>>({});
  const [isBuilderVisualRevealDeferred, setIsBuilderVisualRevealDeferred] = useState(false);
  const [isBuilderEntryRevealImmediate, setIsBuilderEntryRevealImmediate] = useState(false);
  const [isBuilderEntryRunwayVisible, setIsBuilderEntryRunwayVisible] = useState(false);
  const [builderScrollRequestNonce, setBuilderScrollRequestNonce] = useState(0);
  const [executionNotices, setExecutionNotices] = useState<ExecutionNoticeEntry[]>([]);
  const [executionState, setExecutionState] = useState<"pending" | "success" | "failure" | null>(null);
  const exportQueueRef = useRef(Promise.resolve());
  const activeExportCountRef = useRef(0);
  const [providerLoading, setProviderLoading] = useState({ outlook: true, gmail: true });
  const [previewLoading, setPreviewLoading] = useState(true);
  const [lastDynamicFieldsExportAt, setLastDynamicFieldsExportAt] = useState<string | null>(null);
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
  const [isTemplateLibraryOpen, setIsTemplateLibraryOpen] = useState(false);
  const [showAiApplyConfirm, setShowAiApplyConfirm] = useState(false);
  const [aiApplySuccessMessage, setAiApplySuccessMessage] = useState<string | null>(null);
  const [showNewPlanDialog, setShowNewPlanDialog] = useState(false);
  const [isNewPlanSetupPending, setIsNewPlanSetupPending] = useState(false);
  const [planSetupDialogMode, setPlanSetupDialogMode] = useState<PlanSetupDialogMode>("new");
  const [planSetupTemplateId, setPlanSetupTemplateId] = useState<string | null>(null);
  const [isPopupAnimatedIn, setIsPopupAnimatedIn] = useState(false);
  const [recipientGroups, setRecipientGroups] = useState<RecipientGroup[]>([]);
  const [isRecipientGroupsModalOpen, setIsRecipientGroupsModalOpen] = useState(false);
  const [recipientGroupsModalMode, setRecipientGroupsModalMode] = useState<"select" | "create" | "edit">("select");
  const [recipientGroupsEditingGroup, setRecipientGroupsEditingGroup] = useState<RecipientGroup | null>(null);
  const [recipientGroupsModalTarget, setRecipientGroupsModalTarget] = useState<RecipientGroupsModalTarget | null>(null);
  const [newPlanDraft, setNewPlanDraft] = useState<NewPlanDraft>({
    eventName: "",
    anchorDate: "",
    eventTime: "",
    noEventDate: false,
    weekendRule: "none",
  });
  const [newPlanDialogMessage, setNewPlanDialogMessage] = useState<string | null>(null);
  const [builderSourceProvenance, setBuilderSourceProvenance] = useState<BuilderSourceProvenance | null>(null);
  const [hasMounted, setHasMounted] = useState(false);
  const [hasHydratedTemplates, setHasHydratedTemplates] = useState(false);
  const [hasHydratedBuilderDraft, setHasHydratedBuilderDraft] = useState(false);
  const anchorFieldsRegionId = useId();
  const hasLocalTemplateMutationRef = useRef(false);
  const isBuilderDraftPersistencePausedRef = useRef(false);
  const builderTimeInputRefs = useRef<Record<string, HTMLInputElement | HTMLTextAreaElement | null>>({});
  const builderTitleInputRefs = useRef<Record<string, HTMLTextAreaElement | null>>({});
  const builderOffsetInputRefs = useRef<Record<string, HTMLInputElement | null>>({});
  const sortMenuRef = useRef<HTMLDivElement | null>(null);
  const workflowRowMenuRef = useRef<HTMLDivElement | null>(null);
  const workflowRowMenuButtonRefs = useRef<Record<string, HTMLButtonElement | null>>({});
  const anchorMenuRef = useRef<HTMLDivElement | null>(null);
  const anchorMenuButtonRefs = useRef<Record<string, HTMLButtonElement | null>>({});
  const templateLibraryButtonRef = useRef<HTMLButtonElement | null>(null);
  const addReminderButtonRef = useRef<HTMLButtonElement | null>(null);
  const addEmailButtonRef = useRef<HTMLButtonElement | null>(null);
  const addMeetingButtonRef = useRef<HTMLButtonElement | null>(null);
  const rowInsertAnchorRef = useRef<HTMLDivElement | null>(null);
  const rowListStackRef = useRef<HTMLDivElement | null>(null);
  const scheduledAddRowLaunchRef = useRef<number | null>(null);
  const scheduledAddRowInsertRef = useRef<number | null>(null);
  const scheduledBuilderEntryRef = useRef<number | null>(null);
  const scheduledRowEditorAttentionRefs = useRef<number[]>([]);
  const scheduledInlineEditorRevealRef = useRef<number | null>(null);
  const scheduledInlineEditorDimRef = useRef<number | null>(null);
  const aiConversationRef = useRef<HTMLDivElement | null>(null);
  const aiComposerRef = useRef<HTMLTextAreaElement | null>(null);
  const builderSectionRef = useRef<HTMLElement | null>(null);
  const eventHeaderCardRef = useRef<HTMLDivElement | null>(null);
  const eventNameInputRef = useRef<HTMLInputElement | null>(null);
  const eventDateInputRef = useRef<HTMLInputElement | null>(null);
  const eventTimeInputRef = useRef<HTMLInputElement | null>(null);
  const weekendHandlingSelectRef = useRef<HTMLSelectElement | null>(null);
  const useTodayInputRef = useRef<HTMLInputElement | null>(null);
  const shouldFocusEventNameInputRef = useRef(false);
  const shouldScrollBuilderIntoViewRef = useRef(false);
  const builderScrollRequestIdRef = useRef(0);
  const rowsRef = useRef<BuilderRow[]>(rows);
  const rowNodeRefs = useRef<Record<string, HTMLDivElement | null>>({});
  const rowEditorPanelRefs = useRef<Record<string, HTMLDivElement | null>>({});
  const rowEditorFocusReturnRef = useRef<HTMLElement | null>(null);
  const dragInsertionIndexRef = useRef<number | null>(null);
  const [pressedRowId, setPressedRowId] = useState<string | null>(null);
  const activeDragRef = useRef<{
    pointerId: number;
    rowId: string;
    startX: number;
    startY: number;
    isDragging: boolean;
  } | null>(null);

  function getLatestPreviewPlan() {
    return resolvePlanAnchors(
      createPlan({
        name: effectivePlanName,
        type: planType,
        anchorDate: previewAnchorDateForComputation,
        weekendRule,
        template: buildTemplateItemsFromRows(rowsRef.current, recipientGroups),
      }),
      anchorMap
    );
  }

  useEffect(() => {
    setHasMounted(true);
  }, []);

  useEffect(() => {
    if (!isSortMenuOpen) return;

    function handlePointerDown(event: MouseEvent) {
      const target = event.target as Node | null;
      if (sortMenuRef.current?.contains(target)) return;
      setIsSortMenuOpen(false);
    }

    document.addEventListener("mousedown", handlePointerDown);
    return () => {
      document.removeEventListener("mousedown", handlePointerDown);
    };
  }, [isSortMenuOpen]);

  function getFloatingMenuPositionFromRect(rect: DOMRect, menuWidth: number) {
    const viewportPadding = 12;
    const left = Math.min(
      Math.max(viewportPadding, rect.right - menuWidth),
      Math.max(viewportPadding, window.innerWidth - menuWidth - viewportPadding)
    );
    return {
      top: Math.min(rect.bottom + 6, window.innerHeight - viewportPadding),
      left,
    };
  }

  const closeWorkflowRowMenu = useCallback(
    (options?: { restoreFocus?: boolean }) => {
      const currentMenuId = openWorkflowRowMenuId;
      setOpenWorkflowRowMenuId(null);
      setWorkflowRowMenuPosition(null);
      if (options?.restoreFocus && currentMenuId) {
        requestAnimationFrame(() => {
          workflowRowMenuButtonRefs.current[currentMenuId]?.focus();
        });
      }
    },
    [openWorkflowRowMenuId]
  );

  const closeAnchorMenu = useCallback(
    (options?: { restoreFocus?: boolean }) => {
      const currentMenuId = openAnchorMenuId;
      setOpenAnchorMenuId(null);
      setAnchorMenuPosition(null);
      if (options?.restoreFocus && currentMenuId) {
        requestAnimationFrame(() => {
          anchorMenuButtonRefs.current[currentMenuId]?.focus();
        });
      }
    },
    [openAnchorMenuId]
  );

  function openWorkflowRowMenu(event: ReactMouseEvent<HTMLButtonElement>, rowId: string) {
    event.stopPropagation();
    const rect = event.currentTarget.getBoundingClientRect();
    setOpenAnchorMenuId(null);
    setAnchorMenuPosition(null);
    setWorkflowRowMenuPosition(getFloatingMenuPositionFromRect(rect, 184));
    setOpenWorkflowRowMenuId((current) => (current === rowId ? null : rowId));
  }

  function openAnchorMenu(event: ReactMouseEvent<HTMLButtonElement>, anchorId: string) {
    event.stopPropagation();
    const rect = event.currentTarget.getBoundingClientRect();
    setOpenWorkflowRowMenuId(null);
    setWorkflowRowMenuPosition(null);
    setAnchorMenuPosition(getFloatingMenuPositionFromRect(rect, 194));
    setOpenAnchorMenuId((current) => (current === anchorId ? null : anchorId));
  }

  useEffect(() => {
    if (!openWorkflowRowMenuId) return;
    const menuRowId = openWorkflowRowMenuId;

    function handlePointerDown(event: MouseEvent) {
      const target = event.target as Node | null;
      if (workflowRowMenuRef.current?.contains(target)) return;
      if (workflowRowMenuButtonRefs.current[menuRowId]?.contains(target)) return;
      closeWorkflowRowMenu();
    }

    function handleKeyDown(event: KeyboardEvent) {
      if (event.key !== "Escape") return;
      event.preventDefault();
      closeWorkflowRowMenu({ restoreFocus: true });
    }

    document.addEventListener("mousedown", handlePointerDown);
    document.addEventListener("keydown", handleKeyDown);
    return () => {
      document.removeEventListener("mousedown", handlePointerDown);
      document.removeEventListener("keydown", handleKeyDown);
    };
  }, [closeWorkflowRowMenu, openWorkflowRowMenuId]);

  useEffect(() => {
    if (!openAnchorMenuId) return;
    const menuAnchorId = openAnchorMenuId;

    function handlePointerDown(event: MouseEvent) {
      const target = event.target as Node | null;
      if (anchorMenuRef.current?.contains(target)) return;
      if (anchorMenuButtonRefs.current[menuAnchorId]?.contains(target)) return;
      closeAnchorMenu();
    }

    function handleKeyDown(event: KeyboardEvent) {
      if (event.key !== "Escape") return;
      event.preventDefault();
      closeAnchorMenu({ restoreFocus: true });
    }

    document.addEventListener("mousedown", handlePointerDown);
    document.addEventListener("keydown", handleKeyDown);
    return () => {
      document.removeEventListener("mousedown", handlePointerDown);
      document.removeEventListener("keydown", handleKeyDown);
    };
  }, [closeAnchorMenu, openAnchorMenuId]);

  function focusAndSelectOffsetInput(rowId: string) {
    requestAnimationFrame(() => {
      const input = builderOffsetInputRefs.current[rowId];
      if (!input) return;
      input.focus();
      input.select();
    });
  }

  function beginEditingOffsetForRow(rowId: string, offsetDays: number | null | undefined) {
    setEditingOffsetRowId(rowId);
    setOffsetDrafts((current) => ({
      ...current,
      [rowId]: offsetDays == null ? "" : String(offsetDays),
    }));
    focusAndSelectOffsetInput(rowId);
  }

  function focusBuilderTimeInput(rowId: string) {
    requestAnimationFrame(() => {
      const input = builderTimeInputRefs.current[rowId];
      if (!input) return;
      input.focus();
      if (input instanceof HTMLInputElement) {
        input.select();
      } else {
        const valueLength = input.value.length;
        input.setSelectionRange(0, valueLength);
      }
    });
  }

  function focusCollapsedTitleInput(rowId: string) {
    requestAnimationFrame(() => {
      const input = builderTitleInputRefs.current[`collapsed-title-${rowId}`];
      if (!input) return;
      input.focus();
      input.select();
    });
  }

  function applyEventTimeZone(nextTimeZone: string) {
    const normalizedTimeZone = normalizeOutlookTimeZone(nextTimeZone);
    setEventTimeZone(normalizedTimeZone);
    setRows((current) => current.map((row) => ({ ...row, timeZone: normalizedTimeZone })));
  }

  function applyBuilderSort(mode: BuilderSortMode) {
    if (mode !== "type") {
      const previewItemsById = new Map(getLatestPreviewPlan().items.map((item) => [item.id, item]));
      setRows((current) => {
        const rowsWithIndex = current.map((row, index) => {
          const previewItem = previewItemsById.get(row.id);
          return {
            row,
            index,
            scheduledTimestamp: previewItem ? getPreviewItemScheduledTimestamp(previewItem) : null,
          };
        });

        rowsWithIndex.sort((left, right) => {
          if (left.scheduledTimestamp != null && right.scheduledTimestamp != null) {
            if (left.scheduledTimestamp !== right.scheduledTimestamp) {
              return mode === "nearest_first"
                ? left.scheduledTimestamp - right.scheduledTimestamp
                : right.scheduledTimestamp - left.scheduledTimestamp;
            }
          } else if (left.scheduledTimestamp != null) {
            return -1;
          } else if (right.scheduledTimestamp != null) {
            return 1;
          }

          const leftOffset = left.row.offsetDays ?? 0;
          const rightOffset = right.row.offsetDays ?? 0;
          if (leftOffset !== rightOffset) {
            return mode === "nearest_first" ? leftOffset - rightOffset : rightOffset - leftOffset;
          }

          return left.index - right.index;
        });

        return rowsWithIndex.map(({ row }) => row);
      });
      setIsSortMenuOpen(false);
      return;
    }

    setRows((current) => sortBuilderRows(current, mode));
    setIsSortMenuOpen(false);
  }

  function commitOffsetDraft(rowId: string) {
    const draftValue = offsetDrafts[rowId];
    updateRow(rowId, (current) => {
      const parsedValue = Number(draftValue);
      return {
        ...current,
        offsetDays: Number.isNaN(parsedValue) ? current.offsetDays ?? 0 : parsedValue,
      };
    });
    setEditingOffsetRowId((current) => (current === rowId ? null : current));
  }

  function nudgeOffsetDraft(rowId: string, delta: number) {
    setOffsetDrafts((current) => {
      const currentValue = current[rowId];
      const fallbackRow = rowsRef.current.find((row) => row.id === rowId);
      const baseValue =
        currentValue != null && currentValue.trim() !== ""
          ? Number(currentValue)
          : fallbackRow?.offsetDays ?? 0;
      const safeBaseValue = Number.isNaN(baseValue) ? 0 : baseValue;
      return {
        ...current,
        [rowId]: String(safeBaseValue + delta),
      };
    });
    focusAndSelectOffsetInput(rowId);
  }

  function nudgeRowOffset(rowId: string, delta: number) {
    updateRow(rowId, (current) => ({
      ...current,
      offsetDays: (current.offsetDays ?? 0) + delta,
    }));
  }

  const focusNextEventHeaderField = useCallback(
    (field: "eventDate" | "eventTime" | "weekendHandling" | "useToday") => {
      if (field === "eventDate") {
        eventDateInputRef.current?.focus();
        return;
      }
      if (field === "eventTime") {
        eventTimeInputRef.current?.focus();
        return;
      }
      if (field === "weekendHandling") {
        weekendHandlingSelectRef.current?.focus();
        return;
      }
      useTodayInputRef.current?.focus();
    },
    []
  );

  const focusEventNameInput = useCallback((attempt = 0) => {
    const input = eventNameInputRef.current;
    if (input) {
      input.focus({ preventScroll: true });
      input.setSelectionRange(eventName.length, eventName.length);
      shouldFocusEventNameInputRef.current = false;
      return;
    }

    if (attempt >= 5) return;
    window.requestAnimationFrame(() => {
      focusEventNameInput(attempt + 1);
    });
  }, [eventName.length]);

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
    const hasOpenPopup = Boolean(
      plansModal ||
      showAiApplyConfirm ||
      showNewPlanDialog ||
      showAiTemplateSaveDialog ||
      showBuilderTemplateSaveDialog
    );

    if (!hasOpenPopup) {
      setIsPopupAnimatedIn(false);
      return;
    }

    const frameId = window.requestAnimationFrame(() => {
      setIsPopupAnimatedIn(true);
    });

    return () => window.cancelAnimationFrame(frameId);
  }, [plansModal, showAiApplyConfirm, showAiTemplateSaveDialog, showBuilderTemplateSaveDialog, showNewPlanDialog]);

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

  useEffect(() => {
    let active = true;

    async function refreshRecipientGroups() {
      const nextGroups = await hydrateRecipientGroupsFromSupabase();
      if (!active) return;
      setRecipientGroups(nextGroups);
    }

    void refreshRecipientGroups();

    function handleRecipientGroupsUpdated() {
      void refreshRecipientGroups();
    }

    window.addEventListener(RECIPIENT_GROUPS_UPDATED_EVENT, handleRecipientGroupsUpdated);
    return () => {
      active = false;
      window.removeEventListener(RECIPIENT_GROUPS_UPDATED_EVENT, handleRecipientGroupsUpdated);
    };
  }, []);

  // The initial remote template hydration intentionally runs once on mount.
  /* eslint-disable react-hooks/exhaustive-deps */
  useEffect(() => {
    let active = true;

    async function hydrateTemplates() {
      setPreviewLoading(true);
      const seedState: PersistedPlanTemplate[] = [];
      const remoteState = await loadTemplateStateFromSupabase(seedState);
      const rawState = remoteState ?? loadCachedTemplateState(seedState);
      const migratedState = migratePersistedTemplatesToCustom(rawState);
      const nextState = ensureCoreCustomTemplates(
        migratedState,
        loadAppSettings().defaultReminderTime,
        loadAppSettings().defaultPressReleaseTime
      );

      if (!active) return;
      if (hasLocalTemplateMutationRef.current) {
        setHasHydratedTemplates(true);
        setHasHydratedBuilderDraft(true);
        setPreviewLoading(false);
        return;
      }

      saveCachedTemplateState(nextState);
      if (remoteState) {
        void saveTemplateStateToSupabase(nextState);
      }
      const nextTemplates = nextState.templates
        .filter((template) => !template.id.startsWith("seed:"))
        .map((template) => fromPersistedTemplate(template));
      const shouldOpenNewPlanFromEntry =
        consumePlansSidebarNeutralEntry() ||
        (typeof window !== "undefined" && new URLSearchParams(window.location.search).get("new") === "1");
      if (shouldOpenNewPlanFromEntry && typeof window !== "undefined") {
        window.history.replaceState(null, "", window.location.pathname);
      }
      if (shouldOpenNewPlanFromEntry) {
        clearPersistedBuilderDraft();
      }
      const persistedDraftSnapshot = shouldOpenNewPlanFromEntry ? null : loadPersistedBuilderDraft();
      setSavedTemplates(nextTemplates);
      setHasHydratedTemplates(true);

      if (shouldOpenNewPlanFromEntry) {
        void startNewPlan({ keepViewAtTop: true });
      } else if (persistedDraftSnapshot) {
        const restoredDraftSnapshot: BuilderStateSnapshot = {
          ...persistedDraftSnapshot,
          rows: persistedDraftSnapshot.rows,
          anchors:
            persistedDraftSnapshot.anchors.length > 0
              ? persistedDraftSnapshot.anchors
              : createGenericPresetAnchors(),
        };
        applyBuilderSnapshot(restoredDraftSnapshot);
        setIsBuilderSectionVisible(true);
        setBuilderSourceProvenance(
          buildBuilderSourceProvenance(restoredDraftSnapshot, {
            sourceType: "manual",
            sourceLabel: "Recovered draft",
          })
        );
        setLastTemplateSnapshot(restoredDraftSnapshot);
      }
      setHasHydratedBuilderDraft(true);
      setPreviewLoading(false);
    }

    void hydrateTemplates();

    return () => {
      active = false;
    };
  }, []);
  /* eslint-enable react-hooks/exhaustive-deps */

  // Sidebar navigation should always open Plans as a fresh plan workspace.
  /* eslint-disable react-hooks/exhaustive-deps */
  useEffect(() => {
    function handlePlansSidebarNeutralEntry() {
      consumePlansSidebarNeutralEntry();
      void startNewPlan({ keepViewAtTop: true });
    }

    window.addEventListener(PLANS_SIDEBAR_NEUTRAL_EVENT, handlePlansSidebarNeutralEntry);
    return () => {
      window.removeEventListener(PLANS_SIDEBAR_NEUTRAL_EVENT, handlePlansSidebarNeutralEntry);
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
    if (!hasHydratedTemplates || !hasHydratedBuilderDraft) return;

    const timeoutId = window.setTimeout(() => {
      if (isBuilderDraftPersistencePausedRef.current) {
        clearPersistedBuilderDraft();
        return;
      }

      const hasMeaningfulName = Boolean(templateName.trim() || eventName.trim());
      const hasMeaningfulTiming = Boolean(anchorDate.trim() || eventTime.trim() || noEventDate || hasExplicitEventDate);
      const hasMeaningfulAnchorValue = anchors.some((anchor) => anchor.value.trim());
      const hasMeaningfulGuidedForm = Object.values(guidedForm).some((value) =>
        typeof value === "boolean" ? value : Boolean(String(value).trim())
      );
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

      if (
        !(
          hasMeaningfulName ||
          hasMeaningfulTiming ||
          hasMeaningfulAnchorValue ||
          hasMeaningfulGuidedForm ||
          hasMeaningfulRowContent ||
          hasMultipleRows
        )
      ) {
        clearPersistedBuilderDraft();
        return;
      }

      savePersistedBuilderDraft({
        builderMode,
        selectedTemplateId,
        planType,
        templateName,
        eventName,
        anchorDate,
        hasExplicitEventDate,
        eventTime,
        eventTimeZone,
        noEventDate,
        weekendRule,
        rows: cloneTemplateRows(rows),
        anchors: cloneAnchors(anchors),
        guidedForm: { ...guidedForm },
        lastDynamicFieldsExportAt,
      });
    }, 250);

    return () => {
      window.clearTimeout(timeoutId);
    };
  }, [
    anchorDate,
    anchors,
    builderMode,
    eventName,
    eventTime,
    eventTimeZone,
    guidedForm,
    hasExplicitEventDate,
    hasHydratedBuilderDraft,
    hasHydratedTemplates,
    lastDynamicFieldsExportAt,
    noEventDate,
    planType,
    rows,
    selectedTemplateId,
    templateName,
    weekendRule,
  ]);

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
  const currentSelectedTemplateActionCount = currentSelectedTemplate?.items.length ?? 0;
  const currentSelectedTemplateActionLabel =
    currentSelectedTemplateActionCount === 1 ? "1 action" : `${currentSelectedTemplateActionCount} actions`;
  const currentSourceLine = currentSelectedTemplate
    ? `Using ${currentSelectedTemplate.name} template · ${currentSelectedTemplateActionLabel}`
    : "Starting from a blank plan";
  const templateLibraryButtonLabel = currentSelectedTemplate ? "Change template" : "Use template";
  const planSetupTemplate = useMemo(
    () => (planSetupTemplateId ? savedTemplates.find((template) => template.id === planSetupTemplateId) ?? null : null),
    [planSetupTemplateId, savedTemplates]
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
  const genericEventAnchorTime = eventTime.trim();
  const resolvedAnchors = useMemo(
    () =>
      anchors.map((anchor) => {
        const derivedValue = getDerivedAnchorValue(planType, effectivePlanName, anchorDate, anchor.key, guidedForm, {
          eventNameValue: genericEventAnchorName,
          eventDateValue: genericEventAnchorDate,
          eventTimeValue: genericEventAnchorTime,
        });
        const resolvedValue =
          anchor.locked || isCoreEventAnchorKey(anchor.key)
            ? derivedValue ?? anchor.value
            : anchor.value;
        return {
          id: anchor.id,
          key: anchor.key,
          value: resolvedValue,
          displayValue: getAnchorDisplayValue(anchor.key, resolvedValue),
          locked: anchor.locked,
          isImportant: Boolean(anchor.isImportant),
          lastUpdatedAt: anchor.lastUpdatedAt ?? null,
        };
      }),
    [anchorDate, anchors, effectivePlanName, genericEventAnchorDate, genericEventAnchorName, genericEventAnchorTime, guidedForm, planType]
  );
  const anchorMap = useMemo(() => buildAnchorMap(resolvedAnchors), [resolvedAnchors]);
  const knownAnchorKeys = useMemo(
    () => new Set(resolvedAnchors.map((anchor) => normalizeAnchorKey(anchor.key))),
    [resolvedAnchors]
  );
  const resolvedAutoFilledAnchorValueKeys = useMemo(
    () =>
      new Set(
        resolvedAnchors
          .filter((anchor) => Boolean(anchor.locked || isCoreEventAnchorKey(anchor.key)))
          .filter((anchor) => anchor.value.trim())
          .map((anchor) => normalizeAnchorKey(anchor.key))
      ),
    [resolvedAnchors]
  );
  useEffect(() => {
    if (missingFieldHighlights.anchorKeys.length === 0 || resolvedAutoFilledAnchorValueKeys.size === 0) return;

    setMissingFieldHighlights((current) => {
      const nextAnchorKeys = current.anchorKeys.filter(
        (anchorKey) => !resolvedAutoFilledAnchorValueKeys.has(normalizeAnchorKey(anchorKey))
      );

      if (nextAnchorKeys.length === current.anchorKeys.length) return current;
      return { ...current, anchorKeys: nextAnchorKeys };
    });
  }, [missingFieldHighlights.anchorKeys, resolvedAutoFilledAnchorValueKeys]);
  const previewTemplate = useMemo(() => buildTemplateItemsFromRows(rows, recipientGroups), [recipientGroups, rows]);
  const renderedRows = useMemo(() => {
    return draggingRowId
      ? moveBuilderRowToIndex(rows, draggingRowId, dragInsertionIndex ?? rows.findIndex((row) => row.id === draggingRowId))
      : rows;
  }, [dragInsertionIndex, draggingRowId, rows]);
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
  const shouldRenderSimpleEventHeaderFields = true;

  const getDragInsertionIndex = useCallback((pointerY: number, draggedRowId: string) => {
    const rowsExcludingDragged = renderedRows.filter((row) => row.id !== draggedRowId);
    for (let index = 0; index < rowsExcludingDragged.length; index += 1) {
      const row = rowsExcludingDragged[index];
      const node = rowNodeRefs.current[row.id];
      if (!node) continue;
      const rect = node.getBoundingClientRect();
      if (pointerY < rect.top + rect.height / 2) {
        return index;
      }
    }
    return rowsExcludingDragged.length;
  }, [renderedRows]);

  const finishRowDrag = useCallback((shouldCommit: boolean) => {
    const activeDrag = activeDragRef.current;
    if (!activeDrag) return;

    if (shouldCommit && activeDrag.isDragging) {
      setRows((current) => {
        const currentIndex = current.findIndex((row) => row.id === activeDrag.rowId);
        const nextRows = moveBuilderRowToIndex(
          current,
          activeDrag.rowId,
          dragInsertionIndexRef.current ?? dragInsertionIndex ?? currentIndex
        );
        rowsRef.current = nextRows;
        return nextRows;
      });
    }

    activeDragRef.current = null;
    dragInsertionIndexRef.current = null;
    setPressedRowId(null);
    setDraggingRowId(null);
    setDragInsertionIndex(null);
  }, [dragInsertionIndex]);

  useEffect(() => {
    if (!draggingRowId) return;
    const previousUserSelect = document.body.style.userSelect;
    document.body.style.userSelect = "none";
    return () => {
      document.body.style.userSelect = previousUserSelect;
    };
  }, [draggingRowId]);

  useEffect(() => {
    if (!draggingRowId && !pressedRowId) return;

    const handlePointerUp = (event: PointerEvent) => {
      const activeDrag = activeDragRef.current;
      if (!activeDrag || activeDrag.pointerId !== event.pointerId) {
        setPressedRowId(null);
        return;
      }

      const activeRowNode = rowNodeRefs.current[activeDrag.rowId];
      if (activeRowNode?.hasPointerCapture(activeDrag.pointerId)) {
        activeRowNode.releasePointerCapture(activeDrag.pointerId);
      }

      finishRowDrag(activeDrag.isDragging);
    };

    const handlePointerCancel = (event: PointerEvent) => {
      const activeDrag = activeDragRef.current;
      if (!activeDrag || activeDrag.pointerId !== event.pointerId) {
        setPressedRowId(null);
        return;
      }

      const activeRowNode = rowNodeRefs.current[activeDrag.rowId];
      if (activeRowNode?.hasPointerCapture(activeDrag.pointerId)) {
        activeRowNode.releasePointerCapture(activeDrag.pointerId);
      }

      finishRowDrag(false);
    };

    window.addEventListener("pointerup", handlePointerUp);
    window.addEventListener("pointercancel", handlePointerCancel);

    return () => {
      window.removeEventListener("pointerup", handlePointerUp);
      window.removeEventListener("pointercancel", handlePointerCancel);
    };
  }, [draggingRowId, finishRowDrag, pressedRowId]);

  const clearScheduledRowEditorOpen = useCallback(() => {
    if (scheduledRowEditorOpenRef.current == null) return;
    window.clearTimeout(scheduledRowEditorOpenRef.current);
    scheduledRowEditorOpenRef.current = null;
  }, []);

  const clearScheduledAddRowLaunch = useCallback(() => {
    if (scheduledAddRowLaunchRef.current == null) return;
    window.clearTimeout(scheduledAddRowLaunchRef.current);
    scheduledAddRowLaunchRef.current = null;
  }, []);

  const clearScheduledAddRowInsert = useCallback(() => {
    if (scheduledAddRowInsertRef.current == null) return;
    window.clearTimeout(scheduledAddRowInsertRef.current);
    scheduledAddRowInsertRef.current = null;
  }, []);

  const clearScheduledRowEditorAttention = useCallback(() => {
    scheduledRowEditorAttentionRefs.current.forEach((timeoutId) => window.clearTimeout(timeoutId));
    scheduledRowEditorAttentionRefs.current = [];
  }, []);

  const clearScheduledBuilderEntry = useCallback(() => {
    if (scheduledBuilderEntryRef.current == null) return;
    window.clearTimeout(scheduledBuilderEntryRef.current);
    scheduledBuilderEntryRef.current = null;
  }, []);

  const clearScheduledRowEditorClose = useCallback(() => {
    if (scheduledRowEditorCloseRef.current == null) return;
    window.clearTimeout(scheduledRowEditorCloseRef.current);
    scheduledRowEditorCloseRef.current = null;
  }, []);

  const clearScheduledInlineEditorFocus = useCallback(() => {
    if (scheduledInlineEditorRevealRef.current !== null) {
      window.clearTimeout(scheduledInlineEditorRevealRef.current);
      scheduledInlineEditorRevealRef.current = null;
    }
    if (scheduledInlineEditorDimRef.current !== null) {
      window.clearTimeout(scheduledInlineEditorDimRef.current);
      scheduledInlineEditorDimRef.current = null;
    }
  }, []);

  const closeAllRowEditors = useCallback(() => {
    clearScheduledRowEditorOpen();
    clearScheduledRowEditorClose();
    clearScheduledAddRowLaunch();
    clearScheduledAddRowInsert();
    clearScheduledRowEditorAttention();
    clearScheduledInlineEditorFocus();
    setInlineEditorFocusPhase("idle");
    setOpenEmailDraftRowId(null);
    setOpenMeetingEditorRowId(null);
    setOpenDurationEditorRowId(null);
    setOpenBodyEditorRowId(null);
    setClosingRowEditor(null);
  }, [
    clearScheduledAddRowInsert,
    clearScheduledAddRowLaunch,
    clearScheduledInlineEditorFocus,
    clearScheduledRowEditorAttention,
    clearScheduledRowEditorClose,
    clearScheduledRowEditorOpen,
  ]);

  const clearScheduledBuilderCloseCleanup = useCallback(() => {
    if (scheduledBuilderCloseCleanupRef.current == null) return;
    window.clearTimeout(scheduledBuilderCloseCleanupRef.current);
    scheduledBuilderCloseCleanupRef.current = null;
  }, []);

  useEffect(() => {
    return () => {
      clearScheduledAddRowLaunch();
      clearScheduledAddRowInsert();
      clearScheduledBuilderEntry();
      clearScheduledRowEditorAttention();
      clearScheduledInlineEditorFocus();
    };
  }, [
    clearScheduledAddRowInsert,
    clearScheduledAddRowLaunch,
    clearScheduledBuilderEntry,
    clearScheduledInlineEditorFocus,
    clearScheduledRowEditorAttention,
  ]);

  const closeBuilderSectionAfterAnimation = useCallback((afterClose?: () => void) => {
    clearScheduledBuilderCloseCleanup();
    setIsBuilderSectionVisible(false);
    scheduledBuilderCloseCleanupRef.current = window.setTimeout(() => {
      scheduledBuilderCloseCleanupRef.current = null;
      afterClose?.();
    }, ROW_EDITOR_CLOSE_ANIMATION_MS);
  }, [clearScheduledBuilderCloseCleanup]);

  const getCurrentOpenRowEditor = useCallback((): { rowId: string; kind: RenderedRowEditorKind } | null => {
    if (openEmailDraftRowId) return { rowId: openEmailDraftRowId, kind: "email" };
    if (openMeetingEditorRowId) return { rowId: openMeetingEditorRowId, kind: "meeting" };
    if (openBodyEditorRowId || openDurationEditorRowId) {
      return { rowId: openBodyEditorRowId ?? openDurationEditorRowId ?? "", kind: "reminder" };
    }
    if (openBodyEditorRowId) return { rowId: openBodyEditorRowId, kind: "reminderBody" };
    if (openDurationEditorRowId) return { rowId: openDurationEditorRowId, kind: "reminderDuration" };
    return null;
  }, [openBodyEditorRowId, openDurationEditorRowId, openEmailDraftRowId, openMeetingEditorRowId]);

  const applyRowEditorOpen = useCallback((rowId: string, kind: RowEditorKind) => {
    clearScheduledRowEditorOpen();
    clearScheduledRowEditorClose();
    clearScheduledInlineEditorFocus();
    const activeElement = document.activeElement;
    const fallbackFocusTarget =
      kind === "email"
        ? addEmailButtonRef.current
        : kind === "meeting"
          ? addMeetingButtonRef.current
          : addReminderButtonRef.current;
    rowEditorFocusReturnRef.current =
      activeElement instanceof HTMLElement && activeElement !== document.body ? activeElement : fallbackFocusTarget;
    setInlineEditorFocusPhase("expanding");
    scheduledInlineEditorRevealRef.current = window.setTimeout(() => {
      scheduledInlineEditorRevealRef.current = null;
      setInlineEditorFocusPhase("revealing");
    }, INLINE_EDITOR_EXPAND_ANIMATION_MS);
    scheduledInlineEditorDimRef.current = window.setTimeout(() => {
      scheduledInlineEditorDimRef.current = null;
      setInlineEditorFocusPhase("dimmed");
    }, INLINE_EDITOR_EXPAND_ANIMATION_MS + INLINE_EDITOR_REVEAL_ANIMATION_MS);
    setAreAnchorsHidden(false);
    if (kind === "email") {
      setOpenEmailDraftRowId(rowId);
      return;
    }
    if (kind === "meeting") {
      setOpenMeetingEditorRowId(rowId);
      return;
    }
    setOpenBodyEditorRowId(rowId);
    setOpenDurationEditorRowId(rowId);
  }, [clearScheduledInlineEditorFocus, clearScheduledRowEditorClose, clearScheduledRowEditorOpen]);

  const bringRowEditorPanelIntoView = useCallback((rowId: string) => {
    const editorPanel = rowEditorPanelRefs.current[rowId];
    if (!editorPanel) return;

    const rect = editorPanel.getBoundingClientRect();
    const viewportHeight = window.innerHeight;
    const targetTop = Math.max(72, viewportHeight * 0.14);
    const nextScrollTop = window.scrollY + rect.top - targetTop;
    if (Math.abs(rect.top - targetTop) < 8) return;

    window.scrollTo({
      top: Math.max(0, nextScrollTop),
      behavior: "smooth",
    });
  }, []);

  const scheduleRowEditorAttention = useCallback((rowId: string) => {
    clearScheduledRowEditorAttention();
    scheduledRowEditorAttentionRefs.current = [80, 320, 720, INLINE_EDITOR_EXPAND_ANIMATION_MS + 80].map((delayMs) =>
      window.setTimeout(() => {
        bringRowEditorPanelIntoView(rowId);
      }, delayMs)
    );
  }, [bringRowEditorPanelIntoView, clearScheduledRowEditorAttention]);

  const prepareThenOpenRowEditor = useCallback((rowId: string, kind: RowEditorKind) => {
    clearScheduledRowEditorOpen();
    clearScheduledRowEditorAttention();
    applyRowEditorOpen(rowId, kind);
    scheduleRowEditorAttention(rowId);
  }, [applyRowEditorOpen, clearScheduledRowEditorAttention, clearScheduledRowEditorOpen, scheduleRowEditorAttention]);

  const closeRowEditorsAfterAnimation = useCallback((afterClose?: () => void) => {
    clearScheduledRowEditorOpen();
    clearScheduledRowEditorClose();
    clearScheduledRowEditorAttention();
    const currentOpenEditor =
      openEmailDraftRowId
        ? { rowId: openEmailDraftRowId, kind: "email" as const }
        : openMeetingEditorRowId
          ? { rowId: openMeetingEditorRowId, kind: "meeting" as const }
          : openBodyEditorRowId || openDurationEditorRowId
            ? { rowId: openBodyEditorRowId ?? openDurationEditorRowId ?? "", kind: "reminder" as const }
            : null;

    if (!currentOpenEditor) {
      setClosingRowEditor(null);
      clearScheduledInlineEditorFocus();
      setInlineEditorFocusPhase("idle");
      afterClose?.();
      return;
    }

    clearScheduledInlineEditorFocus();
    setInlineEditorFocusPhase("idle");
    setClosingRowEditor(currentOpenEditor);
    setOpenEmailDraftRowId(null);
    setOpenMeetingEditorRowId(null);
    setOpenDurationEditorRowId(null);
    setOpenBodyEditorRowId(null);
    scheduledRowEditorCloseRef.current = window.setTimeout(() => {
      scheduledRowEditorCloseRef.current = null;
      setClosingRowEditor(null);
      const focusReturnTarget = rowEditorFocusReturnRef.current;
      const fallbackFocusTarget =
        currentOpenEditor.kind === "email"
          ? addEmailButtonRef.current
          : currentOpenEditor.kind === "meeting"
            ? addMeetingButtonRef.current
            : addReminderButtonRef.current;
      rowEditorFocusReturnRef.current = null;
      const safeFocusReturnTarget =
        focusReturnTarget && document.contains(focusReturnTarget) ? focusReturnTarget : fallbackFocusTarget;
      if (safeFocusReturnTarget && document.contains(safeFocusReturnTarget)) {
        requestAnimationFrame(() => {
          safeFocusReturnTarget.focus({ preventScroll: true });
        });
      }
      afterClose?.();
    }, ROW_EDITOR_CLOSE_ANIMATION_MS);
  }, [
    clearScheduledInlineEditorFocus,
    clearScheduledRowEditorAttention,
    clearScheduledRowEditorClose,
    clearScheduledRowEditorOpen,
    openBodyEditorRowId,
    openDurationEditorRowId,
    openEmailDraftRowId,
    openMeetingEditorRowId,
  ]);

  const openExclusiveRowEditor = useCallback((rowId: string, kind: RowEditorKind) => {
    const currentOpenEditor = getCurrentOpenRowEditor();

    if (currentOpenEditor && currentOpenEditor.rowId === rowId && currentOpenEditor.kind === kind) {
      closeRowEditorsAfterAnimation();
      return;
    }

    if (currentOpenEditor) {
      closeRowEditorsAfterAnimation(() => {
        prepareThenOpenRowEditor(rowId, kind);
      });
      return;
    }

    prepareThenOpenRowEditor(rowId, kind);
  }, [closeRowEditorsAfterAnimation, getCurrentOpenRowEditor, prepareThenOpenRowEditor]);

  const scheduleRowEditorOpen = useCallback((rowId: string, kind: RowEditorKind) => {
    clearScheduledRowEditorOpen();
    window.requestAnimationFrame(() => {
      openExclusiveRowEditor(rowId, kind);
    });
  }, [clearScheduledRowEditorOpen, openExclusiveRowEditor]);

  const isAnyBuilderRowEditorVisible = Boolean(
    openEmailDraftRowId || openMeetingEditorRowId || openBodyEditorRowId || openDurationEditorRowId || closingRowEditor
  );
  const inactiveEditorDimTransitionClass = "plans-editor-dim-target";
  const builderInactiveControlClass = inactiveEditorDimTransitionClass;

  useEffect(() => {
    if (!isAnyBuilderRowEditorVisible) return;

    const previousBodyOverflow = document.body.style.overflow;
    const previousDocumentOverflow = document.documentElement.style.overflow;
    document.body.style.overflow = "hidden";
    document.documentElement.style.overflow = "hidden";

    return () => {
      document.body.style.overflow = previousBodyOverflow;
      document.documentElement.style.overflow = previousDocumentOverflow;
    };
  }, [isAnyBuilderRowEditorVisible]);

  useEffect(() => {
    if (!isAnyBuilderRowEditorVisible) return;
    setIsSortMenuOpen(false);
    setOpenTimeZoneRowId(null);
  }, [isAnyBuilderRowEditorVisible]);

  useEffect(
    () => () => {
      clearScheduledRowEditorOpen();
      clearScheduledRowEditorClose();
      clearScheduledBuilderCloseCleanup();
      plansModalResolverRef.current = null;
    },
    [clearScheduledBuilderCloseCleanup, clearScheduledRowEditorClose, clearScheduledRowEditorOpen]
  );

  const openPlansModal = useCallback((config: PlansModalConfig) => {
    setPlansModalInputValue(config.defaultValue ?? "");
    setPlansModal(config);
    return new Promise<boolean | string | null>((resolve) => {
      plansModalResolverRef.current = resolve;
    });
  }, []);

  const closePlansModal = useCallback((result: boolean | string | null) => {
    const resolver = plansModalResolverRef.current;
    plansModalResolverRef.current = null;
    setPlansModal(null);
    setPlansModalInputValue("");
    resolver?.(result);
  }, []);

  const showAlertModal = useCallback(
    async (config: Omit<PlansModalConfig, "kind">) => {
      await openPlansModal({ ...config, kind: "alert" });
    },
    [openPlansModal]
  );

  const showConfirmModal = useCallback(
    async (config: Omit<PlansModalConfig, "kind">) => {
      const result = await openPlansModal({ ...config, kind: "confirm" });
      return result === true;
    },
    [openPlansModal]
  );

  const showPromptModal = useCallback(
    async (config: Omit<PlansModalConfig, "kind">) => {
      const result = await openPlansModal({ ...config, kind: "prompt" });
      return typeof result === "string" ? result : null;
    },
    [openPlansModal]
  );

  function clearMissingFieldHighlights() {
    setMissingFieldHighlights({ eventName: false, eventDate: false, eventTime: false, anchorKeys: [], rowIds: [], fieldTargets: [] });
  }

  function isValidationFieldHighlighted(rowId: string, field: ValidationFieldName) {
    return missingFieldHighlights.fieldTargets.some((target) => target.rowId === rowId && target.field === field);
  }

  function getValidationFieldHighlightClass(rowId: string, field: ValidationFieldName) {
    return isValidationFieldHighlighted(rowId, field) ? "!border-red-400 !ring-2 !ring-red-200" : "";
  }

  function clearValidationFieldHighlight(rowId: string, field: ValidationFieldName) {
    setMissingFieldHighlights((current) => ({
      ...current,
      fieldTargets: current.fieldTargets.filter((target) => target.rowId !== rowId || target.field !== field),
      rowIds: current.fieldTargets.some((target) => target.rowId === rowId && target.field !== field)
        ? current.rowIds
        : current.rowIds.filter((id) => id !== rowId),
    }));
  }

  function openEditorForValidationTarget(target: ValidationFieldTarget) {
    const row = rowsRef.current.find((entry) => entry.id === target.rowId);
    if (!row) return;
    const rowKind = classifyPlanRow(row);
    if (rowKind === "email") {
      applyRowEditorOpen(target.rowId, "email");
      return;
    }
    if (rowKind === "meeting") {
      applyRowEditorOpen(target.rowId, "meeting");
      return;
    }
    applyRowEditorOpen(target.rowId, "reminder");
  }

  function focusValidationTarget(target: ValidationFieldTarget) {
    const selector = `[data-validation-field="${target.rowId}:${target.field}"]`;
    const focusField = () => {
      const element =
        document.querySelector<HTMLElement>(`${selector}[data-validation-field-scope="editor"]`) ??
        document.querySelector<HTMLElement>(selector);
      if (element) {
        element.scrollIntoView({ behavior: "smooth", block: "center" });
        if (element instanceof HTMLInputElement || element instanceof HTMLTextAreaElement || element instanceof HTMLButtonElement) {
          element.focus({ preventScroll: true });
        } else {
          element.querySelector<HTMLElement>("input, textarea, button")?.focus({ preventScroll: true });
        }
        return;
      }
      rowNodeRefs.current[target.rowId]?.scrollIntoView({ behavior: "smooth", block: "center" });
    };

    window.setTimeout(focusField, INLINE_EDITOR_EXPAND_ANIMATION_MS + 80);
  }

  function togglePreviewDetail(rowId: string, kind: "reminder" | "email" | "meeting") {
    setOpenPreviewDetail((current) =>
      current?.rowId === rowId && current.kind === kind
        ? null
        : { rowId, kind }
    );
  }

  function isPreviewItemIncluded(itemId: string) {
    return !excludedPreviewItemIds.includes(itemId);
  }

  function togglePreviewItemIncluded(itemId: string, checked: boolean) {
    if (checked) {
      setShowNoPreviewItemsSelectedCallout(false);
    }
    setExcludedPreviewItemIds((current) =>
      checked
        ? current.filter((id) => id !== itemId)
        : current.includes(itemId)
          ? current
          : [...current, itemId]
    );
  }

  function getIncludedPreviewItemIds() {
    return getLatestPreviewPlan()
      .items
      .filter((item) => isPreviewItemIncluded(item.id))
      .map((item) => item.id);
  }

  function getPreviewItemScheduledTimestamp(item: Plan["items"][number]) {
    const rowKind = classifyPlanRow(item);

    if (rowKind === "meeting") {
      const timing = buildPreviewItemGraphTiming(item);
      const scheduledAt = new Date(timing.startISO);
      return Number.isNaN(scheduledAt.getTime()) ? null : scheduledAt.getTime();
    }

    const dueDate = getEffectivePreviewItemDate(item);
    const reminderTime = getUsableReminderTime(item.reminderTime, anchorMap);
    if (!dueDate || !reminderTime) return null;

    const scheduledAt = new Date(buildLocalDateTimeIso(dueDate, reminderTime));
    return Number.isNaN(scheduledAt.getTime()) ? null : scheduledAt.getTime();
  }

  function isPreviewItemScheduledInPast(item: Plan["items"][number]) {
    const scheduledTimestamp = getPreviewItemScheduledTimestamp(item);
    return scheduledTimestamp != null && scheduledTimestamp < Date.now();
  }

  function formatPastScheduledItemDateTime(item: Plan["items"][number]) {
    const scheduledTimestamp = getPreviewItemScheduledTimestamp(item);
    if (scheduledTimestamp != null) {
      const scheduledAt = new Date(scheduledTimestamp);
      return `${scheduledAt.toLocaleDateString("en-US", {
        month: "short",
        day: "numeric",
        year: "numeric",
      })} · ${scheduledAt.toLocaleTimeString("en-US", {
        hour: "numeric",
        minute: "2-digit",
      })}`;
    }

    const dueDate = getEffectivePreviewItemDate(item);
    const timeLabel =
      item.meetingDraft?.isAllDay || item.durationDraft?.isAllDay
        ? "All day"
        : formatPreviewTime(getUsableReminderTime(item.reminderTime, anchorMap) ?? "");
    if (!dueDate) return timeLabel || "No scheduled time";
    const [year, month, day] = dueDate.split("-").map(Number);
    const date = new Date((year ?? 2000), (month ?? 1) - 1, day ?? 1);
    return `${date.toLocaleDateString("en-US", {
      month: "short",
      day: "numeric",
      year: "numeric",
    })}${timeLabel ? ` · ${timeLabel}` : ""}`;
  }

  function renderPastScheduledItemsContent(itemIds?: string[]) {
    const previewItems = getLatestPreviewPlan().items;
    const scopedItems = itemIds ? previewItems.filter((item) => itemIds.includes(item.id)) : previewItems;
    const pastItems = scopedItems.filter(isPreviewItemScheduledInPast);
    const itemCountLabel = pastItems.length === 1 ? "1 selected item" : `${pastItems.length} selected items`;
    const visibleItems = pastItems.slice(0, 5);
    const hiddenCount = Math.max(0, pastItems.length - visibleItems.length);

    return (
      <div className="space-y-3">
        <div className="text-[13px] font-semibold text-amber-700">{itemCountLabel}</div>
        <div className="max-h-[32dvh] overflow-y-auto rounded-[10px] border border-amber-200/80 bg-amber-50/45">
          <div className="divide-y divide-amber-200/70">
            {visibleItems.map((item) => {
              const rowKind = classifyPlanRow(item);
              const typeLabel = rowKind === "email" ? "Email" : rowKind === "meeting" ? "Meeting" : "Reminder";
              const title = item.customTitle ?? item.title ?? `Untitled ${typeLabel.toLowerCase()}`;
              return (
                <div key={item.id} className="grid gap-1 px-3.5 py-3 min-[560px]:grid-cols-[92px_minmax(0,1fr)_190px] min-[560px]:gap-3">
                  <div className="text-[13px] font-semibold text-amber-700">{typeLabel}</div>
                  <div className="min-w-0 text-[14px] font-semibold leading-5 text-slate-900">{title}</div>
                  <div className="min-w-0 text-[13px] leading-5 text-slate-600 tabular-nums">
                    {formatPastScheduledItemDateTime(item)}
                  </div>
                </div>
              );
            })}
            {hiddenCount > 0 ? (
              <div className="px-3.5 py-3 text-[13px] font-semibold text-amber-700">+ {hiddenCount} more</div>
            ) : null}
          </div>
        </div>
      </div>
    );
  }

  function getReviewExportItemsForRender(items: Plan["items"]) {
    return items
      .map((item, index) => ({ item, index }))
      .sort((left, right) => {
        const leftTimestamp = getPreviewItemScheduledTimestamp(left.item);
        const rightTimestamp = getPreviewItemScheduledTimestamp(right.item);
        if (leftTimestamp != null && rightTimestamp != null && leftTimestamp !== rightTimestamp) {
          return leftTimestamp - rightTimestamp;
        }
        if (leftTimestamp != null && rightTimestamp == null) return -1;
        if (leftTimestamp == null && rightTimestamp != null) return 1;

        const leftOffset = left.item.offsetDays ?? 0;
        const rightOffset = right.item.offsetDays ?? 0;
        if (leftOffset !== rightOffset) return leftOffset - rightOffset;

        return left.index - right.index;
      })
      .map(({ item }) => item);
  }

  useEffect(() => {
    if (!isBuilderPreviewOpen || !openPreviewDetail) return;

    let frameId = 0;
    const timeoutIds: number[] = [];

    const scrollOpenDetailToCenter = () => {
      const scroller = previewModalScrollRef.current;
      const row = previewDetailRowRefs.current[openPreviewDetail.rowId];
      if (!scroller || !row) return;

      const scrollerRect = scroller.getBoundingClientRect();
      const rowRect = row.getBoundingClientRect();
      const centeredTop = scroller.scrollTop + rowRect.top - scrollerRect.top - (scroller.clientHeight - rowRect.height) / 2;
      const maxScrollTop = scroller.scrollHeight - scroller.clientHeight;

      scroller.scrollTo({
        top: Math.max(0, Math.min(centeredTop, maxScrollTop)),
        behavior: "smooth",
      });
    };

    frameId = window.requestAnimationFrame(scrollOpenDetailToCenter);
    timeoutIds.push(window.setTimeout(scrollOpenDetailToCenter, ROW_EDITOR_OPEN_ANIMATION_MS + 80));

    return () => {
      window.cancelAnimationFrame(frameId);
      timeoutIds.forEach((timeoutId) => window.clearTimeout(timeoutId));
    };
  }, [isBuilderPreviewOpen, openPreviewDetail]);

  function loadPersistedBuilderDraft() {
    if (typeof window === "undefined") return null;

    const raw = readPersistedValue("localStorage", getBuilderDraftStorageKey());
    if (!raw) return null;

    try {
      const parsed = JSON.parse(raw) as { version?: number; snapshot?: unknown } | unknown;
      const snapshotValue =
        isObject(parsed) && "snapshot" in parsed
          ? (parsed as { snapshot?: unknown }).snapshot
          : parsed;
      return normalizeBuilderDraftSnapshot(snapshotValue);
    } catch {
      removePersistedValue("localStorage", getBuilderDraftStorageKey());
      return null;
    }
  }

  function savePersistedBuilderDraft(snapshot: BuilderStateSnapshot) {
    if (typeof window === "undefined") return;
    writePersistedValue(
      "localStorage",
      getBuilderDraftStorageKey(),
      JSON.stringify({
        version: 1,
        savedAt: new Date().toISOString(),
        snapshot,
      })
    );
  }

  function clearPersistedBuilderDraft() {
    if (typeof window === "undefined") return;
    removePersistedValue("localStorage", getBuilderDraftStorageKey());
  }

  function applyMissingFieldHighlights(issues: MissingFieldIssue[]) {
    const anchorKeys = new Set<string>();
    const rowIds = new Set<string>();
    const fieldTargetMap = new Map<string, ValidationFieldTarget>();

    issues.forEach((issue) => {
      if (issue.anchorKey) anchorKeys.add(normalizeAnchorKey(issue.anchorKey));
      if (issue.rowId) rowIds.add(issue.rowId);
      issue.rowIds?.forEach((rowId) => rowIds.add(rowId));
      issue.fieldTargets?.forEach((target) => {
        rowIds.add(target.rowId);
        fieldTargetMap.set(`${target.rowId}:${target.field}`, target);
      });
    });

    const fieldTargets = Array.from(fieldTargetMap.values());

    setMissingFieldHighlights({
      eventName: issues.some((issue) => issue.eventName),
      eventDate: issues.some((issue) => issue.eventDate),
      eventTime: issues.some((issue) => issue.eventTime),
      anchorKeys: Array.from(anchorKeys),
      rowIds: Array.from(rowIds),
      fieldTargets,
    });

    if (anchorKeys.size > 0) {
      setAreAnchorsHidden(false);
    }

    window.setTimeout(() => {
      const firstFieldTarget = fieldTargets[0];
      if (firstFieldTarget) {
        openEditorForValidationTarget(firstFieldTarget);
        focusValidationTarget(firstFieldTarget);
        return;
      }
      if (issues.some((issue) => issue.eventName)) {
        eventNameInputRef.current?.scrollIntoView({ behavior: "smooth", block: "center" });
        eventNameInputRef.current?.focus({ preventScroll: true });
        return;
      }
      if (issues.some((issue) => issue.eventDate)) {
        eventDateInputRef.current?.scrollIntoView({ behavior: "smooth", block: "center" });
        eventDateInputRef.current?.focus({ preventScroll: true });
        return;
      }
      if (issues.some((issue) => issue.eventTime)) {
        eventTimeInputRef.current?.scrollIntoView({ behavior: "smooth", block: "center" });
        eventTimeInputRef.current?.focus({ preventScroll: true });
        return;
      }
      const firstAnchorKey = Array.from(anchorKeys)[0];
      if (firstAnchorKey) {
        const matchingAnchor = anchors.find((anchor) => normalizeAnchorKey(anchor.key) === firstAnchorKey);
        window.requestAnimationFrame(() => {
          window.setTimeout(() => {
            const anchorTarget = matchingAnchor
              ? document.querySelector<HTMLInputElement>(
                  `[data-plan-anchor-value="${matchingAnchor.id}"], [data-plan-anchor-key="${matchingAnchor.id}"]`
                )
              : document.querySelector<HTMLElement>('[data-plan-anchor-surface="true"]');
            anchorTarget?.scrollIntoView({ behavior: "smooth", block: "center" });
            anchorTarget?.focus?.({ preventScroll: true });
          }, 80);
        });
        return;
      }
      builderSectionRef.current?.scrollIntoView({ behavior: "smooth", block: "start" });
    }, 0);
  }

  function persistTemplateStateImmediately(nextTemplates: SavedPlanTemplate[], nextSelectedTemplateId: string | null) {
    const nextState = buildPersistedTemplateState(nextTemplates, nextSelectedTemplateId);
    saveCachedTemplateState(nextState);
    if (hasHydratedTemplates) {
      void saveTemplateStateToSupabase(nextState);
    }
  }

  function persistSelectedTemplateDynamicFieldTimestamp(nextTimestamp: string | null) {
    setLastDynamicFieldsExportAt(nextTimestamp);
    if (!selectedTemplateId) return;
    if (!savedTemplates.some((template) => template.id === selectedTemplateId)) return;

    hasLocalTemplateMutationRef.current = true;
    const nextTemplates = savedTemplates.map((template) =>
      template.id === selectedTemplateId
        ? { ...template, lastDynamicFieldsExportAt: nextTimestamp }
        : template
    );
    setSavedTemplates(nextTemplates);
    persistTemplateStateImmediately(nextTemplates, selectedTemplateId);
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

  function openRecipientGroupsModal(target: RecipientGroupsModalTarget) {
    setRecipientGroupsModalTarget(target);
    setRecipientGroupsEditingGroup(null);
    setRecipientGroupsModalMode("select");
    setIsRecipientGroupsModalOpen(true);
  }

  function applyRecipientGroupToTarget(group: RecipientGroup) {
    if (!recipientGroupsModalTarget) return;

    const { rowId, field } = recipientGroupsModalTarget;
    const groupEntry = createGroupRecipientEntry(group);
    if (field === "email_to") {
      updateRow(rowId, (current) => ({
        ...current,
        emailDraft: {
          ...normalizeEmailDraft(current.emailDraft),
          to: mergeRecipientEntries(normalizeEmailDraft(current.emailDraft).to, [groupEntry]),
        },
      }));
    } else {
      updateRow(rowId, (current) => ({
        ...current,
        meetingDraft: {
          ...normalizeMeetingDraft(current.meetingDraft),
          attendees: mergeRecipientEntries(normalizeMeetingDraft(current.meetingDraft)?.attendees ?? [], [groupEntry]),
        },
      }));
    }

    setIsRecipientGroupsModalOpen(false);
    setRecipientGroupsEditingGroup(null);
    setRecipientGroupsModalTarget(null);
  }

  async function handleSaveRecipientGroup(input: { id?: string; name: string; emails: string[] }) {
    const savedGroup = await saveRecipientGroup(input);
    setRecipientGroupsEditingGroup(savedGroup);
    return savedGroup;
  }

  async function handleDeleteRecipientGroup(group: RecipientGroup) {
    const confirmed = await showConfirmModal({
      title: "Delete recipient group?",
      message: `“${group.name}” will be removed. Plans that reference this group will no longer resolve its recipients.`,
      confirmLabel: "Delete recipient group",
      cancelLabel: "Cancel",
      destructive: true,
      severity: "destructive",
    });
    if (!confirmed) return;
    await deleteRecipientGroup(group.id);
    setRecipientGroupsEditingGroup(null);
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

  function renderSortControl() {
    if (renderedRows.length === 0) return null;

    return (
      <div ref={sortMenuRef} className="relative">
        <button
          type="button"
          onClick={() => setIsSortMenuOpen((current) => !current)}
          className={`${plansToolbarButtonClass} min-w-[86px] gap-2 border-slate-200 text-slate-700 hover:border-slate-300 hover:bg-slate-50`}
          aria-label="Sort rows"
          aria-expanded={isSortMenuOpen}
        >
          <svg
            viewBox="0 0 20 20"
            aria-hidden="true"
            className="h-4 w-4"
            fill="none"
            stroke="currentColor"
            strokeWidth="1.8"
            strokeLinecap="round"
            strokeLinejoin="round"
          >
            <path d="M4 5h8" />
            <path d="M4 10h12" />
            <path d="M4 15h6" />
          </svg>
          <span>Sort</span>
        </button>
        {isSortMenuOpen ? (
          <div className="plans-menu-enter absolute right-0 top-full z-[170] mt-1.5 min-w-[220px] rounded-[10px] border border-slate-200 bg-white p-[7px] shadow-[0_18px_46px_rgba(21,40,66,0.18)]">
            <button
              type="button"
              onClick={() => applyBuilderSort("nearest_first")}
              className="flex h-10 w-full items-center justify-between rounded-[8px] px-3 text-left text-[14px] font-medium text-slate-700 transition hover:bg-slate-50 focus:outline-none focus:ring-2 focus:ring-[#6f9fd1]/20"
            >
              <span>Closest to now</span>
            </button>
            <button
              type="button"
              onClick={() => applyBuilderSort("latest_first")}
              className="flex h-10 w-full items-center justify-between rounded-[8px] px-3 text-left text-[14px] font-medium text-slate-700 transition hover:bg-slate-50 focus:outline-none focus:ring-2 focus:ring-[#6f9fd1]/20"
            >
              <span>Furthest from now</span>
            </button>
            <button
              type="button"
              onClick={() => applyBuilderSort("type")}
              className="flex h-10 w-full items-center justify-between rounded-[8px] px-3 text-left text-[14px] font-medium text-slate-700 transition hover:bg-slate-50 focus:outline-none focus:ring-2 focus:ring-[#6f9fd1]/20"
            >
              <span>By reminder type</span>
            </button>
          </div>
        ) : null}
      </div>
    );
  }

  function renderDynamicFieldsSection(options?: { inlineEditor?: boolean }) {
    const inlineEditor = options?.inlineEditor ?? false;

    if (inlineEditor) {
      return (
        <section className="space-y-3">
          <div className="flex flex-wrap items-center justify-between gap-2">
            <div className="text-[14px] font-semibold text-slate-700">Anchor Fields</div>
            <button
              type="button"
              data-inspector-anchor-add="true"
              onClick={() => {
                setAnchors((current) => [...current, createEmptyAnchor()]);
                setAreAnchorsHidden(false);
              }}
              className="inline-flex h-[36px] shrink-0 items-center justify-center whitespace-nowrap rounded-lg border border-slate-200 bg-white px-3 text-[13px] font-semibold text-slate-700 transition hover:border-slate-300 hover:bg-slate-50 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#6f9fd1]/30 max-[420px]:text-[12px]"
            >
              + Add Anchor Field
            </button>
          </div>

          <div className="divide-y divide-slate-200/80 overflow-hidden rounded-xl border border-slate-200/80 bg-white">
            {anchors.map((anchor) => {
              const resolvedAnchor = resolvedAnchors.find((entry) => entry.id === anchor.id);
              const isCoreEventAnchor = isCoreEventAnchorKey(anchor.key);
              const isReadOnlyAnchor = Boolean(anchor.locked || isCoreEventAnchor);
              const isMissingFieldHighlighted = missingFieldHighlights.anchorKeys.includes(normalizeAnchorKey(anchor.key));
              const displayedAnchorValue = isReadOnlyAnchor
                ? getAnchorDisplayValue(
                    anchor.key,
                    getDerivedAnchorValue(planType, effectivePlanName, anchorDate, anchor.key, guidedForm, {
                      eventNameValue: eventName.trim(),
                      eventDateValue: noEventDate ? "" : anchorDate,
                      eventTimeValue: eventTime.trim(),
                    }) ??
                      resolvedAnchor?.value ??
                      anchor.value
                  )
                : anchor.value;

              return (
                <div
                  key={anchor.id}
                  className={`grid grid-cols-[28px_minmax(0,1fr)_32px] items-center gap-2 px-3 py-2.5 min-[620px]:min-h-[58px] min-[620px]:grid-cols-[28px_minmax(150px,0.75fr)_minmax(180px,1.25fr)_72px_72px_32px] ${
                    isMissingFieldHighlighted ? "bg-red-50/45 ring-2 ring-inset ring-red-200" : ""
                  }`}
                >
                  <div className="order-1 flex h-8 items-center justify-center">
                    {isReadOnlyAnchor ? (
                      <span
                        className="group/lock relative inline-flex h-7 w-7 shrink-0 items-center justify-center rounded-lg border border-slate-200 bg-slate-50 text-slate-400"
                        title="This anchor is filled automatically from the event details."
                      >
                        <svg
                          aria-hidden="true"
                          viewBox="0 0 16 16"
                          className="h-3.5 w-3.5"
                          fill="none"
                          stroke="currentColor"
                          strokeWidth="1.5"
                          strokeLinecap="round"
                          strokeLinejoin="round"
                        >
                          <rect x="3.5" y="7" width="9" height="6" rx="1.5" />
                          <path d="M5.5 7V5.75a2.5 2.5 0 1 1 5 0V7" />
                        </svg>
                        <span className="pointer-events-none absolute left-[calc(100%+0.5rem)] top-1/2 z-20 hidden w-56 -translate-y-1/2 rounded-xl border border-slate-200 bg-white p-3 text-left text-[11px] font-normal leading-4 text-slate-600 shadow-lg group-hover/lock:block">
                          This anchor is locked because it is filled automatically from the event details.
                        </span>
                      </span>
                    ) : (
                      <span className="h-2 w-2 rounded-full bg-slate-300" aria-hidden="true" />
                    )}
                  </div>

                  <div className="order-2 flex h-[40px] min-w-0 items-center rounded-lg border border-slate-200 bg-white px-2.5 font-mono text-[13px] text-slate-700">
                    <span className="text-slate-400">[</span>
                    <input
                      data-inspector-anchor-key="true"
                      className={`min-w-0 flex-1 border-0 bg-transparent px-1 text-center uppercase focus:outline-none focus:ring-0 ${
                        isReadOnlyAnchor ? "cursor-not-allowed text-slate-700" : ""
                      }`}
                      placeholder="KEY"
                      value={anchor.key}
                      readOnly={isReadOnlyAnchor}
                      onChange={(e) => {
                        const previousNormalizedKey = normalizeAnchorKey(anchor.key);
                        if (missingFieldHighlights.anchorKeys.includes(previousNormalizedKey)) {
                          setMissingFieldHighlights((current) => ({
                            ...current,
                            anchorKeys: current.anchorKeys.filter((key) => key !== previousNormalizedKey),
                          }));
                        }
                        setAnchors((current) =>
                          current.map((entry) =>
                            entry.id === anchor.id ? { ...entry, key: e.target.value.toUpperCase() } : entry
                          )
                        );
                      }}
                    />
                    <span className="text-slate-400">]</span>
                  </div>

                  <input
                    data-inspector-anchor-value="true"
                    className={`order-4 col-span-3 h-[40px] w-full rounded-lg border border-slate-200 bg-white px-3 text-[14px] text-slate-700 placeholder:text-slate-400 focus:border-[#6f9fd1] focus:outline-none focus:ring-2 focus:ring-[#6f9fd1]/20 min-[620px]:order-3 min-[620px]:col-span-1 ${
                      isReadOnlyAnchor ? "cursor-not-allowed" : ""
                    }`}
                    placeholder="Value"
                    value={displayedAnchorValue}
                    readOnly={isReadOnlyAnchor}
                    onChange={(e) => {
                      const normalizedAnchorKey = normalizeAnchorKey(anchor.key);
                      setAnchors((current) =>
                        current.map((entry) =>
                          entry.id === anchor.id
                            ? { ...entry, value: e.target.value, lastUpdatedAt: new Date().toISOString() }
                            : entry
                        )
                      );
                      if (missingFieldHighlights.anchorKeys.includes(normalizedAnchorKey)) {
                        setMissingFieldHighlights((current) => ({
                          ...current,
                          anchorKeys: current.anchorKeys.filter((key) => key !== normalizedAnchorKey),
                        }));
                      }
                    }}
                  />

                  <div className="order-5 col-span-3 grid grid-cols-2 gap-2 min-[620px]:col-span-2 min-[620px]:grid-cols-2">
                    <button
                      type="button"
                      data-inspector-anchor-insert="true"
                      onMouseDown={(e) => e.preventDefault()}
                      className="inline-flex h-[40px] items-center justify-center rounded-lg border border-slate-200 bg-white px-2 text-[13px] font-semibold text-slate-700 transition hover:border-slate-300 hover:bg-slate-50 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#6f9fd1]/25 min-[620px]:h-[38px]"
                      onClick={() => onInsertAnchor(anchor.key)}
                    >
                      Insert
                    </button>
                    <button
                      type="button"
                      data-inspector-anchor-delete="true"
                      disabled={isReadOnlyAnchor}
                      onClick={() => setAnchors((current) => current.filter((entry) => entry.id !== anchor.id))}
                      className={`inline-flex h-[40px] items-center justify-center rounded-lg border px-2 text-[13px] font-semibold transition focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-red-200 min-[620px]:h-[38px] ${
                        isReadOnlyAnchor
                          ? "cursor-not-allowed border-slate-200 bg-slate-50 text-slate-300"
                          : "border-red-200 bg-white text-red-600 hover:border-red-300 hover:bg-red-50 hover:text-red-700"
                      }`}
                    >
                      Delete
                    </button>
                  </div>

                  <div className="order-3 flex justify-end min-[620px]:order-6">
                    <button
                      type="button"
                      data-inspector-anchor-important="true"
                      aria-label="Mark as important"
                      title="Mark as important"
                      onClick={() =>
                        setAnchors((current) =>
                          current.map((entry) =>
                            entry.id === anchor.id ? { ...entry, isImportant: !entry.isImportant } : entry
                          )
                        )
                      }
                      className={`peer inline-flex h-8 w-8 items-center justify-center rounded-lg border transition focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-red-200 ${
                        anchor.isImportant
                          ? "border-red-200 bg-red-50/70 text-red-600"
                          : "border-slate-200 bg-white text-slate-400 hover:border-slate-300 hover:text-slate-600"
                      }`}
                    >
                      <svg
                        aria-hidden="true"
                        viewBox="0 0 16 16"
                        className="h-3.5 w-3.5"
                        fill="none"
                        stroke="currentColor"
                        strokeWidth="1.6"
                        strokeLinecap="round"
                        strokeLinejoin="round"
                      >
                        <path d="M8 3.25v5.4" />
                        <circle cx="8" cy="12.15" r="0.95" fill="currentColor" stroke="none" />
                      </svg>
                    </button>
                  </div>
                </div>
              );
            })}
          </div>
        </section>
      );
    }

    const getDisplayValueForAnchor = (anchor: BuilderAnchor) => {
      const resolvedAnchor = resolvedAnchors.find((entry) => entry.id === anchor.id);
      const isReadOnlyAnchor = Boolean(anchor.locked || isCoreEventAnchorKey(anchor.key));
      if (!isReadOnlyAnchor) return anchor.value;
      return getAnchorDisplayValue(
        anchor.key,
        getDerivedAnchorValue(planType, effectivePlanName, anchorDate, anchor.key, guidedForm, {
          eventNameValue: eventName.trim(),
          eventDateValue: noEventDate ? "" : anchorDate,
          eventTimeValue: eventTime.trim(),
        }) ??
          resolvedAnchor?.value ??
          anchor.value
      );
    };
    const coreAnchorCount = anchors.filter((anchor) => anchor.locked || isCoreEventAnchorKey(anchor.key)).length;
    const customAnchorCount = anchors.length - coreAnchorCount;
    const requiredAnchorNeedsValues = anchors.filter((anchor) => anchor.isImportant && !getDisplayValueForAnchor(anchor).trim()).length;
    const anchorSummaryParts = [
      coreAnchorCount ? `${coreAnchorCount} core ${coreAnchorCount === 1 ? "field" : "fields"}` : null,
      customAnchorCount ? `${customAnchorCount} custom ${customAnchorCount === 1 ? "field" : "fields"}` : null,
      requiredAnchorNeedsValues
        ? `${requiredAnchorNeedsValues} ${requiredAnchorNeedsValues === 1 ? "field needs a value" : "fields need values"}`
        : null,
    ].filter(Boolean);
    const anchorSummary =
      anchorSummaryParts.length > 0
        ? anchorSummaryParts.join(" · ")
        : `${anchors.length} ${anchors.length === 1 ? "field" : "fields"}`;

    function addBaseAnchorField() {
      const nextAnchor = createEmptyAnchor();
      setAreAnchorsHidden(false);
      setAnchors((current) => [...current, nextAnchor]);
      requestAnimationFrame(() => {
        requestAnimationFrame(() => {
          const target = document.querySelector<HTMLInputElement>(`[data-plan-anchor-key="${nextAnchor.id}"]`);
          target?.scrollIntoView({ behavior: "smooth", block: "center" });
          target?.focus({ preventScroll: true });
        });
      });
    }

    return (
      <section className={plansSurfaceClass} data-plan-anchor-surface="true">
        <div className={plansSectionHeaderClass}>
          <div className="flex flex-col gap-[12px] min-[720px]:flex-row min-[720px]:items-start min-[720px]:justify-between">
            <div className="min-w-0">
              <h2 className={plansPanelHeadingClass}>Anchor Fields</h2>
              <p className="mt-[3px] text-[14px] leading-[1.4] text-slate-600">{anchorSummary}</p>
            </div>
            <div className="grid grid-cols-2 gap-[8px] min-[480px]:flex min-[480px]:justify-end">
              <button type="button" onClick={addBaseAnchorField} className={`${plansSecondaryButtonClass} min-w-0 px-3`}>
                Add Anchor Field
              </button>
              <button
                type="button"
                aria-expanded={!areAnchorsHidden}
                aria-controls={anchorFieldsRegionId}
                onClick={() => setAreAnchorsHidden((current) => !current)}
                className={`${plansSecondaryButtonClass} min-w-0 px-3`}
              >
                {areAnchorsHidden ? "Show fields" : "Hide fields"}
              </button>
            </div>
          </div>
        </div>

        {!areAnchorsHidden ? (
          <div
            id={anchorFieldsRegionId}
            aria-label="Anchor Fields"
            className="grid transition-[grid-template-rows,opacity] duration-150 ease-[cubic-bezier(0.16,1,0.3,1)] motion-reduce:transition-none"
          >
            <div className="divide-y divide-slate-200/80 overflow-hidden bg-white">
              {anchors.map((anchor) => {
                const isCoreEventAnchor = isCoreEventAnchorKey(anchor.key);
                const isReadOnlyAnchor = Boolean(anchor.locked || isCoreEventAnchor);
                const isMissingFieldHighlighted = missingFieldHighlights.anchorKeys.includes(normalizeAnchorKey(anchor.key));
                const displayedAnchorValue = getDisplayValueForAnchor(anchor);
                const normalizedAnchorKey = normalizeAnchorKey(anchor.key);

                return (
                  <div
                    key={anchor.id}
                    className={`relative grid grid-cols-[28px_minmax(0,1fr)_32px] items-center gap-x-[8px] gap-y-[8px] px-[14px] py-[10px] min-[640px]:grid-cols-[28px_minmax(150px,0.8fr)_minmax(180px,1.2fr)_76px_32px] min-[640px]:gap-x-[10px] min-[640px]:gap-y-0 min-[640px]:px-[14px] min-[640px]:py-[10px] min-[900px]:min-h-[60px] min-[900px]:grid-cols-[28px_minmax(190px,0.8fr)_minmax(260px,1.2fr)_76px_32px] min-[900px]:gap-x-[12px] min-[900px]:px-[16px] ${
                      isMissingFieldHighlighted ? "bg-red-50/40 ring-2 ring-inset ring-red-200" : ""
                    }`}
                  >
                    <div className="flex h-[32px] items-center justify-center">
                      {isReadOnlyAnchor ? (
                        <span
                          className="group/lock relative inline-flex h-7 w-7 shrink-0 items-center justify-center rounded-[8px] border border-slate-200 bg-slate-50 text-slate-400"
                          title="This anchor is filled automatically from the event details."
                        >
                          <svg
                            aria-hidden="true"
                            viewBox="0 0 16 16"
                            className="h-3.5 w-3.5"
                            fill="none"
                            stroke="currentColor"
                            strokeWidth="1.5"
                            strokeLinecap="round"
                            strokeLinejoin="round"
                          >
                            <rect x="3.5" y="7" width="9" height="6" rx="1.5" />
                            <path d="M5.5 7V5.75a2.5 2.5 0 1 1 5 0V7" />
                          </svg>
                        </span>
                      ) : (
                        <span className="h-2 w-2 rounded-full bg-slate-300" aria-hidden="true" />
                      )}
                    </div>

                    <div className="min-w-0">
                      {isReadOnlyAnchor ? (
                        <div className="flex min-w-0 flex-col gap-[4px] min-[640px]:flex-row min-[640px]:items-center">
                          <div
                            className="flex h-[38px] w-full min-w-0 items-center rounded-[9px] border border-slate-200 bg-slate-50 px-3 font-mono text-[13px] font-semibold uppercase text-slate-700 min-[640px]:flex-1"
                            title={formatAnchorTokenDisplay(anchor.key)}
                            aria-label={anchor.isImportant ? `${formatAnchorTokenDisplay(anchor.key)} required` : formatAnchorTokenDisplay(anchor.key)}
                          >
                            <span className="truncate">{formatAnchorTokenDisplay(anchor.key)}</span>
                          </div>
                          {anchor.isImportant ? (
                            <span className="inline-flex h-[20px] shrink-0 items-center self-start rounded-[6px] border border-amber-200 bg-amber-50 px-2 text-[11px] font-semibold text-amber-800 min-[640px]:self-center">
                              Required
                            </span>
                          ) : null}
                        </div>
                      ) : (
                        <div className="flex min-w-0 flex-col gap-[4px] min-[640px]:flex-row min-[640px]:items-center">
                          <div className="flex h-[38px] w-full min-w-0 items-center rounded-[9px] border border-slate-200 bg-white px-3 font-mono text-[13px] text-slate-700 shadow-sm min-[640px]:flex-1">
                            <span className="text-gray-500">[</span>
                            <input
                              data-plan-anchor-key={anchor.id}
                              aria-label={anchor.isImportant ? "Anchor key, required" : "Anchor key"}
                              className="min-w-0 flex-1 border-0 bg-transparent px-1 text-center uppercase focus:outline-none focus:ring-0"
                              placeholder="KEY"
                              value={anchor.key}
                              onChange={(e) => {
                                if (missingFieldHighlights.anchorKeys.includes(normalizedAnchorKey)) {
                                  setMissingFieldHighlights((current) => ({
                                    ...current,
                                    anchorKeys: current.anchorKeys.filter((key) => key !== normalizedAnchorKey),
                                  }));
                                }
                                setAnchors((current) =>
                                  current.map((entry) =>
                                    entry.id === anchor.id ? { ...entry, key: e.target.value.toUpperCase() } : entry
                                  )
                                );
                              }}
                            />
                            <span className="text-gray-500">]</span>
                          </div>
                          {anchor.isImportant ? (
                            <span className="inline-flex h-[20px] shrink-0 items-center self-start rounded-[6px] border border-amber-200 bg-amber-50 px-2 text-[11px] font-semibold text-amber-800 min-[640px]:self-center">
                              Required
                            </span>
                          ) : null}
                        </div>
                      )}
                    </div>

                    <div className="col-span-3 grid grid-cols-[minmax(0,1fr)_82px] gap-[8px] min-[640px]:contents">
                      <input
                        data-plan-anchor-value={anchor.id}
                        aria-label={`Value for ${formatAnchorTokenDisplay(anchor.key)}`}
                        className={`h-[38px] w-full rounded-[9px] border border-slate-200 bg-white px-3 text-[14px] text-slate-700 shadow-sm placeholder:text-slate-400 focus:border-[#6f9fd1] focus:outline-none focus:ring-2 focus:ring-[#6f9fd1]/20 ${
                          isReadOnlyAnchor ? "cursor-not-allowed" : ""
                        }`}
                        placeholder="Value"
                        value={displayedAnchorValue}
                        readOnly={isReadOnlyAnchor}
                        onChange={(e) => {
                          setAnchors((current) =>
                            current.map((entry) =>
                              entry.id === anchor.id ? { ...entry, value: e.target.value, lastUpdatedAt: new Date().toISOString() } : entry
                            )
                          );
                          if (missingFieldHighlights.anchorKeys.includes(normalizedAnchorKey)) {
                            setMissingFieldHighlights((current) => ({
                              ...current,
                              anchorKeys: current.anchorKeys.filter((key) => key !== normalizedAnchorKey),
                            }));
                          }
                        }}
                      />

                      <button
                        type="button"
                        aria-label={`Insert ${formatAnchorTokenDisplay(anchor.key)}`}
                        onMouseDown={(e) => e.preventDefault()}
                        className="inline-flex h-[38px] w-[82px] items-center justify-center rounded-[9px] border border-slate-200 bg-white px-2 text-[13px] font-semibold text-slate-700 shadow-sm transition hover:border-slate-300 hover:bg-slate-50 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#6f9fd1]/25 min-[640px]:w-[76px]"
                        onClick={() => onInsertAnchor(anchor.key)}
                      >
                        Insert
                      </button>
                    </div>

                    <button
                      ref={(node) => {
                        anchorMenuButtonRefs.current[anchor.id] = node;
                      }}
                      type="button"
                      aria-haspopup="menu"
                      aria-expanded={openAnchorMenuId === anchor.id}
                      aria-label={`More actions for ${formatAnchorTokenDisplay(anchor.key)}`}
                      onClick={(event) => openAnchorMenu(event, anchor.id)}
                      className="col-start-3 row-start-1 inline-flex h-[32px] w-[32px] items-center justify-center justify-self-end rounded-[8px] border border-slate-200 bg-white text-slate-500 transition hover:border-slate-300 hover:bg-slate-50 hover:text-slate-700 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#6f9fd1]/25 min-[640px]:col-start-auto min-[640px]:row-start-auto min-[640px]:justify-self-center"
                    >
                      <span className="text-[15px] leading-none">•••</span>
                    </button>

                    {hasMounted && openAnchorMenuId === anchor.id && anchorMenuPosition
                      ? createPortal(
                          <div
                            ref={anchorMenuRef}
                            role="menu"
                            className="plans-menu-enter fixed z-[170] min-w-[194px] rounded-[10px] border border-slate-200 bg-white p-[7px] shadow-[0_18px_46px_rgba(21,40,66,0.18)]"
                            style={{ top: anchorMenuPosition.top, left: anchorMenuPosition.left }}
                          >
                            <button
                              type="button"
                              role="menuitem"
                              onClick={() => {
                                setAnchors((current) =>
                                  current.map((entry) =>
                                    entry.id === anchor.id ? { ...entry, isImportant: !entry.isImportant } : entry
                                  )
                                );
                                closeAnchorMenu({ restoreFocus: true });
                              }}
                              className="block h-10 w-full rounded-[8px] px-3 text-left text-[14px] font-medium text-slate-700 transition hover:bg-slate-50 focus:outline-none focus:ring-2 focus:ring-[#6f9fd1]/20"
                            >
                              {anchor.isImportant ? "Mark as optional" : "Mark as important"}
                            </button>
                            {!isReadOnlyAnchor ? (
                              <button
                                type="button"
                                role="menuitem"
                                onClick={() => {
                                  setAnchors((current) => current.filter((entry) => entry.id !== anchor.id));
                                  closeAnchorMenu();
                                }}
                                className="block h-10 w-full rounded-[8px] px-3 text-left text-[14px] font-medium text-red-700 transition hover:bg-red-50 focus:outline-none focus:ring-2 focus:ring-red-200"
                              >
                                Delete Anchor Field
                              </button>
                            ) : null}
                          </div>,
                          document.body
                        )
                      : null}
                  </div>
                );
              })}
            </div>
          </div>
        ) : null}
      </section>
    );
  }

  function renderPlanActionsSection() {
    const canSaveAsTemplate = shouldRenderBuilderSourceBanner;
    const handleReviewAndExport = () => {
      void (async () => {
        if (await validatePreviewBeforeOpen()) return;
        setOpenPreviewDetail(null);
        setOpenPreviewRowMenuId(null);
        setExcludedPreviewItemIds([]);
        setShowNoPreviewItemsSelectedCallout(false);
        setIsBuilderPreviewOpen(true);
      })();
    };
    const handleSave = () => {
      void saveCurrentTemplate();
    };
    const handleCancel = () => {
      void requestCancelEditing();
    };
    const handleSaveAsTemplate = () => {
      if (canSaveAsTemplate) {
        openBuilderTemplateSaveDialog();
      }
    };
    const saveAsTemplateDisabledClass =
      "disabled:cursor-not-allowed disabled:border-slate-200 disabled:bg-slate-100 disabled:text-slate-400 disabled:shadow-none";

    return (
      <section className={`${plansSurfaceClass} px-[16px] py-[14px] sm:px-[20px]`} data-plan-action-bar="true">
        <div className="hidden items-center justify-between gap-[10px] min-[900px]:flex min-[900px]:flex-nowrap" data-plan-action-layout="desktop">
          <button
            type="button"
            onClick={handleCancel}
            className={`${plansSecondaryButtonClass} h-[40px] whitespace-nowrap border-red-200 text-red-600 hover:border-red-300 hover:bg-red-50`}
          >
            Cancel
          </button>
          <div className="flex min-w-0 items-center justify-end gap-[10px]">
            <button
              type="button"
              onClick={handleSave}
              className={`${plansSecondaryButtonClass} h-[40px] whitespace-nowrap`}
            >
              Save
            </button>
            <button
              type="button"
              onClick={handleSaveAsTemplate}
              disabled={!canSaveAsTemplate}
              aria-disabled={!canSaveAsTemplate}
              className={`${plansSecondaryButtonClass} h-[40px] min-w-[138px] whitespace-nowrap ${saveAsTemplateDisabledClass}`}
              title={canSaveAsTemplate ? "Save as Template" : "Add event details, actions, or Anchor Fields before saving as a template."}
            >
              Save as Template
            </button>
            {executionNotices.some((entry) => entry.notice.tone === "success") ? <ExportDoneBadge /> : null}
            <button
              type="button"
              onClick={handleReviewAndExport}
              className={`${plansPrimaryButtonClass} h-[40px] min-w-[150px] whitespace-nowrap`}
            >
              Review &amp; Export
            </button>
          </div>
        </div>
        <div className="grid gap-[10px] min-[900px]:hidden" data-plan-action-layout="narrow">
          <button
            type="button"
            onClick={handleReviewAndExport}
            className={`${plansPrimaryButtonClass} h-[40px] w-full whitespace-nowrap`}
          >
            Review &amp; Export
          </button>
          <div className="grid grid-cols-2 gap-[8px]">
            <button type="button" onClick={handleSave} className={`${plansSecondaryButtonClass} h-[40px] w-full whitespace-nowrap px-3`}>
              Save
            </button>
            <button
              type="button"
              onClick={handleSaveAsTemplate}
              disabled={!canSaveAsTemplate}
              aria-disabled={!canSaveAsTemplate}
              className={`${plansSecondaryButtonClass} h-[40px] w-full min-w-0 whitespace-nowrap px-3 ${saveAsTemplateDisabledClass}`}
              title={canSaveAsTemplate ? "Save as Template" : "Add event details, actions, or Anchor Fields before saving as a template."}
            >
              Save as Template
            </button>
          </div>
          {executionNotices.some((entry) => entry.notice.tone === "success") ? <ExportDoneBadge /> : null}
          <button
            type="button"
            onClick={handleCancel}
            className={`${plansSecondaryButtonClass} h-[40px] w-full whitespace-nowrap border-red-200 text-red-600 hover:border-red-300 hover:bg-red-50`}
          >
            Cancel
          </button>
        </div>
      </section>
    );
  }

  async function confirmExport() {
    return showConfirmModal({
      title: "Confirm export",
      message: "Are you sure you want to export?",
      confirmLabel: "Export",
      cancelLabel: "Cancel",
    });
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
    closeAllRowEditors();
    if (affectedIds[0]) {
      applyRowEditorOpen(affectedIds[0], "meeting");
    }
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

  async function warnIfMissingReminderTimes(options?: { usePopup?: boolean; itemIds?: string[] }) {
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
    if (options?.usePopup) {
      await showAlertModal({
        title: "Missing reminder times",
        message: "One or more reminder rows are missing a reminder time.",
        confirmLabel: "OK",
      });
    }
    return true;
  }

  function getPastScheduledItemsMessage(itemIds?: string[]) {
    const previewItems = getLatestPreviewPlan().items;
    const scopedItems = itemIds ? previewItems.filter((item) => itemIds.includes(item.id)) : previewItems;
    const pastItems = scopedItems.filter(isPreviewItemScheduledInPast);

    if (pastItems.length === 0) return null;
    return "One or more selected items are scheduled before now. Exporting them may create outdated reminders, meetings, or email drafts.";
  }

  async function warnIfPastScheduledItems(options?: { usePopup?: boolean; itemIds?: string[] }) {
    const message = getPastScheduledItemsMessage(options?.itemIds);
    if (!message) return false;
    if (options?.usePopup) {
      const confirmed = await showConfirmModal({
        title: "Some selected items are in the past",
        message,
        content: renderPastScheduledItemsContent(options?.itemIds),
        confirmLabel: "Continue export",
        cancelLabel: "Go back",
        severity: "warning",
      });
      return !confirmed;
    }
    return true;
  }

  async function validateMeetingRowsForExport(itemIds?: string[], options?: { usePopup?: boolean }) {
    const message = getMeetingValidationMessage(itemIds);
    if (!message) return false;
    if (options?.usePopup) {
      const [intro, ...rest] = message.split("\n\n");
      await showAlertModal({
        title: "Meeting information required",
        message: intro,
        items: rest.join("\n\n").split("\n").filter(Boolean),
        confirmLabel: "OK",
      });
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

  async function validateEmailRowsForExport(itemIds?: string[], options?: { usePopup?: boolean }) {
    const message = getEmailValidationMessage(itemIds);
    if (!message) return false;
    if (options?.usePopup) {
      const [intro, ...rest] = message.split("\n\n");
      await showAlertModal({
        title: "Email information required",
        message: intro,
        items: rest.join("\n\n").split("\n").filter(Boolean),
        confirmLabel: "OK",
      });
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
    const draft = {
      to: Array.isArray(item.emailDraft?.to) ? item.emailDraft.to.filter((entry): entry is string => typeof entry === "string") : [],
      cc: Array.isArray(item.emailDraft?.cc) ? item.emailDraft.cc.filter((entry): entry is string => typeof entry === "string") : [],
      bcc: Array.isArray(item.emailDraft?.bcc) ? item.emailDraft.bcc.filter((entry): entry is string => typeof entry === "string") : [],
      subject: typeof item.emailDraft?.subject === "string" ? item.emailDraft.subject : "",
      body: typeof item.emailDraft?.body === "string" ? item.emailDraft.body : "",
    };
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
    return Array.isArray(item.meetingDraft?.attendees)
      ? item.meetingDraft.attendees.filter((entry): entry is string => typeof entry === "string")
      : [];
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
          attendees: Array.isArray(item.meetingDraft?.attendees)
            ? item.meetingDraft.attendees.filter((entry): entry is string => typeof entry === "string")
            : [],
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
    const outlookTimeZone = normalizeOutlookTimeZone(item.timeZone);
    const providerTimeZone = provider === "gmail" ? getIanaTimeZoneForProvider(outlookTimeZone) : outlookTimeZone;
    if (provider === "gmail") {
      const result = await createGoogleCalendarEvent({
        subject: item.customTitle ?? item.title,
        bodyText: item.body?.trim() || "",
        startISO: timing.startISO,
        endISO: timing.endISO,
        timeZone: providerTimeZone,
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
      timeZone: providerTimeZone,
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
      if (await validateEmailRowsForExport([itemId], { usePopup: true })) return;
      if (await warnIfMissingReminderTimes({ usePopup: true, itemIds: [itemId] })) return;
      if (await warnIfPastScheduledItems({ usePopup: true, itemIds: [itemId] })) return;
      if (!(await confirmImportantDynamicFieldsForExport())) return;
      if (!(await confirmExport())) return;
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
          persistSelectedTemplateDynamicFieldTimestamp(new Date().toISOString());
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
            await showAlertModal({
              title: "Action required",
              message: userFixableMessage,
              confirmLabel: "OK",
            });
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
      if (await validateMeetingRowsForExport([itemId], { usePopup: true })) return;
      if (await warnIfMissingReminderTimes({ usePopup: true, itemIds: [itemId] })) return;
      if (await warnIfPastScheduledItems({ usePopup: true, itemIds: [itemId] })) return;
      if (!(await confirmImportantDynamicFieldsForExport())) return;
      if (!(await confirmExport())) return;
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
          persistSelectedTemplateDynamicFieldTimestamp(new Date().toISOString());
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
    closeAllRowEditors();
    setForcedOpenMeetingEditorRowIds([]);
    setMeetingValidationErrors({});
    setEditingOffsetRowId(null);
    setOffsetDrafts({});
    setFocusedTimeInputRowId(null);
    setOpenTimeZoneRowId(null);
    setTimeZoneSearch("");
    setTimeInputDrafts({});
    setEmailFieldVisibility({});
    setIsBuilderPreviewOpen(false);
    setOpenPreviewDetail(null);
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
    const nextAnchorDate =
      templateMode === "template" && template.baseType === "press_release"
        ? nextGuidedForm.releaseDate || ""
        : templateMode === "template" && template.baseType === "earnings"
          ? nextGuidedForm.earningsDate || ""
          : templateMode === "template" && template.baseType === "conference"
            ? nextGuidedForm.conferenceDate || ""
            : "";
    return {
      builderMode: templateMode,
      selectedTemplateId: template.id,
      planType: template.baseType,
      templateName: template.name,
      eventName: "",
      anchorDate: nextAnchorDate,
      hasExplicitEventDate: Boolean(nextAnchorDate),
      eventTime: "",
      eventTimeZone: getDefaultOutlookTimeZone(),
      noEventDate: Boolean(template.noEventDate),
      weekendRule: template.weekendRule,
      rows: cloneTemplateRows(template.items),
      anchors:
        templateMode === "template"
          ? buildAnchorStateForType(
              template.baseType,
              template.anchors.map(buildFreshBuilderAnchorFromTemplate)
            )
          : buildAnchorStateWithCoreEventFields(
              template.anchors.map(buildFreshBuilderAnchorFromTemplate)
            ),
      guidedForm: nextGuidedForm,
      lastDynamicFieldsExportAt: template.lastDynamicFieldsExportAt ?? null,
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
      hasExplicitEventDate,
      eventTime,
      eventTimeZone,
      noEventDate,
      weekendRule,
      rows: cloneTemplateRows(rows),
      anchors: cloneAnchors(anchors),
      guidedForm: { ...guidedForm },
      lastDynamicFieldsExportAt,
    };
  }

  function applyBuilderSnapshot(snapshot: BuilderStateSnapshot) {
    setIsBuilderSectionVisible(true);
    setBuilderMode(snapshot.builderMode);
    setSelectedTemplateId(snapshot.selectedTemplateId);
    setPlanType(snapshot.planType);
    setTemplateName(snapshot.templateName);
    setEventName(snapshot.eventName);
    setAnchorDate(snapshot.anchorDate);
    setHasExplicitEventDate(snapshot.hasExplicitEventDate);
    setEventTime(snapshot.eventTime);
    const nextEventTimeZone = normalizeOutlookTimeZone(snapshot.eventTimeZone);
    setEventTimeZone(nextEventTimeZone);
    setNoEventDate(snapshot.noEventDate);
    setWeekendRule(snapshot.weekendRule);
    setRows(cloneTemplateRows(snapshot.rows).map((row) => ({ ...row, timeZone: row.timeZone || nextEventTimeZone })));
    setAnchors(cloneAnchors(snapshot.anchors));
    setAreAnchorsHidden(true);
    setGuidedForm({ ...snapshot.guidedForm });
    setLastDynamicFieldsExportAt(snapshot.lastDynamicFieldsExportAt);
    clearTransientEditingState();
  }

  function applyTemplateRecord(template: SavedPlanTemplate, options?: { skipBuilderScroll?: boolean; immediateReveal?: boolean }) {
    isBuilderDraftPersistencePausedRef.current = false;
    setTemplateActionMessage("");
    const shouldScrollBuilder = !(options?.skipBuilderScroll ?? false);
    shouldFocusEventNameInputRef.current = true;
    shouldScrollBuilderIntoViewRef.current = shouldScrollBuilder;
    if (shouldScrollBuilder) {
      setBuilderScrollRequestNonce((current) => current + 1);
    }
    setIsBuilderVisualRevealDeferred(!(options?.immediateReveal ?? false));
    setIsBuilderEntryRevealImmediate(Boolean(options?.immediateReveal));
    setHasActivePlanSession(true);
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
      if (normalizedKey === normalizeAnchorKey("Event Time")) {
        return { ...anchor, value: draft.eventTime || "" };
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
      hasExplicitEventDate: Boolean(draft.anchorDate),
      eventTime: draft.eventTime || "",
      eventTimeZone: normalizeOutlookTimeZone(draft.timezone),
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
        timeZone: draft.timezone || "",
        emailDraft:
          row.rowType === "email"
            ? normalizeEmailDraft(row.emailDraft)
            : { to: [], cc: [], bcc: [], subject: "", body: "" },
        durationDraft: row.durationDraft ? { ...row.durationDraft } : undefined,
        meetingDraft: row.meetingDraft ? { ...normalizeMeetingDraft(row.meetingDraft) } : undefined,
      })),
      anchors: buildAiAnchorState(draft),
      guidedForm: createEmptyGuidedForm(),
      lastDynamicFieldsExportAt: null,
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
        timeZone: draft.timezone || "",
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
    const hasMeaningfulTiming = Boolean(anchorDate.trim() || eventTime.trim() || noEventDate || hasExplicitEventDate);
    const hasMeaningfulAnchorValue = anchors.some((anchor) => anchor.value.trim());
    const hasMeaningfulGuidedForm = Object.values(guidedForm).some((value) =>
      typeof value === "boolean" ? value : Boolean(String(value).trim())
    );
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

    return (
      hasMeaningfulName ||
      hasMeaningfulTiming ||
      hasMeaningfulAnchorValue ||
      hasMeaningfulGuidedForm ||
      hasMeaningfulRowContent ||
      hasMultipleRows
    );
  }

  function applyAiDraftToBuilder() {
    if (!aiChatDraft) return;
    isBuilderDraftPersistencePausedRef.current = false;
    setHasActivePlanSession(true);
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
      setAiTemplateSaveMessage("Enter a template name.");
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
      setBuilderTemplateSaveMessage("Enter a template name.");
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
        .map((anchor) => ({
          key: anchor.key.trim(),
          value: "",
          isImportant: Boolean(anchor.isImportant),
          lastUpdatedAt: null,
        }))
        .filter((anchor) => anchor.key),
      items: cloneTemplateRows(rows),
      lastDynamicFieldsExportAt,
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

  function cancelPendingBuilderEntryTransition() {
    if (scheduledBuilderEntryRef.current !== null) {
      window.clearTimeout(scheduledBuilderEntryRef.current);
      scheduledBuilderEntryRef.current = null;
    }
    setIsBuilderEntryRunwayVisible(false);
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
        .map((anchor) => ({
          key: anchor.key.trim(),
          value: "",
          isImportant: Boolean(anchor.isImportant),
          lastUpdatedAt: null,
        }))
        .filter((anchor) => anchor.key),
      items: cloneTemplateRows(rows),
      lastDynamicFieldsExportAt,
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
      hasExplicitEventDate,
      eventTime,
      eventTimeZone,
      noEventDate: Boolean(nextTemplate.noEventDate),
      weekendRule,
      rows: cloneTemplateRows(rows),
      anchors:
        nextTemplateMode === "template"
          ? cloneAnchors(
              buildAnchorStateForType(
                nextTemplate.baseType,
                nextTemplate.anchors.map(buildFreshBuilderAnchorFromTemplate)
              )
            )
          : cloneAnchors(
              nextTemplate.anchors.map(buildFreshBuilderAnchorFromTemplate)
            ),
      guidedForm: { ...guidedForm },
      lastDynamicFieldsExportAt,
    };
    persistTemplateStateImmediately(nextTemplates, nextTemplate.id);
    applyBuilderSnapshot(snapshot);
    setLastTemplateSnapshot(snapshot);
  }

  async function promptForNewTemplateName(defaultName: string) {
    const nextName = await showPromptModal({
      title: "Save as template",
      message: "Save this workflow structure so it can be reused for another event.",
      defaultValue: defaultName,
      inputLabel: "Template name",
      helperText: "Event details and run-specific Anchor Field values are not saved.",
      placeholder: "Template name",
      confirmLabel: "Save template",
      cancelLabel: "Cancel",
      severity: "information",
    });
    if (!nextName) return null;
    const trimmedName = nextName.trim();
    if (!trimmedName) {
      setTemplateActionMessage("Enter a template name.");
      return null;
    }
    if (hasDuplicateTemplateName(trimmedName)) {
      setTemplateActionMessage("That name is already taken. Please choose another name.");
      return null;
    }
    return trimmedName;
  }

  async function onSelectSavedTemplate(templateId: string) {
    const template = savedTemplates.find((entry) => entry.id === templateId);
    if (!template) return;
    cancelPendingBuilderEntryTransition();
    applyTemplateRecord(template, { immediateReveal: true });
  }

  async function duplicateSavedTemplate(templateId: string) {
    const template = savedTemplates.find((entry) => entry.id === templateId);
    if (!template) return;

    const confirmed = await showConfirmModal({
      title: "Copy template",
      message: `Are you sure you want to copy "${template.name}"?`,
      confirmLabel: "Copy",
      cancelLabel: "Cancel",
    });
    if (!confirmed) return;

    const nextName = await promptForNewTemplateName(`${template.name} Copy`);
    if (!nextName) return;

    const duplicatedTemplate: SavedPlanTemplate = {
      ...template,
      id: makeId("template"),
      name: nextName,
      anchors: template.anchors.map((anchor) => ({ ...anchor })),
      items: cloneTemplateRows(template.items),
    };
    setTemplateActionMessage("");
    setSavedTemplates((current) => [...current, duplicatedTemplate]);
    applyTemplate(duplicatedTemplate.id);
  }

  async function saveCurrentTemplate() {
    setTemplateActionMessage("");

    const promptDefaultName =
      templateName.trim() || selectedEditableTemplate?.name || getSeedTemplateName(planType) || "Untitled Template";

    const promptForNewTemplate = async () => {
      const nextName = await promptForNewTemplateName(promptDefaultName);
      if (!nextName) return;
      const nextTemplate = buildTemplateToSave(makeId("template"), nextName);
      saveTemplateEntry(nextTemplate);
    };

    const chooseSave = await showConfirmModal({
      title: "Save template",
      message: "Choose how you want to save this template.",
      confirmLabel: "Save",
      cancelLabel: "Save as New",
    });

    if (chooseSave) {
      if (!selectedEditableTemplate) {
        await promptForNewTemplate();
        return;
      }

      const nextName = templateName.trim() || selectedEditableTemplate.name;
      if (hasDuplicateTemplateName(nextName, { excludeTemplateId: selectedEditableTemplate.id })) {
        setTemplateActionMessage("That name is already taken. Please choose another name.");
        return;
      }

      const confirmedOverwrite = await showConfirmModal({
        title: "Overwrite template",
        message: "Are you sure you want to overwrite the existing template?",
        confirmLabel: "Overwrite",
        cancelLabel: "Cancel",
      });
      if (!confirmedOverwrite) return;

      const nextTemplate = buildTemplateToSave(selectedEditableTemplate.id, nextName);
      saveTemplateEntry(nextTemplate);
      return;
    }

    await promptForNewTemplate();
  }

  function resetBuilderToNewPlan(
    nextTemplateName = "",
    options?: {
      eventName?: string;
      anchorDate?: string;
      eventTime?: string;
      eventTimeZone?: string;
      noEventDate?: boolean;
      weekendRule?: WeekendRule;
      revealBuilder?: boolean;
      includeStarterRows?: boolean;
    }
  ) {
    if (builderMode !== "new") {
      setLastTemplateSnapshot(buildCurrentTemplateSnapshot());
      setLastBuilderSourceProvenance(builderSourceProvenance);
    }
    const nextEventName = options?.eventName?.trim() ?? "";
    const nextNoEventDate = options?.noEventDate ?? false;
    const nextAnchorDate = nextNoEventDate ? "" : options?.anchorDate ?? "";
    const nextEventTime = options?.eventTime ?? "";
    const nextEventTimeZone = normalizeOutlookTimeZone(options?.eventTimeZone ?? getDefaultOutlookTimeZone());
    const nextWeekendRule = options?.weekendRule ?? "none";
    setIsBuilderSectionVisible(options?.revealBuilder ?? true);
    setBuilderSourceProvenance(null);
    setBuilderMode("new");
    setSelectedTemplateId(null);
    setTemplateName(nextTemplateName);
    setEventName(nextEventName);
    setAnchorDate(nextAnchorDate);
    setHasExplicitEventDate(Boolean(nextAnchorDate));
    setEventTime(nextEventTime);
    setEventTimeZone(nextEventTimeZone);
    setNoEventDate(nextNoEventDate);
    setWeekendRule(nextWeekendRule);
    setRows(
      options?.includeStarterRows
        ? [createBuilderRowForKind("reminder", nextEventTimeZone), createBuilderRowForKind("meeting", nextEventTimeZone), createBuilderRowForKind("email", nextEventTimeZone)]
        : [],
    );
    setAnchors(createGenericPresetAnchors());
    setAreAnchorsHidden(true);
    setGuidedForm(createEmptyGuidedForm());
    setLastDynamicFieldsExportAt(null);
    clearPersistedBuilderDraft();
    clearTransientEditingState();
  }

  async function startNewPlan(options?: { keepViewAtTop?: boolean }) {
    cancelPendingBuilderEntryTransition();
    isBuilderDraftPersistencePausedRef.current = false;
    setTemplateActionMessage("");
    shouldFocusEventNameInputRef.current = !options?.keepViewAtTop;
    shouldScrollBuilderIntoViewRef.current = !options?.keepViewAtTop;
    if (!options?.keepViewAtTop) {
      setBuilderScrollRequestNonce((current) => current + 1);
    }
    setIsBuilderVisualRevealDeferred(true);
    setIsNewPlanSetupPending(false);
    setPlanSetupDialogMode("new");
    setPlanSetupTemplateId(null);
    setHasActivePlanSession(true);
    resetBuilderToNewPlan("", { revealBuilder: true, includeStarterRows: true });
    setNewPlanDialogMessage(null);
    setNewPlanDraft({
      eventName: "",
      anchorDate: "",
      eventTime: "",
      noEventDate: false,
      weekendRule: "none",
    });
    setShowNewPlanDialog(false);
    if (options?.keepViewAtTop) {
      window.requestAnimationFrame(() => {
        window.scrollTo({ top: 0, behavior: "auto" });
      });
    }
  }

  function confirmStartNewPlan() {
    const trimmedEventName = newPlanDraft.eventName.trim();

    setTemplateActionMessage("");
    setNewPlanDialogMessage(null);
    setShowNewPlanDialog(false);
    if (planSetupDialogMode === "template") {
      setIsNewPlanSetupPending(false);
      setEventName(trimmedEventName);
      setAnchorDate(newPlanDraft.noEventDate ? "" : newPlanDraft.anchorDate);
      setHasExplicitEventDate(Boolean(!newPlanDraft.noEventDate && newPlanDraft.anchorDate.trim()));
      setEventTime(newPlanDraft.eventTime);
      setNoEventDate(newPlanDraft.noEventDate);
      setWeekendRule(newPlanDraft.weekendRule);
      setPlanSetupTemplateId(null);
      return;
    }

    setIsNewPlanSetupPending(false);
    setHasActivePlanSession(true);
    setIsBuilderVisualRevealDeferred(true);
    resetBuilderToNewPlan(trimmedEventName, {
      eventName: trimmedEventName,
      anchorDate: newPlanDraft.anchorDate,
      eventTime: newPlanDraft.eventTime,
      noEventDate: newPlanDraft.noEventDate,
      weekendRule: newPlanDraft.weekendRule,
      revealBuilder: true,
      includeStarterRows: true,
    });
  }

  function getImportantDynamicFieldWarnings() {
    return resolvedAnchors
      .filter((anchor) => anchor.isImportant)
      .filter((anchor) => {
        const isEmpty = !anchor.value.trim();
        if (isEmpty) return true;
        if (anchor.locked || isCoreEventAnchorKey(anchor.key)) return false;
        if (!lastDynamicFieldsExportAt) return false;
        if (!anchor.lastUpdatedAt) return true;
        return new Date(anchor.lastUpdatedAt).getTime() <= new Date(lastDynamicFieldsExportAt).getTime();
      });
  }

  function getMissingAnchorUsageWarnings() {
    const anchorValueMap = new Map(
      resolvedAnchors.map((anchor) => [normalizeAnchorKey(anchor.key), anchor.value.trim()])
    );
    const knownAnchorKeys = new Set(anchorValueMap.keys());
    const warnings = new Map<
      string,
      {
        anchorKey: string;
        isUndefinedAnchor: boolean;
        fieldTargets: ValidationFieldTarget[];
        locations: string[];
        rowIds: Set<string>;
      }
    >();

    const collectMissingAnchorWarnings = (
      value: string | Array<string | RecipientEntry> | undefined,
      locationLabel: string,
      options?: { rowId?: string; field?: ValidationFieldName; anchorKeyPrefix?: string }
    ) => {
      const values = Array.isArray(value)
        ? value.map((entry) => (typeof entry === "string" ? entry : entry.type === "email" ? entry.email : entry.name))
        : [value ?? ""];

      values.forEach((entry) => {
        const matches = entry.matchAll(/\[([^\]]+)\]/g);
        for (const match of matches) {
          const normalizedKey = normalizeAnchorKey(match[1] ?? "");
          if (!normalizedKey) continue;
          const isUndefinedAnchor = !knownAnchorKeys.has(normalizedKey);
          if (!isUndefinedAnchor && (anchorValueMap.get(normalizedKey) ?? "").trim()) continue;

          const warningKey = `${isUndefinedAnchor ? "undefined" : "missing"}:${normalizedKey}`;
          const existing = warnings.get(warningKey) ?? {
            anchorKey: normalizedKey,
            isUndefinedAnchor,
            fieldTargets: [],
            locations: [],
            rowIds: new Set<string>(),
          };

          if (!existing.locations.includes(locationLabel)) {
            existing.locations.push(locationLabel);
          }
          if (options?.rowId) {
            existing.rowIds.add(options.rowId);
          }
          if (options?.rowId && options.field) {
            const hasTarget = existing.fieldTargets.some(
              (target) => target.rowId === options.rowId && target.field === options.field
            );
            if (!hasTarget) {
              existing.fieldTargets.push({ rowId: options.rowId, field: options.field });
            }
          }

          warnings.set(warningKey, existing);
        }
      });
    };

    getLatestPreviewPlan().items.forEach((item) => {
      const rowKind = classifyPlanRow(item);
      const rowLabel =
        rowKind === "email" ? "Email" : rowKind === "meeting" ? "Meeting" : "Reminder";
      const rowTitle =
        (item.customTitle ?? item.title).trim() ||
        (rowKind === "email" ? "Email row name" : rowKind === "meeting" ? "Untitled meeting" : "Untitled reminder");

      collectMissingAnchorWarnings(item.customTitle ?? item.title, `${rowLabel} "${rowTitle}" title`, {
        rowId: item.id,
        field: "title",
      });

      if (rowKind === "email") {
        const emailDraft = normalizeEmailDraft(item.emailDraft);
        collectMissingAnchorWarnings(emailDraft.to, `Email "${rowTitle}" To`, { rowId: item.id, field: "emailTo" });
        collectMissingAnchorWarnings(emailDraft.cc, `Email "${rowTitle}" CC`, { rowId: item.id, field: "emailCc" });
        collectMissingAnchorWarnings(emailDraft.bcc, `Email "${rowTitle}" BCC`, { rowId: item.id, field: "emailBcc" });
        collectMissingAnchorWarnings(emailDraft.subject, `Email "${rowTitle}" subject`, {
          rowId: item.id,
          field: "emailSubject",
        });
        collectMissingAnchorWarnings(emailDraft.body, `Email "${rowTitle}" body`, { rowId: item.id, field: "emailBody" });
        return;
      }

      collectMissingAnchorWarnings(item.body ?? "", `${rowLabel} "${rowTitle}" body`, { rowId: item.id, field: "body" });
      collectMissingAnchorWarnings(item.reminderTime ?? "", `${rowLabel} "${rowTitle}" time`, {
        rowId: item.id,
        field: "reminderTime",
      });

      if (rowKind === "meeting") {
        const meetingDraft = normalizeMeetingDraft(item.meetingDraft);
        collectMissingAnchorWarnings(meetingDraft?.attendees ?? [], `Meeting "${rowTitle}" attendees`, {
          rowId: item.id,
          field: "meetingAttendees",
        });
        collectMissingAnchorWarnings(meetingDraft?.location ?? "", `Meeting "${rowTitle}" location`, {
          rowId: item.id,
          field: "meetingLocation",
        });
      }
    });

    return Array.from(warnings.values()).map((warning) => {
      const locationText =
        warning.locations.length === 1
          ? warning.locations[0]
          : `${warning.locations.slice(0, -1).join("; ")}; and ${warning.locations.at(-1)}`;

      return {
        message: warning.isUndefinedAnchor
          ? `Anchor [${warning.anchorKey}] is used in ${locationText} but is not listed in Anchor Fields.`
          : `Anchor [${warning.anchorKey}] is used in ${locationText} but has no value.`,
        anchorKey: warning.anchorKey,
        fieldTargets: warning.fieldTargets,
        rowIds: Array.from(warning.rowIds),
        isUndefinedAnchor: warning.isUndefinedAnchor,
        issueType: "anchor_usage" as const,
        severity: warning.isUndefinedAnchor ? ("warning" as const) : ("error" as const),
      };
    });
  }

  async function confirmImportantDynamicFieldsForExport() {
    const flaggedFields = getImportantDynamicFieldWarnings();
    if (flaggedFields.length === 0) return true;
    const flaggedFieldNames = flaggedFields.map((anchor) => anchor.key.trim() || "Untitled field");

    const confirmed = await showConfirmModal({
      title: "Important field needs review",
      message:
        flaggedFieldNames.length === 1
          ? `The field "${flaggedFieldNames[0]}" is marked as important, but it is empty or has not been updated since the last export.`
          : "These fields are marked as important, but they are empty or have not been updated since the last export.",
      items: flaggedFieldNames.length > 1 ? flaggedFieldNames : undefined,
      confirmLabel: "Export anyway",
      cancelLabel: "Go back",
    });

    return confirmed;
  }

  function cancelEditing() {
    isBuilderDraftPersistencePausedRef.current = true;
    clearPersistedBuilderDraft();

    if (isNewPlanSetupPending) {
      setIsNewPlanSetupPending(false);
      closeBuilderSectionAfterAnimation(() => {
        setHasActivePlanSession(false);
        setSelectedTemplateId(null);
        setBuilderSourceProvenance(null);
        clearTransientEditingState();
      });
      return;
    }
    if (builderMode === "new" && lastTemplateSnapshot) {
      closeBuilderSectionAfterAnimation(() => {
        setHasActivePlanSession(false);
        setSelectedTemplateId(null);
        setBuilderSourceProvenance(lastBuilderSourceProvenance);
        clearTransientEditingState();
      });
      return;
    }
    if (selectedTemplateId) {
      closeBuilderSectionAfterAnimation(() => {
        setHasActivePlanSession(false);
        setSelectedTemplateId(null);
        setBuilderSourceProvenance(null);
        clearTransientEditingState();
      });
      return;
    }
    closeBuilderSectionAfterAnimation(() => {
      setHasActivePlanSession(false);
      setSelectedTemplateId(null);
      setBuilderSourceProvenance(null);
      clearTransientEditingState();
    });
  }

  async function requestCancelEditing() {
    const confirmed = await showConfirmModal({
      title: "Discard this plan?",
      message: "Unsaved changes in this event plan will be cleared. Saved templates are not affected.",
      confirmLabel: "Discard plan",
      cancelLabel: "Keep editing",
      destructive: true,
      severity: "destructive",
    });
    if (!confirmed) return;
    cancelEditing();
  }

  async function exportPreviewReminderItem(itemId: string) {
    await runQueuedExport(async () => {
      const item = getLatestPreviewPlan().items.find((entry) => entry.id === itemId);
      if (!item) return;
      if (await warnIfMissingReminderTimes({ usePopup: true, itemIds: [itemId] })) return;
      if (await warnIfPastScheduledItems({ usePopup: true, itemIds: [itemId] })) return;
      if (!(await confirmImportantDynamicFieldsForExport())) return;
      if (!(await confirmExport())) return;
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
          persistSelectedTemplateDynamicFieldTimestamp(new Date().toISOString());
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

  async function exportCurrentPlan(options?: { skipValidation?: boolean; skipConfirm?: boolean; itemIds?: string[] }) {
    if (!options?.skipValidation && (await validateExportBeforeRun())) return;
    if (await warnIfPastScheduledItems({ usePopup: true, itemIds: options?.itemIds })) return;
    if (!options?.skipConfirm && !(await confirmExport())) return;

    const latestPreviewPlan = getLatestPreviewPlan();
    const selectedItemIds = options?.itemIds ? new Set(options.itemIds) : null;
    const includedItems = selectedItemIds
      ? latestPreviewPlan.items.filter((item) => selectedItemIds.has(item.id))
      : latestPreviewPlan.items;
    if (includedItems.length === 0) {
      await showAlertModal({
        title: "No items selected",
        message: "Select at least one item to export.",
        confirmLabel: "OK",
      });
      return;
    }
    const { emailItems, calendarItems } = partitionPlanItemsByKind(includedItems);
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
    const executableItems = includedItems.filter((item) => {
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
    persistSelectedTemplateDynamicFieldTimestamp(new Date().toISOString());
  }

  function getMissingPreviewRequirements() {
    const missing: MissingFieldIssue[] = [];

    if (effectiveTemplateMode === "template") {
      if (planType === "press_release") {
        if (!guidedForm.releaseName.trim()) missing.push({ message: "Press Release Name", anchorKey: "Press Release Name", issueType: "required", severity: "error" });
        if (!guidedForm.releaseDate) missing.push({ message: "Dissemination Date", anchorKey: "Dissemination Date", issueType: "required", severity: "error" });
        if (!guidedForm.releaseTime) missing.push({ message: "Dissemination Time", anchorKey: "Dissemination Time", issueType: "required", severity: "error" });
      } else if (planType === "earnings") {
        if (!guidedForm.quarter) missing.push({ message: "Quarter", anchorKey: "Quarter", issueType: "required", severity: "error" });
        if (!guidedForm.year.trim()) missing.push({ message: "Year", anchorKey: "Year / Fiscal Year", issueType: "required", severity: "error" });
        if (!guidedForm.earningsDate) missing.push({ message: "Earnings Call Date", anchorKey: "Earnings Call Date", issueType: "required", severity: "error" });
        if (!guidedForm.earningsTime) missing.push({ message: "Earnings Call Time", anchorKey: "Earnings Call Time", issueType: "required", severity: "error" });
      } else if (planType === "conference") {
        if (!guidedForm.conferenceName.trim()) missing.push({ message: "Conference Name", anchorKey: "Conference Name", issueType: "required", severity: "error" });
        if (!guidedForm.conferenceLocation.trim()) missing.push({ message: "Conference Location", anchorKey: "Conference Location", issueType: "required", severity: "error" });
        if (!guidedForm.conferenceDate) missing.push({ message: "Conference Start Date", anchorKey: "Conference Start Date", issueType: "required", severity: "error" });
        if (!guidedForm.conferenceEndDate) missing.push({ message: "Conference End Date", anchorKey: "Conference End Date", issueType: "required", severity: "error" });
      }
    } else {
      if (!eventName.trim()) missing.push({ message: "Event Name", eventName: true, issueType: "required", severity: "error" });
      if (!noEventDate && !hasExplicitEventDate) missing.push({ message: "Event Date", eventDate: true, issueType: "required", severity: "error" });
      if (!eventTime.trim()) missing.push({ message: "Event Time", eventTime: true, issueType: "required", severity: "error" });
    }

    const importantFieldWarnings = getImportantDynamicFieldWarnings().map((anchor) => ({
      message: `Value is missing for anchor field: ${anchor.key.trim() || "Untitled anchor field"}`,
      anchorKey: anchor.key,
      issueType: "important_anchor" as const,
      severity: "warning" as const,
    }));
    missing.push(...importantFieldWarnings);

    const missingAnchorKeys = new Set(
      missing
        .map((issue) => issue.anchorKey ? normalizeAnchorKey(issue.anchorKey) : "")
        .filter(Boolean)
    );

    const dedupedAnchorUsageWarnings = getMissingAnchorUsageWarnings().filter(
      (issue) => issue.isUndefinedAnchor || !issue.anchorKey || !missingAnchorKeys.has(normalizeAnchorKey(issue.anchorKey))
    );

    missing.push(...dedupedAnchorUsageWarnings);

    return missing;
  }

  function formatMissingFieldIssueForModal(issue: MissingFieldIssue) {
    if (issue.isUndefinedAnchor && issue.anchorKey) {
      const displayMessage = issue.message.replace(
        `Anchor [${normalizeAnchorKey(issue.anchorKey)}]`,
        `Warning: Suspected anchor ${formatAnchorTokenDisplay(issue.anchorKey)}`
      );

      return renderTextWithBoldAnchors(displayMessage, {
        anchorClassName: "font-bold text-slate-950",
        invalidAnchorClassName: "font-bold text-red-500",
        knownAnchorKeys,
      });
    }

    if (issue.anchorKey && issue.message.startsWith("Value is missing for anchor field:")) {
      return (
        <>
          <span>Value is missing for anchor field: </span>
          <span className="font-bold text-slate-950">{formatAnchorTokenDisplay(issue.anchorKey)}</span>
        </>
      );
    }

    return renderTextWithBoldAnchors(issue.message, {
      anchorClassName: "font-bold text-slate-950",
      invalidAnchorClassName: "font-bold text-red-500",
      knownAnchorKeys,
    });
  }

  function getValidationFieldLabel(field?: ValidationFieldName) {
    switch (field) {
      case "title":
        return "Title";
      case "body":
        return "Body";
      case "reminderTime":
        return "Time";
      case "emailTo":
        return "Recipients";
      case "emailCc":
        return "CC";
      case "emailBcc":
        return "BCC";
      case "emailSubject":
        return "Subject";
      case "emailBody":
        return "Message";
      case "meetingAttendees":
        return "Attendees";
      case "meetingLocation":
        return "Location";
      default:
        return "Field";
    }
  }

  function getValidationIssueGroupLabel(issue: MissingFieldIssue) {
    if (issue.eventName || issue.eventDate || issue.eventTime) return "Event details";
    const targetRowId = issue.fieldTargets?.[0]?.rowId ?? issue.rowId ?? issue.rowIds?.[0];
    const row = targetRowId ? rows.find((entry) => entry.id === targetRowId) : null;
    if (row) {
      const rowKind = classifyPlanRow(row);
      const rowTypeLabel = rowKind === "email" ? "Email" : rowKind === "meeting" ? "Meeting" : "Reminder";
      const rowTitle =
        rowKind === "email"
          ? normalizeEmailDraft(row.emailDraft).subject.trim() || row.title.trim() || "Untitled email"
          : row.title.trim() || `Untitled ${rowTypeLabel.toLowerCase()}`;
      return `${rowTypeLabel} · ${rowTitle}`;
    }
    return "Anchor Fields";
  }

  function getValidationIssueFieldLabel(issue: MissingFieldIssue) {
    if (issue.eventName) return "Event Name";
    if (issue.eventDate) return "Event Date";
    if (issue.eventTime) return "Event Time";
    if (issue.anchorKey) return formatAnchorTokenDisplay(issue.anchorKey);
    return getValidationFieldLabel(issue.fieldTargets?.[0]?.field);
  }

  function getValidationIssueExplanation(issue: MissingFieldIssue) {
    if (issue.eventName) return "Enter a name for this event.";
    if (issue.eventDate) return "Choose the date this event occurs.";
    if (issue.eventTime) return "Enter the time this event starts.";
    if (issue.isUndefinedAnchor && issue.anchorKey) {
      return `${formatAnchorTokenDisplay(issue.anchorKey)} is used here but is not listed in Anchor Fields.`;
    }
    if (issue.anchorKey) {
      return `Add a value for ${formatAnchorTokenDisplay(issue.anchorKey)} before continuing.`;
    }
    if (issue.message) return issue.message;
    return "Review this field before continuing.";
  }

  function renderValidationIssueGroups(issues: MissingFieldIssue[]) {
    const issueCountLabel = issues.length === 1 ? "1 issue" : `${issues.length} issues`;
    const groups = new Map<string, MissingFieldIssue[]>();
    issues.forEach((issue) => {
      const groupLabel = getValidationIssueGroupLabel(issue);
      groups.set(groupLabel, [...(groups.get(groupLabel) ?? []), issue]);
    });

    return (
      <div className="space-y-3">
        <div className="text-[13px] font-semibold text-slate-600">{issueCountLabel}</div>
        <div className="overflow-hidden rounded-[10px] border border-slate-200 bg-white">
          {Array.from(groups.entries()).map(([groupLabel, groupIssues], groupIndex) => (
            <div key={groupLabel} className={groupIndex > 0 ? "border-t border-slate-200/80" : ""}>
              <div className="bg-slate-50/70 px-3.5 py-2 text-[13px] font-semibold text-slate-700">
                {groupLabel}
              </div>
              <div className="divide-y divide-slate-200/80">
                {groupIssues.map((issue, index) => {
                  const isWarning = issue.severity === "warning" || issue.isUndefinedAnchor;
                  return (
                    <div key={`${groupLabel}-${index}`} className="grid gap-1 px-3.5 py-3 min-[560px]:grid-cols-[92px_minmax(130px,0.48fr)_minmax(0,1fr)] min-[560px]:gap-3">
                      <div className={`text-[13px] font-semibold ${isWarning ? "text-amber-700" : "text-red-600"}`}>
                        {isWarning ? "Check this" : "Required"}
                      </div>
                      <div className="min-w-0 text-[14px] font-semibold leading-5 text-slate-900">
                        {getValidationIssueFieldLabel(issue)}
                      </div>
                      <div className="min-w-0 text-[14px] leading-5 text-slate-600">
                        {renderTextWithBoldAnchors(getValidationIssueExplanation(issue), {
                          anchorClassName: "font-bold text-slate-950",
                          invalidAnchorClassName: "font-bold text-red-500",
                          knownAnchorKeys,
                        })}
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          ))}
        </div>
      </div>
    );
  }

  async function validatePreviewBeforeOpen() {
    const missingRequirements = getMissingPreviewRequirements();
    if (missingRequirements.length > 0) {
      await showAlertModal({
        title: "Resolve plan issues",
        message: "Review the highlighted fields before continuing.",
        content: renderValidationIssueGroups(missingRequirements),
        secondaryLabel: "Back to plan",
        confirmLabel: "Fix first issue",
        onConfirmAction: () => applyMissingFieldHighlights(missingRequirements),
        severity: "validation",
        maxWidthClassName: "max-w-[640px]",
      });
      return true;
    }
    clearMissingFieldHighlights();
    if (await validateEmailRowsForExport(undefined, { usePopup: true })) return true;
    if (await validateMeetingRowsForExport(undefined, { usePopup: true })) return true;
    if (await warnIfMissingReminderTimes({ usePopup: true })) return true;
    return false;
  }

  async function validateExportBeforeRun() {
    const missingRequirements = getMissingPreviewRequirements();
    if (missingRequirements.length > 0) {
      await showAlertModal({
        title: "Resolve plan issues",
        message: "Review the highlighted fields before continuing.",
        content: renderValidationIssueGroups(missingRequirements),
        secondaryLabel: "Back to plan",
        confirmLabel: "Fix first issue",
        onConfirmAction: () => applyMissingFieldHighlights(missingRequirements),
        severity: "validation",
        maxWidthClassName: "max-w-[640px]",
      });
      return true;
    }
    clearMissingFieldHighlights();
    if (await validateEmailRowsForExport(undefined, { usePopup: true })) return true;
    if (await validateMeetingRowsForExport(undefined, { usePopup: true })) return true;
    if (await warnIfMissingReminderTimes({ usePopup: true })) return true;
    if (!(await confirmImportantDynamicFieldsForExport())) return true;
    return false;
  }

  async function deleteSavedTemplate(templateId: string) {
    const template = savedTemplates.find((entry) => entry.id === templateId);
    if (!template) return;
    const confirmed = await showConfirmModal({
      title: "Delete template?",
      message: `“${template.name}” will be removed from your saved templates. This cannot be undone.`,
      confirmLabel: "Delete template",
      cancelLabel: "Cancel",
      destructive: true,
      severity: "destructive",
    });
    if (!confirmed) {
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
        setHasActivePlanSession(false);
        resetBuilderToNewPlan("", { revealBuilder: false });
      }
    }
  }

  async function renameSavedTemplate(templateId: string) {
    const template = savedTemplates.find((entry) => entry.id === templateId);
    if (!template) return;

    const nextName = await showPromptModal({
      title: "Rename template",
      message: "Choose a clear name for this reusable workflow.",
      defaultValue: template.name,
      inputLabel: "Template name",
      placeholder: "Template name",
      confirmLabel: "Rename template",
      cancelLabel: "Cancel",
      severity: "information",
    });
    if (!nextName) return;

    const trimmedName = nextName.trim();
    if (!trimmedName) {
      setTemplateActionMessage("Enter a template name.");
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

  function createBuilderRowForKind(kind: RowEditorKind, defaultTimeZone = eventTimeZone) {
    const normalizedTimeZone = normalizeOutlookTimeZone(defaultTimeZone);
    if (kind === "email") {
      return {
        ...createEmptyBuilderRow(),
        reminderTime: normalizedDefaultReminderTime,
        timeZone: normalizedTimeZone,
        rowType: "email" as const,
        emailDraft: { to: [], cc: [], bcc: [], subject: "", body: "" },
      };
    }
    if (kind === "meeting") {
      return {
        ...createEmptyBuilderRow(),
        reminderTime: normalizedDefaultReminderTime,
        timeZone: normalizedTimeZone,
        rowType: "calendar_event" as const,
        meetingDraft: {
          attendees: [],
          location: "",
          durationMinutes: 30,
        },
      };
    }
    return {
      ...createEmptyBuilderRow(normalizedDefaultReminderTime),
      timeZone: normalizedTimeZone,
    };
  }

  function addBuilderRow(kind: RowEditorKind) {
    clearScheduledAddRowLaunch();
    clearScheduledAddRowInsert();
    setHiddenAddRowId(null);
    setAddRowSettlingRowId(null);
    setInstantlyVisibleRowId(null);
    setIsSortMenuOpen(false);

    const nextRow = createBuilderRowForKind(kind);
    setRows((current) => [nextRow, ...current]);
    setInstantlyVisibleRowId(nextRow.id);
    scheduledAddRowLaunchRef.current = window.setTimeout(() => {
      scheduledAddRowLaunchRef.current = null;
      setInstantlyVisibleRowId((current) => (current === nextRow.id ? null : current));
    }, 420);
    setFocusedTitleInputId(`collapsed-title-${nextRow.id}`);
    scheduleRowEditorOpen(nextRow.id, kind);
  }

  function addReminderRow() {
    const currentOpenEditor = getCurrentOpenRowEditor();
    if (currentOpenEditor || closingRowEditor) {
      closeRowEditorsAfterAnimation(() => {
        addBuilderRow("reminder");
      });
      return;
    }
    addBuilderRow("reminder");
  }

  function addEmailRow() {
    const currentOpenEditor = getCurrentOpenRowEditor();
    if (currentOpenEditor || closingRowEditor) {
      closeRowEditorsAfterAnimation(() => {
        addBuilderRow("email");
      });
      return;
    }
    addBuilderRow("email");
  }

  function addMeetingRow() {
    const currentOpenEditor = getCurrentOpenRowEditor();
    if (currentOpenEditor || closingRowEditor) {
      closeRowEditorsAfterAnimation(() => {
        addBuilderRow("meeting");
      });
      return;
    }
    addBuilderRow("meeting");
  }

  function deleteBuilderRowImmediately(rowId: string) {
    clearScheduledRowEditorOpen();
    clearScheduledRowEditorClose();
    setRows((current) => current.filter((entry) => entry.id !== rowId));
    setOpenEmailDraftRowId((current) => (current === rowId ? null : current));
    setOpenMeetingEditorRowId((current) => (current === rowId ? null : current));
    setOpenDurationEditorRowId((current) => (current === rowId ? null : current));
    setOpenBodyEditorRowId((current) => (current === rowId ? null : current));
    setClosingRowEditor((current) => (current?.rowId === rowId ? null : current));
    setMeetingValidationErrors((current) => {
      const next = { ...current };
      delete next[rowId];
      return next;
    });
    setEmailFieldVisibility((current) => {
      const next = { ...current };
      delete next[rowId];
      return next;
    });
    setOffsetDrafts((current) => {
      const next = { ...current };
      delete next[rowId];
      return next;
    });
    setTimeInputDrafts((current) => clearReminderTimeDraft(current, rowId));
    setOpenTimeZoneRowId((current) => (current === rowId ? null : current));
    setTimeZoneSearch("");
    setFocusedTimeInputRowId((current) => (current === rowId ? null : current));
    setFocusedTitleInputId((current) =>
      current &&
      (current === `collapsed-title-${rowId}` ||
        current === `email-title-${rowId}` ||
        current === `reminder-title-${rowId}` ||
        current === `meeting-title-${rowId}`)
        ? null
        : current
    );
  }

  function deleteBuilderRow(rowId: string) {
    const currentOpenEditor = getCurrentOpenRowEditor();

    if (currentOpenEditor?.rowId === rowId) {
      closeRowEditorsAfterAnimation(() => deleteBuilderRowImmediately(rowId));
      return;
    }

    if (closingRowEditor?.rowId === rowId) {
      window.setTimeout(() => deleteBuilderRowImmediately(rowId), ROW_EDITOR_CLOSE_ANIMATION_MS);
      return;
    }

    deleteBuilderRowImmediately(rowId);
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
  const normalizedGlobalTimeZoneSearch = timeZoneSearch.trim().toLowerCase();
  const filteredGlobalTimeZoneOptions = normalizedGlobalTimeZoneSearch
    ? timeZoneOptions.filter((timeZone) =>
        `${timeZone.label} ${timeZone.value}`.toLowerCase().includes(normalizedGlobalTimeZoneSearch)
      )
    : timeZoneOptions;

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
    eventTimeZone,
    anchors,
    rows,
  });
  const shouldRenderBuilderSourceBanner = hasMounted && builderHasMeaningfulContent;
  const isBuilderWorkspaceVisible = true;
  const isBuilderEntryRevealActive = isBuilderWorkspaceVisible && !isBuilderVisualRevealDeferred;

  useLayoutEffect(() => {
    if (!focusedTitleInputId) return;
    const input = builderTitleInputRefs.current[focusedTitleInputId];
    if (!input) return;
    if (document.activeElement !== input) {
      input.focus({ preventScroll: true });
      input.select();
    }
    input.style.height = "0px";
    input.style.height = `${input.scrollHeight}px`;
  }, [focusedTitleInputId, rows]);

  useEffect(() => {
    setEventDateInputValue(noEventDate ? "Today" : formatEventDateForEditor(anchorDate));
  }, [anchorDate, noEventDate]);

  useEffect(() => {
    setEventTimeInputValue(eventTime ? formatTimeForEditor(eventTime) : "");
  }, [eventTime]);

  useEffect(() => {
    if (!isBuilderSectionVisible) return;
    if (!shouldRenderSimpleEventHeaderFields) return;
    if (!shouldFocusEventNameInputRef.current) return;
    if (shouldScrollBuilderIntoViewRef.current) return;

    const timeoutIds: number[] = [];
    const focusHandle = window.requestAnimationFrame(() => {
      focusEventNameInput();
      timeoutIds.push(window.setTimeout(() => focusEventNameInput(), 80));
      timeoutIds.push(window.setTimeout(() => focusEventNameInput(), 180));
      timeoutIds.push(window.setTimeout(() => focusEventNameInput(), 320));
    });

    return () => {
      window.cancelAnimationFrame(focusHandle);
      timeoutIds.forEach((timeoutId) => window.clearTimeout(timeoutId));
    };
  }, [focusEventNameInput, isBuilderSectionVisible, shouldRenderSimpleEventHeaderFields]);

  useEffect(() => {
    if (!isBuilderSectionVisible) return;
    if (!shouldScrollBuilderIntoViewRef.current) return;

    shouldScrollBuilderIntoViewRef.current = false;
    const requestId = builderScrollRequestIdRef.current + 1;
    builderScrollRequestIdRef.current = requestId;

    const scrollToEventHeader = (behavior: ScrollBehavior) => {
      if (builderScrollRequestIdRef.current !== requestId) return;
      const scrollTarget = eventHeaderCardRef.current ?? builderSectionRef.current;
      if (!scrollTarget) return;
      const topOffset = 12;
      const nextScrollTop = Math.max(0, window.scrollY + scrollTarget.getBoundingClientRect().top - topOffset);
      window.scrollTo({
        top: nextScrollTop,
        behavior,
      });
    };

    const scrollTimeoutIds = [
      window.setTimeout(() => {
        scrollToEventHeader("smooth");
      }, ROW_EDITOR_OPEN_ANIMATION_MS + 80),
      window.setTimeout(() => {
        scrollToEventHeader("auto");
      }, ROW_EDITOR_OPEN_ANIMATION_MS + 420),
      window.setTimeout(() => {
        scrollToEventHeader("auto");
      }, ROW_EDITOR_OPEN_ANIMATION_MS + 760),
    ];

    const focusTimeoutId = window.setTimeout(() => {
      if (builderScrollRequestIdRef.current !== requestId) return;
      if (shouldFocusEventNameInputRef.current) {
        focusEventNameInput();
      }
    }, ROW_EDITOR_OPEN_ANIMATION_MS + 460);

    return () => {
      scrollTimeoutIds.forEach((timeoutId) => window.clearTimeout(timeoutId));
      window.clearTimeout(focusTimeoutId);
    };
  }, [builderScrollRequestNonce, focusEventNameInput, isBuilderSectionVisible]);

  useEffect(() => {
    if (!isBuilderWorkspaceVisible || !isBuilderVisualRevealDeferred) return;

    const timeoutId = window.setTimeout(() => {
      setIsBuilderVisualRevealDeferred(false);
    }, 420);

    return () => window.clearTimeout(timeoutId);
  }, [isBuilderVisualRevealDeferred, isBuilderWorkspaceVisible]);

  useEffect(() => {
    if (isBuilderSectionVisible) return;
    setIsBuilderEntryRevealImmediate(false);
  }, [isBuilderSectionVisible]);

  useEffect(() => {
    if (selectedTemplateId) return;
    if (hasActivePlanSession) return;
    if (showNewPlanDialog || isNewPlanSetupPending) return;
    if (!isBuilderSectionVisible) return;
    setIsBuilderSectionVisible(false);
  }, [hasActivePlanSession, isBuilderSectionVisible, isNewPlanSetupPending, selectedTemplateId, showNewPlanDialog]);

  useEffect(() => {
    if (selectedTemplateId) return;
    if (hasActivePlanSession) return;
    if (showNewPlanDialog || isNewPlanSetupPending) return;
    if (builderHasMeaningfulContent) return;
    if (!isBuilderSectionVisible) return;
    setIsBuilderSectionVisible(false);
  }, [
    builderHasMeaningfulContent,
    hasActivePlanSession,
    isBuilderSectionVisible,
    isNewPlanSetupPending,
    selectedTemplateId,
    showNewPlanDialog,
  ]);

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
    if (selectedTemplateId) return;
    if (builderMode !== "template") return;
    if (builderSourceProvenance?.sourceType !== "saved_template") return;
    setIsBuilderSectionVisible(false);
  }, [builderMode, builderSourceProvenance, selectedTemplateId]);

  useEffect(() => {
    if (AI_ENABLED) return;
    if (!isAiPanelOpen) return;
    setIsAiPanelOpen(false);
  }, [isAiPanelOpen]);

  return (
    <div className={plansWorkspaceClass}>
      <AppShellSidebarContent>
        <div>
          <div className="text-[11px] font-semibold uppercase tracking-[0.14em] text-slate-500">Connected Email Account</div>
          <div className={`mt-2 break-words text-[11px] font-medium leading-[1.5] ${accountStatusClass}`}>{accountPrimaryText}</div>
          <div className={`mt-1 text-xs font-semibold uppercase tracking-wide ${accountStatusClass}`}>
            {accountStatusLabel}
          </div>
          {providerLoading.outlook || providerLoading.gmail ? (
            <div className="mt-1 text-xs text-gray-500">Refreshing provider status…</div>
          ) : null}
          <Link
            href="/settings"
            className="mt-3 inline-flex rounded-lg border border-gray-200 px-3 py-1.5 text-xs text-gray-700 hover:bg-gray-50"
          >
            {accountButtonLabel}
          </Link>
        </div>
      </AppShellSidebarContent>

      {AI_ENABLED ? (
        <section className="space-y-2">
          <button
            type="button"
            onClick={openAiPanel}
            className="rounded-lg border border-gray-300 bg-white px-4 py-2 text-sm text-gray-900 hover:bg-gray-50"
          >
            Generate with AI
          </button>
        </section>
      ) : null}

      {executionNotices.length > 0 && !isBuilderPreviewOpen ? (
        <div className="space-y-3">
          {executionNotices.map((entry) => (
            <OutlookExecutionNoticeCard key={entry.id} notice={entry.notice} onDismiss={() => dismissExecutionNotice(entry.id)} />
          ))}
        </div>
      ) : null}

      {AI_ENABLED && aiApplySuccessMessage ? (
        <div className="rounded-2xl border border-green-200 bg-green-50 px-4 py-3 text-sm text-green-900 shadow-sm">
          {aiApplySuccessMessage}
        </div>
      ) : null}

      {isBuilderEntryRunwayVisible && !isBuilderWorkspaceVisible ? (
        <div className="h-[76vh] min-h-[680px]" aria-hidden="true" />
      ) : null}

      <AnimatedRowEditor
        open={isBuilderWorkspaceVisible}
      >
        <section
          ref={builderSectionRef}
          data-plan-builder="true"
          className="min-w-0"
        >
          <div className={`space-y-5 transition-opacity duration-300 ${isBuilderVisualRevealDeferred ? "opacity-0" : ""}`}>
              <div className="transition-opacity duration-200">
                <div>
                  <div>
                    {shouldRenderSimpleEventHeaderFields ? (
                      <div className="space-y-[18px]">
			                        <StaggeredInlineEditorItem active={isBuilderEntryRevealActive} immediate={isBuilderEntryRevealImmediate} delayMs={20} className={builderInactiveControlClass}>
			                          <div className="flex flex-col gap-[18px] md:flex-row md:items-start md:justify-between md:gap-6">
			                            <div className="min-w-0">
			                              <h1 className="text-[30px] font-bold leading-[1.08] tracking-[-0.02em] text-slate-950 md:text-[34px]">Plans</h1>
			                              <p className="mt-1.5 text-[15px] leading-[1.4] text-slate-600">
			                                Create a new event plan or reuse a saved workflow.
			                              </p>
                                      <p className="mt-[7px] text-[13px] font-medium leading-5 text-slate-500">
                                        {currentSourceLine}
                                      </p>
			                            </div>
                              <div className="grid w-full grid-cols-2 gap-[10px] md:w-auto md:flex md:shrink-0 md:justify-end">
                                <button
                                  ref={templateLibraryButtonRef}
                                  type="button"
                                  onClick={() => setIsTemplateLibraryOpen(true)}
                                  className={`${plansSecondaryButtonClass} min-w-0 whitespace-nowrap rounded-[10px] px-3 md:min-w-[128px] md:px-4`}
                                >
                                  {templateLibraryButtonLabel}
                                </button>
			                              <button
			                                type="button"
                                  onClick={() => {
                                    setIsTemplateLibraryOpen(false);
                                    void startNewPlan();
                                  }}
                                  className={`${plansPrimaryButtonClass} min-w-0 whitespace-nowrap rounded-[10px] px-3 md:min-w-[118px] md:px-4`}
			                              >
			                                + New Event
				                              </button>
                              </div>
				                          </div>
				                        </StaggeredInlineEditorItem>
			                        <StaggeredInlineEditorItem active={isBuilderEntryRevealActive} immediate={isBuilderEntryRevealImmediate} delayMs={40} className={`relative w-full transition-opacity duration-200 ${openTimeZoneRowId === EVENT_TIME_ZONE_MENU_ID ? "z-[120]" : "z-30"} ${builderInactiveControlClass}`}>
			                          <section ref={eventHeaderCardRef} className={plansSurfaceClass}>
		                            <div className={plansSectionHeaderClass}>
		                              <h2 className={plansPanelHeadingClass}>Event details</h2>
		                              <p className={plansPanelHelperClass}>Name and schedule this run.</p>
		                            </div>
		                            <div className={plansCanvasSectionClass}>
		                            <div>
		                              <div>
		                                <label className={plansFieldLabelClass}>Event Name</label>
		                                <input
	                                  ref={eventNameInputRef}
	                                  className={`${plansInputClass} ${
	                                    missingFieldHighlights.eventName ? "border-red-300 ring-2 ring-red-200" : ""
	                                  }`}
	                                  placeholder="Enter event name..."
	                                  value={eventName}
	                                  onKeyDown={(e) => {
	                                    if (e.key === "Enter") {
	                                      e.preventDefault();
	                                      focusNextEventHeaderField("eventDate");
	                                      return;
	                                    }
	                                    if (e.key === "Tab" && !e.shiftKey) {
	                                      e.preventDefault();
	                                      focusNextEventHeaderField("eventDate");
	                                    }
	                                  }}
	                                  onChange={(e) => {
	                                    setEventName(e.target.value);
	                                    if (missingFieldHighlights.eventName) {
	                                      setMissingFieldHighlights((current) => ({ ...current, eventName: false }));
	                                    }
		                                  }}
		                                />
		                              </div>
		                            </div>
		                            <div className="mt-4 grid gap-4 min-[520px]:grid-cols-2 min-[900px]:grid-cols-3 min-[900px]:items-start">
			                              <div className={eventDetailFieldClass}>
		                                <label className={eventDetailLabelClass}>Event Date</label>
		                                <input
	                                  ref={eventDateInputRef}
	                                  type="text"
	                                  inputMode="numeric"
	                                  className={`${plansInputClass} ${
	                                    missingFieldHighlights.eventDate ? "border-red-300 ring-2 ring-red-200" : ""
	                                  }`}
	                                  value={noEventDate ? "Today" : eventDateInputValue}
	                                  readOnly={noEventDate}
	                                  placeholder="mm/dd/yyyy"
	                                  onKeyDown={(e) => {
	                                    if (e.key === "Enter") {
	                                      e.preventDefault();
	                                      focusNextEventHeaderField("eventTime");
	                                    }
	                                  }}
	                                  onChange={(e) => setEventDateInputValue(e.target.value)}
	                                  onBlur={(e) => {
	                                    const parsed = parseEventDateInput(e.target.value);
	                                    if (parsed === null) {
	                                      setEventDateInputValue(formatEventDateForEditor(anchorDate));
	                                      return;
	                                    }
	                                    setAnchorDate(parsed);
	                                    setHasExplicitEventDate(Boolean(parsed.trim()));
	                                    setEventDateInputValue(formatEventDateForEditor(parsed));
	                                    if (missingFieldHighlights.eventDate) {
	                                      setMissingFieldHighlights((current) => ({ ...current, eventDate: false }));
	                                    }
	                                  }}
	                                />
		                                <label className={`${eventDetailSupplementClass} gap-2`}>
		                                  <input
	                                    ref={useTodayInputRef}
	                                    type="checkbox"
	                                    className="h-3.5 w-3.5 rounded border-slate-300 text-[#315f92] focus:ring-[#6f9fd1]/30"
	                                    checked={noEventDate}
	                                    onChange={(e) => {
	                                      setNoEventDate(e.target.checked);
	                                      if (missingFieldHighlights.eventDate) {
	                                        setMissingFieldHighlights((current) => ({ ...current, eventDate: false }));
	                                      }
	                                    }}
	                                  />
	                                  <span>Today</span>
	                                  <span className="group relative inline-flex h-4 w-4 items-center justify-center rounded-full border border-slate-300 bg-white text-[10px] font-semibold text-slate-500">
	                                    i
	                                    <span className="pointer-events-none absolute bottom-full left-1/2 z-10 mb-2 hidden w-48 -translate-x-1/2 rounded-lg border border-slate-200 bg-white px-2 py-1 text-[11px] font-normal leading-4 text-slate-600 shadow-lg group-hover:block">
	                                      Schedules this plan from today instead of using a specific event date.
	                                    </span>
	                                  </span>
	                                </label>
	                              </div>
		                              <div className={eventDetailFieldClass}>
		                                <label className={eventDetailLabelClass}>Event Time</label>
		                                <div
	                                  className={`flex h-[42px] w-full items-center gap-2 rounded-xl border bg-white px-3.5 text-slate-950 shadow-sm ${
	                                    missingFieldHighlights.eventTime ? "border-red-300 ring-2 ring-red-200" : "border-slate-200"
	                                  }`}
	                                >
	                                  <input
	                                    ref={eventTimeInputRef}
	                                    type="text"
	                                    inputMode="text"
	                                    className="min-w-0 flex-1 border-0 bg-transparent p-0 text-[14px] font-semibold text-slate-950 placeholder:font-normal placeholder:text-slate-400 focus:outline-none focus:ring-0"
	                                    value={eventTimeInputValue}
	                                    placeholder="--:-- --"
	                                    onKeyDown={(e) => {
	                                      if (e.key === "Enter") {
	                                        e.preventDefault();
	                                        focusNextEventHeaderField("weekendHandling");
	                                      }
	                                    }}
	                                    onChange={(e) => {
	                                      const rawValue = e.target.value;
	                                      const selectionEnd = e.target.selectionEnd ?? rawValue.length;
	                                      const meaningfulCount = countReminderTimeMeaningfulChars(rawValue, selectionEnd);
	                                      const nextMaskedValue = maskReminderTimeDraftInput(rawValue);
	                                      const nextCursor = findReminderTimeCursorFromMeaningfulCount(nextMaskedValue, meaningfulCount);

	                                      setEventTimeInputValue(nextMaskedValue);

	                                      requestAnimationFrame(() => {
	                                        const input = eventTimeInputRef.current;
	                                        if (input && document.activeElement === input) {
	                                          input.setSelectionRange(nextCursor, nextCursor);
	                                        }
	                                      });
	                                    }}
	                                    onBlur={(e) => {
	                                      const normalized = normalizeReminderTimeInput(e.target.value);
	                                      setEventTime(normalized);
	                                      setEventTimeInputValue(normalized || "");
	                                      if (missingFieldHighlights.eventTime) {
	                                        setMissingFieldHighlights((current) => ({ ...current, eventTime: false }));
	                                      }
	                                    }}
	                                  />
	                                  <span className="relative inline-flex">
	                                    <button
	                                      type="button"
	                                      aria-label="Set event time zone"
	                                      title="Set time zone"
	                                      onClick={() => {
	                                        setTimeZoneSearch("");
	                                        setOpenTimeZoneRowId((current) => (current === EVENT_TIME_ZONE_MENU_ID ? null : EVENT_TIME_ZONE_MENU_ID));
	                                      }}
	                                      className="group/timezone-globe inline-flex h-[22px] w-[22px] items-center justify-center rounded-full text-slate-500 transition hover:bg-slate-100 hover:text-slate-700"
	                                    >
	                                      <svg
	                                        viewBox="0 0 24 24"
	                                        aria-hidden="true"
	                                        className="h-3.5 w-3.5"
	                                        fill="none"
	                                        stroke="currentColor"
	                                        strokeWidth="1.9"
	                                        strokeLinecap="round"
	                                        strokeLinejoin="round"
	                                      >
	                                        <circle cx="12" cy="12" r="9" />
	                                        <path d="M3 12h18" />
	                                        <path d="M12 3c2.35 2.46 3.55 5.46 3.55 9S14.35 18.54 12 21" />
	                                        <path d="M12 3c-2.35 2.46-3.55 5.46-3.55 9S9.65 18.54 12 21" />
	                                      </svg>
	                                      <span className="pointer-events-none absolute bottom-full left-1/2 z-40 mb-2 hidden w-24 -translate-x-1/2 rounded-lg border border-slate-200 bg-white px-2 py-1 text-center text-[11px] font-medium text-slate-600 shadow-lg group-hover/timezone-globe:block">
	                                        Set time zone
	                                      </span>
	                                    </button>
	                                    {openTimeZoneRowId === EVENT_TIME_ZONE_MENU_ID ? (
	                                      <div className="absolute right-0 top-[calc(100%+0.5rem)] z-[170] w-72 rounded-2xl border border-slate-600 bg-slate-800 p-2 text-left shadow-[0_18px_36px_-16px_rgba(15,23,42,0.65)]">
	                                        <div className="mb-2 px-2 text-[11px] font-semibold text-slate-200">
	                                          {getOutlookTimeZoneLabel(eventTimeZone)}
	                                        </div>
	                                        <input
	                                          type="search"
	                                          autoFocus
	                                          value={timeZoneSearch}
	                                          onChange={(e) => setTimeZoneSearch(e.target.value)}
	                                          onKeyDown={(e) => {
	                                            if (e.key === "Escape") {
	                                              e.preventDefault();
	                                              setOpenTimeZoneRowId(null);
	                                              setTimeZoneSearch("");
	                                            }
	                                          }}
	                                          placeholder="Search time zones"
	                                          className="mb-2 w-full rounded-xl border border-slate-600 bg-slate-950 px-3 py-2 text-[12px] text-white placeholder:text-slate-400 focus:border-blue-300 focus:outline-none focus:ring-2 focus:ring-blue-400/30"
	                                        />
	                                        <div className="max-h-64 overflow-y-auto rounded-xl border border-slate-600 bg-slate-900 p-1">
	                                          {filteredGlobalTimeZoneOptions.length > 0 ? (
	                                            filteredGlobalTimeZoneOptions.map((timeZone) => {
	                                              const isSelectedTimeZone = timeZone.value === eventTimeZone;
	                                              return (
	                                                <button
	                                                  key={timeZone.value}
	                                                  type="button"
	                                                  onClick={() => {
	                                                    applyEventTimeZone(timeZone.value);
	                                                    setOpenTimeZoneRowId(null);
	                                                    setTimeZoneSearch("");
	                                                  }}
	                                                  className={`block w-full rounded-lg px-2 py-1.5 text-left text-[12px] leading-5 transition ${
	                                                    isSelectedTimeZone
	                                                      ? "bg-blue-500/30 text-white"
	                                                      : "text-slate-200 hover:bg-slate-700"
	                                                  }`}
	                                                >
	                                                  <span className="block font-medium">{timeZone.label}</span>
	                                                  <span className="mt-0.5 block text-[10px] text-slate-400">{timeZone.value}</span>
	                                                </button>
	                                              );
	                                            })
	                                          ) : (
	                                            <div className="px-3 py-4 text-center text-[12px] text-slate-400">
	                                              No time zones found
	                </div>
	              )}
	            </div>
		          </div>
		                              ) : null}
		                                  </span>
		                                </div>
		                                <div className={eventDetailSupplementSpacerClass} aria-hidden="true" />
		                              </div>
			                              <div className={`${eventDetailFieldClass} min-[520px]:col-span-2 min-[900px]:col-span-1`}>
			                                <label className={eventDetailLabelClass}>
		                                  <span>Weekend Handling</span>
	                                  <span className="group relative inline-flex h-4 w-4 items-center justify-center rounded-full border border-slate-300 bg-white text-[10px] font-semibold text-slate-500">
	                                    i
	                                    <span className="pointer-events-none absolute bottom-full left-1/2 z-10 mb-2 hidden w-48 -translate-x-1/2 rounded-lg border border-slate-200 bg-white px-2 py-1 text-[11px] font-normal leading-4 text-slate-600 shadow-lg group-hover:block">
	                                      Choose whether dates that land on weekends should stay there or move to Friday.
	                                    </span>
	                                  </span>
	                                </label>
	                                <select
	                                  ref={weekendHandlingSelectRef}
	                                  className={plansInputClass}
	                                  value={weekendRule}
	                                  onKeyDown={(e) => {
	                                    if (e.key === "Enter") {
	                                      e.preventDefault();
	                                      focusNextEventHeaderField("useToday");
	                                    }
	                                  }}
	                                  onChange={(e) => setWeekendRule(e.target.value as WeekendRule)}
	                                >
	                                  <option value="none">Keep weekend date</option>
		                                  <option value="prior_business_day">Move to Friday</option>
		                                </select>
		                                <div className={eventDetailSupplementSpacerClass} aria-hidden="true" />
		                              </div>
	                            </div>
	                            </div>
	                          </section>
	                        </StaggeredInlineEditorItem>
			              <StaggeredInlineEditorItem
			                active={isBuilderEntryRevealActive}
			                immediate={isBuilderEntryRevealImmediate}
			                delayMs={180}
			                className="relative z-20 transition-opacity duration-200"
			              >
			                <section className={plansSurfaceClass}>
		                  <div className={plansSectionHeaderClass}>
		                  <div className="flex flex-col gap-4 min-[900px]:flex-row min-[900px]:items-start min-[900px]:justify-between">
		                    <div className="min-w-0">
		                        <h2 className={plansPanelHeadingClass}>Workflow actions</h2>
		                        <p className={plansPanelHelperClass}>Add the reminders, emails, and meetings that should run for this event.</p>
		                        {renderedRows.length > 0 ? (
		                          <p className="mt-2 text-[13px] font-medium text-slate-500">{renderedRows.length} actions</p>
		                        ) : null}
		                    </div>
			                    <div className="grid grid-cols-2 gap-2 min-[520px]:flex min-[520px]:flex-wrap min-[900px]:justify-end">
		                      <button
		                        ref={addReminderButtonRef}
		                        type="button"
		                        onClick={addReminderRow}
		                        className={`${plansToolbarButtonClass} border-blue-200 text-blue-700 hover:border-blue-300 hover:bg-blue-50/70`}
		                      >
		                        + Reminder
		                      </button>
	                      <button
		                        ref={addEmailButtonRef}
		                        type="button"
		                        onClick={addEmailRow}
		                        className={`${plansToolbarButtonClass} border-emerald-200 text-emerald-700 hover:border-emerald-300 hover:bg-emerald-50/70`}
		                      >
		                        + Email
		                      </button>
	                      <button
		                        ref={addMeetingButtonRef}
		                        type="button"
		                        onClick={addMeetingRow}
		                        className={`${plansToolbarButtonClass} border-violet-200 text-violet-700 hover:border-violet-300 hover:bg-violet-50/70`}
		                      >
		                        + Meeting
		                      </button>
	                      {renderSortControl()}
	                    </div>
	                  </div>
	                  </div>
	                  <div className={plansCanvasSectionClass}>
			                  <div className="overflow-hidden border-t border-slate-200/80 bg-white">
			                    <div ref={rowListStackRef}>
			                      <div className="relative">
			                        <div ref={rowInsertAnchorRef} className="h-0" aria-hidden="true" />
                                  {renderedRows.length > 0 ? (
			                        <div className="divide-y divide-slate-200/80">
	                  {renderedRows.map((row, index) => {
                    const rowMeta = getBuilderRowTypeMeta(row);
                    const rowKind = classifyPlanRow(row);
                    const shouldHideFollowingRowDuringAddRowSettle =
                      addRowSettlingRowId !== null &&
                      renderedRows[0]?.id === addRowSettlingRowId &&
                      row.id !== addRowSettlingRowId &&
                      index === 1;
                    const collapsedTitleInputId = `collapsed-title-${row.id}`;
                    const reminderTitleInputId = `reminder-title-${row.id}`;
                    const meetingTitleInputId = `meeting-title-${row.id}`;
                    const emailUsesTimingFields = row.rowType === "email" && appSettings.emailHandlingMode === "schedule";
                    const meetingErrors = meetingValidationErrors[row.id];
                    const isDraggingRow = draggingRowId === row.id;
                    const isPressedRow = pressedRowId === row.id;
                    const isActiveDragRow = isDraggingRow || isPressedRow;
                    const isClosingEmailEditor = closingRowEditor?.rowId === row.id && closingRowEditor.kind === "email";
                    const isClosingReminderBodyEditor =
                      closingRowEditor?.rowId === row.id &&
                      (closingRowEditor.kind === "reminderBody" || closingRowEditor.kind === "reminder");
                    const isClosingReminderDurationEditor =
                      closingRowEditor?.rowId === row.id &&
                      (closingRowEditor.kind === "reminderDuration" || closingRowEditor.kind === "reminder");
                    const isClosingMeetingEditor = closingRowEditor?.rowId === row.id && closingRowEditor.kind === "meeting";
                    const isEmailEditorOpen = openEmailDraftRowId === row.id;
                    const isTimeZoneMenuOpen = openTimeZoneRowId === row.id;
                    const isReminderInlineEditorOpen =
                      rowKind !== "email" &&
                      rowKind !== "meeting" &&
                      (openBodyEditorRowId === row.id || openDurationEditorRowId === row.id);
                    const isReminderInlineEditorRendered =
                      rowKind !== "email" &&
                      rowKind !== "meeting" &&
                      (isReminderInlineEditorOpen || isClosingReminderBodyEditor || isClosingReminderDurationEditor);
                    const isMeetingEditorOpen = openMeetingEditorRowId === row.id || isClosingMeetingEditor;
                    const isThisRowEditorLayered =
                      openEmailDraftRowId === row.id ||
                      openBodyEditorRowId === row.id ||
                      openDurationEditorRowId === row.id ||
                      openMeetingEditorRowId === row.id ||
                      closingRowEditor?.rowId === row.id;
                    const isThisRowEditorClosing = closingRowEditor?.rowId === row.id;
                    const isInlineEditorRevealActive = isThisRowEditorLayered;
                    const inspectorRevealDelay = (delayMs: number) => (isThisRowEditorLayered ? Math.min(delayMs, 120) : delayMs);
                    const isMissingRowHighlighted = missingFieldHighlights.rowIds.includes(row.id);
                    const inspectorTitleId = `row-inspector-title-${row.id}`;
                    const inspectorKindLabel =
                      rowKind === "email" ? "email" : rowKind === "meeting" ? "meeting" : "reminder";
                    const inspectorWidthClass =
                      rowKind === "email" || rowKind === "meeting"
                        ? "lg:w-[min(720px,calc(100vw-48px))]"
                        : "lg:w-[min(680px,calc(100vw-48px))]";
                    const handleInspectorKeyDown = (event: ReactKeyboardEvent<HTMLDivElement>) => {
                      if (!isThisRowEditorLayered) return;
                      if (event.key === "Escape") {
                        event.preventDefault();
                        closeRowEditorsAfterAnimation(() => {
                          if (rowKind === "meeting") {
                            setForcedOpenMeetingEditorRowIds((current) => current.filter((id) => id !== row.id));
                          }
                        });
                        return;
                      }
                      if (event.key !== "Tab") return;

                      const focusableElements = Array.from(
                        event.currentTarget.querySelectorAll<HTMLElement>(
                          "a[href], button:not([disabled]), textarea:not([disabled]), input:not([disabled]), select:not([disabled]), [tabindex]:not([tabindex='-1'])"
                        )
                      ).filter((element) => !element.hasAttribute("disabled") && element.offsetParent !== null);
                      if (focusableElements.length === 0) return;

                      const firstElement = focusableElements[0];
                      const lastElement = focusableElements[focusableElements.length - 1];
                      if (!firstElement || !lastElement) return;
                      if (event.shiftKey && document.activeElement === firstElement) {
                        event.preventDefault();
                        lastElement.focus();
                      } else if (!event.shiftKey && document.activeElement === lastElement) {
                        event.preventDefault();
                        firstElement.focus();
                      }
                    };
                    const reminderEditorIndentClass = "";
                    const reminderSurfaceInputClass = plansEditorInputClass;
                    const rowDateBasisControl =
                      !emailUsesTimingFields && row.rowType === "email" ? null : (
	                        <label className="inline-flex w-fit items-center gap-2 rounded-xl border border-slate-200 bg-white px-3 py-2 text-[12px] font-medium text-slate-600 shadow-sm">
                          <input
                            type="checkbox"
                            className="h-4 w-4 rounded border-slate-300 text-slate-700 focus:ring-slate-300"
                            checked={row.dateBasis === "today"}
                            onChange={(e) =>
                              updateRow(row.id, (current) => ({
                                ...current,
                                dateBasis: e.target.checked ? "today" : "event",
                              }))
                            }
                          />
                          <span>Base on Today&apos;s Date</span>
                          <span
                            className="relative inline-flex h-4 w-4 shrink-0 items-center justify-center rounded-full border border-slate-300 bg-white text-[10px] font-semibold text-slate-500"
                            onMouseEnter={() => setTodayBasisTooltipRowId(row.id)}
                            onMouseLeave={() => setTodayBasisTooltipRowId((current) => (current === row.id ? null : current))}
                          >
                            i
                            {todayBasisTooltipRowId === row.id ? (
                              <span className="pointer-events-none absolute left-full top-1/2 z-10 ml-2 w-44 -translate-y-1/2 rounded-lg border border-slate-200 bg-white px-2 py-1 text-left text-[11px] font-normal leading-4 text-slate-600 shadow-lg">
                                Uses today as the anchor for this row instead of the event date.
                              </span>
                            ) : null}
                          </span>
                        </label>
                      );
	                    const detailEditorIndentClass = "";
	                    const detailInputClass = plansEditorInputClass;
                    const reminderInlineEditor = isReminderInlineEditorRendered ? (
                      <AnimatedRowEditor
                        open={isReminderInlineEditorOpen}
                        activeLayer={isThisRowEditorLayered}
                        animate={!isThisRowEditorLayered}
                        openMs={INLINE_EDITOR_EXPAND_ANIMATION_MS}
                        closeMs={ROW_EDITOR_CLOSE_ANIMATION_MS}
                        className={isThisRowEditorLayered ? "min-h-0 flex-1" : ""}
                        contentClassName={isThisRowEditorLayered ? "flex h-full min-h-0 flex-col" : undefined}
                        innerClassName={isThisRowEditorLayered ? "flex h-full min-h-0 w-full flex-col" : undefined}
                      >
                        <div
                          data-inline-row-editor="true"
	                          ref={(node) => {
	                            rowEditorPanelRefs.current[row.id] = node;
	                          }}
	                          className={`${reminderEditorIndentClass} ${plansInspectorBodyClass} ${inactiveEditorDimTransitionClass} opacity-100`}
	                        >
                        <StaggeredInlineEditorItem active={isInlineEditorRevealActive} delayMs={0}>
                        <div className={plansInspectorSectionClass}>
                          <div className={plansInspectorSectionHeadingClass}>Schedule</div>
                          <div>{rowDateBasisControl}</div>
                        </div>
                        </StaggeredInlineEditorItem>
                        <StaggeredInlineEditorItem active={isInlineEditorRevealActive} delayMs={inspectorRevealDelay(360)}>
                        <div className={plansInspectorSectionClass}>
                          <div className={plansInspectorSectionHeadingClass}>Reminder content</div>
                          <div className="mb-2 text-sm font-medium text-slate-700">Reminder Title</div>
                          <div className="relative">
                            <div
                              aria-hidden="true"
                              className={`pointer-events-none absolute inset-0 flex items-center overflow-hidden rounded-[16px] px-3.5 py-2.5 text-[14px] text-slate-900 ${
                                focusedTitleInputId === reminderTitleInputId ? "opacity-0" : ""
                              }`}
                            >
                              <span className="block w-full truncate">
                                {row.title ? (
                                  renderTextWithBoldAnchors(row.title, {
                                    anchorClassName: "font-bold text-black",
                                    knownAnchorKeys,
                                  })
                                ) : (
                                  <span className="text-slate-400">Enter Reminder Title</span>
                                )}
                              </span>
                            </div>
                            <input
                              data-validation-field={`${row.id}:title`}
                              data-validation-field-scope="editor"
                              className={`${reminderSurfaceInputClass} caret-slate-950 ${getValidationFieldHighlightClass(row.id, "title")} ${
                                focusedTitleInputId === reminderTitleInputId
                                  ? "text-slate-950"
                                  : "text-transparent placeholder:text-transparent"
                              }`}
                              value={row.title}
                              placeholder="Enter Reminder Title"
                              onChange={(e) => {
                                clearValidationFieldHighlight(row.id, "title");
                                updateRow(row.id, (current) => ({ ...current, title: e.target.value }));
                              }}
                              onFocus={() => setFocusedTitleInputId(reminderTitleInputId)}
                              onBlur={() =>
                                setFocusedTitleInputId((current) => (current === reminderTitleInputId ? null : current))
                              }
                            />
                          </div>
                        </div>
                        </StaggeredInlineEditorItem>

                        <StaggeredInlineEditorItem active={isInlineEditorRevealActive} delayMs={inspectorRevealDelay(720)}>
                        <div className={plansInspectorSectionClass}>
                          <div className="mb-2 text-sm font-medium text-slate-700">Body</div>
                          <textarea
                            data-validation-field={`${row.id}:body`}
                            data-validation-field-scope="editor"
                            className={`${reminderSurfaceInputClass} min-h-36 resize-y ${getValidationFieldHighlightClass(row.id, "body")}`}
                            value={row.body ?? ""}
                            onChange={(e) => {
                              clearValidationFieldHighlight(row.id, "body");
                              updateRow(row.id, (current) => ({ ...current, body: e.target.value }));
                            }}
                          />
                        </div>
                        </StaggeredInlineEditorItem>

                        <StaggeredInlineEditorItem active={isInlineEditorRevealActive} delayMs={inspectorRevealDelay(1080)}>
                        <div className={plansInspectorSectionClass}>
                          <div className={plansInspectorSectionHeadingClass}>Schedule options</div>
                          {(() => {
                            const previewDurationItem = previewPlan.items.find((preview) => preview.id === row.id);
                            const computedStartDate =
                              (previewDurationItem ? getEffectivePreviewItemDate(previewDurationItem) : null) || todayYYYYMMDD();
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
                              <div className="space-y-3">
                                  <div className="grid gap-3 lg:grid-cols-[minmax(0,220px)_auto] lg:items-end">
                                  <div>
                                    <div className="mb-2 text-sm font-medium text-slate-700">Duration</div>
                                    <select
                                      className={`${reminderSurfaceInputClass} h-[28px]`}
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
                                  <label className="inline-flex h-8 w-fit items-center gap-1.5 rounded-[16px] border border-slate-200 bg-white px-2.5 text-[12px] font-medium text-slate-700 shadow-sm">
                                    <input
                                      type="checkbox"
                                      className="m-0 h-4 w-4"
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
                                {normalizedDuration?.useCustomEnd ? (
                                  <div className="grid gap-3 lg:grid-cols-2">
                                    <div>
                                      <div className="mb-2 text-sm font-medium text-slate-700">End Date</div>
                                      <input
                                        type="date"
                                        className={reminderSurfaceInputClass}
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
                                      <div className="mb-2 text-sm font-medium text-slate-700">End Time</div>
                                      <input
                                        type="time"
                                        className={reminderSurfaceInputClass}
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
                                  </div>
                                ) : null}
                              </div>
                            );
                          })()}
                        </div>
                        </StaggeredInlineEditorItem>

	                        <StaggeredInlineEditorItem active={isInlineEditorRevealActive} delayMs={inspectorRevealDelay(1440)}>
	                          <div className={plansInspectorSectionClass}>
	                            {renderDynamicFieldsSection({ inlineEditor: true })}
	                          </div>
	                        </StaggeredInlineEditorItem>
                      </div>
	                        <div className={plansInspectorFooterClass}>
                          <button
                            type="button"
                            onClick={() => {
                              closeRowEditorsAfterAnimation();
                            }}
	                            className={`${plansPrimaryButtonClass} !h-[40px] min-w-[96px]`}
                          >
                            Done
                          </button>
                        </div>
                      </AnimatedRowEditor>
                    ) : null;
                    const emailInlineEditorRendered = isEmailEditorOpen || isClosingEmailEditor;
                    const emailInlineEditor = emailInlineEditorRendered ? (
	                      <AnimatedRowEditor
	                        open={isEmailEditorOpen}
	                        activeLayer={isThisRowEditorLayered}
	                        animate={!isThisRowEditorLayered}
	                        openMs={INLINE_EDITOR_EXPAND_ANIMATION_MS}
	                        closeMs={ROW_EDITOR_CLOSE_ANIMATION_MS}
	                        className={isThisRowEditorLayered ? "min-h-0 flex-1" : ""}
	                        contentClassName={isThisRowEditorLayered ? "flex h-full min-h-0 flex-col" : undefined}
	                        innerClassName={isThisRowEditorLayered ? "flex h-full min-h-0 w-full flex-col" : undefined}
	                      >
                        <div
                          data-inline-row-editor="true"
	                          ref={(node) => {
	                            rowEditorPanelRefs.current[row.id] = node;
	                          }}
	                          className={`${detailEditorIndentClass} ${plansInspectorBodyClass} ${inactiveEditorDimTransitionClass} opacity-100`}
	                        >
                          {emailUsesTimingFields ? (
                            <StaggeredInlineEditorItem active={isInlineEditorRevealActive} delayMs={0}>
                              <div className={plansInspectorSectionClass}>
                                <div className={plansInspectorSectionHeadingClass}>Schedule</div>
                                <div>{rowDateBasisControl}</div>
                              </div>
                            </StaggeredInlineEditorItem>
                          ) : null}
                          <StaggeredInlineEditorItem active={isInlineEditorRevealActive} delayMs={inspectorRevealDelay(360)}>
                            <div className={plansInspectorSectionClass}>
                              <div className={plansInspectorSectionHeadingClass}>Message</div>
                              <div className="mb-2 text-sm font-medium text-slate-700">Subject</div>
                              <input
                                data-validation-field={`${row.id}:emailSubject`}
                                data-validation-field-scope="editor"
                                className={`${detailInputClass} ${getValidationFieldHighlightClass(row.id, "emailSubject")}`}
                                value={normalizeEmailDraft(row.emailDraft).subject}
                                placeholder="Email Subject Line"
                                onChange={(e) => {
                                  clearValidationFieldHighlight(row.id, "emailSubject");
                                  updateRow(row.id, (current) => ({
                                    ...current,
                                    emailDraft: { ...normalizeEmailDraft(current.emailDraft), subject: e.target.value },
                                  }))
                                }}
                              />
                            </div>
                          </StaggeredInlineEditorItem>
                          <StaggeredInlineEditorItem active={isInlineEditorRevealActive} delayMs={inspectorRevealDelay(720)}>
                            <div className={plansInspectorSectionClass}>
                              <div className={plansInspectorSectionHeadingClass}>Recipients</div>
                              {(() => {
                                const normalizedDraft = normalizeEmailDraft(row.emailDraft);
                                const showCc = Boolean(emailFieldVisibility[row.id]?.cc || normalizedDraft.cc.length);
                                const showBcc = Boolean(emailFieldVisibility[row.id]?.bcc || normalizedDraft.bcc.length);

                                return (
                                  <div className="space-y-3">
                                    <div>
                                      <div className="mb-2 flex items-center justify-between gap-3">
                                        <div className="text-sm font-medium text-slate-700">To</div>
                                        <div className="flex items-center gap-3 text-xs font-medium text-amber-900">
                                          <button
                                            type="button"
                                            onClick={() => openRecipientGroupsModal({ rowId: row.id, field: "email_to" })}
                                            className="rounded-full border border-slate-300 bg-white px-2.5 py-1 text-xs font-medium text-slate-700 hover:bg-slate-50"
                                          >
                                            Group
                                          </button>
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
                                      <div data-validation-field={`${row.id}:emailTo`} data-validation-field-scope="editor">
                                        <EmailTokensInput
                                          label=""
                                          values={normalizedDraft.to}
                                          recipientGroups={recipientGroups}
                                          placeholder="Type an email and press Enter or comma"
                                          hasError={isValidationFieldHighlighted(row.id, "emailTo")}
                                          onChange={(nextValues) => {
                                            clearValidationFieldHighlight(row.id, "emailTo");
                                            updateRow(row.id, (current) => ({
                                              ...current,
                                              emailDraft: { ...normalizeEmailDraft(current.emailDraft), to: nextValues },
                                            }));
                                          }}
                                        />
                                      </div>
                                    </div>
                                    {showCc ? (
                                      <div>
                                        <div data-validation-field={`${row.id}:emailCc`} data-validation-field-scope="editor">
                                          <EmailTokensInput
                                            label="CC"
                                            values={normalizedDraft.cc}
                                            recipientGroups={recipientGroups}
                                            placeholder="Add CC recipients"
                                            hasError={isValidationFieldHighlighted(row.id, "emailCc")}
                                            onChange={(nextValues) => {
                                              clearValidationFieldHighlight(row.id, "emailCc");
                                              updateRow(row.id, (current) => ({
                                                ...current,
                                                emailDraft: { ...normalizeEmailDraft(current.emailDraft), cc: nextValues },
                                              }));
                                            }}
                                          />
                                        </div>
                                      </div>
                                    ) : null}
                                    {showBcc ? (
                                      <div>
                                        <div data-validation-field={`${row.id}:emailBcc`} data-validation-field-scope="editor">
                                          <EmailTokensInput
                                            label="BCC"
                                            values={normalizedDraft.bcc}
                                            recipientGroups={recipientGroups}
                                            placeholder="Add BCC recipients"
                                            hasError={isValidationFieldHighlighted(row.id, "emailBcc")}
                                            onChange={(nextValues) => {
                                              clearValidationFieldHighlight(row.id, "emailBcc");
                                              updateRow(row.id, (current) => ({
                                                ...current,
                                                emailDraft: { ...normalizeEmailDraft(current.emailDraft), bcc: nextValues },
                                              }));
                                            }}
                                          />
                                        </div>
                                      </div>
                                    ) : null}
                                  </div>
                                );
                              })()}
                            </div>
                          </StaggeredInlineEditorItem>
                          <StaggeredInlineEditorItem active={isInlineEditorRevealActive} delayMs={inspectorRevealDelay(1080)}>
                            <div className={plansInspectorSectionClass}>
                              <div className="mb-2 text-sm font-medium text-slate-700">Message</div>
                              <textarea
                                data-validation-field={`${row.id}:emailBody`}
                                data-validation-field-scope="editor"
                                className={`${detailInputClass} min-h-32 resize-y ${getValidationFieldHighlightClass(row.id, "emailBody")}`}
                                value={normalizeEmailDraft(row.emailDraft).body}
                                onChange={(e) => {
                                  clearValidationFieldHighlight(row.id, "emailBody");
                                  updateRow(row.id, (current) => ({
                                    ...current,
                                    emailDraft: { ...normalizeEmailDraft(current.emailDraft), body: e.target.value },
                                  }))
                                }}
                              />
                            </div>
                          </StaggeredInlineEditorItem>
	                          <StaggeredInlineEditorItem active={isInlineEditorRevealActive} delayMs={inspectorRevealDelay(1440)}>
	                            <div className={plansInspectorSectionClass}>
	                              {renderDynamicFieldsSection({ inlineEditor: true })}
	                            </div>
	                          </StaggeredInlineEditorItem>
                        </div>
	                        <div className={plansInspectorFooterClass}>
                            <button
                              type="button"
                              onClick={() => {
                                collapseEmptyEmailFields(row.id, row.emailDraft);
                                closeRowEditorsAfterAnimation();
                              }}
	                            className={`${plansPrimaryButtonClass} !h-[40px] min-w-[96px]`}
                            >
                              Done
                            </button>
                          </div>
                      </AnimatedRowEditor>
                    ) : null;
                    const meetingInlineEditorRendered = isMeetingEditorOpen || isClosingMeetingEditor;
                    const meetingInlineEditor = meetingInlineEditorRendered ? (
	                      <AnimatedRowEditor
	                        open={isMeetingEditorOpen}
	                        activeLayer={isThisRowEditorLayered}
	                        animate={!isThisRowEditorLayered}
	                        openMs={INLINE_EDITOR_EXPAND_ANIMATION_MS}
	                        closeMs={ROW_EDITOR_CLOSE_ANIMATION_MS}
	                        className={isThisRowEditorLayered ? "min-h-0 flex-1" : ""}
	                        contentClassName={isThisRowEditorLayered ? "flex h-full min-h-0 flex-col" : undefined}
	                        innerClassName={isThisRowEditorLayered ? "flex h-full min-h-0 w-full flex-col" : undefined}
	                      >
                        <div
                          data-inline-row-editor="true"
	                          ref={(node) => {
	                            rowEditorPanelRefs.current[row.id] = node;
	                          }}
	                          className={`${detailEditorIndentClass} ${plansInspectorBodyClass} ${inactiveEditorDimTransitionClass} opacity-100`}
	                        >
                          <StaggeredInlineEditorItem active={isInlineEditorRevealActive} delayMs={0}>
                            <div className={plansInspectorSectionClass}>
                              <div className={plansInspectorSectionHeadingClass}>Schedule</div>
                              <div>{rowDateBasisControl}</div>
                            </div>
                          </StaggeredInlineEditorItem>
                          <StaggeredInlineEditorItem active={isInlineEditorRevealActive} delayMs={inspectorRevealDelay(360)}>
                            <div className={plansInspectorSectionClass}>
                              <div className={plansInspectorSectionHeadingClass}>Meeting content</div>
                              <div className="mb-2 text-sm font-medium text-slate-700">Meeting Title</div>
                              <div className="relative">
                                <div
                                  aria-hidden="true"
                                  className={`pointer-events-none absolute inset-0 flex items-center overflow-hidden rounded-[16px] px-3.5 py-2.5 text-[14px] text-slate-900 ${
                                    focusedTitleInputId === meetingTitleInputId ? "opacity-0" : ""
                                  }`}
                                >
                                  <span className="block w-full truncate">
                                    {row.title ? (
                                      renderTextWithBoldAnchors(row.title, {
                                        anchorClassName: "font-bold text-black",
                                        knownAnchorKeys,
                                      })
                                    ) : (
                                      <span className="text-slate-400">Meeting Title</span>
                                    )}
                                  </span>
                                </div>
                                <input
                                  data-validation-field={`${row.id}:title`}
                                  data-validation-field-scope="editor"
                                  className={`${detailInputClass} caret-slate-950 ${getValidationFieldHighlightClass(row.id, "title")} ${
                                    focusedTitleInputId === meetingTitleInputId
                                      ? "text-slate-950"
                                      : "text-transparent placeholder:text-transparent"
                                  }`}
                                  value={row.title}
                                  placeholder="Meeting Title"
                                  onChange={(e) => {
                                    clearValidationFieldHighlight(row.id, "title");
                                    updateRow(row.id, (current) => ({ ...current, title: e.target.value }));
                                  }}
                                  onFocus={() => setFocusedTitleInputId(meetingTitleInputId)}
                                  onBlur={() =>
                                    setFocusedTitleInputId((current) => (current === meetingTitleInputId ? null : current))
                                  }
                                />
                              </div>
                            </div>
                          </StaggeredInlineEditorItem>
                          <StaggeredInlineEditorItem active={isInlineEditorRevealActive} delayMs={inspectorRevealDelay(720)}>
                            <div className={`${plansInspectorSectionClass} space-y-3`}>
                              <div className={plansInspectorSectionHeadingClass}>Attendees and location</div>
                              <div>
                                <div className="mb-2 flex items-center justify-between gap-3">
                                  <div className="text-sm font-medium text-slate-700">
                                    Attendees
                                    {meetingErrors?.attendees ? <span className="ml-1 text-red-500">*</span> : null}
                                  </div>
                                  <button
                                    type="button"
                                    onClick={() => openRecipientGroupsModal({ rowId: row.id, field: "meeting_attendees" })}
                                    className="rounded-full border border-slate-300 bg-white px-2.5 py-1 text-xs font-medium text-slate-700 hover:bg-slate-50"
                                  >
                                    Group
                                  </button>
                                </div>
                                <div data-validation-field={`${row.id}:meetingAttendees`} data-validation-field-scope="editor">
                                  <EmailTokensInput
                                    label=""
                                    values={normalizeMeetingDraft(row.meetingDraft)?.attendees ?? []}
                                    recipientGroups={recipientGroups}
                                    onChange={(nextValues) => {
                                      clearValidationFieldHighlight(row.id, "meetingAttendees");
                                      updateRow(row.id, (current) => ({
                                        ...current,
                                        meetingDraft: { ...normalizeMeetingDraft(current.meetingDraft), attendees: nextValues },
                                      }));
                                    }}
                                    placeholder="Add attendee emails"
                                    hasError={Boolean(meetingErrors?.attendees) || isValidationFieldHighlighted(row.id, "meetingAttendees")}
                                  />
                                </div>
                              </div>
                              <div className="grid gap-3 lg:grid-cols-2">
                                <div>
                                  <div className="mb-2 text-sm font-medium text-slate-700">Location</div>
                                  {(() => {
                                    const normalizedMeetingDraft = normalizeMeetingDraft(row.meetingDraft);
                                    const isGoogleProviderActive = activeAccountProvider === "gmail";
                                    const isProviderManagedMeeting =
                                      normalizedMeetingDraft?.teamsMeeting || (isGoogleProviderActive && normalizedMeetingDraft?.addGoogleMeet);
                                    const locationValue = getMeetingLocationValue(normalizedMeetingDraft, activeAccountProvider);
                                    return (
                                      <input
                                        data-validation-field={`${row.id}:meetingLocation`}
                                        data-validation-field-scope="editor"
                                        className={`${detailInputClass} ${getValidationFieldHighlightClass(row.id, "meetingLocation")} ${
                                          isProviderManagedMeeting ? "bg-gray-100 text-gray-600" : ""
                                        }`}
                                        value={locationValue}
                                        readOnly={Boolean(isProviderManagedMeeting)}
                                        onChange={(e) => {
                                          clearValidationFieldHighlight(row.id, "meetingLocation");
                                          updateRow(row.id, (current) => ({
                                            ...current,
                                            meetingDraft: { ...normalizeMeetingDraft(current.meetingDraft), location: e.target.value },
                                          }))
                                        }}
                                      />
                                    );
                                  })()}
                                </div>
                                <div className="space-y-3">
                                  <div>
                                    <div className="mb-2 text-sm font-medium text-slate-700">
                                      Meeting Duration
                                      {meetingErrors?.duration ? <span className="ml-1 text-red-500">*</span> : null}
                                    </div>
                                    <select
                                      className={`${detailInputClass} h-[28px] ${
                                        meetingErrors?.duration ? "border-red-400 ring-1 ring-red-100" : ""
                                      }`}
                                      value={
                                        normalizeMeetingDraft(row.meetingDraft)?.isAllDay
                                          ? "__all_day__"
                                          : normalizeMeetingDraft(row.meetingDraft)?.useCustomEnd
                                            ? "custom"
                                            : String(normalizeMeetingDraft(row.meetingDraft)?.durationMinutes ?? 30)
                                      }
                                      onChange={(e) =>
                                        e.target.value === "custom"
                                          ? (() => {
                                              const previewMeetingItem = previewPlan.items.find((preview) => preview.id === row.id);
                                              const effectiveStartDate =
                                                (previewMeetingItem ? getEffectivePreviewItemDate(previewMeetingItem) : null) || todayYYYYMMDD();
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
                                      {normalizeMeetingDraft(row.meetingDraft)?.isAllDay ? (
                                        <option value="__all_day__">All Day</option>
                                      ) : null}
                                      {MEETING_DURATION_OPTIONS.map((option) => (
                                        <option key={option.value} value={option.value}>
                                          {option.label}
                                        </option>
                                      ))}
                                    </select>
                                  </div>
                                  <label className="inline-flex h-8 w-fit items-center gap-1.5 rounded-[16px] border border-slate-200 bg-white px-2.5 text-[12px] font-medium text-slate-700 shadow-sm">
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
                                    <span>All day</span>
                                  </label>
                                </div>
                              </div>
                            </div>
                          </StaggeredInlineEditorItem>
                          <StaggeredInlineEditorItem active={isInlineEditorRevealActive} delayMs={inspectorRevealDelay(1080)}>
                            <div className={`${plansInspectorSectionClass} space-y-3`}>
                              <div className={plansInspectorSectionHeadingClass}>Meeting options</div>
                              {activeAccountProvider === "gmail" ? (
                                <label className="inline-flex items-center gap-2 text-sm text-slate-700">
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
                                </label>
                              ) : (
                                <label className="inline-flex items-center gap-2 text-sm text-slate-700">
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
                                </label>
                              )}
                              <div>
                                <div className="mb-2 text-sm font-medium text-slate-700">Body</div>
                                {(() => {
                                  const normalizedMeetingDraft = normalizeMeetingDraft(row.meetingDraft);
                                  const isTeamsManaged = activeAccountProvider !== "gmail" && Boolean(normalizedMeetingDraft?.teamsMeeting);
                                  const isGoogleMeetManaged = activeAccountProvider === "gmail" && Boolean(normalizedMeetingDraft?.addGoogleMeet);
                                  const isProviderManagedBody = isTeamsManaged || isGoogleMeetManaged;
                                  const providerManagedBodyText = isTeamsManaged
                                    ? "Teams join info will appear here after the Outlook event is created."
                                    : isGoogleMeetManaged
                                      ? "Google Meet link will be generated after export."
                                      : "";

                                  return (
                                    <textarea
                                      data-validation-field={`${row.id}:body`}
                                      data-validation-field-scope="editor"
                                      className={`${detailInputClass} min-h-32 resize-y ${getValidationFieldHighlightClass(row.id, "body")} ${
                                        isProviderManagedBody ? "bg-gray-100 text-gray-500" : ""
                                      }`}
                                      value={isProviderManagedBody ? providerManagedBodyText : row.body ?? ""}
                                      readOnly={isProviderManagedBody}
                                      onChange={(e) => {
                                        clearValidationFieldHighlight(row.id, "body");
                                        updateRow(row.id, (current) => ({ ...current, body: e.target.value }));
                                      }}
                                    />
                                  );
                                })()}
                              </div>
                              {normalizeMeetingDraft(row.meetingDraft)?.useCustomEnd &&
                              !normalizeMeetingDraft(row.meetingDraft)?.isAllDay ? (
                                <div className="grid gap-3 lg:grid-cols-2">
                                  <div>
                                    <div className="mb-2 text-sm font-medium text-slate-700">
                                      End Date
                                      {meetingErrors?.duration ? <span className="ml-1 text-red-500">*</span> : null}
                                    </div>
                                    <input
                                      type="date"
                                      className={`${detailInputClass} ${
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
                                  <div>
                                    <div className="mb-2 text-sm font-medium text-slate-700">
                                      End Time
                                      {meetingErrors?.duration ? <span className="ml-1 text-red-500">*</span> : null}
                                    </div>
                                    <input
                                      type="time"
                                      className={`${detailInputClass} ${
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
                                </div>
                              ) : null}
                            </div>
                          </StaggeredInlineEditorItem>
	                          <StaggeredInlineEditorItem active={isInlineEditorRevealActive} delayMs={inspectorRevealDelay(1440)}>
	                            <div className={plansInspectorSectionClass}>
	                              {renderDynamicFieldsSection({ inlineEditor: true })}
	                            </div>
	                          </StaggeredInlineEditorItem>
                        </div>
	                          <div className={plansInspectorFooterClass}>
                            <button
                              type="button"
                              onClick={() => {
                                closeRowEditorsAfterAnimation(() => {
                                  setForcedOpenMeetingEditorRowIds((current) => current.filter((id) => id !== row.id));
                                });
                              }}
		                              className={`${plansPrimaryButtonClass} !h-[40px] min-w-[96px]`}
                            >
                              Done
                            </button>
                          </div>
                      </AnimatedRowEditor>
                    ) : null;
		                    const rowTimingChipClass =
				                      "inline-flex h-[36px] items-center justify-center whitespace-nowrap rounded-[9px] border border-slate-200 bg-white text-center text-[13px] font-medium leading-none text-slate-600 shadow-sm";
		                    const inspectorTimingChipClass =
				                      "inline-flex h-[34px] items-center justify-center whitespace-nowrap rounded-xl border border-slate-200 bg-white px-3 text-center text-[13px] font-medium leading-none text-slate-600 shadow-sm";
                    const rowTimeZone = normalizeOutlookTimeZone(row.timeZone || eventTimeZone);
                    const normalizedTimeZoneSearch = timeZoneSearch.trim().toLowerCase();
                    const filteredTimeZoneOptions = normalizedTimeZoneSearch
                      ? timeZoneOptions.filter((timeZone) =>
                          `${timeZone.label} ${timeZone.value}`.toLowerCase().includes(normalizedTimeZoneSearch)
                        )
                      : timeZoneOptions;
                    const rowTimeZoneControl = (
                      <span className="relative inline-flex">
                        <button
                          type="button"
                          data-no-row-drag="true"
                          aria-label="Set time zone"
                          title="Set time zone"
                          onClick={() => {
                            setTimeZoneSearch("");
                            setOpenTimeZoneRowId((current) => (current === row.id ? null : row.id));
                          }}
                          className="group/timezone-globe inline-flex h-[18px] w-[18px] items-center justify-center rounded-full text-slate-500 transition hover:bg-slate-100 hover:text-slate-700"
                        >
                          <svg
                            viewBox="0 0 24 24"
                            aria-hidden="true"
                            className="h-3.5 w-3.5"
                            fill="none"
                            stroke="currentColor"
                            strokeWidth="1.9"
                            strokeLinecap="round"
                            strokeLinejoin="round"
                          >
                            <circle cx="12" cy="12" r="9" />
                            <path d="M3 12h18" />
                            <path d="M12 3c2.35 2.46 3.55 5.46 3.55 9S14.35 18.54 12 21" />
                            <path d="M12 3c-2.35 2.46-3.55 5.46-3.55 9S9.65 18.54 12 21" />
                          </svg>
                          <span className="pointer-events-none absolute bottom-full left-1/2 z-40 mb-2 hidden w-24 -translate-x-1/2 rounded-lg border border-slate-200 bg-white px-2 py-1 text-center text-[11px] font-medium text-slate-600 shadow-lg group-hover/timezone-globe:block">
                            Set time zone
                          </span>
                        </button>
                        {openTimeZoneRowId === row.id ? (
	                          <div className="absolute right-0 top-[calc(100%+0.5rem)] z-[170] w-72 rounded-2xl border border-slate-600 bg-slate-800 p-2 text-left shadow-[0_18px_36px_-16px_rgba(15,23,42,0.65)]">
                            <div className="mb-2 px-2 text-[11px] font-semibold text-slate-200">
                              {getOutlookTimeZoneLabel(rowTimeZone)}
                            </div>
                            <input
                              type="search"
                              autoFocus
                              data-no-row-drag="true"
                              value={timeZoneSearch}
                              onChange={(e) => setTimeZoneSearch(e.target.value)}
                              onKeyDown={(e) => {
                                if (e.key === "Escape") {
                                  e.preventDefault();
                                  setOpenTimeZoneRowId(null);
                                  setTimeZoneSearch("");
                                }
                              }}
                              placeholder="Search time zones"
                              className="mb-2 w-full rounded-xl border border-slate-600 bg-slate-950 px-3 py-2 text-[12px] text-white placeholder:text-slate-400 focus:border-blue-300 focus:outline-none focus:ring-2 focus:ring-blue-400/30"
                            />
                            <div className="max-h-64 overflow-y-auto rounded-xl border border-slate-600 bg-slate-900 p-1">
                              {filteredTimeZoneOptions.length > 0 ? (
                              filteredTimeZoneOptions.map((timeZone) => {
                                const isSelectedTimeZone = timeZone.value === rowTimeZone;
                                return (
                                <button
                                  key={timeZone.value}
                                  type="button"
                                  data-no-row-drag="true"
                                  onClick={() => {
                                    updateRow(row.id, (current) => ({ ...current, timeZone: timeZone.value }));
                                    setOpenTimeZoneRowId(null);
                                    setTimeZoneSearch("");
                                  }}
                                  className={`block w-full rounded-lg px-2 py-1.5 text-left text-[12px] leading-5 transition ${
                                    isSelectedTimeZone
                                      ? "bg-blue-500/30 text-white"
                                      : "text-slate-200 hover:bg-slate-700"
                                  }`}
                                >
                                  <span className="block font-medium">{timeZone.label}</span>
                                  <span className="mt-0.5 block text-[10px] text-slate-400">{timeZone.value}</span>
                                </button>
                                );
                              })
                              ) : (
                                <div className="px-3 py-4 text-center text-[12px] text-slate-400">
                                  No time zones found
                                </div>
                              )}
                            </div>
                          </div>
                        ) : null}
                      </span>
                    );
                    const inspectorSubtitle =
                      row.title ||
                      (rowKind === "email"
                        ? "Email Subject Line"
                        : rowKind === "meeting"
                          ? "Meeting Title"
                          : "Enter Reminder Title");
                    const inspectorDotClass = `h-3 w-3 shrink-0 rounded-full ${
                      rowKind === "email" ? "bg-green-500" : rowKind === "meeting" ? "bg-violet-500" : "bg-blue-500"
                    }`;
                    const handleInspectorClose = () => {
                      closeRowEditorsAfterAnimation(() => {
                        if (rowKind === "meeting") {
                          setForcedOpenMeetingEditorRowIds((current) => current.filter((id) => id !== row.id));
                        }
                      });
                    };
	                    const renderInspectorTimingChips = () => (
	                      <>
	                        <span className={`${inspectorTimingChipClass} min-w-[92px]`}>
	                          {formatOffsetLabel(row.offsetDays, {
	                            relativeToToday: noEventDate,
	                            dateBasis: row.dateBasis,
	                          })}
	                        </span>
	                        <span className={`${inspectorTimingChipClass} min-w-[86px] gap-1 px-2`}>
	                          {row.meetingDraft?.isAllDay || row.durationDraft?.isAllDay
	                            ? "All day"
	                            : getReminderTimeDisplayValue(row.reminderTime ?? "") || "No time"}
	                          {rowTimeZoneControl}
	                        </span>
	                      </>
	                    );
	                    const emailProviderNote =
	                      row.rowType === "email" ? getBuilderEmailModeMessage(appSettings.emailHandlingMode) : null;
	                    const collapsedOffsetChipClass = `${rowTimingChipClass} w-full max-w-full min-[900px]:w-[136px]`;
	                    const collapsedTimeChipClass = `${rowTimingChipClass} w-full max-w-full tabular-nums min-[900px]:w-[112px]`;
	                    const collapsedOffsetControl = (
	                      <div data-workflow-row-part="offset" className="min-w-0 max-w-full min-[900px]:justify-self-start">
	                        {editingOffsetRowId === row.id ? (
	                          <div className={`${collapsedOffsetChipClass} gap-2 px-2 py-1`}>
	                            <input
	                              ref={(node) => {
	                                builderOffsetInputRefs.current[row.id] = node;
	                              }}
	                              type="text"
	                              inputMode="numeric"
	                              className="min-w-0 flex-1 border-0 bg-transparent px-1 text-slate-600 focus:outline-none focus:ring-0"
	                              value={offsetDrafts[row.id] ?? (row.offsetDays == null ? "" : String(row.offsetDays))}
	                              onChange={(e) =>
	                                setOffsetDrafts((current) => ({
	                                  ...current,
	                                  [row.id]: e.target.value,
	                                }))
	                              }
	                              onFocus={(e) => {
	                                e.currentTarget.select();
	                              }}
	                              onBlur={() => {
	                                commitOffsetDraft(row.id);
	                              }}
	                              onKeyDown={(e) => {
	                                if (e.key === "ArrowUp") {
	                                  e.preventDefault();
	                                  nudgeOffsetDraft(row.id, 1);
	                                  return;
	                                }
	                                if (e.key === "ArrowDown") {
	                                  e.preventDefault();
	                                  nudgeOffsetDraft(row.id, -1);
	                                  return;
	                                }
	                                if (e.key === "Enter") {
	                                  e.preventDefault();
	                                  commitOffsetDraft(row.id);
	                                  focusBuilderTimeInput(row.id);
	                                  return;
	                                }
	                                if (e.key === "Tab" && !e.shiftKey) {
	                                  e.preventDefault();
	                                  commitOffsetDraft(row.id);
	                                  focusBuilderTimeInput(row.id);
	                                  return;
	                                }
	                                if (e.key === "Tab" && e.shiftKey) {
	                                  e.preventDefault();
	                                  commitOffsetDraft(row.id);
	                                  focusCollapsedTitleInput(row.id);
	                                }
	                              }}
	                              autoFocus
	                            />
	                            <div className="flex flex-col gap-0.5">
	                              <button
	                                type="button"
	                                data-no-row-drag="true"
	                                tabIndex={-1}
	                                aria-label="Increase offset"
	                                className="flex h-4 w-4 items-center justify-center rounded-full text-slate-400 transition hover:bg-slate-100 hover:text-slate-700"
	                                onMouseDown={(e) => e.preventDefault()}
	                                onClick={() => nudgeOffsetDraft(row.id, 1)}
	                              >
	                                <svg
	                                  viewBox="0 0 16 16"
	                                  aria-hidden="true"
	                                  className="h-3 w-3"
	                                  fill="none"
	                                  stroke="currentColor"
	                                  strokeWidth="1.8"
	                                  strokeLinecap="round"
	                                  strokeLinejoin="round"
	                                >
	                                  <path d="M4 10l4-4 4 4" />
	                                </svg>
	                              </button>
	                              <button
	                                type="button"
	                                data-no-row-drag="true"
	                                tabIndex={-1}
	                                aria-label="Decrease offset"
	                                className="flex h-4 w-4 items-center justify-center rounded-full text-slate-400 transition hover:bg-slate-100 hover:text-slate-700"
	                                onMouseDown={(e) => e.preventDefault()}
	                                onClick={() => nudgeOffsetDraft(row.id, -1)}
	                              >
	                                <svg
	                                  viewBox="0 0 16 16"
	                                  aria-hidden="true"
	                                  className="h-3 w-3"
	                                  fill="none"
	                                  stroke="currentColor"
	                                  strokeWidth="1.8"
	                                  strokeLinecap="round"
	                                  strokeLinejoin="round"
	                                >
	                                  <path d="M4 6l4 4 4-4" />
	                                </svg>
	                              </button>
	                            </div>
	                          </div>
	                        ) : (
	                          <span className={`${collapsedOffsetChipClass} gap-1 py-0.5 pl-2 pr-1 transition-colors hover:border-slate-300 hover:text-slate-600`}>
	                            <button
	                              type="button"
	                              data-no-row-drag="true"
	                              className="min-w-0 flex-1 bg-transparent px-0 text-center focus:outline-none"
	                              onFocus={() => beginEditingOffsetForRow(row.id, row.offsetDays)}
	                              onClick={() => beginEditingOffsetForRow(row.id, row.offsetDays)}
	                            >
	                              {formatOffsetLabel(row.offsetDays, {
	                                relativeToToday: noEventDate,
	                                dateBasis: row.dateBasis,
	                              })}
	                            </button>
	                            <span className="flex flex-col gap-0.5">
	                              <button
	                                type="button"
	                                data-no-row-drag="true"
	                                tabIndex={-1}
	                                aria-label="Increase offset"
	                                className="flex h-3.5 w-3.5 items-center justify-center rounded-full text-slate-400 transition hover:bg-slate-100 hover:text-slate-700"
	                                onMouseDown={(e) => e.preventDefault()}
	                                onClick={() => nudgeRowOffset(row.id, 1)}
	                              >
	                                <svg
	                                  viewBox="0 0 16 16"
	                                  aria-hidden="true"
	                                  className="h-2.5 w-2.5"
	                                  fill="none"
	                                  stroke="currentColor"
	                                  strokeWidth="1.9"
	                                  strokeLinecap="round"
	                                  strokeLinejoin="round"
	                                >
	                                  <path d="M4 10l4-4 4 4" />
	                                </svg>
	                              </button>
	                              <button
	                                type="button"
	                                data-no-row-drag="true"
	                                tabIndex={-1}
	                                aria-label="Decrease offset"
	                                className="flex h-3.5 w-3.5 items-center justify-center rounded-full text-slate-400 transition hover:bg-slate-100 hover:text-slate-700"
	                                onMouseDown={(e) => e.preventDefault()}
	                                onClick={() => nudgeRowOffset(row.id, -1)}
	                              >
	                                <svg
	                                  viewBox="0 0 16 16"
	                                  aria-hidden="true"
	                                  className="h-2.5 w-2.5"
	                                  fill="none"
	                                  stroke="currentColor"
	                                  strokeWidth="1.9"
	                                  strokeLinecap="round"
	                                  strokeLinejoin="round"
	                                >
	                                  <path d="M4 6l4 4 4-4" />
	                                </svg>
	                              </button>
	                            </span>
	                          </span>
	                        )}
	                      </div>
	                    );
	                    const collapsedTimeControl = (
	                      <div data-workflow-row-part="time" className="relative min-w-0 max-w-full min-[900px]:justify-self-start">
	                        {meetingErrors?.time ? <div className="absolute -top-[18px] left-0 text-[12px] font-medium text-red-500">Time *</div> : null}
	                        {(() => {
	                          const shouldWrapTimeField =
	                            Boolean(row.reminderTime?.includes("[")) || Boolean(row.reminderTime?.includes("\n"));
	                          const isAnchorTimeValue = isReminderTimeAnchorValue(row.reminderTime ?? "");
	                          const literalTimeEditorValue =
	                            focusedTimeInputRowId === row.id
	                              ? (timeInputDrafts[row.id] ?? buildReminderTimeMaskedValue(row.reminderTime ?? ""))
	                              : getReminderTimeDisplayValue(row.reminderTime ?? "");

	                          if (row.meetingDraft?.isAllDay || row.durationDraft?.isAllDay) {
	                            return <span className={`${collapsedTimeChipClass} px-2`}>All day</span>;
	                          }

	                          if (!shouldWrapTimeField && !isAnchorTimeValue) {
	                            return (
	                              <span className={`${collapsedTimeChipClass} gap-0 px-2`}>
	                                <input
	                                  ref={(node) => {
	                                    builderTimeInputRefs.current[row.id] = node;
	                                  }}
	                                  data-validation-field={`${row.id}:reminderTime`}
	                                  type="text"
	                                  inputMode="text"
	                                  className={`w-[56px] min-w-0 border-0 bg-transparent p-0 text-center placeholder:text-slate-400 focus:outline-none focus:ring-0 ${getValidationFieldHighlightClass(row.id, "reminderTime")}`}
	                                  value={literalTimeEditorValue}
	                                  placeholder={focusedTimeInputRowId === row.id ? REMINDER_TIME_INPUT_MASK : "No time"}
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
	                                    const nextCursor = findReminderTimeCursorFromMeaningfulCount(nextMaskedValue, meaningfulCount);

	                                    setTimeInputDrafts((current) => ({
	                                      ...current,
	                                      [row.id]: nextMaskedValue,
	                                    }));

	                                    requestAnimationFrame(() => {
	                                      const input = builderTimeInputRefs.current[row.id];
	                                      if (input && input instanceof HTMLInputElement && document.activeElement === input) {
	                                        input.setSelectionRange(nextCursor, nextCursor);
	                                      }
	                                    });
	                                  }}
	                                  onBlur={(e) => {
	                                    clearValidationFieldHighlight(row.id, "reminderTime");
	                                    const nextValue = normalizeReminderTimeInput(e.target.value);
	                                    updateRow(row.id, (current) => ({
	                                      ...current,
	                                      reminderTime: nextValue,
	                                    }));
	                                    setFocusedTimeInputRowId((current) => (current === row.id ? null : current));
	                                    setTimeInputDrafts((current) => clearReminderTimeDraft(current, row.id));
	                                  }}
	                                />
	                                {rowTimeZoneControl}
	                              </span>
	                            );
	                          }

	                          return (
	                            <span className={`${collapsedTimeChipClass} gap-0 px-2 py-1.5`}>
	                              <textarea
	                                ref={(node) => {
	                                  builderTimeInputRefs.current[row.id] = node;
	                                }}
	                                data-validation-field={`${row.id}:reminderTime`}
	                                rows={shouldWrapTimeField ? 2 : 1}
	                                className={`w-[66px] min-w-0 resize-none border-0 bg-transparent p-0 placeholder:text-slate-400 [overflow-wrap:anywhere] focus:outline-none focus:ring-0 ${getValidationFieldHighlightClass(row.id, "reminderTime")}`}
	                                value={row.reminderTime ?? ""}
	                                placeholder={REMINDER_TIME_INPUT_MASK}
	                                onChange={(e) => {
	                                  clearValidationFieldHighlight(row.id, "reminderTime");
	                                  updateRow(row.id, (current) => ({
	                                    ...current,
	                                    reminderTime: maskReminderTimeDraftInput(e.target.value),
	                                  }))
	                                }}
	                                onBlur={(e) =>
	                                  updateRow(row.id, (current) => ({
	                                    ...current,
	                                    reminderTime: normalizeReminderTimeInput(e.target.value),
	                                  }))
	                                }
	                              />
	                              {rowTimeZoneControl}
	                            </span>
	                          );
	                        })()}
	                      </div>
	                    );
                    const rowEditorElement = (
	                      <StaggeredInlineEditorItem
	                        key={row.id}
	                        active={isBuilderEntryRevealActive}
	                        immediate={isBuilderEntryRevealImmediate || instantlyVisibleRowId === row.id}
		                        delayMs={360 + index * 70}
		                        role={isThisRowEditorLayered ? "dialog" : undefined}
		                        aria-modal={isThisRowEditorLayered ? true : undefined}
		                        aria-labelledby={isThisRowEditorLayered ? inspectorTitleId : undefined}
		                        data-plans-inspector-panel={isThisRowEditorLayered ? "true" : undefined}
		                        tabIndex={isThisRowEditorLayered ? -1 : undefined}
		                        onKeyDown={handleInspectorKeyDown}
			                        className={`${isThisRowEditorLayered ? "absolute" : "relative"} space-y-0 ${isTimeZoneMenuOpen ? "z-[80]" : ""} ${instantlyVisibleRowId === row.id ? "plans-new-row-enter" : ""} ${
		                          isThisRowEditorLayered
		                            ? `${isThisRowEditorClosing ? "plans-inspector-exit pointer-events-none" : "plans-inspector-enter pointer-events-auto"} inset-3 z-10 mx-auto flex h-[calc(100dvh-24px)] max-h-[calc(100dvh-24px)] max-w-[calc(100vw-24px)] flex-col overflow-hidden rounded-xl border border-slate-200/80 bg-white shadow-[-18px_0_50px_rgba(30,50,75,0.12)] max-[480px]:inset-0 max-[480px]:h-[100dvh] max-[480px]:max-h-[100dvh] max-[480px]:max-w-none max-[480px]:rounded-none lg:inset-y-0 lg:left-auto lg:right-0 lg:top-0 lg:mx-0 lg:h-[100dvh] lg:max-h-[100dvh] lg:max-w-none lg:rounded-none lg:border-y-0 lg:border-r-0 ${inspectorWidthClass}`
		                            : ""
		                        }`}
	                      >
                        <div
                          data-builder-row-card="true"
                          data-layered-row-editor={isThisRowEditorLayered ? "true" : undefined}
	                          ref={(node) => {
	                            rowNodeRefs.current[row.id] = node;
	                          }}
	                          onPointerDown={(e) => {
	                            if (isThisRowEditorLayered) return;
	                            if (e.button !== 0) return;
                            const target = e.target as HTMLElement | null;
                            if (target?.closest("[data-no-row-drag='true']")) return;
                            activeDragRef.current = {
                              pointerId: e.pointerId,
                              rowId: row.id,
                              startX: e.clientX,
                              startY: e.clientY,
                              isDragging: false,
                            };
                            setPressedRowId(row.id);
                          }}
	                          onPointerMove={(e) => {
	                            if (isThisRowEditorLayered) return;
	                            const activeDrag = activeDragRef.current;
                            if (!activeDrag || activeDrag.pointerId !== e.pointerId) return;

                            if (!activeDrag.isDragging) {
                              const movedX = Math.abs(e.clientX - activeDrag.startX);
                              const movedY = Math.abs(e.clientY - activeDrag.startY);
                              if (movedX < 6 && movedY < 6) return;
                              activeDrag.isDragging = true;
                              if (!e.currentTarget.hasPointerCapture(e.pointerId)) {
                                e.currentTarget.setPointerCapture(e.pointerId);
                              }
                              setDraggingRowId(activeDrag.rowId);
                              const nextInsertionIndex = getDragInsertionIndex(e.clientY, activeDrag.rowId);
                              dragInsertionIndexRef.current = nextInsertionIndex;
                              setDragInsertionIndex(nextInsertionIndex);
                            }

                            e.preventDefault();
                            const nextInsertionIndex = getDragInsertionIndex(e.clientY, activeDrag.rowId);
                            dragInsertionIndexRef.current = nextInsertionIndex;
                            setDragInsertionIndex(nextInsertionIndex);
                          }}
	                          onPointerUp={(e) => {
	                            if (isThisRowEditorLayered) return;
	                            const activeDrag = activeDragRef.current;
                            if (!activeDrag || activeDrag.pointerId !== e.pointerId) return;
                            if (e.currentTarget.hasPointerCapture(e.pointerId)) {
                              e.currentTarget.releasePointerCapture(e.pointerId);
                            }
                            finishRowDrag(activeDrag.isDragging);
                          }}
	                          onPointerCancel={(e) => {
	                            if (isThisRowEditorLayered) return;
	                            const activeDrag = activeDragRef.current;
                            if (!activeDrag || activeDrag.pointerId !== e.pointerId) return;
                            if (e.currentTarget.hasPointerCapture(e.pointerId)) {
                              e.currentTarget.releasePointerCapture(e.pointerId);
                            }
                            finishRowDrag(false);
                          }}
			                        className={`group relative overflow-hidden bg-white transition duration-150 ${inactiveEditorDimTransitionClass} ${
			                          isThisRowEditorLayered
			                            ? "shrink-0 cursor-default border-0 shadow-none focus-within:!border-transparent focus-within:!shadow-none"
			                            : "min-h-[72px] cursor-grab px-4 py-3 hover:bg-slate-50/70 active:cursor-grabbing"
			                        } ${isActiveDragRow ? "relative z-40 select-none bg-slate-50 shadow-[0_12px_24px_-20px_rgba(15,23,42,0.25)]" : ""} ${
	                            draggingRowId && !isDraggingRow ? "transition-transform duration-150" : ""
		                          } ${isMissingRowHighlighted ? "bg-red-50/35 ring-2 ring-inset ring-red-200" : ""} opacity-100 ${shouldHideFollowingRowDuringAddRowSettle ? "pointer-events-none invisible opacity-0" : ""} ${hiddenAddRowId === row.id ? "pointer-events-none invisible" : ""}`}
	                        >
	                          {isThisRowEditorLayered ? (
	                            <div className="grid shrink-0 gap-2 border-b border-slate-200/80 bg-white px-4 py-4 sm:grid-cols-[minmax(0,1fr)_auto] sm:items-start sm:gap-x-5 sm:gap-y-0 sm:px-6 sm:py-5">
	                              <div className="flex min-w-0 items-center justify-between gap-2 sm:col-start-2 sm:row-span-3 sm:row-start-1 sm:min-w-fit sm:justify-end">
	                                <div className="flex min-w-0 items-center gap-2 sm:hidden">
	                                  <span aria-hidden="true" className={inspectorDotClass} />
	                                  <div className="truncate text-[12px] font-semibold uppercase tracking-[0.09em] text-slate-600">
	                                    {rowMeta.label}
	                                  </div>
	                                </div>
	                                <div className="flex shrink-0 items-center gap-1.5 sm:gap-2">
	                                  <div className="hidden items-center gap-2 sm:flex">{renderInspectorTimingChips()}</div>
	                                  <button
	                                    type="button"
	                                    data-no-row-drag="true"
	                                    aria-label={`More actions for row ${index + 1}`}
	                                    className="flex h-[32px] w-[32px] items-center justify-center rounded-xl border border-slate-200 bg-white text-slate-500 transition hover:border-slate-300 hover:bg-slate-50 hover:text-slate-700 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#6f9fd1]/30 sm:h-[36px] sm:w-[36px]"
	                                  >
	                                    <span className="text-base leading-none">•••</span>
	                                  </button>
	                                  <button
	                                    type="button"
	                                    data-no-row-drag="true"
	                                    onClick={() => deleteBuilderRow(row.id)}
	                                    aria-label={`Delete row ${index + 1}`}
	                                    className="flex h-[32px] w-[32px] items-center justify-center rounded-xl border border-red-200 bg-white text-red-500 transition hover:border-red-300 hover:bg-red-50 hover:text-red-600 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-red-200 sm:h-[36px] sm:w-[36px]"
	                                  >
	                                    <svg
	                                      viewBox="0 0 24 24"
	                                      aria-hidden="true"
	                                      className="h-4 w-4"
	                                      fill="none"
	                                      stroke="currentColor"
	                                      strokeWidth="2"
	                                      strokeLinecap="round"
	                                      strokeLinejoin="round"
	                                    >
	                                      <path d="M4 7h16" />
	                                      <path d="M9 7V5h6v2" />
	                                      <path d="M7 7l1 12h8l1-12" />
	                                      <path d="M10 11v5M14 11v5" />
	                                    </svg>
	                                  </button>
	                                  <button
	                                    type="button"
	                                    data-no-row-drag="true"
	                                    autoFocus
	                                    onClick={handleInspectorClose}
	                                    aria-label={`Close ${inspectorKindLabel} inspector`}
	                                    className="flex h-[32px] w-[32px] items-center justify-center rounded-xl border border-slate-200 bg-white text-slate-500 transition hover:border-slate-300 hover:bg-slate-50 hover:text-slate-700 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#6f9fd1]/30 sm:h-[36px] sm:w-[36px]"
	                                  >
	                                    <svg
	                                      viewBox="0 0 20 20"
	                                      aria-hidden="true"
	                                      className="h-4 w-4"
	                                      fill="none"
	                                      stroke="currentColor"
	                                      strokeWidth="2"
	                                      strokeLinecap="round"
	                                    >
	                                      <path d="M5 5l10 10" />
	                                      <path d="M15 5 5 15" />
	                                    </svg>
	                                  </button>
	                                </div>
	                              </div>
	                              <div className="hidden min-w-0 items-center gap-4 sm:col-start-1 sm:row-start-1 sm:flex">
	                                <span aria-hidden="true" className={inspectorDotClass} />
	                                <div className="min-w-0 text-[12px] font-semibold uppercase tracking-[0.09em] text-slate-600">
	                                  {rowMeta.label}
	                                </div>
	                              </div>
	                              <h2
	                                id={inspectorTitleId}
	                                className="min-w-0 whitespace-nowrap text-[25px] font-semibold leading-[1.05] text-slate-950 sm:col-start-1 sm:row-start-2 sm:mt-1 sm:pl-7 sm:leading-7"
	                              >
	                                Edit {inspectorKindLabel}
	                              </h2>
	                              <div className="min-w-0 sm:col-start-1 sm:row-start-3 sm:mt-1 sm:pl-7">
	                                <p
	                                  className="truncate text-[14px] font-medium text-slate-600"
	                                  title={inspectorSubtitle}
	                                >
	                                  {inspectorSubtitle}
	                                </p>
	                                <div className="mt-2 flex flex-wrap items-center gap-2 sm:hidden">
	                                  {renderInspectorTimingChips()}
	                                </div>
	                              </div>
	                            </div>
	                          ) : (
	                            <>
	                        <span
	                          aria-hidden="true"
		                          className={`absolute bottom-0 left-0 top-0 w-[3px] ${
                            rowKind === "email"
                              ? "bg-green-500"
                              : rowKind === "meeting"
                                ? "bg-violet-500"
                                : "bg-blue-500"
                          }`}
                        />
			                        <div className="relative grid grid-cols-2 gap-x-2 gap-y-3 min-[900px]:grid-cols-[minmax(0,1fr)_136px_112px_32px] min-[900px]:items-center min-[900px]:gap-x-3 min-[900px]:gap-y-0">
	                          <div data-workflow-row-part="content" className="col-span-2 min-w-0 max-w-full pr-[40px] min-[900px]:col-span-1 min-[900px]:pr-0">
	                            <div className="mb-1 flex items-center gap-2 pl-2 text-[12px] font-semibold text-slate-500">
	                              <span>{rowMeta.label}</span>
	                            </div>
		                            <div className="min-w-0">
		                              <div className="min-w-0">
		                                <div className="inline-grid min-h-6 w-full min-w-0 items-center px-2 [grid-template-areas:'stack']">
	                                  <div
	                                    aria-hidden="true"
	                                    className="[grid-area:stack] invisible whitespace-normal px-0 py-0 text-left text-[14px] font-semibold leading-5 [overflow-wrap:anywhere]"
	                                  >
                                    {row.title ||
                                      (rowKind === "email"
                                        ? "Email Subject Line"
                                        : rowKind === "meeting"
                                          ? "Meeting Title"
                                          : rowMeta.label === "Meeting"
                                            ? "Meeting title"
                                            : "Enter Reminder Title")}
                                  </div>
	                                  <div
	                                    aria-hidden="true"
	                                    className={`pointer-events-none [grid-area:stack] whitespace-normal text-left text-[15px] font-semibold leading-5 text-slate-900 [overflow-wrap:anywhere] ${focusedTitleInputId === collapsedTitleInputId ? "opacity-0" : ""
	                                    }`}
                                  >
                                    {row.title ? (
                                      renderTextWithBoldAnchors(row.title, {
                                        anchorClassName: "font-bold text-black",
                                        knownAnchorKeys,
                                      })
                                    ) : (
	                                      <span className="font-medium text-slate-500">
                                        {rowKind === "email"
                                          ? "Email Subject Line"
                                          : rowKind === "meeting"
                                            ? "Meeting Title"
                                            : rowMeta.label === "Meeting"
                                              ? "Meeting title"
                                              : "Enter Reminder Title"}
                                      </span>
                                    )}
                                  </div>
                                  <textarea
                                    data-builder-row-title-input="true"
                                    ref={(node) => {
                                      builderTitleInputRefs.current[collapsedTitleInputId] = node;
                                    }}
                                    rows={1}
		                                    className={`[grid-area:stack] min-w-0 max-w-full resize-none border-0 bg-transparent px-0 py-0 text-left text-[15px] font-semibold leading-5 focus:outline-none focus:ring-0 ${
                                      focusedTitleInputId === collapsedTitleInputId
                                        ? "h-auto overflow-y-hidden whitespace-pre-wrap text-slate-950 caret-slate-950 [overflow-wrap:anywhere] placeholder:text-slate-400"
                                        : "h-full overflow-hidden whitespace-pre-wrap text-transparent caret-slate-950 placeholder:text-transparent"
                                    }`}
                                    value={row.title}
                                    placeholder={
                                      rowKind === "email"
                                        ? "Email Subject Line"
                                        : rowKind === "meeting"
                                          ? "Meeting Title"
                                        : rowMeta.label === "Meeting"
                                          ? "Meeting title"
                                            : "Enter Reminder Title"
                                    }
                                    onChange={(e) =>
                                      updateRow(row.id, (current) => ({
                                        ...current,
                                        title: e.target.value,
                                      }))
                                    }
                                    onKeyDown={(e) => {
                                      if (e.shiftKey) return;
                                      if (e.key === "Enter" || e.key === "Tab") {
                                        e.preventDefault();
                                        beginEditingOffsetForRow(row.id, row.offsetDays);
                                      }
                                    }}
                                    onFocus={() => setFocusedTitleInputId(collapsedTitleInputId)}
                                    onBlur={(e) => {
                                      e.currentTarget.style.height = "28px";
                                      setFocusedTitleInputId((current) =>
                                        current === collapsedTitleInputId ? null : current
                                      );
                                    }}
                                  />
                                </div>
                              </div>
                              {emailProviderNote ? (
                                <div className="mt-[3px] px-2 text-[12px] font-medium leading-[1.35] text-slate-500">
                                  {emailProviderNote}
                                </div>
                              ) : null}
                            </div>
                          </div>

			                          {collapsedOffsetControl}
			                          {collapsedTimeControl}

                            <button
                              ref={(node) => {
                                workflowRowMenuButtonRefs.current[row.id] = node;
                              }}
                              type="button"
                              data-no-row-drag="true"
                              data-workflow-row-part="menu"
                              aria-haspopup="menu"
                              aria-expanded={openWorkflowRowMenuId === row.id}
                              onClick={(event) => openWorkflowRowMenu(event, row.id)}
                              title="More actions"
                              aria-label={`More actions for row ${index + 1}`}
			                              className="absolute right-0 top-0 flex h-[32px] w-[32px] flex-none items-center justify-center rounded-[8px] border border-slate-200 bg-white text-slate-500 transition hover:border-slate-300 hover:bg-slate-50 hover:text-slate-700 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#6f9fd1]/30 min-[900px]:static min-[900px]:justify-self-center"
                            >
                              <span className="text-base leading-none">•••</span>
                            </button>
                            {hasMounted && openWorkflowRowMenuId === row.id && workflowRowMenuPosition
                              ? createPortal(
                                  <div
                                    ref={workflowRowMenuRef}
                                    role="menu"
                                    className="plans-menu-enter fixed z-[170] min-w-[184px] rounded-[10px] border border-slate-200 bg-white p-[7px] shadow-[0_18px_46px_rgba(21,40,66,0.18)]"
                                    style={{ top: workflowRowMenuPosition.top, left: workflowRowMenuPosition.left }}
                                  >
                                    <button
                                      type="button"
                                      role="menuitem"
                                      onClick={() => {
                                        closeWorkflowRowMenu();
                                        openExclusiveRowEditor(
                                          row.id,
                                          rowKind === "email" ? "email" : rowKind === "meeting" ? "meeting" : "reminder"
                                        );
                                      }}
                                      className="block h-10 w-full rounded-[8px] px-3 text-left text-[14px] font-medium text-slate-700 transition hover:bg-slate-50 focus:outline-none focus:ring-2 focus:ring-[#6f9fd1]/20"
                                    >
                                      Edit details
                                    </button>
                                    <button
                                      type="button"
                                      role="menuitem"
                                      onClick={() => {
                                        closeWorkflowRowMenu();
                                        deleteBuilderRow(row.id);
                                      }}
                                      className="block h-10 w-full rounded-[8px] px-3 text-left text-[14px] font-medium text-red-700 transition hover:bg-red-50 focus:outline-none focus:ring-2 focus:ring-red-200"
                                    >
                                      Delete {inspectorKindLabel}
                                    </button>
                                  </div>,
                                  document.body
                                )
                              : null}
	                        </div>
	                            </>
	                          )}
	                        </div>
	                        {reminderInlineEditor}
                        {emailInlineEditor}
                        {meetingInlineEditor}
	                      </StaggeredInlineEditorItem>
                    );

                    if (isThisRowEditorLayered && hasMounted) {
                      return createPortal(
                        <div
                          data-plans-inspector-overlay-root="true"
                          className="pointer-events-none fixed inset-0 z-[200]"
                        >
                          <div
                            aria-hidden="true"
                            data-plans-inspector-scrim="true"
                            className={`absolute inset-0 z-0 bg-slate-950/[0.14] ${
                              isThisRowEditorClosing ? "plans-scrim-exit pointer-events-none" : "plans-scrim-enter pointer-events-auto"
                            }`}
                          />
                          {rowEditorElement}
                        </div>,
                        document.body,
                        `plans-inspector-${row.id}`
                      );
                    }

                    return rowEditorElement;
			                  })}
		                        </div>
                                  ) : (
                                    <div className="flex min-h-[102px] flex-col items-center justify-center bg-slate-50/50 px-5 py-6 text-center">
                                      <h3 className="text-[15px] font-semibold leading-5 text-slate-950">No actions yet</h3>
                                      <p className="mt-2 max-w-[360px] text-[14px] leading-5 text-slate-600">
                                        Add a reminder, email, or meeting above.
                                      </p>
                                    </div>
                                  )}
		                      </div>
		                    </div>
		                  </div>
		                  </div>
	                </section>
	              </StaggeredInlineEditorItem>
		                        {!isAnyBuilderRowEditorVisible ? renderDynamicFieldsSection() : null}
		                        {!shouldRenderSimpleEventHeaderFields ? (
		                          <section className={`${plansSurfaceClass} ${plansCanvasSectionClass}`}>
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
		                              className={`${plansInputClass} mt-2 max-w-md`}
		                              value={weekendRule}
		                              onChange={(e) => setWeekendRule(e.target.value as WeekendRule)}
		                            >
		                              <option value="none">Allow weekends (no adjustment)</option>
		                              <option value="prior_business_day">Adjust to prior business day (Fri)</option>
		                            </select>
		                          </section>
		                        ) : null}
		                        {renderPlanActionsSection()}
		                      </div>
		                  ) : null}
          </div>
                </div>
              </div>
            </div>
        </section>
      </AnimatedRowEditor>

      {isBuilderWorkspaceVisible ? (
        <div className="h-16" aria-hidden="true" />
      ) : null}

      {isBuilderPreviewOpen && hasMounted ? createPortal(
		        <div className="fixed inset-0 z-[180] flex items-stretch justify-center bg-slate-950/[0.12] p-0 min-[640px]:items-center min-[640px]:p-5">
          <div
            role="dialog"
            aria-modal="true"
            aria-labelledby="review-export-title"
            aria-describedby="review-export-description"
            className="flex h-[100dvh] w-screen flex-col overflow-hidden bg-white min-[640px]:h-auto min-[640px]:max-h-[calc(100dvh-40px)] min-[640px]:w-[min(1100px,calc(100vw-40px))] min-[640px]:rounded-[16px] min-[640px]:border min-[640px]:border-slate-200 min-[640px]:shadow-[0_28px_80px_rgba(21,40,66,0.20)]"
          >
            <div className="flex flex-none items-start justify-between gap-4 border-b border-slate-200/80 bg-white px-4 py-4 min-[640px]:px-5 min-[640px]:py-[18px]">
              <div className="min-w-0">
                <h2 id="review-export-title" className="text-[22px] font-semibold leading-7 text-slate-950 min-[640px]:text-[24px]">Review &amp; Export</h2>
                <p id="review-export-description" className="mt-1.5 text-[14px] leading-[1.4] text-slate-600">
                  Review the selected actions before creating them in your connected tools.
                </p>
                {previewPlanForRender ? (
                  <p className="mt-[5px] text-[13px] font-medium leading-5 text-slate-500">
                    {previewPlanForRender.name} · {previewPlanForRender.items.length} {previewPlanForRender.items.length === 1 ? "action" : "actions"}
                  </p>
                ) : null}
              </div>
              <button
                type="button"
                onClick={() => {
		                  setIsBuilderPreviewOpen(false);
		                  setOpenPreviewDetail(null);
		                  setOpenPreviewRowMenuId(null);
		                  setExcludedPreviewItemIds([]);
		                  setShowNoPreviewItemsSelectedCallout(false);
                }}
                aria-label="Close Review & Export"
                className="flex h-[36px] w-[36px] shrink-0 items-center justify-center rounded-[9px] border border-slate-200 bg-white text-slate-500 transition hover:border-slate-300 hover:bg-slate-50 hover:text-slate-700 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#6f9fd1]/30"
              >
                <span aria-hidden="true" className="text-[20px] leading-none">×</span>
              </button>
            </div>

	            <div ref={previewModalScrollRef} className="min-h-0 flex-1 space-y-4 overflow-y-auto overscroll-contain bg-slate-50/70 p-3 pb-5 scroll-pb-[96px] scroll-pt-4 min-[640px]:p-4 min-[640px]:pb-5">
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
	                  <div className="overflow-hidden rounded-[14px] border border-slate-200 bg-white shadow-sm">
	                    <div ref={previewItemListHeadingRef} tabIndex={-1} className="rounded-t-[14px] border-b border-slate-200/80 bg-white px-4 py-3 outline-none focus-visible:ring-2 focus-visible:ring-[#6f9fd1]/30">
	                      <div className="text-[15px] font-semibold leading-5 text-slate-950">{previewPlanForRender.name}</div>
	                      {previewLoading ? <div className="mt-1 text-xs text-gray-500">Refreshing preview…</div> : null}
	                      <div className="mt-1 text-[13px] leading-5 text-slate-600">
	                        {previewPlanForRender.items.filter((item) => classifyPlanRow(item) !== "email").length} scheduled items
	                      </div>
	                    </div>
	                    <div className="divide-y divide-slate-200/80">
                      {getReviewExportItemsForRender(previewPlanForRender.items).map((item) => {
                    const rowKind = classifyPlanRow(item);
                    const builderItem = rows.find((entry) => entry.id === item.id);
                    const previewReminderBody = item.body ?? builderItem?.body ?? "";
                    const hasReminderPreview = rowKind === "reminder";
                    const hasEmailPreview = rowKind === "email";
                    const hasMeetingPreview = rowKind === "meeting";
                    const isReminderExpanded = openPreviewDetail?.rowId === item.id && openPreviewDetail.kind === "reminder";
                    const isEmailExpanded = openPreviewDetail?.rowId === item.id && openPreviewDetail.kind === "email";
                    const isMeetingExpanded = openPreviewDetail?.rowId === item.id && openPreviewDetail.kind === "meeting";
                    const previewEmailDraft = normalizeEmailDraft(item.emailDraft);
                    const previewEmailSubject = previewEmailDraft.subject.trim() || item.title || "Email draft";
                    const previewEmailFinalBody = buildFinalEmailBody(previewEmailDraft.body, {
                      signature: appSettings.emailSignatureText,
                    });
                    const previewMeetingDraft = normalizeMeetingDraft(item.meetingDraft);
                    const isPreviewTeamsMeetingEnabled = Boolean(previewMeetingDraft?.teamsMeeting);
                    const isPreviewGoogleMeetEnabled = Boolean(previewMeetingDraft?.addGoogleMeet);
                    const rowTypeLabel = rowKind === "email" ? "Email" : rowKind === "meeting" ? "Meeting" : "Reminder";
                    const isItemIncluded = isPreviewItemIncluded(item.id);
                    const isItemScheduledInPast = isPreviewItemScheduledInPast(item);
                    const rowTypeDotClass = !isItemIncluded
                      ? "bg-gray-300"
                      : rowKind === "email"
                        ? "bg-amber-400"
                        : rowKind === "meeting"
                          ? "bg-violet-400"
                          : "bg-blue-400";
                    const itemTimeLabel =
                      item.meetingDraft?.isAllDay || item.durationDraft?.isAllDay
                        ? "All day"
                        : item.reminderTime
                          ? formatPreviewTime(item.reminderTime)
                          : "All day";
                    const detailLabel =
                      rowKind === "email"
                        ? isEmailExpanded
                          ? "Hide Email"
                          : "View Email"
                        : rowKind === "meeting"
                          ? isMeetingExpanded
                            ? "Hide Meeting"
                            : "View Meeting"
                          : isReminderExpanded
                            ? "Hide Reminder"
                            : "View Reminder";
                    const exportLabel = hasEmailPreview
                      ? getPreviewEmailActionLabel(appSettings.emailHandlingMode)
                      : hasMeetingPreview
                        ? "Add Meeting"
                        : "Export Reminder";

                    return (
                      <div
                        key={item.id}
                        ref={(node) => {
                          if (node) {
                            previewDetailRowRefs.current[item.id] = node;
                            return;
                          }
                          delete previewDetailRowRefs.current[item.id];
                        }}
                      >
	                        <div className={`grid gap-3 px-4 py-3 transition min-[900px]:grid-cols-[32px_minmax(0,1fr)_136px_192px_36px] min-[900px]:items-center min-[900px]:gap-3 ${
	                          isItemIncluded ? "" : "bg-slate-100/80 text-slate-400 opacity-65 grayscale"
	                        }`}>
	                            <div className="flex h-[32px] items-start pt-1 min-[900px]:items-center min-[900px]:justify-center min-[900px]:pt-0">
	                              <input
	                                type="checkbox"
	                                checked={isItemIncluded}
                                onChange={(e) => togglePreviewItemIncluded(item.id, e.target.checked)}
                                aria-label={`${isItemIncluded ? "Exclude" : "Include"} ${item.customTitle ?? item.title}`}
	                                className="h-4 w-4 rounded border-gray-300 text-blue-600 focus:ring-blue-500"
	                              />
	                            </div>
	                            <div className="min-w-0">
	                            <div className="flex items-center gap-2">
                              <span className={`h-2.5 w-2.5 shrink-0 rounded-full ${rowTypeDotClass}`} />
	                              <span className={`text-[11px] font-semibold uppercase tracking-[0.06em] ${
                                isItemIncluded ? "text-gray-500" : "text-gray-400"
                              }`}>{rowTypeLabel}</span>
                            </div>
	                            <div className={`mt-1 whitespace-normal break-words text-[14px] font-semibold leading-5 ${
                              isItemIncluded ? "text-gray-900" : "text-gray-400"
                            }`}>
                              {renderTextWithBoldAnchors(item.customTitle ?? item.title, {
                                anchorClassName: isItemIncluded ? "font-bold text-black" : "font-bold text-gray-400",
                                knownAnchorKeys,
                              })}
                            </div>
	                            <div className={`mt-1 flex flex-wrap items-center gap-x-2 gap-y-1 text-[12px] leading-4 ${
                              isItemIncluded ? "text-gray-500" : "text-gray-400"
                            }`}>
                              <span>{formatOffsetLabel(item.offsetDays, { relativeToToday: noEventDate, dateBasis: item.dateBasis })}</span>
                              <span aria-hidden="true">•</span>
                              <span>{itemTimeLabel}</span>
                              {isItemScheduledInPast ? (
                                <>
                                  <span aria-hidden="true">•</span>
                                  <span className="inline-flex items-center gap-1 font-semibold text-red-600">
                                    <span aria-hidden="true">!</span>
                                    <span>In the Past</span>
                                  </span>
                                </>
                              ) : null}
                              {hasEmailPreview ? (
                                <>
                                  <span aria-hidden="true">•</span>
                                  <span className={`font-medium ${isItemIncluded ? "text-gray-600" : "text-gray-400"}`}>
                                    {getPreviewEmailModeMessage(appSettings.emailHandlingMode)}
                                  </span>
                                </>
                              ) : null}
                            </div>
	                            </div>
	                          <div className="flex min-w-0 items-center min-[900px]:justify-self-stretch">
	                            {hasReminderPreview || hasEmailPreview || hasMeetingPreview ? (
	                              <button
                                type="button"
                                onClick={() => {
                                  togglePreviewDetail(
                                    item.id,
                                    hasEmailPreview ? "email" : hasMeetingPreview ? "meeting" : "reminder"
                                  );
                                }}
	                                className={`inline-flex h-[40px] w-full min-w-[136px] items-center justify-center rounded-[10px] border px-3 text-center text-[14px] font-semibold ${
	                                  isItemIncluded
	                                    ? "border-gray-200 bg-white text-gray-700 hover:border-gray-300 hover:bg-gray-50"
	                                    : "border-gray-200 bg-gray-50 text-gray-400 hover:border-gray-300 hover:bg-gray-100"
                                }`}
                              >
                                {detailLabel}
	                              </button>
	                            ) : null}
	                          </div>
	                          <div className="flex min-w-0 items-center min-[900px]:justify-self-stretch">
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
                                disabled={executionState === "pending" || !isItemIncluded}
                                title={
                                  hasReminderPreview
                                    ? "Export this reminder"
                                    : hasEmailPreview
                                      ? "Run this email action"
                                      : "Create this meeting"
                                }
	                                className="inline-flex h-[40px] w-full min-w-[192px] items-center justify-center rounded-[10px] border border-blue-600 bg-blue-600 px-3 text-center text-[14px] font-semibold text-white hover:bg-blue-700 disabled:cursor-not-allowed disabled:border-gray-300 disabled:bg-gray-100 disabled:text-gray-400"
	                              >
	                                {exportLabel}
	                              </button>
	                          </div>
	                          <div className="flex items-center min-[900px]:justify-self-center">
	                            {hasReminderPreview || hasEmailPreview || hasMeetingPreview ? (
	                              <div className="relative">
                                <button
                                  type="button"
                                  onClick={() => setOpenPreviewRowMenuId((prev) => (prev === item.id ? null : item.id))}
                                  title="Actions"
                                  aria-label="Actions"
	                                  className="flex h-[36px] w-[36px] items-center justify-center rounded-[10px] border border-gray-200 bg-white text-gray-500 hover:border-gray-300 hover:bg-gray-100 hover:text-gray-700"
                                >
                                  <span className="text-base leading-none">•••</span>
                                </button>
                                {openPreviewRowMenuId === item.id ? (
	                                  <div className="plans-menu-enter absolute right-0 top-[calc(100%+0.5rem)] z-[190] w-44 rounded-[10px] border border-slate-200 bg-white p-[7px] text-left shadow-[0_18px_46px_rgba(21,40,66,0.18)]">
                                    {hasReminderPreview ? (
                                      <button
                                        type="button"
                                        onClick={() => {
                                          togglePreviewDetail(item.id, "reminder");
                                          setOpenPreviewRowMenuId(null);
                                        }}
	                                        className="h-10 w-full rounded-[8px] px-3 text-left text-[13px] font-medium text-slate-700 hover:bg-slate-50 focus:outline-none focus:ring-2 focus:ring-[#6f9fd1]/20"
                                      >
                                        {isReminderExpanded ? "Hide Reminder" : "View Reminder"}
                                      </button>
                                    ) : null}
                                    {hasEmailPreview ? (
                                      <button
                                        type="button"
                                        onClick={() => {
                                          togglePreviewDetail(item.id, "email");
                                          setOpenPreviewRowMenuId(null);
                                        }}
	                                        className="h-10 w-full rounded-[8px] px-3 text-left text-[13px] font-medium text-slate-700 hover:bg-slate-50 focus:outline-none focus:ring-2 focus:ring-[#6f9fd1]/20"
                                      >
                                        {isEmailExpanded ? "Hide Email" : "View Email"}
                                      </button>
                                    ) : null}
                                    {hasMeetingPreview ? (
                                      <button
                                        type="button"
                                        onClick={() => {
                                          togglePreviewDetail(item.id, "meeting");
                                          setOpenPreviewRowMenuId(null);
                                        }}
	                                        className="h-10 w-full rounded-[8px] px-3 text-left text-[13px] font-medium text-slate-700 hover:bg-slate-50 focus:outline-none focus:ring-2 focus:ring-[#6f9fd1]/20"
                                      >
                                        {isMeetingExpanded ? "Hide Meeting" : "View Meeting"}
                                      </button>
                                    ) : null}
	                              </div>
	                            ) : null}
	                          </div>
                            ) : null}
                          </div>
                        </div>

                        {rowKind === "reminder" ? (
                          <AnimatedRowEditor open={isReminderExpanded}>
	                          <div className="mb-5 border-t border-blue-100 bg-blue-50/65 px-4 py-3">
                            <div className="grid grid-cols-1 gap-3">
                              <div>
                                <label className="mb-1 block text-sm font-medium text-blue-950">Reminder Body</label>
                                <textarea
                                  rows={5}
                                  className="w-full rounded-lg border border-blue-200 bg-white px-4 py-2 text-sm text-gray-900"
                                  value={previewReminderBody}
                                  onChange={(e) =>
                                    builderItem ? updateRow(builderItem.id, (current) => ({ ...current, body: e.target.value })) : undefined
                                  }
                                />
                              </div>
                            </div>
                          </div>
                          </AnimatedRowEditor>
                        ) : null}

                        {rowKind === "meeting" ? (
                          <AnimatedRowEditor open={isMeetingExpanded}>
	                          <div className="mb-5 border-t border-violet-100 bg-violet-50/65 px-4 py-3">
                            <div className="mb-3 text-sm font-medium text-violet-900">
                              {activeAccountProvider === "gmail"
                                ? "This meeting will be created in Google Calendar when Google is connected."
                                : "This meeting will be created in Outlook when connected."}
                            </div>
                            <div className="grid grid-cols-1 gap-3 md:grid-cols-2">
                              <div className="md:col-span-2">
                                <label className="mb-1 block text-sm font-medium text-violet-950">Meeting Title</label>
                                <input
                                  className="w-full rounded-lg border border-violet-200 bg-white px-4 py-2 text-sm text-gray-900"
                                  value={builderItem?.title ?? item.title}
                                  onChange={(e) =>
                                    builderItem ? updateRow(builderItem.id, (current) => ({ ...current, title: e.target.value })) : undefined
                                  }
                                />
                              </div>
                              <div className="md:col-span-2">
                                <div className="mb-1 flex items-center justify-between gap-3">
                                  <label className="block text-sm font-medium text-violet-950">Attendees</label>
                                  {builderItem ? (
                                    <button
                                      type="button"
                                      onClick={() => openRecipientGroupsModal({ rowId: builderItem.id, field: "meeting_attendees" })}
                                      className="rounded-full border border-slate-300 bg-white px-2.5 py-1 text-xs font-medium text-slate-700 hover:bg-slate-50"
                                    >
                                      Group
                                    </button>
                                  ) : null}
                                </div>
                                <EmailTokensInput
                                  label=""
                                  values={previewMeetingDraft?.attendees ?? []}
                                  recipientGroups={recipientGroups}
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
                                  className={`w-full rounded-lg border border-violet-200 px-4 py-2 text-sm ${
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
                                  className="w-full max-w-[220px] rounded-lg border border-violet-200 bg-white px-4 py-2 text-sm text-gray-900"
                                  value={previewMeetingDraft?.isAllDay ? "__all_day__" : String(previewMeetingDraft?.durationMinutes ?? 30)}
                                  onChange={(e) =>
                                    builderItem
                                      ? updateRow(builderItem.id, (current) => ({
                                          ...current,
                                          meetingDraft: { ...normalizeMeetingDraft(current.meetingDraft), durationMinutes: Number(e.target.value) },
                                        }))
                                      : undefined
                                  }
                                >
                                  {previewMeetingDraft?.isAllDay ? <option value="__all_day__">All Day</option> : null}
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
                                          className="w-full rounded-lg border border-violet-200 bg-white px-4 py-2 text-sm text-gray-900"
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
                                          className="w-full rounded-lg border border-violet-200 bg-white px-4 py-2 text-sm text-gray-900"
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
                          </AnimatedRowEditor>
                        ) : null}

                        {rowKind === "email" ? (
                          <AnimatedRowEditor open={isEmailExpanded}>
	                          <div className="mb-5 border-t border-amber-100 bg-amber-50/65 px-4 py-3">
                            <div className="font-medium text-amber-950">
                              <span>Subject: </span>
                              {renderTextWithBoldAnchors(previewEmailSubject, {
                                anchorClassName: "font-bold text-black",
                                knownAnchorKeys,
                              })}
                            </div>
                            <div className="mt-2 text-sm font-medium text-amber-900">
                              {getPreviewEmailModeMessage(appSettings.emailHandlingMode)}
                            </div>
                            <div className="mt-2 grid grid-cols-1 gap-3">
                              <div>
                                <div className="mb-1 flex items-center justify-between gap-3">
                                  <label className="block text-sm font-medium text-amber-950">To</label>
                                  {builderItem ? (
                                    <button
                                      type="button"
                                      onClick={() => openRecipientGroupsModal({ rowId: builderItem.id, field: "email_to" })}
                                      className="rounded-full border border-slate-300 bg-white px-2.5 py-1 text-xs font-medium text-slate-700 hover:bg-slate-50"
                                    >
                                      Group
                                    </button>
                                  ) : null}
                                </div>
                                <EmailTokensInput
                                  label=""
                                  values={previewEmailDraft.to}
                                  recipientGroups={recipientGroups}
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
                                  recipientGroups={recipientGroups}
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
                                  recipientGroups={recipientGroups}
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
                                  className="w-full rounded-lg border border-amber-200 bg-white px-4 py-2 text-sm text-gray-900"
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
                                  className="w-full rounded-lg border border-amber-200 bg-white px-4 py-2 text-sm text-gray-900"
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
                                  <pre className="whitespace-pre-wrap rounded-lg border border-amber-200 bg-white px-4 py-2 text-sm text-gray-900">
                                    {renderTextWithBoldAnchors(previewEmailFinalBody, { knownAnchorKeys })}
                                  </pre>
                                </div>
                              ) : null}
                            </div>
                          </div>
                          </AnimatedRowEditor>
                        ) : null}
                      </div>
                    );
                  })}
                    </div>
                  </div>

	                  {showNoPreviewItemsSelectedCallout && getIncludedPreviewItemIds().length === 0 ? (
	                    <div className="rounded-[12px] border border-amber-200 bg-amber-50 px-4 py-3 text-[14px] leading-5 text-amber-900">
	                      <div className="font-semibold">Select at least one action.</div>
	                      <p className="mt-1 text-amber-800">Choose one or more actions before exporting this plan.</p>
	                    </div>
	                  ) : null}

		                  <div className="mb-4 flex flex-col gap-4 rounded-[14px] border border-slate-200 bg-white p-4 shadow-sm">
		                    <div className="max-w-md">
	                      <label className="mb-1 block text-[14px] font-semibold text-slate-700">Weekend handling</label>
	                      <select
	                        className="h-[42px] w-full rounded-[10px] border border-slate-200 bg-white px-3 text-[14px] text-slate-900 shadow-sm focus:border-[#6f9fd1] focus:outline-none focus:ring-2 focus:ring-[#6f9fd1]/20"
	                        value={weekendRule}
	                        onChange={(e) => setWeekendRule(e.target.value as WeekendRule)}
                      >
                        <option value="none">Allow weekends (no adjustment)</option>
                        <option value="prior_business_day">Adjust to prior business day (Fri)</option>
                      </select>
	                      <p className="mt-1.5 text-[13px] leading-5 text-slate-500">
	                        If a computed date lands on Sat/Sun, we either move it to Friday or leave it as-is.
	                      </p>
	                    </div>

	                  </div>
	                </>
	              ) : (
	                <div className="rounded-xl border border-dashed px-4 py-8 text-sm text-gray-500">
	                  {noEventDate
	                    ? "Review is not ready for the current event plan yet."
	                    : "Add an event date to review the current event plan."}
	                </div>
	              )}
	            </div>
	            {previewPlanForRender ? (
	              <div className="flex flex-none flex-col gap-3 border-t border-slate-200/80 bg-white px-[14px] py-3 [padding-bottom:max(12px,env(safe-area-inset-bottom))] min-[640px]:flex-row min-[640px]:items-center min-[640px]:justify-between min-[640px]:gap-4 min-[640px]:px-5 min-[640px]:py-[14px] min-[640px]:[padding-bottom:14px]">
	                <div className="text-[13px] font-semibold leading-5 text-slate-600">
	                  {getIncludedPreviewItemIds().length} of {previewPlanForRender.items.length} selected
	                </div>
	                <div className="flex w-full flex-col gap-2 min-[640px]:w-auto min-[640px]:flex-row min-[640px]:items-center min-[640px]:justify-end">
	                  {executionState === "success" ? <ExportDoneBadge /> : null}
	                  <button
	                    type="button"
	                    onClick={() => {
	                      const includedItemIds = getIncludedPreviewItemIds();
	                      if (includedItemIds.length === 0) {
	                        setShowNoPreviewItemsSelectedCallout(true);
	                        window.setTimeout(() => previewItemListHeadingRef.current?.focus(), 0);
	                        return;
	                      }
	                      void exportCurrentPlan({ skipConfirm: true, itemIds: includedItemIds });
	                    }}
	                    disabled={executionState === "pending"}
	                    className="inline-flex h-[40px] w-full min-w-[128px] items-center justify-center rounded-[10px] border border-blue-600 bg-blue-600 px-4 text-[14px] font-semibold text-white shadow-sm hover:bg-blue-700 disabled:cursor-not-allowed disabled:border-slate-200 disabled:bg-slate-100 disabled:text-slate-400 min-[640px]:w-auto"
	                  >
	                    {executionState === "pending" ? "Exporting..." : "Export Plan"}
	                  </button>
	                </div>
	              </div>
	            ) : null}
	          </div>
	        </div>,
	        document.body
	      ) : null}

	      {AI_ENABLED && isAiPanelOpen ? (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/35 px-4 py-8">
          <div className="max-h-[90vh] w-full max-w-6xl overflow-hidden rounded-2xl border bg-white shadow-xl">
            <div className="flex items-center justify-between border-b px-4 py-3">
              <div>
                <h2 className="text-lg font-semibold text-gray-900">Draft with AI</h2>
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
                <div ref={aiConversationRef} className="flex-1 space-y-4 overflow-y-auto p-4">
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

              <div className="min-h-0 overflow-y-auto bg-gray-50/60 p-4">
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
                              <div className="whitespace-pre-wrap text-sm text-gray-700">
                                {renderTextWithBoldAnchors(row.body, { knownAnchorKeys })}
                              </div>
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

      {hasMounted ? (
        <div data-recipient-groups-modal-host="true">
          <RecipientGroupsModal
            open={isRecipientGroupsModalOpen}
            groups={recipientGroups}
            initialMode={recipientGroupsModalMode}
            initialEditingGroup={recipientGroupsEditingGroup}
            onClose={() => {
              setIsRecipientGroupsModalOpen(false);
              setRecipientGroupsEditingGroup(null);
            }}
            onSelect={applyRecipientGroupToTarget}
            onDelete={handleDeleteRecipientGroup}
            onSave={handleSaveRecipientGroup}
          />
        </div>
      ) : null}

      <TemplateLibrary
        open={isTemplateLibraryOpen}
        mounted={hasMounted}
        templates={savedTemplates}
        selectedTemplateId={selectedTemplateId}
        highlightedTemplateId={highlightedTemplateId}
        actionMessage={templateActionMessage}
        onSelectTemplate={(templateId) => void onSelectSavedTemplate(templateId)}
        onDuplicateTemplate={(templateId) => void duplicateSavedTemplate(templateId)}
        onRenameTemplate={(templateId) => void renameSavedTemplate(templateId)}
        onDeleteTemplate={(templateId) => void deleteSavedTemplate(templateId)}
        onClose={() => setIsTemplateLibraryOpen(false)}
        returnFocus={() => templateLibraryButtonRef.current?.focus()}
      />

      {plansModal ? (
        <PlansDialogShell
          title={plansModal.title}
          description={plansModal.message}
          animatedIn={isPopupAnimatedIn}
          severity={plansModal.severity ?? (plansModal.destructive ? "destructive" : "information")}
          maxWidthClassName={plansModal.maxWidthClassName}
          onEscape={() => closePlansModal(plansModal.kind === "prompt" ? null : plansModal.kind === "alert" ? true : false)}
          footer={
            <>
              {plansModal.secondaryLabel ? (
                <button
                  type="button"
                  onClick={() => {
                    plansModal.onSecondaryAction?.();
                    closePlansModal(false);
                  }}
                  className={`${plansDialogSecondaryButtonClass} w-full whitespace-nowrap sm:w-auto`}
                >
                  {plansModal.secondaryLabel}
                </button>
              ) : null}
              {plansModal.kind !== "alert" ? (
                <button
                  type="button"
                  onClick={() => closePlansModal(plansModal.kind === "prompt" ? null : false)}
                  className={`${plansDialogSecondaryButtonClass} w-full whitespace-nowrap sm:w-auto`}
                >
                  {plansModal.cancelLabel ?? "Cancel"}
                </button>
              ) : null}
              <button
                type="button"
                autoFocus={plansModal.kind !== "prompt"}
                onClick={() => {
                  plansModal.onConfirmAction?.();
                  closePlansModal(
                    plansModal.kind === "prompt" ? plansModalInputValue : true
                  );
                }}
                className={`${plansModal.destructive || plansModal.severity === "destructive" ? plansDialogDangerButtonClass : plansModal.severity === "warning" ? plansDialogWarningButtonClass : plansDialogPrimaryButtonClass} w-full whitespace-nowrap sm:w-auto`}
              >
                {plansModal.confirmLabel ?? "OK"}
              </button>
            </>
          }
        >
          {plansModal.content ? plansModal.content : null}
          {plansModal.items?.length ? (
            <ul className="max-h-[36dvh] space-y-2 overflow-y-auto rounded-xl border border-slate-200/80 bg-slate-50/70 p-3 text-[14px] leading-5 text-slate-700">
              {plansModal.items.map((item, index) => (
                <li key={index} className="flex items-start gap-2">
                  <span className={`mt-[7px] h-1.5 w-1.5 shrink-0 rounded-full ${plansModal.destructive ? "bg-red-400" : "bg-[#4f7fb8]"}`} />
                  <span className="min-w-0">{item}</span>
                </li>
              ))}
            </ul>
          ) : null}
          {plansModal.kind === "prompt" ? (
            <div className={plansModal.items?.length || plansModal.content ? "mt-4" : ""}>
              <label className="mb-1.5 block text-[13px] font-medium text-slate-600" htmlFor="plans-modal-prompt-input">
                {plansModal.inputLabel ?? plansModal.placeholder ?? "Value"}
              </label>
                <input
                  id="plans-modal-prompt-input"
                  autoFocus
                  value={plansModalInputValue}
                  onChange={(e) => setPlansModalInputValue(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === "Enter") {
                      e.preventDefault();
                      closePlansModal(plansModalInputValue);
                    } else if (e.key === "Escape") {
                      e.preventDefault();
                      closePlansModal(null);
                    }
                  }}
                  placeholder={plansModal.placeholder}
                  className={plansInputClass}
                />
                {plansModal.helperText ? (
                  <div className="mt-2 text-[13px] leading-5 text-slate-500">{plansModal.helperText}</div>
                ) : null}
              </div>
          ) : null}
        </PlansDialogShell>
      ) : null}

      {AI_ENABLED && showAiApplyConfirm && aiChatDraft ? (
        <div className="fixed inset-0 z-[60] flex items-center justify-center bg-black/45 px-4">
          <div
            className={`w-full max-w-md origin-top transform-gpu rounded-2xl border bg-white p-4 shadow-xl transition-all duration-300 ease-out ${
              isPopupAnimatedIn
                ? "translate-y-0 scale-100 opacity-100"
                : "-translate-y-3 scale-95 opacity-0"
            }`}
          >
            <h3 className="text-lg font-semibold text-gray-900">Replace current event plan?</h3>
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
                Replace current plan
              </button>
            </div>
          </div>
        </div>
      ) : null}

      {showNewPlanDialog ? (
        <div className="fixed inset-0 z-[60] flex items-center justify-center bg-black/45 px-4">
          <div
            className={`w-full max-w-md origin-top transform-gpu rounded-2xl border bg-white p-4 shadow-xl transition-all duration-300 ease-out ${
              isPopupAnimatedIn
                ? "translate-y-0 scale-100 opacity-100"
                : "-translate-y-3 scale-95 opacity-0"
            }`}
          >
            <h3 className="text-lg font-semibold text-gray-900">
              {planSetupDialogMode === "template" ? `Use ${planSetupTemplate?.name ?? "Template"}` : "New Event"}
            </h3>
            <p className="mt-2 text-sm text-gray-600">
              {planSetupDialogMode === "template"
                ? "Confirm the event details for this template before continuing."
                : "Start a new event and prefill the builder with the key timing details."}
            </p>
            <div className="mt-4 space-y-4">
              <div>
                <label className="mb-1 block text-sm font-medium text-gray-700">Event Name</label>
                <input
                  autoFocus
                  value={newPlanDraft.eventName}
                  onChange={(e) => {
                    const nextValue = e.target.value;
                    setNewPlanDraft((current) => ({ ...current, eventName: nextValue }));
                    if (newPlanDialogMessage) {
                      setNewPlanDialogMessage(null);
                    }
                  }}
                  className="w-full rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900"
                  placeholder="Event name"
                />
              </div>
              <div>
                <label className="mb-1 block text-sm font-medium text-gray-700">Event Date</label>
                {newPlanDraft.noEventDate ? (
                  <input
                    type="text"
                    className="w-full rounded-lg border border-gray-300 bg-gray-100 px-3 py-2 text-sm text-gray-500"
                    value="Today"
                    readOnly
                  />
                ) : (
                  <input
                    type="date"
                    value={newPlanDraft.anchorDate}
                    onChange={(e) =>
                      setNewPlanDraft((current) => ({
                        ...current,
                        anchorDate: e.target.value,
                      }))
                    }
                    className="w-full rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900"
                  />
                )}
              </div>
              <div>
                <label className="mb-1 block text-sm font-medium text-gray-700">Event Time</label>
                <input
                  type="time"
                  value={newPlanDraft.eventTime}
                  onChange={(e) =>
                    setNewPlanDraft((current) => ({
                      ...current,
                      eventTime: e.target.value,
                    }))
                  }
                  className="w-full rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900"
                />
              </div>
              <label className="inline-flex w-full items-center gap-3 rounded-xl border border-slate-200 bg-white px-4 py-3 text-sm font-medium text-slate-700">
                <input
                  type="checkbox"
                  checked={newPlanDraft.noEventDate}
                  onChange={(e) =>
                    setNewPlanDraft((current) => ({
                      ...current,
                      noEventDate: e.target.checked,
                      anchorDate: e.target.checked ? "" : current.anchorDate,
                    }))
                  }
                />
                <span>Use today instead of an event date</span>
              </label>
              <div>
                <label className="mb-1 block text-sm font-medium text-gray-700">Weekend Handling</label>
                <select
                  value={newPlanDraft.weekendRule}
                  onChange={(e) =>
                    setNewPlanDraft((current) => ({
                      ...current,
                      weekendRule: e.target.value as WeekendRule,
                    }))
                  }
                  className="w-full rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900"
                >
                  <option value="none">Allow weekends (no adjustment)</option>
                  <option value="prior_business_day">Adjust to prior business day (Fri)</option>
                </select>
              </div>
            </div>
            {newPlanDialogMessage ? (
              <div className="mt-3 rounded-lg border border-gray-200 bg-gray-50 px-3 py-2 text-sm text-gray-700">
                {newPlanDialogMessage}
              </div>
            ) : null}
            <div className="mt-5 flex flex-wrap justify-end gap-3">
              <button
                type="button"
                onClick={() => {
                  if (planSetupDialogMode === "new") {
                    setIsNewPlanSetupPending(true);
                  }
                  setPlanSetupTemplateId(null);
                  setShowNewPlanDialog(false);
                  setNewPlanDialogMessage(null);
                }}
                className="rounded-lg border border-gray-300 bg-white px-4 py-2 text-sm text-gray-900 hover:bg-gray-50"
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={confirmStartNewPlan}
                className="rounded-lg bg-blue-600 px-4 py-2 text-sm font-medium text-white hover:bg-blue-700"
              >
                {planSetupDialogMode === "template" ? "Use Template" : "Create Event"}
              </button>
            </div>
          </div>
        </div>
      ) : null}

      {AI_ENABLED && showAiTemplateSaveDialog && aiChatDraft ? (
        <div className="fixed inset-0 z-[60] flex items-center justify-center bg-black/45 px-4">
          <div
            className={`w-full max-w-md origin-top transform-gpu rounded-2xl border bg-white p-4 shadow-xl transition-all duration-300 ease-out ${
              isPopupAnimatedIn
                ? "translate-y-0 scale-100 opacity-100"
                : "-translate-y-3 scale-95 opacity-0"
            }`}
          >
            <h3 className="text-lg font-semibold text-gray-900">Save AI draft as template</h3>
            <p className="mt-2 text-sm text-gray-600">
              Save this draft as a reusable template. Your current builder will stay unchanged.
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
        <PlansDialogShell
          title="Save as template"
          description={
            <>
              <div>Save this workflow structure so it can be reused for another event.</div>
              <div className="mt-1 text-[13px] text-slate-500">
                Event details and run-specific Anchor Field values are not saved.
              </div>
            </>
          }
          animatedIn={isPopupAnimatedIn}
          severity="information"
          onEscape={() => setShowBuilderTemplateSaveDialog(false)}
          footer={
            <>
              <button
                type="button"
                onClick={() => setShowBuilderTemplateSaveDialog(false)}
                className={`${plansDialogSecondaryButtonClass} w-full sm:w-auto`}
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={saveCurrentBuilderAsTemplate}
                className={`${plansDialogPrimaryButtonClass} w-full sm:w-auto`}
              >
                Save template
              </button>
            </>
          }
        >
            <div>
              <label className="mb-1.5 block text-[13px] font-medium text-slate-600" htmlFor="builder-template-name">
                Template name
              </label>
              <input
                id="builder-template-name"
                autoFocus
                value={builderTemplateNameDraft}
                onChange={(e) => setBuilderTemplateNameDraft(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === "Enter") {
                    e.preventDefault();
                    saveCurrentBuilderAsTemplate();
                  }
                }}
                className={plansInputClass}
                placeholder="Template name"
              />
            </div>
            {builderTemplateSaveMessage ? (
              <div className="mt-3 rounded-xl border border-amber-200 bg-amber-50 px-3 py-2 text-[13px] font-medium leading-5 text-amber-900">
                {builderTemplateSaveMessage}
              </div>
            ) : null}
        </PlansDialogShell>
      ) : null}
    </div>
  );
}
