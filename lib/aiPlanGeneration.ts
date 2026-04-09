import type { PlanDateBasis, PlanRowType, PlanType, WeekendRule } from "../types/plan";

export type AIPlanDraftEmail = {
  to: string[];
  cc: string[];
  bcc: string[];
  subject: string;
  body: string;
};

export type AIPlanDraftDuration = {
  durationMinutes?: number;
  useCustomEnd?: boolean;
  endDate?: string;
  endTime?: string;
  isAllDay?: boolean;
};

export type AIPlanDraftMeeting = {
  attendees: string[];
  location: string;
  durationMinutes?: number;
  useCustomEnd?: boolean;
  endDate?: string;
  endTime?: string;
  isAllDay?: boolean;
  teamsMeeting?: boolean;
  addGoogleMeet?: boolean;
};

export type AIPlanDraftRow = {
  title: string;
  body?: string;
  rowType: PlanRowType;
  offsetDays: number;
  dateBasis: PlanDateBasis;
  reminderTime?: string;
  emailDraft?: AIPlanDraftEmail;
  durationDraft?: AIPlanDraftDuration;
  meetingDraft?: AIPlanDraftMeeting;
  rationale?: string;
};

export type AIPlanDraft = {
  baseType: PlanType;
  templateName: string;
  eventName: string;
  anchorDate?: string;
  eventTime?: string;
  timezone?: string;
  noEventDate: boolean;
  weekendRule: WeekendRule;
  anchors: Array<{ key: string; value: string }>;
  rows: AIPlanDraftRow[];
};

export type AIChatMessage = {
  role: "user" | "assistant";
  text: string;
};

export type AIPlanBuilderContext = {
  title: string;
  planType: PlanType;
  noEventDate: boolean;
  anchorDate: string;
  weekendRule: WeekendRule;
  anchors: Array<{ key: string; value: string }>;
  rows: Array<{
    rowType: PlanRowType;
    title: string;
    offsetDays: number;
    dateBasis: PlanDateBasis;
    reminderTime?: string;
    emailSubject?: string;
    recipientCount?: number;
    attendeeCount?: number;
  }>;
};

export type AIPlanChatRequest = {
  messages: AIChatMessage[];
  currentSummary?: string;
  currentDraft?: AIPlanDraft | null;
  builderContextMode?: "refine_current" | "start_new";
  currentBuilderContext?: AIPlanBuilderContext | null;
};

export type AIPlanChatTurnResult = {
  assistantMessage: string;
  summary: string;
  followUpQuestions: string[];
  changeSummary: string[];
  confidenceNote: string;
  suggestedNextActions: string[];
  draft: AIPlanDraft | null;
  status: "needs_more_info" | "ready_to_apply";
};

export const aiPlanChatSchema = {
  name: "plans_ai_chat_turn",
  strict: true,
  schema: {
    type: "object",
    additionalProperties: false,
    properties: {
      assistantMessage: { type: "string" },
      summary: { type: "string" },
      followUpQuestions: {
        type: "array",
        items: { type: "string" },
      },
      changeSummary: {
        type: "array",
        items: { type: "string" },
      },
      confidenceNote: { type: "string" },
      suggestedNextActions: {
        type: "array",
        items: { type: "string" },
      },
      status: {
        type: "string",
        enum: ["needs_more_info", "ready_to_apply"],
      },
      draft: {
        type: "object",
        additionalProperties: false,
        properties: {
          baseType: {
            type: "string",
            enum: ["earnings", "conference", "press_release"],
          },
          templateName: { type: "string" },
          eventName: { type: "string" },
          anchorDate: { type: "string" },
          eventTime: { type: "string" },
          timezone: { type: "string" },
          noEventDate: { type: "boolean" },
          weekendRule: {
            type: "string",
            enum: ["prior_business_day", "none"],
          },
          anchors: {
            type: "array",
            items: {
              type: "object",
              additionalProperties: false,
              properties: {
                key: { type: "string" },
                value: { type: "string" },
              },
              required: ["key", "value"],
            },
          },
          rows: {
            type: "array",
            items: {
              type: "object",
              additionalProperties: false,
              properties: {
                title: { type: "string" },
                body: { type: "string" },
                rowType: {
                  type: "string",
                  enum: ["reminder", "email", "calendar_event"],
                },
                offsetDays: { type: "number" },
                dateBasis: {
                  type: "string",
                  enum: ["event", "today"],
                },
                reminderTime: { type: "string" },
                rationale: { type: "string" },
                emailDraft: {
                  type: "object",
                  additionalProperties: false,
                  properties: {
                    to: { type: "array", items: { type: "string" } },
                    cc: { type: "array", items: { type: "string" } },
                    bcc: { type: "array", items: { type: "string" } },
                    subject: { type: "string" },
                    body: { type: "string" },
                  },
                  required: ["to", "cc", "bcc", "subject", "body"],
                },
                durationDraft: {
                  type: "object",
                  additionalProperties: false,
                  properties: {
                    durationMinutes: { type: "number" },
                    useCustomEnd: { type: "boolean" },
                    endDate: { type: "string" },
                    endTime: { type: "string" },
                    isAllDay: { type: "boolean" },
                  },
                  required: ["durationMinutes", "useCustomEnd", "endDate", "endTime", "isAllDay"],
                },
                meetingDraft: {
                  type: "object",
                  additionalProperties: false,
                  properties: {
                    attendees: { type: "array", items: { type: "string" } },
                    location: { type: "string" },
                    durationMinutes: { type: "number" },
                    useCustomEnd: { type: "boolean" },
                    endDate: { type: "string" },
                    endTime: { type: "string" },
                    isAllDay: { type: "boolean" },
                    teamsMeeting: { type: "boolean" },
                  },
                  required: [
                    "attendees",
                    "location",
                    "durationMinutes",
                    "useCustomEnd",
                    "endDate",
                    "endTime",
                    "isAllDay",
                    "teamsMeeting",
                  ],
                },
              },
              required: [
                "title",
                "body",
                "rowType",
                "offsetDays",
                "dateBasis",
                "reminderTime",
                "rationale",
                "emailDraft",
                "durationDraft",
                "meetingDraft",
              ],
            },
          },
        },
        required: [
          "baseType",
          "templateName",
          "eventName",
          "anchorDate",
          "eventTime",
          "timezone",
          "noEventDate",
          "weekendRule",
          "anchors",
          "rows",
        ],
      },
    },
    required: [
      "assistantMessage",
      "summary",
      "followUpQuestions",
      "changeSummary",
      "confidenceNote",
      "suggestedNextActions",
      "status",
      "draft",
    ],
  },
} as const;

function isObject(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function normalizeStringList(value: unknown) {
  if (!Array.isArray(value)) return [];
  return value
    .filter((entry): entry is string => typeof entry === "string")
    .map((entry) => entry.trim())
    .filter(Boolean);
}

function readNormalizedString(value: unknown) {
  return typeof value === "string" ? value.trim() : "";
}

function readNumber(value: unknown, fallback = 0) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value === "string" && value.trim()) {
    const parsed = Number(value);
    if (Number.isFinite(parsed)) return parsed;
  }
  return fallback;
}

function normalizeBaseType(value: unknown): PlanType {
  return value === "conference" || value === "press_release" ? value : "earnings";
}

function normalizeWeekendRule(value: unknown): WeekendRule {
  return value === "none" ? "none" : "prior_business_day";
}

function normalizeDateBasis(value: unknown): PlanDateBasis {
  return value === "today" ? "today" : "event";
}

function normalizeRowType(value: unknown): PlanRowType {
  return value === "email" || value === "calendar_event" ? value : "reminder";
}

function normalizeDurationDraft(value: unknown): AIPlanDraftDuration | undefined {
  if (!isObject(value)) return undefined;
  return {
    durationMinutes:
      typeof value.durationMinutes === "number" && value.durationMinutes > 0 ? value.durationMinutes : undefined,
    useCustomEnd: Boolean(value.useCustomEnd),
    endDate: typeof value.endDate === "string" ? value.endDate : "",
    endTime: typeof value.endTime === "string" ? value.endTime : "",
    isAllDay: Boolean(value.isAllDay),
  };
}

function normalizeEmailDraft(value: unknown): AIPlanDraftEmail | undefined {
  if (!isObject(value)) return undefined;
  return {
    to: normalizeStringList(value.to),
    cc: normalizeStringList(value.cc),
    bcc: normalizeStringList(value.bcc),
    subject: typeof value.subject === "string" ? value.subject : "",
    body: typeof value.body === "string" ? value.body : "",
  };
}

function normalizeMeetingDraft(value: unknown): AIPlanDraftMeeting | undefined {
  if (!isObject(value)) return undefined;
  return {
    attendees: normalizeStringList(value.attendees),
    location: typeof value.location === "string" ? value.location : "",
    durationMinutes:
      typeof value.durationMinutes === "number" && value.durationMinutes > 0 ? value.durationMinutes : undefined,
    useCustomEnd: Boolean(value.useCustomEnd),
    endDate: typeof value.endDate === "string" ? value.endDate : "",
    endTime: typeof value.endTime === "string" ? value.endTime : "",
    isAllDay: Boolean(value.isAllDay),
    teamsMeeting: Boolean(value.teamsMeeting),
    addGoogleMeet: Boolean(value.addGoogleMeet),
  };
}

function normalizeRow(value: unknown): AIPlanDraftRow | null {
  if (!isObject(value)) return null;
  const rowType = normalizeRowType(value.rowType);
  const emailDraft = normalizeEmailDraft(value.emailDraft);
  const fallbackTitle =
    readNormalizedString(value.title) ||
    readNormalizedString(emailDraft?.subject) ||
    readNormalizedString(value.body).split("\n")[0]?.trim() ||
    (rowType === "email" ? "Email" : rowType === "calendar_event" ? "Meeting" : "Reminder");
  return {
    title: fallbackTitle,
    body: typeof value.body === "string" ? value.body : "",
    rowType,
    offsetDays: readNumber(value.offsetDays, 0),
    dateBasis: normalizeDateBasis(value.dateBasis),
    reminderTime: typeof value.reminderTime === "string" ? value.reminderTime : "",
    rationale: typeof value.rationale === "string" ? value.rationale : "",
    emailDraft,
    durationDraft: normalizeDurationDraft(value.durationDraft),
    meetingDraft: normalizeMeetingDraft(value.meetingDraft),
  };
}

export function parseAIPlanDraft(value: unknown): AIPlanDraft | null {
  if (!isObject(value)) return null;
  const rows = Array.isArray(value.rows)
    ? value.rows.map(normalizeRow).filter((row): row is AIPlanDraftRow => Boolean(row))
    : [];

  return {
    baseType: normalizeBaseType(value.baseType),
    templateName: typeof value.templateName === "string" ? value.templateName.trim() : "",
    eventName: typeof value.eventName === "string" ? value.eventName.trim() : "",
    anchorDate: typeof value.anchorDate === "string" ? value.anchorDate : "",
    eventTime: typeof value.eventTime === "string" ? value.eventTime : "",
    timezone: typeof value.timezone === "string" ? value.timezone : "",
    noEventDate: Boolean(value.noEventDate),
    weekendRule: normalizeWeekendRule(value.weekendRule),
    anchors: Array.isArray(value.anchors)
      ? value.anchors
          .filter((entry): entry is Record<string, unknown> => isObject(entry))
          .map((anchor) => ({
            key: typeof anchor.key === "string" ? anchor.key.trim() : "",
            value: typeof anchor.value === "string" ? anchor.value : "",
          }))
          .filter((anchor) => anchor.key)
      : [],
    rows,
  };
}

export function parseAIPlanChatTurnResult(value: unknown): AIPlanChatTurnResult | null {
  if (!isObject(value)) return null;
  const assistantMessage = readNormalizedString(value.assistantMessage) || readNormalizedString(value.summary);
  const summary = readNormalizedString(value.summary) || assistantMessage;
  if (!assistantMessage || !summary) return null;

  const draft = parseAIPlanDraft(value.draft);
  return {
    assistantMessage,
    summary,
    followUpQuestions: normalizeStringList(value.followUpQuestions).slice(0, 3),
    changeSummary: normalizeStringList(value.changeSummary).slice(0, 5),
    confidenceNote: typeof value.confidenceNote === "string" ? value.confidenceNote.trim() : "",
    suggestedNextActions: normalizeStringList(value.suggestedNextActions).slice(0, 4),
    draft,
    status: value.status === "ready_to_apply" ? "ready_to_apply" : "needs_more_info",
  };
}

export function diagnoseAIPlanChatTurnResult(value: unknown) {
  const issues: string[] = [];

  if (!isObject(value)) {
    issues.push("Top-level response was not an object.");
    return issues;
  }

  if (!readNormalizedString(value.assistantMessage) && !readNormalizedString(value.summary)) {
    issues.push("Missing assistantMessage and summary.");
  } else if (!readNormalizedString(value.assistantMessage)) {
    issues.push("Missing assistantMessage.");
  } else if (!readNormalizedString(value.summary)) {
    issues.push("Missing summary.");
  }

  if (!isObject(value.draft)) {
    issues.push("Missing draft object.");
    return issues;
  }

  const rawRows = Array.isArray(value.draft.rows) ? value.draft.rows : [];
  if (!Array.isArray(value.draft.rows)) {
    issues.push("Draft rows were missing or not an array.");
  } else if (rawRows.length === 0) {
    issues.push("Draft rows array was empty.");
  }

  const normalizedRows = rawRows.map(normalizeRow).filter((row): row is AIPlanDraftRow => Boolean(row));
  if (rawRows.length > 0 && normalizedRows.length === 0) {
    issues.push("No rows survived normalization.");
  } else if (normalizedRows.length < rawRows.length) {
    issues.push(`Some rows were dropped during normalization (${normalizedRows.length}/${rawRows.length} kept).`);
  }

  return issues;
}

export function parseAIPlanChatRequest(value: unknown): AIPlanChatRequest | null {
  if (!isObject(value)) return null;
  const messages = Array.isArray(value.messages)
    ? value.messages
        .filter((entry): entry is Record<string, unknown> => isObject(entry))
        .map((message) => ({
          role: (message.role === "assistant" ? "assistant" : "user") as AIChatMessage["role"],
          text: typeof message.text === "string" ? message.text.trim() : "",
        }))
        .filter((message) => message.text)
    : [];

  if (messages.length === 0) return null;

  const currentBuilderContext = isObject(value.currentBuilderContext)
    ? {
        title: typeof value.currentBuilderContext.title === "string" ? value.currentBuilderContext.title.trim() : "",
        planType: normalizeBaseType(value.currentBuilderContext.planType),
        noEventDate: Boolean(value.currentBuilderContext.noEventDate),
        anchorDate: typeof value.currentBuilderContext.anchorDate === "string" ? value.currentBuilderContext.anchorDate : "",
        weekendRule: normalizeWeekendRule(value.currentBuilderContext.weekendRule),
        anchors: Array.isArray(value.currentBuilderContext.anchors)
          ? value.currentBuilderContext.anchors
              .filter((entry): entry is Record<string, unknown> => isObject(entry))
              .map((anchor) => ({
                key: typeof anchor.key === "string" ? anchor.key.trim() : "",
                value: typeof anchor.value === "string" ? anchor.value : "",
              }))
              .filter((anchor) => anchor.key)
          : [],
        rows: Array.isArray(value.currentBuilderContext.rows)
          ? value.currentBuilderContext.rows
              .filter((entry): entry is Record<string, unknown> => isObject(entry))
              .map((row) => ({
                rowType: normalizeRowType(row.rowType),
                title: typeof row.title === "string" ? row.title.trim() : "",
                offsetDays: typeof row.offsetDays === "number" ? row.offsetDays : 0,
                dateBasis: normalizeDateBasis(row.dateBasis),
                reminderTime: typeof row.reminderTime === "string" ? row.reminderTime : "",
                emailSubject: typeof row.emailSubject === "string" ? row.emailSubject : "",
                recipientCount: typeof row.recipientCount === "number" ? row.recipientCount : 0,
                attendeeCount: typeof row.attendeeCount === "number" ? row.attendeeCount : 0,
              }))
          : [],
      }
    : null;

  return {
    messages,
    currentSummary: typeof value.currentSummary === "string" ? value.currentSummary.trim() : "",
    currentDraft: parseAIPlanDraft(value.currentDraft),
    builderContextMode: value.builderContextMode === "refine_current" ? "refine_current" : "start_new",
    currentBuilderContext,
  };
}
