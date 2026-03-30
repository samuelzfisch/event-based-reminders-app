import type { PlanDateBasis, PlanRowType, PlanType, WeekendRule } from "../types/plan";

export type TemplateSnapshotRow = {
  id: string;
  title: string;
  body?: string;
  offsetDays: number;
  dateBasis?: PlanDateBasis;
  rowType: PlanRowType;
  reminderTime?: string;
  emailDraft?: {
    to?: string[];
    cc?: string[];
    bcc?: string[];
    subject?: string;
    body?: string;
  };
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

export type TemplateSnapshotTemplate = {
  id: string;
  name: string;
  baseType: PlanType;
  templateMode?: "template" | "custom";
  noEventDate?: boolean;
  weekendRule: WeekendRule;
  anchors: Array<{ key: string; value: string }>;
  items: TemplateSnapshotRow[];
};

export type TemplateSnapshotFile = {
  version: 1;
  exportedAt: string;
  selectedTemplateId?: string | null;
  templates: TemplateSnapshotTemplate[];
};

function isObject(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function normalizeRowType(value: unknown): PlanRowType {
  return value === "email" || value === "calendar_event" ? value : "reminder";
}

function normalizeDateBasis(value: unknown): PlanDateBasis {
  return value === "today" ? "today" : "event";
}

function normalizePlanType(value: unknown): PlanType {
  if (value === "conference" || value === "press_release") return value;
  return "earnings";
}

function normalizeWeekendRule(value: unknown): WeekendRule {
  return value === "none" ? "none" : "prior_business_day";
}

function normalizeStringArray(value: unknown) {
  return Array.isArray(value) ? value.filter((entry): entry is string => typeof entry === "string") : [];
}

function parseSnapshotRow(value: unknown): TemplateSnapshotRow | null {
  if (!isObject(value) || typeof value.id !== "string" || typeof value.title !== "string") return null;
  return {
    id: value.id,
    title: value.title,
    body: typeof value.body === "string" ? value.body : undefined,
    offsetDays: typeof value.offsetDays === "number" ? value.offsetDays : 0,
    dateBasis: normalizeDateBasis(value.dateBasis),
    rowType: normalizeRowType(value.rowType),
    reminderTime: typeof value.reminderTime === "string" ? value.reminderTime : undefined,
    emailDraft: isObject(value.emailDraft)
      ? {
          to: normalizeStringArray(value.emailDraft.to),
          cc: normalizeStringArray(value.emailDraft.cc),
          bcc: normalizeStringArray(value.emailDraft.bcc),
          subject: typeof value.emailDraft.subject === "string" ? value.emailDraft.subject : "",
          body: typeof value.emailDraft.body === "string" ? value.emailDraft.body : "",
        }
      : undefined,
    durationDraft: isObject(value.durationDraft)
      ? {
          durationMinutes:
            typeof value.durationDraft.durationMinutes === "number" ? value.durationDraft.durationMinutes : undefined,
          useCustomEnd:
            typeof value.durationDraft.useCustomEnd === "boolean" ? value.durationDraft.useCustomEnd : undefined,
          endDate: typeof value.durationDraft.endDate === "string" ? value.durationDraft.endDate : undefined,
          endTime: typeof value.durationDraft.endTime === "string" ? value.durationDraft.endTime : undefined,
          isAllDay: typeof value.durationDraft.isAllDay === "boolean" ? value.durationDraft.isAllDay : undefined,
        }
      : undefined,
    meetingDraft: isObject(value.meetingDraft)
      ? {
          attendees: normalizeStringArray(value.meetingDraft.attendees),
          location: typeof value.meetingDraft.location === "string" ? value.meetingDraft.location : "",
          durationMinutes:
            typeof value.meetingDraft.durationMinutes === "number" ? value.meetingDraft.durationMinutes : undefined,
          useCustomEnd:
            typeof value.meetingDraft.useCustomEnd === "boolean" ? value.meetingDraft.useCustomEnd : undefined,
          endDate: typeof value.meetingDraft.endDate === "string" ? value.meetingDraft.endDate : undefined,
          endTime: typeof value.meetingDraft.endTime === "string" ? value.meetingDraft.endTime : undefined,
          isAllDay: typeof value.meetingDraft.isAllDay === "boolean" ? value.meetingDraft.isAllDay : undefined,
          teamsMeeting:
            typeof value.meetingDraft.teamsMeeting === "boolean" ? value.meetingDraft.teamsMeeting : undefined,
        }
      : undefined,
  };
}

function parseSnapshotTemplate(value: unknown): TemplateSnapshotTemplate | null {
  if (!isObject(value) || typeof value.id !== "string" || typeof value.name !== "string") return null;
  const anchors = Array.isArray(value.anchors)
    ? value.anchors
        .filter((entry): entry is Record<string, unknown> => isObject(entry))
        .map((anchor) => ({
          key: typeof anchor.key === "string" ? anchor.key : "",
          value: typeof anchor.value === "string" ? anchor.value : "",
        }))
        .filter((anchor) => anchor.key)
    : [];
  const items = Array.isArray(value.items)
    ? value.items.map(parseSnapshotRow).filter((row): row is TemplateSnapshotRow => Boolean(row))
    : [];
  return {
    id: value.id,
    name: value.name,
    baseType: normalizePlanType(value.baseType),
    templateMode: value.templateMode === "custom" ? "custom" : value.templateMode === "template" ? "template" : undefined,
    noEventDate: Boolean(value.noEventDate),
    weekendRule: normalizeWeekendRule(value.weekendRule),
    anchors,
    items,
  };
}

export function parseTemplateSnapshotFile(value: unknown): TemplateSnapshotFile {
  if (!isObject(value) || !Array.isArray(value.templates)) {
    throw new Error("Template snapshot must contain a templates array.");
  }

  const templates = value.templates
    .map(parseSnapshotTemplate)
    .filter((template): template is TemplateSnapshotTemplate => Boolean(template));

  if (templates.length === 0) {
    throw new Error("Template snapshot did not contain any valid templates.");
  }

  return {
    version: 1,
    exportedAt: typeof value.exportedAt === "string" ? value.exportedAt : new Date().toISOString(),
    selectedTemplateId:
      typeof value.selectedTemplateId === "string" || value.selectedTemplateId === null
        ? value.selectedTemplateId
        : null,
    templates,
  };
}
