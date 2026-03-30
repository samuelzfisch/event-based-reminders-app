import type { Plan, PlanItem, PlanRowType } from "../types/plan";

export type AnchorMap = Map<string, string>;
export type InterpretedPlanRowKind = "email" | "meeting" | "reminder";

type AnchorPair = {
  key: string;
  value: string;
};

type RowLike = {
  rowType?: PlanRowType;
  meetingDraft?: PlanItem["meetingDraft"] | null;
};

type EmailDraftLike = NonNullable<PlanItem["emailDraft"]>;
type MeetingDraftLike = NonNullable<PlanItem["meetingDraft"]>;

export function normalizeAnchorKey(value: string) {
  return value.trim().replace(/^\[(.*)\]$/, "$1").trim().toUpperCase();
}

export function buildAnchorMap<T extends AnchorPair>(anchors: T[]): AnchorMap {
  const map = new Map<string, string>();
  for (const anchor of anchors) {
    const key = normalizeAnchorKey(anchor.key);
    if (!key) continue;
    map.set(key, anchor.value);
  }
  return map;
}

export function replaceAnchorsInText(value: unknown, anchorMap?: AnchorMap) {
  const safeValue = typeof value === "string" ? value : value == null ? "" : String(value);
  if (!safeValue) return safeValue;
  if (!anchorMap || anchorMap.size === 0) return safeValue;

  return safeValue.replace(/\[([^\]]+)\]/g, (match, rawKey) => {
    const replacement = anchorMap.get(normalizeAnchorKey(String(rawKey ?? "")));
    return replacement ?? match;
  });
}

export function resolveReminderTimeValue(value: string | undefined, anchorMap?: AnchorMap) {
  if (typeof value !== "string") return "";
  return replaceAnchorsInText(value, anchorMap).trim();
}

export function resolveEmailDraftAnchors(
  draft: EmailDraftLike | undefined,
  anchorMap?: AnchorMap
): EmailDraftLike | undefined {
  if (!draft) return draft;
  return {
    ...draft,
    to: (draft.to ?? []).map((value) => replaceAnchorsInText(value, anchorMap)),
    cc: (draft.cc ?? []).map((value) => replaceAnchorsInText(value, anchorMap)),
    bcc: (draft.bcc ?? []).map((value) => replaceAnchorsInText(value, anchorMap)),
    subject: draft.subject ? replaceAnchorsInText(draft.subject, anchorMap) : draft.subject,
    body: draft.body ? replaceAnchorsInText(draft.body, anchorMap) : draft.body,
  };
}

export function resolveMeetingDraftAnchors(
  draft: MeetingDraftLike | undefined,
  anchorMap?: AnchorMap
): MeetingDraftLike | undefined {
  if (!draft) return draft;
  return {
    ...draft,
    attendees: (draft.attendees ?? []).map((value) => replaceAnchorsInText(value, anchorMap)),
    location: draft.location ? replaceAnchorsInText(draft.location, anchorMap) : draft.location,
  };
}

export function resolvePlanItemAnchors(item: PlanItem, anchorMap?: AnchorMap): PlanItem {
  return {
    ...item,
    title: replaceAnchorsInText(item.title, anchorMap),
    customTitle: item.customTitle ? replaceAnchorsInText(item.customTitle, anchorMap) : item.customTitle,
    body: item.body ? replaceAnchorsInText(item.body, anchorMap) : item.body,
    reminderTime: item.reminderTime ? resolveReminderTimeValue(item.reminderTime, anchorMap) : item.reminderTime,
    emailDraft: resolveEmailDraftAnchors(item.emailDraft, anchorMap),
    meetingDraft: resolveMeetingDraftAnchors(item.meetingDraft, anchorMap),
  };
}

export function resolvePlanAnchors(plan: Plan, anchorMap?: AnchorMap): Plan {
  if (!anchorMap || anchorMap.size === 0) return plan;
  return {
    ...plan,
    name: replaceAnchorsInText(plan.name, anchorMap),
    items: plan.items.map((item) => resolvePlanItemAnchors(item, anchorMap)),
  };
}

export function classifyPlanRow(row: RowLike): InterpretedPlanRowKind {
  if (row.rowType === "email") return "email";
  if (row.meetingDraft) return "meeting";
  return "reminder";
}

export function partitionPlanItemsByKind(items: Plan["items"]) {
  const emailItems: Plan["items"] = [];
  const meetingItems: Plan["items"] = [];
  const calendarItems: Plan["items"] = [];

  for (const item of items) {
    const rowKind = classifyPlanRow(item);
    if (rowKind === "email") {
      emailItems.push(item);
      continue;
    }
    if (rowKind === "meeting") {
      meetingItems.push(item);
      calendarItems.push(item);
      continue;
    }
    calendarItems.push(item);
  }

  return {
    emailItems,
    meetingItems,
    calendarItems,
    reminderItems: calendarItems.filter((item) => classifyPlanRow(item) === "reminder"),
  };
}
