import type {
  Plan,
  PlanDateBasis,
  PlanItem,
  PlanItemStatus,
  PlanType,
  WeekendRule,
} from "../types/plan";
import { addDaysISO, applyWeekendRule, todayYYYYMMDD } from "./dateUtils";

export type TemplateItem = {
  id?: string;
  title: string;
  body?: string;
  offsetDays: number;
  dateBasis?: PlanDateBasis;
  rowType?: "reminder" | "email" | "calendar_event";
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

export const TEMPLATES: Record<PlanType, TemplateItem[]> = {
  earnings: [
    { title: "Draft Earnings Curtain Raiser Press Release", offsetDays: -21 },
    { title: "Draft Earnings Press Release", offsetDays: -14 },
    { title: "Prepare Press Release Distribution For Tomorrow [Curtain Raiser]", offsetDays: -8 },
    {
      title: "Distribute Press Release & Post To Social Media Channels [Curtain Raiser]",
      offsetDays: -7,
      reminderTime: "08:30",
    },
    { title: "Send Chorus Call Introductory Script & Authorized Callers", offsetDays: -1 },
    { title: "Prepare Earnings Press Release Distribution For Tomorrow", offsetDays: -1 },
    { title: "Turn Off Morning Alarm & Set Alarm For 5:30 AM", offsetDays: -1 },
    { title: "Earnings Call Script Walk Through", offsetDays: -1 },
    { title: "Distribute Press Release [Earnings] & Post To Social Media Channels", offsetDays: 0 },
    {
      title: "Print Earnings Scripts & Press Release For Team",
      offsetDays: 0,
      reminderTime: "06:45",
    },
    { title: "Final Walk Through Earnings Call Script", offsetDays: 0 },
  ],
  conference: [
    { title: "Draft Upcoming Conference Press Release", offsetDays: -21 },
    { title: "Register For Webcast & Send Webcast Link To Website Designer", offsetDays: -21 },
    { title: "Set Up Flights & Hotels", offsetDays: -20 },
    { title: "Confirm Corporate Presentation With Mitch", offsetDays: -14 },
    { title: "Prepare Investor Briefs", offsetDays: -7 },
    { title: "Print Deck", offsetDays: -7 },
    { title: "Review Conference Agenda / Presentation Time", offsetDays: -7 },
    { title: "Pack Business Cards & Presentations", offsetDays: -1 },
    { title: "Add New Corporate Presentation / Webcast To Website", offsetDays: 1 },
    {
      title: "Draft Email To Dawn Fitzpatrick For Hotel/Flight Support",
      offsetDays: -20,
      rowType: "email",
      emailDraft: { to: [], cc: [], bcc: [], subject: "", body: "" },
    },
  ],
  press_release: [
    {
      title: "Prepare Press Release Distribution For Tomorrow",
      body: `Check Forward Looks Statement with Legal
Submit To Notified
Submit To Nasdaq IssuerEntry
Send Press Release (Word & PDF to Web Designer)
Prepare Constant Contact Email Blast
Review Proof of Announcement & Schedule for Release`,
      offsetDays: -1,
    },
    {
      title: "Review Proof Of Announcement & Schedule For Release",
      offsetDays: -1,
      reminderTime: "15:00",
    },
    {
      title: "Distribute Press Release & Post To Social Media Channels",
      offsetDays: 0,
    },
  ],
};

function makeId(prefix: string) {
  return `${prefix}_${Math.random().toString(36).slice(2, 10)}`;
}

function nowISO() {
  return new Date().toISOString();
}

function diffDays(fromDate: string, toDate: string): number {
  const [fromY, fromM, fromD] = fromDate.split("-").map(Number);
  const [toY, toM, toD] = toDate.split("-").map(Number);
  const from = new Date(fromY, (fromM ?? 1) - 1, fromD ?? 1);
  const to = new Date(toY, (toM ?? 1) - 1, toD ?? 1);
  const msPerDay = 24 * 60 * 60 * 1000;
  return Math.round((to.getTime() - from.getTime()) / msPerDay);
}

export function computeItems(
  template: TemplateItem[],
  anchorDate: string,
  weekendRule: WeekendRule,
  preserveFrom?: PlanItem[]
): PlanItem[] {
  return template.map((tpl) => {
    const effectiveAnchorDate = tpl.dateBasis === "today" ? todayYYYYMMDD() : anchorDate;
    const raw = addDaysISO(effectiveAnchorDate, tpl.offsetDays);
    const due = applyWeekendRule(raw, weekendRule);

    const preserved = preserveFrom?.find(
      (x) =>
        x.title === tpl.title &&
        x.offsetDays === tpl.offsetDays &&
        (x.dateBasis ?? "event") === (tpl.dateBasis ?? "event")
    );

    return {
      id: tpl.id ?? preserved?.id ?? makeId("item"),
      title: tpl.title,
      customTitle: preserved?.customTitle,
      body: preserved?.body ?? tpl.body,
      rowType: preserved?.rowType ?? tpl.rowType ?? "reminder",
      offsetDays: tpl.offsetDays,
      dateBasis: preserved?.dateBasis ?? tpl.dateBasis ?? "event",
      reminderTime: preserved?.reminderTime ?? tpl.reminderTime,
      rawDueDate: raw,
      dueDate: due,
      customDueDate: preserved?.customDueDate,
      wasAdjusted: due !== raw,
      emailDraft: preserved?.emailDraft ?? tpl.emailDraft,
      durationDraft: preserved?.durationDraft ?? tpl.durationDraft,
      meetingDraft: preserved?.meetingDraft ?? tpl.meetingDraft,
      status: preserved?.status ?? ("not_started" as PlanItemStatus),
    };
  });
}

export function createPlan(args: {
  id?: string;
  name: string;
  type: PlanType;
  anchorDate: string;
  weekendRule: WeekendRule;
  template?: TemplateItem[];
}): Plan {
  const template = args.template ?? TEMPLATES[args.type];
  const items = computeItems(template, args.anchorDate, args.weekendRule);

  const ts = nowISO();
  return {
    id: args.id ?? makeId("plan"),
    name: args.name.trim() || "Untitled Plan",
    type: args.type,
    anchorDate: args.anchorDate,
    weekendRule: args.weekendRule,
    version: 1,
    items,
    createdAt: ts,
    updatedAt: ts,
  };
}

export function updateItemStatus(plan: Plan, itemId: string, status: PlanItemStatus): Plan {
  const ts = nowISO();
  return {
    ...plan,
    updatedAt: ts,
    items: plan.items.map((it) => (it.id === itemId ? { ...it, status } : it)),
  };
}

export function regeneratePlan(plan: Plan, updates: { anchorDate?: string; weekendRule?: WeekendRule }): Plan {
  const nextAnchor = updates.anchorDate ?? plan.anchorDate;
  const nextRule = updates.weekendRule ?? plan.weekendRule;

  const didChange = nextAnchor !== plan.anchorDate || nextRule !== plan.weekendRule;
  const anchorDeltaDays = diffDays(plan.anchorDate, nextAnchor);

  const baseTemplate: TemplateItem[] = plan.items.map((it) => ({
    title: it.title,
    body: it.body,
    offsetDays: it.offsetDays,
    dateBasis: it.dateBasis,
  }));

  const items = computeItems(baseTemplate, nextAnchor, nextRule, plan.items).map((item) => {
    const preserved = plan.items.find((existing) => existing.id === item.id);
    if (!preserved?.customDueDate) return item;
    return {
      ...item,
      customDueDate: addDaysISO(preserved.customDueDate, anchorDeltaDays),
    };
  });

  const ts = nowISO();
  return {
    ...plan,
    anchorDate: nextAnchor,
    weekendRule: nextRule,
    version: didChange ? plan.version + 1 : plan.version,
    items,
    updatedAt: ts,
  };
}
