import { parseTemplateSnapshotFile } from "./templateSnapshots";
import { migrateLegacyPersistedValue, readPersistedValue, writePersistedValue } from "./browserStorage";
import { getCachedOrgContext } from "./orgBootstrap";
import { getSupabaseBrowserClient, isSupabaseConfigured } from "./supabaseClient";
import { getLocalUserKey } from "./userKey";
import type { PlanType, WeekendRule } from "../types/plan";

export type PersistedTemplateRow = {
  id: string;
  title: string;
  body?: string;
  offsetDays: number;
  dateBasis?: "event" | "today";
  rowType: "reminder" | "email" | "calendar_event";
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

export type PersistedPlanTemplate = {
  id: string;
  name: string;
  baseType: PlanType;
  templateMode?: "template" | "custom";
  noEventDate?: boolean;
  weekendRule: WeekendRule;
  anchors: Array<{ key: string; value: string }>;
  items: PersistedTemplateRow[];
  isProtected: boolean;
  sortOrder: number;
};

export type PersistedTemplateState = {
  selectedTemplateId: string | null;
  templates: PersistedPlanTemplate[];
};

const TEMPLATE_CACHE_KEY = "event-based-reminders-app:template-cache-v1";
const LEGACY_TEMPLATE_CACHE_KEYS = ["standalone-plans:template-cache-v1"];
const NO_EVENT_DATE_ANCHOR_KEY = "__event_based_reminders_no_event_date__";

function migrateTemplateCache() {
  migrateLegacyPersistedValue("localStorage", TEMPLATE_CACHE_KEY, LEGACY_TEMPLATE_CACHE_KEYS);
}

/**
 * Recommended Supabase schema for this layer:
 *
 * user_settings:
 * - id uuid primary key default gen_random_uuid()
 * - user_key text not null unique
 * - default_reminder_time text
 * - default_press_release_time text
 * - email_signature_enabled boolean
 * - email_signature_text text
 * - outlook_account_email text
 * - email_handling_mode text
 * - updated_at timestamptz default now()
 *
 * plan_templates:
 * - id text primary key
 * - user_key text not null
 * - name text not null
 * - base_type text not null
 * - template_mode text
 * - is_protected boolean not null default false
 * - weekend_rule text not null
 * - anchors jsonb not null default '[]'::jsonb
 * - items jsonb not null default '[]'::jsonb
 * - sort_order integer not null default 0
 * - updated_at timestamptz default now()
 *
 * Note: the app already uses stable string template IDs like `seed:press_release`,
 * so this implementation keeps `plan_templates.id` text-based to preserve current behavior.
 */

function isObject(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function normalizePlanType(value: unknown): PlanType {
  if (value === "press_release" || value === "conference") return value;
  return "earnings";
}

function normalizeWeekendRule(value: unknown): WeekendRule {
  return value === "none" ? "none" : "prior_business_day";
}

function normalizeTemplateAnchors(value: unknown) {
  const rawAnchors = Array.isArray(value)
    ? value
        .filter((entry): entry is Record<string, unknown> => isObject(entry))
        .map((anchor) => ({
          key: typeof anchor.key === "string" ? anchor.key : "",
          value: typeof anchor.value === "string" ? anchor.value : "",
        }))
        .filter((anchor) => anchor.key)
    : [];

  const noEventDate = rawAnchors.some((anchor) => anchor.key === NO_EVENT_DATE_ANCHOR_KEY && anchor.value === "true");
  const anchors = rawAnchors.filter((anchor) => anchor.key !== NO_EVENT_DATE_ANCHOR_KEY);

  return { anchors, noEventDate };
}

function serializeTemplateAnchors(template: PersistedPlanTemplate) {
  const anchors = template.anchors.map((anchor) => ({ ...anchor }));
  if (template.noEventDate) {
    anchors.push({ key: NO_EVENT_DATE_ANCHOR_KEY, value: "true" });
  }
  return anchors;
}

function normalizeTemplateRow(value: unknown): PersistedTemplateRow | null {
  if (!isObject(value) || typeof value.id !== "string" || typeof value.title !== "string") {
    return null;
  }

  return {
    id: value.id,
    title: value.title,
    body: typeof value.body === "string" ? value.body : "",
    offsetDays: typeof value.offsetDays === "number" ? value.offsetDays : 0,
    dateBasis: value.dateBasis === "today" ? "today" : "event",
    rowType: value.rowType === "email" || value.rowType === "calendar_event" ? value.rowType : "reminder",
    reminderTime: typeof value.reminderTime === "string" ? value.reminderTime : "",
    emailDraft: isObject(value.emailDraft)
      ? {
          to: Array.isArray(value.emailDraft.to) ? value.emailDraft.to.filter((entry): entry is string => typeof entry === "string") : [],
          cc: Array.isArray(value.emailDraft.cc) ? value.emailDraft.cc.filter((entry): entry is string => typeof entry === "string") : [],
          bcc: Array.isArray(value.emailDraft.bcc) ? value.emailDraft.bcc.filter((entry): entry is string => typeof entry === "string") : [],
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
          attendees:
            Array.isArray(value.meetingDraft.attendees)
              ? value.meetingDraft.attendees.filter((entry): entry is string => typeof entry === "string")
              : [],
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

function normalizePersistedTemplate(value: unknown, fallbackSortOrder = 0): PersistedPlanTemplate | null {
  if (!isObject(value) || typeof value.id !== "string" || typeof value.name !== "string") {
    return null;
  }

  const normalizedAnchors = normalizeTemplateAnchors(value.anchors);
  const noEventDate = typeof value.noEventDate === "boolean" ? value.noEventDate : normalizedAnchors.noEventDate;
  const anchors = normalizedAnchors.anchors;

  const items = Array.isArray(value.items)
    ? value.items.map(normalizeTemplateRow).filter((row): row is PersistedTemplateRow => Boolean(row))
    : [];

  return {
    id: value.id,
    name: value.name,
    baseType: normalizePlanType(value.baseType),
    templateMode: value.templateMode === "custom" ? "custom" : value.templateMode === "template" ? "template" : undefined,
    noEventDate,
    weekendRule: normalizeWeekendRule(value.weekendRule),
    anchors,
    items,
    isProtected: Boolean(value.isProtected),
    sortOrder: typeof value.sortOrder === "number" ? value.sortOrder : fallbackSortOrder,
  };
}

export function loadCachedTemplateState(seedTemplates: PersistedPlanTemplate[]): PersistedTemplateState {
  if (typeof window === "undefined") {
    return {
      selectedTemplateId: seedTemplates[0]?.id ?? null,
      templates: seedTemplates,
    };
  }

  migrateTemplateCache();
  const raw = readPersistedValue("localStorage", TEMPLATE_CACHE_KEY, LEGACY_TEMPLATE_CACHE_KEYS);
  if (!raw) {
    return {
      selectedTemplateId: seedTemplates[0]?.id ?? null,
      templates: seedTemplates,
    };
  }

  try {
    const parsed = parseTemplateSnapshotFile(JSON.parse(raw));
    const cachedTemplates = parsed.templates
      .map((template, index) =>
        normalizePersistedTemplate(
          {
            ...template,
            isProtected: template.id.startsWith("seed:"),
            sortOrder: index,
          },
          index
        )
      )
      .filter((template): template is PersistedPlanTemplate => Boolean(template));

    return mergeTemplateStates(seedTemplates, {
      selectedTemplateId:
        typeof parsed.selectedTemplateId === "string" || parsed.selectedTemplateId === null ? parsed.selectedTemplateId : null,
      templates: cachedTemplates,
    });
  } catch {
    return {
      selectedTemplateId: seedTemplates[0]?.id ?? null,
      templates: seedTemplates,
    };
  }
}

export function saveCachedTemplateState(state: PersistedTemplateState) {
  if (typeof window === "undefined") return;
  migrateTemplateCache();

  const snapshot = {
    version: 1 as const,
    exportedAt: new Date().toISOString(),
    selectedTemplateId: state.selectedTemplateId,
    templates: state.templates.map((template) => ({
      id: template.id,
      name: template.name,
      baseType: template.baseType,
      templateMode: template.templateMode,
      noEventDate: Boolean(template.noEventDate),
      weekendRule: template.weekendRule,
      anchors: template.anchors.map((anchor) => ({ ...anchor })),
      items: template.items.map((item) => ({
        ...item,
        body: item.body ?? "",
      })),
    })),
  };

  writePersistedValue("localStorage", TEMPLATE_CACHE_KEY, JSON.stringify(snapshot));
}

export function mergeTemplateStates(
  seedTemplates: PersistedPlanTemplate[],
  remoteState: PersistedTemplateState | null
): PersistedTemplateState {
  if (!remoteState) {
    return {
      selectedTemplateId: seedTemplates[0]?.id ?? null,
      templates: seedTemplates,
    };
  }

  const mergedById = new Map<string, PersistedPlanTemplate>();
  const sortOrderById = new Map<string, number>();

  seedTemplates.forEach((template, index) => {
    mergedById.set(template.id, template);
    sortOrderById.set(template.id, template.sortOrder ?? index);
  });

  remoteState.templates.forEach((template, index) => {
    mergedById.set(template.id, {
      ...template,
      sortOrder: typeof template.sortOrder === "number" ? template.sortOrder : index,
    });
    sortOrderById.set(template.id, typeof template.sortOrder === "number" ? template.sortOrder : index);
  });

  const templates = Array.from(mergedById.values()).sort((left, right) => {
    const leftSort = sortOrderById.get(left.id) ?? Number.MAX_SAFE_INTEGER;
    const rightSort = sortOrderById.get(right.id) ?? Number.MAX_SAFE_INTEGER;
    if (leftSort !== rightSort) return leftSort - rightSort;
    return left.name.localeCompare(right.name);
  });

  const selectedTemplateId =
    remoteState.selectedTemplateId && templates.some((template) => template.id === remoteState.selectedTemplateId)
      ? remoteState.selectedTemplateId
      : templates[0]?.id ?? null;

  return {
    selectedTemplateId,
    templates,
  };
}

function hasMeaningfulTemplateState(state: PersistedTemplateState | null, seedTemplates: PersistedPlanTemplate[]) {
  if (!state) return false;
  const normalizedCurrent = JSON.stringify(
    state.templates.map((template) => ({
      id: template.id,
      name: template.name,
      baseType: template.baseType,
      templateMode: template.templateMode ?? null,
      noEventDate: Boolean(template.noEventDate),
      weekendRule: template.weekendRule,
      anchors: template.anchors,
      items: template.items,
      isProtected: template.isProtected,
      sortOrder: template.sortOrder,
    }))
  );
  const normalizedSeed = JSON.stringify(
    seedTemplates.map((template) => ({
      id: template.id,
      name: template.name,
      baseType: template.baseType,
      templateMode: template.templateMode ?? null,
      noEventDate: Boolean(template.noEventDate),
      weekendRule: template.weekendRule,
      anchors: template.anchors,
      items: template.items,
      isProtected: template.isProtected,
      sortOrder: template.sortOrder,
    }))
  );
  return normalizedCurrent !== normalizedSeed || state.selectedTemplateId !== (seedTemplates[0]?.id ?? null);
}

export async function loadTemplateStateFromSupabase(seedTemplates: PersistedPlanTemplate[]) {
  if (!isSupabaseConfigured()) return null;

  const supabase = getSupabaseBrowserClient();
  const orgId = getCachedOrgContext()?.orgId ?? "";
  const userKey = getLocalUserKey();

  if (!supabase) return null;

  const cachedState = loadCachedTemplateState(seedTemplates);

  if (orgId) {
    const { data, error } = await supabase
      .from("org_plan_templates")
      .select("id,name,base_type,template_mode,is_protected,weekend_rule,anchors,items,sort_order")
      .eq("org_id", orgId)
      .order("sort_order", { ascending: true });

    if (!error && data && data.length > 0) {
      const remoteTemplates = data
        .map((row, index) =>
          normalizePersistedTemplate(
            {
              id: row.id,
              name: row.name,
              baseType: row.base_type,
              templateMode: row.template_mode,
              isProtected: row.is_protected,
              weekendRule: row.weekend_rule,
              anchors: row.anchors,
              items: row.items,
              sortOrder: row.sort_order,
            },
            index
          )
        )
        .filter((template): template is PersistedPlanTemplate => Boolean(template));

      return mergeTemplateStates(seedTemplates, {
        selectedTemplateId: cachedState.selectedTemplateId,
        templates: remoteTemplates,
      });
    }

    const legacyRemoteState = userKey
      ? await (async () => {
          const { data: legacyData, error: legacyError } = await supabase
            .from("plan_templates")
            .select("id,name,base_type,template_mode,is_protected,weekend_rule,anchors,items,sort_order")
            .eq("user_key", userKey)
            .order("sort_order", { ascending: true });

          if (legacyError || !legacyData) return null;

          const legacyTemplates = legacyData
            .map((row, index) =>
              normalizePersistedTemplate(
                {
                  id: row.id,
                  name: row.name,
                  baseType: row.base_type,
                  templateMode: row.template_mode,
                  isProtected: row.is_protected,
                  weekendRule: row.weekend_rule,
                  anchors: row.anchors,
                  items: row.items,
                  sortOrder: row.sort_order,
                },
                index
              )
            )
            .filter((template): template is PersistedPlanTemplate => Boolean(template));

          return mergeTemplateStates(seedTemplates, {
            selectedTemplateId: cachedState.selectedTemplateId,
            templates: legacyTemplates,
          });
        })()
      : null;

    const sourceState = hasMeaningfulTemplateState(cachedState, seedTemplates)
      ? cachedState
      : legacyRemoteState ?? cachedState;
    if (sourceState) {
      await saveTemplateStateToSupabase(sourceState);
      return sourceState;
    }

    return mergeTemplateStates(seedTemplates, cachedState);
  }

  if (!userKey) return null;

  const { data, error } = await supabase
    .from("plan_templates")
    .select("id,name,base_type,template_mode,is_protected,weekend_rule,anchors,items,sort_order")
    .eq("user_key", userKey)
    .order("sort_order", { ascending: true });

  if (error || !data) return null;

  const remoteTemplates = data
    .map((row, index) =>
      normalizePersistedTemplate(
        {
          id: row.id,
          name: row.name,
          baseType: row.base_type,
          templateMode: row.template_mode,
          isProtected: row.is_protected,
          weekendRule: row.weekend_rule,
          anchors: row.anchors,
          items: row.items,
          sortOrder: row.sort_order,
        },
        index
      )
    )
    .filter((template): template is PersistedPlanTemplate => Boolean(template));

  return mergeTemplateStates(seedTemplates, {
    selectedTemplateId: cachedState.selectedTemplateId,
    templates: remoteTemplates,
  });
}

export async function saveTemplateStateToSupabase(state: PersistedTemplateState) {
  if (!isSupabaseConfigured()) return;

  const supabase = getSupabaseBrowserClient();
  const orgId = getCachedOrgContext()?.orgId ?? "";
  const userKey = getLocalUserKey();

  if (!supabase) return;

  if (orgId) {
    const rows = state.templates.map((template, index) => ({
      id: template.id,
      org_id: orgId,
      name: template.name,
      base_type: template.baseType,
      template_mode: template.templateMode ?? null,
      is_protected: template.isProtected,
      weekend_rule: template.weekendRule,
      anchors: serializeTemplateAnchors(template),
      items: template.items,
      sort_order: index,
    }));

    const { error: upsertError } = await supabase.from("org_plan_templates").upsert(rows, {
      onConflict: "org_id,id",
    });

    if (upsertError) return;

    const { data: existingRows, error: existingRowsError } = await supabase
      .from("org_plan_templates")
      .select("id")
      .eq("org_id", orgId);

    if (existingRowsError || !existingRows) return;

    const currentIds = new Set(state.templates.map((template) => template.id));
    const idsToDelete = existingRows
      .map((row) => row.id)
      .filter((id): id is string => typeof id === "string" && !currentIds.has(id));

    if (idsToDelete.length === 0) return;

    await supabase.from("org_plan_templates").delete().eq("org_id", orgId).in("id", idsToDelete);
    return;
  }

  if (!userKey) return;

  const rows = state.templates.map((template, index) => ({
    id: template.id,
    user_key: userKey,
    name: template.name,
    base_type: template.baseType,
    template_mode: template.templateMode ?? null,
    is_protected: template.isProtected,
    weekend_rule: template.weekendRule,
    anchors: serializeTemplateAnchors(template),
    items: template.items,
    sort_order: index,
  }));

  const { error: upsertError } = await supabase.from("plan_templates").upsert(rows, {
    onConflict: "id",
  });

  if (upsertError) return;

  const { data: existingRows, error: existingRowsError } = await supabase
    .from("plan_templates")
    .select("id")
    .eq("user_key", userKey);

  if (existingRowsError || !existingRows) return;

  const currentIds = new Set(state.templates.map((template) => template.id));
  const idsToDelete = existingRows
    .map((row) => row.id)
    .filter((id): id is string => typeof id === "string" && !currentIds.has(id));

  if (idsToDelete.length === 0) return;

  await supabase.from("plan_templates").delete().eq("user_key", userKey).in("id", idsToDelete);
}
