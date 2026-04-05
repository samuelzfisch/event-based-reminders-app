import { getSupabaseBrowserClient, isSupabaseConfigured } from "./supabaseClient";
import { migrateLegacyPersistedValue, readPersistedValue, writePersistedValue } from "./browserStorage";
import { getCachedOrgContext } from "./orgBootstrap";
import { getLocalUserKey } from "./userKey";

export type EmailHandlingMode = "draft" | "schedule" | "send";
export type OutlookConnectionStatus = "connected" | "reconnect_required" | "not_connected";

export type AppSettings = {
  defaultReminderTime: string;
  defaultPressReleaseTime: string;
  emailSignatureEnabled: boolean;
  emailSignatureText: string;
  outlookAccountEmail: string;
  outlookConnectionStatus: OutlookConnectionStatus;
  emailHandlingMode: EmailHandlingMode;
};

type StoredAppSettingsRecord = {
  settings: AppSettings;
  updatedAt: string;
};

export const APP_SETTINGS_STORAGE_KEY = "event-based-reminders-app:app-settings";
export const APP_SETTINGS_UPDATED_EVENT = "event-based-reminders-app:app-settings-updated";
const LEGACY_APP_SETTINGS_STORAGE_KEYS = ["standalone-plans:app-settings"];

export const DEFAULT_APP_SETTINGS: AppSettings = {
  defaultReminderTime: "",
  defaultPressReleaseTime: "",
  emailSignatureEnabled: false,
  emailSignatureText: "",
  outlookAccountEmail: "",
  outlookConnectionStatus: "not_connected",
  emailHandlingMode: "draft",
};

function normalizeEmailHandlingMode(value: unknown): EmailHandlingMode {
  return value === "schedule" || value === "send" ? value : "draft";
}

function normalizeOutlookConnectionStatus(value: unknown): OutlookConnectionStatus {
  return value === "connected" || value === "reconnect_required" ? value : "not_connected";
}

export function normalizeAppSettings(value: unknown): AppSettings {
  const record = value && typeof value === "object" ? (value as Record<string, unknown>) : {};
  return {
    defaultReminderTime:
      typeof record.defaultReminderTime === "string" ? record.defaultReminderTime : DEFAULT_APP_SETTINGS.defaultReminderTime,
    defaultPressReleaseTime:
      typeof record.defaultPressReleaseTime === "string"
        ? record.defaultPressReleaseTime
        : DEFAULT_APP_SETTINGS.defaultPressReleaseTime,
    emailSignatureEnabled:
      typeof record.emailSignatureEnabled === "boolean"
        ? record.emailSignatureEnabled
        : DEFAULT_APP_SETTINGS.emailSignatureEnabled,
    emailSignatureText:
      typeof record.emailSignatureText === "string" ? record.emailSignatureText : DEFAULT_APP_SETTINGS.emailSignatureText,
    outlookAccountEmail:
      typeof record.outlookAccountEmail === "string" ? record.outlookAccountEmail : DEFAULT_APP_SETTINGS.outlookAccountEmail,
    outlookConnectionStatus: normalizeOutlookConnectionStatus(record.outlookConnectionStatus),
    emailHandlingMode: normalizeEmailHandlingMode(record.emailHandlingMode),
  };
}

export function areAppSettingsEqual(left: AppSettings, right: AppSettings) {
  return JSON.stringify(normalizeAppSettings(left)) === JSON.stringify(normalizeAppSettings(right));
}

function migrateAppSettingsStorage() {
  migrateLegacyPersistedValue("localStorage", APP_SETTINGS_STORAGE_KEY, LEGACY_APP_SETTINGS_STORAGE_KEYS);
}

function getTimestamp(value: string | null | undefined) {
  const timestamp = Date.parse(String(value ?? ""));
  return Number.isFinite(timestamp) ? timestamp : 0;
}

function parseStoredAppSettingsRecord(raw: string): StoredAppSettingsRecord | null {
  try {
    const parsed = JSON.parse(raw) as Record<string, unknown>;
    return {
      settings: normalizeAppSettings(parsed),
      updatedAt: typeof parsed.updatedAt === "string" ? parsed.updatedAt : "",
    };
  } catch {
    return null;
  }
}

function loadStoredAppSettingsRecord() {
  if (typeof window === "undefined") return null;
  migrateAppSettingsStorage();
  const raw = readPersistedValue("localStorage", APP_SETTINGS_STORAGE_KEY, LEGACY_APP_SETTINGS_STORAGE_KEYS);
  if (!raw) return null;
  return parseStoredAppSettingsRecord(raw);
}

export function loadAppSettings(): AppSettings {
  return loadStoredAppSettingsRecord()?.settings ?? DEFAULT_APP_SETTINGS;
}

function cacheAppSettings(normalized: AppSettings, updatedAt = new Date().toISOString()) {
  if (typeof window === "undefined") return false;
  migrateAppSettingsStorage();
  const serialized = JSON.stringify({ ...normalized, updatedAt });
  const didWrite = writePersistedValue("localStorage", APP_SETTINGS_STORAGE_KEY, serialized);
  if (!didWrite) return false;
  window.dispatchEvent(new CustomEvent(APP_SETTINGS_UPDATED_EVENT, { detail: normalized }));
  return true;
}

async function persistAppSettingsToSupabase(normalized: AppSettings, updatedAt: string) {
  if (!isSupabaseConfigured()) return;

  const supabase = getSupabaseBrowserClient();
  const orgId = getCachedOrgContext()?.orgId ?? "";
  const userKey = getLocalUserKey();

  if (!supabase) return;

  if (orgId) {
    await supabase.from("org_settings").upsert(
      {
        org_id: orgId,
        default_reminder_time: normalized.defaultReminderTime,
        default_press_release_time: normalized.defaultPressReleaseTime,
        email_signature_enabled: normalized.emailSignatureEnabled,
        email_signature_text: normalized.emailSignatureText,
        outlook_account_email: normalized.outlookAccountEmail,
        email_handling_mode: normalized.emailHandlingMode,
        updated_at: updatedAt,
      },
      { onConflict: "org_id" }
    );
    return;
  }

  if (!userKey) return;

  await supabase.from("user_settings").upsert(
    {
      user_key: userKey,
      default_reminder_time: normalized.defaultReminderTime,
      default_press_release_time: normalized.defaultPressReleaseTime,
      email_signature_enabled: normalized.emailSignatureEnabled,
      email_signature_text: normalized.emailSignatureText,
      outlook_account_email: normalized.outlookAccountEmail,
      email_handling_mode: normalized.emailHandlingMode,
      updated_at: updatedAt,
    },
    { onConflict: "user_key" }
  );
}

async function loadLegacyAppSettingsFromSupabase() {
  const supabase = getSupabaseBrowserClient();
  const userKey = getLocalUserKey();

  if (!supabase || !userKey) return null;

  const { data, error } = await supabase
    .from("user_settings")
    .select(
      "default_reminder_time,default_press_release_time,email_signature_enabled,email_signature_text,outlook_account_email,email_handling_mode,updated_at"
    )
    .eq("user_key", userKey)
    .maybeSingle();

  if (error || !data) return null;

  return {
    settings: normalizeAppSettings({
      defaultReminderTime: data.default_reminder_time,
      defaultPressReleaseTime: data.default_press_release_time,
      emailSignatureEnabled: data.email_signature_enabled,
      emailSignatureText: data.email_signature_text,
      outlookAccountEmail: data.outlook_account_email,
      outlookConnectionStatus: loadAppSettings().outlookConnectionStatus,
      emailHandlingMode: data.email_handling_mode,
    }),
    updatedAt: typeof data.updated_at === "string" ? data.updated_at : "",
  } satisfies StoredAppSettingsRecord;
}

function hasMeaningfulSettings(record: StoredAppSettingsRecord | null) {
  if (!record) return false;
  return !areAppSettingsEqual(record.settings, DEFAULT_APP_SETTINGS);
}

export function saveAppSettings(nextSettings: AppSettings) {
  const normalized = normalizeAppSettings(nextSettings);
  if (typeof window === "undefined") return;
  const updatedAt = new Date().toISOString();
  const didChange = cacheAppSettings(normalized, updatedAt);
  if (!didChange) return;
  void persistAppSettingsToSupabase(normalized, updatedAt);
}

export async function hydrateAppSettingsFromSupabase() {
  if (typeof window === "undefined" || !isSupabaseConfigured()) {
    return loadAppSettings();
  }

  const supabase = getSupabaseBrowserClient();
  const orgId = getCachedOrgContext()?.orgId ?? "";
  const userKey = getLocalUserKey();
  const localRecord = loadStoredAppSettingsRecord();

  if (!supabase) return loadAppSettings();

  if (orgId) {
    const { data, error } = await supabase
      .from("org_settings")
      .select(
        "default_reminder_time,default_press_release_time,email_signature_enabled,email_signature_text,outlook_account_email,email_handling_mode,updated_at"
      )
      .eq("org_id", orgId)
      .maybeSingle();

    if (!error && data) {
      const normalized = normalizeAppSettings({
        defaultReminderTime: data.default_reminder_time,
        defaultPressReleaseTime: data.default_press_release_time,
        emailSignatureEnabled: data.email_signature_enabled,
        emailSignatureText: data.email_signature_text,
        outlookAccountEmail: data.outlook_account_email,
        outlookConnectionStatus: loadAppSettings().outlookConnectionStatus,
        emailHandlingMode: data.email_handling_mode,
      });

      cacheAppSettings(normalized, typeof data.updated_at === "string" ? data.updated_at : new Date().toISOString());
      return normalized;
    }

    const legacyRemoteRecord = await loadLegacyAppSettingsFromSupabase();
    const localUpdatedAt = getTimestamp(localRecord?.updatedAt);
    const legacyRemoteUpdatedAt = getTimestamp(legacyRemoteRecord?.updatedAt);
    const sourceRecord =
      hasMeaningfulSettings(localRecord) && localUpdatedAt >= legacyRemoteUpdatedAt
        ? localRecord
        : hasMeaningfulSettings(legacyRemoteRecord)
          ? legacyRemoteRecord
          : hasMeaningfulSettings(localRecord)
            ? localRecord
            : null;

    if (sourceRecord) {
      await persistAppSettingsToSupabase(sourceRecord.settings, sourceRecord.updatedAt || new Date().toISOString());
      return sourceRecord.settings;
    }

    return loadAppSettings();
  }

  if (!userKey) return loadAppSettings();

  const { data, error } = await supabase
    .from("user_settings")
    .select(
      "default_reminder_time,default_press_release_time,email_signature_enabled,email_signature_text,outlook_account_email,email_handling_mode,updated_at"
    )
    .eq("user_key", userKey)
    .maybeSingle();

  if (error || !data) {
    return loadAppSettings();
  }

  const localUpdatedAt = getTimestamp(localRecord?.updatedAt);
  const remoteUpdatedAt = getTimestamp(typeof data.updated_at === "string" ? data.updated_at : "");

  if (localRecord && localUpdatedAt > remoteUpdatedAt) {
    void persistAppSettingsToSupabase(localRecord.settings, localRecord.updatedAt || new Date().toISOString());
    return localRecord.settings;
  }

  const normalized = normalizeAppSettings({
    defaultReminderTime: data.default_reminder_time,
    defaultPressReleaseTime: data.default_press_release_time,
    emailSignatureEnabled: data.email_signature_enabled,
    emailSignatureText: data.email_signature_text,
    outlookAccountEmail: data.outlook_account_email,
    outlookConnectionStatus: loadAppSettings().outlookConnectionStatus,
    emailHandlingMode: data.email_handling_mode,
  });

  cacheAppSettings(normalized, typeof data.updated_at === "string" ? data.updated_at : new Date().toISOString());
  return normalized;
}
