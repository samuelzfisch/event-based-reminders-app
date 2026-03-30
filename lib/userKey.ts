import { migrateLegacyPersistedValue, readPersistedValue, writePersistedValue } from "./browserStorage";

const USER_KEY_STORAGE_KEY = "event-based-reminders-app:user-key";
const LEGACY_USER_KEY_STORAGE_KEYS = ["standalone-plans:user-key"];

function generateUserKey() {
  if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") {
    return crypto.randomUUID();
  }

  return `user_${Math.random().toString(36).slice(2, 10)}${Date.now().toString(36)}`;
}

export function getLocalUserKey() {
  if (typeof window === "undefined") return "";

  migrateLegacyPersistedValue("localStorage", USER_KEY_STORAGE_KEY, LEGACY_USER_KEY_STORAGE_KEYS);
  const existing = readPersistedValue("localStorage", USER_KEY_STORAGE_KEY);
  if (existing) return existing;

  const nextKey = generateUserKey();
  writePersistedValue("localStorage", USER_KEY_STORAGE_KEY, nextKey);
  return nextKey;
}
