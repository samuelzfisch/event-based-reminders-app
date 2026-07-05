import { getCachedOrgContext } from "./orgBootstrap";
import { getLocalUserKey } from "./userKey";

const LOCAL_STORAGE_PREFIXES = [
  "event-based-reminders-app:",
  "standalone-plans:",
  "event_based_reminders_app_",
  "standalone_plans_",
];

function buildPersistenceScope() {
  if (typeof window === "undefined") return "server";

  const orgContext = getCachedOrgContext();
  if (orgContext?.orgId && orgContext.userId) {
    return `org:${orgContext.orgId}:user:${orgContext.userId}`;
  }

  const localUserKey = getLocalUserKey();
  return localUserKey ? `local:${localUserKey}` : "anonymous";
}

export function getScopedStorageKey(baseKey: string) {
  return `${baseKey}:${buildPersistenceScope()}`;
}

export function clearAllLocalAppState() {
  if (typeof window === "undefined") return;

  for (const storage of [window.localStorage, window.sessionStorage]) {
    const keysToRemove: string[] = [];
    for (let index = 0; index < storage.length; index += 1) {
      const key = storage.key(index);
      if (!key) continue;
      if (LOCAL_STORAGE_PREFIXES.some((prefix) => key.startsWith(prefix))) {
        keysToRemove.push(key);
      }
    }
    keysToRemove.forEach((key) => storage.removeItem(key));
  }

  if (typeof document !== "undefined") {
    document.cookie = "event_based_reminders_app_outlook_oauth_verifier=; Path=/; Max-Age=0; SameSite=Lax; Secure";
    document.cookie = "event_based_reminders_app_gmail_oauth_verifier=; Path=/; Max-Age=0; SameSite=Lax; Secure";
  }
}
