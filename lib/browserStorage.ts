type BrowserStorageKind = "localStorage" | "sessionStorage";

function getStorage(kind: BrowserStorageKind): Storage | null {
  if (typeof window === "undefined") return null;
  return window[kind];
}

export function readPersistedValue(kind: BrowserStorageKind, primaryKey: string, legacyKeys: string[] = []) {
  const storage = getStorage(kind);
  if (!storage) return null;

  const primaryValue = storage.getItem(primaryKey);
  if (primaryValue !== null) return primaryValue;

  for (const legacyKey of legacyKeys) {
    const legacyValue = storage.getItem(legacyKey);
    if (legacyValue !== null) return legacyValue;
  }

  return null;
}

export function migrateLegacyPersistedValue(kind: BrowserStorageKind, primaryKey: string, legacyKeys: string[] = []) {
  const storage = getStorage(kind);
  if (!storage) return null;

  const primaryValue = storage.getItem(primaryKey);
  if (primaryValue !== null) return primaryValue;

  for (const legacyKey of legacyKeys) {
    const legacyValue = storage.getItem(legacyKey);
    if (legacyValue !== null) {
      storage.setItem(primaryKey, legacyValue);
      return legacyValue;
    }
  }

  return null;
}

export function writePersistedValue(
  kind: BrowserStorageKind,
  primaryKey: string,
  value: string
) {
  const storage = getStorage(kind);
  if (!storage) return false;

  if (storage.getItem(primaryKey) === value) return false;

  storage.setItem(primaryKey, value);
  return true;
}

export function removePersistedValue(kind: BrowserStorageKind, primaryKey: string, legacyKeys: string[] = []) {
  const storage = getStorage(kind);
  if (!storage) return false;

  let removed = false;

  for (const key of [primaryKey, ...legacyKeys]) {
    if (storage.getItem(key) === null) continue;
    storage.removeItem(key);
    removed = true;
  }

  return removed;
}
