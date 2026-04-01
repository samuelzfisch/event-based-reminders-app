"use client";

import {
  migrateLegacyPersistedValue,
  readPersistedValue,
  removePersistedValue,
  writePersistedValue,
} from "./browserStorage";

export type OutlookConnectionStatus = "connected" | "reconnect_required" | "not_connected";
export type OutlookEmailExecutionMode = "draft" | "schedule" | "send";

export type OutlookSession = {
  accessToken: string;
  refreshToken?: string;
  expiresAt: number;
  scope?: string;
  obtainedAt: string;
};

export type OutlookConnectedIdentity = {
  id: string;
  mail?: string;
  userPrincipalName?: string;
  displayName?: string;
  connectedAt: string;
  normalizedEmail: string;
  mailboxEligible: boolean;
  mailboxEligibilityReason: string | null;
};

export type OutlookConnectionState = {
  status: OutlookConnectionStatus;
  connected: boolean;
  stale: boolean;
  supportedMailbox: boolean;
  expectedEmail: string;
  identity: OutlookConnectedIdentity | null;
  normalizedConnectedEmail: string;
  debug: {
    source: string;
    hasSessionObject: boolean;
    hasAccessToken: boolean;
    hasRefreshToken: boolean;
    hasConnectedIdentity: boolean;
    connectedIdentityEmail: string;
    reasonNotConnected: string | null;
    mailboxEligible: boolean;
    mailboxEligibilityReason: string | null;
    storedExpiresAt: number | string | null;
    nowTimestamp: number;
    computedIsExpired: boolean;
    expiryReason: string | null;
    refreshAttempted: boolean;
    refreshSucceeded: boolean;
    reconnectRequired: boolean;
  };
};

type OutlookEmailDraft = {
  to?: string[];
  cc?: string[];
  bcc?: string[];
  subject?: string;
  body?: string;
};

type MeetingEventInput = {
  subject: string;
  bodyText?: string;
  startISO: string;
  endISO: string;
  timeZone: string;
  isAllDay?: boolean;
  location?: string;
  attendees?: string[];
  teamsMeeting?: boolean;
  expectedEmail?: string;
};

export type OutlookDraftResult = {
  id: string;
  webLink: string;
};

export type OutlookCalendarEventResult = {
  id: string;
  webLink: string;
  joinUrl: string;
  hasOnlineMeeting: boolean;
};

const OUTLOOK_SESSION_STORAGE_KEY = "event_based_reminders_app_outlook_session_v1";
const LEGACY_OUTLOOK_SESSION_STORAGE_KEYS = ["standalone_plans_outlook_session_v1"];
const OUTLOOK_IDENTITY_STORAGE_KEY = "event_based_reminders_app_outlook_identity_v1";
const LEGACY_OUTLOOK_IDENTITY_STORAGE_KEYS = ["standalone_plans_outlook_identity_v1"];
const OUTLOOK_OAUTH_STATE_KEY = "event_based_reminders_app_outlook_oauth_state_v1";
const LEGACY_OUTLOOK_OAUTH_STATE_KEYS = ["standalone_plans_outlook_oauth_state_v1"];
const OUTLOOK_OAUTH_VERIFIER_KEY = "event_based_reminders_app_outlook_oauth_verifier_v1";
const LEGACY_OUTLOOK_OAUTH_VERIFIER_KEYS = ["standalone_plans_outlook_oauth_verifier_v1"];
const OUTLOOK_OAUTH_VERIFIER_COOKIE = "event_based_reminders_app_outlook_oauth_verifier";
const OUTLOOK_LOCAL_REDIRECT_URI = "http://localhost:8664/api/auth/microsoft/callback";
const OUTLOOK_HOSTED_REDIRECT_URI = "https://event-based-reminders-app.vercel.app/api/auth/microsoft/callback";
export const OUTLOOK_CONNECTION_UPDATED_EVENT = "event-based-reminders-app:outlook-connection-updated";
export const OUTLOOK_OAUTH_MESSAGE_TYPE = "event_based_reminders_app_outlook_oauth_result";

export const OUTLOOK_SCOPES = [
  "User.Read",
  "offline_access",
  "Mail.ReadWrite",
  "Mail.Send",
  "Calendars.ReadWrite",
] as const;

function emitOutlookConnectionUpdated() {
  if (typeof window === "undefined") return;
  window.dispatchEvent(new CustomEvent(OUTLOOK_CONNECTION_UPDATED_EVENT));
}

function migrateOutlookStorage() {
  migrateLegacyPersistedValue("localStorage", OUTLOOK_SESSION_STORAGE_KEY, LEGACY_OUTLOOK_SESSION_STORAGE_KEYS);
  migrateLegacyPersistedValue("localStorage", OUTLOOK_IDENTITY_STORAGE_KEY, LEGACY_OUTLOOK_IDENTITY_STORAGE_KEYS);
}

function migrateOutlookOAuthStorage() {
  migrateLegacyPersistedValue("sessionStorage", OUTLOOK_OAUTH_STATE_KEY, LEGACY_OUTLOOK_OAUTH_STATE_KEYS);
  migrateLegacyPersistedValue("sessionStorage", OUTLOOK_OAUTH_VERIFIER_KEY, LEGACY_OUTLOOK_OAUTH_VERIFIER_KEYS);
}

function getOutlookClientId() {
  return String(process.env.NEXT_PUBLIC_MICROSOFT_CLIENT_ID ?? "").trim();
}

function getOutlookTenantId() {
  return "common";
}

function getOutlookRedirectUri() {
  if (typeof window === "undefined") return "";
  return window.location.hostname === "localhost" ? OUTLOOK_LOCAL_REDIRECT_URI : OUTLOOK_HOSTED_REDIRECT_URI;
}

function writeOutlookOAuthVerifierCookie(value: string) {
  if (typeof document === "undefined") return;
  document.cookie = `${OUTLOOK_OAUTH_VERIFIER_COOKIE}=${encodeURIComponent(value)}; Path=/; Max-Age=600; SameSite=Lax; Secure`;
}

function clearOutlookOAuthVerifierCookie() {
  if (typeof document === "undefined") return;
  document.cookie = `${OUTLOOK_OAUTH_VERIFIER_COOKIE}=; Path=/; Max-Age=0; SameSite=Lax; Secure`;
}

function normalizeOutlookEmail(value: string | null | undefined) {
  return String(value ?? "")
    .trim()
    .replace(/^mailto:/i, "")
    .trim()
    .toLowerCase();
}

function randomString(length = 64) {
  const bytes = new Uint8Array(length);
  crypto.getRandomValues(bytes);
  return Array.from(bytes, (byte) => (byte % 36).toString(36)).join("");
}

function base64UrlEncode(bytes: Uint8Array) {
  let binary = "";
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
}

async function sha256(value: string) {
  const encoded = new TextEncoder().encode(value);
  const digest = await crypto.subtle.digest("SHA-256", encoded);
  return base64UrlEncode(new Uint8Array(digest));
}

function hasRequiredScopes(session: OutlookSession, requiredScopes: string[]) {
  const grantedScopes = new Set(String(session.scope ?? "").split(/\s+/).filter(Boolean));
  return requiredScopes.every((scope) => grantedScopes.has(scope));
}

function getOutlookSessionExpiryDebug(rawSession: Partial<OutlookSession> | null) {
  const nowTimestamp = Date.now();
  const storedExpiresAt: unknown =
    rawSession && "expiresAt" in rawSession ? (rawSession as Record<string, unknown>).expiresAt : null;

  if (storedExpiresAt == null || storedExpiresAt === "") {
    return {
      storedExpiresAt,
      nowTimestamp,
      normalizedExpiresAt: null,
      computedIsExpired: false,
      expiryReason: "missing_expires_at_treated_as_valid",
    };
  }

  const numericExpiresAt = Number(storedExpiresAt);
  if (!Number.isFinite(numericExpiresAt)) {
    return {
      storedExpiresAt,
      nowTimestamp,
      normalizedExpiresAt: null,
      computedIsExpired: false,
      expiryReason: "invalid_expires_at_treated_as_valid",
    };
  }

  const normalizedExpiresAt =
    numericExpiresAt > 0 && numericExpiresAt < 1_000_000_000_000 ? numericExpiresAt * 1000 : numericExpiresAt;
  const computedIsExpired = nowTimestamp >= normalizedExpiresAt - 60_000;

  return {
    storedExpiresAt,
    nowTimestamp,
    normalizedExpiresAt,
    computedIsExpired,
    expiryReason: computedIsExpired ? "expires_at_in_past" : "expires_at_valid",
  };
}

function loadRawStoredOutlookSession() {
  if (typeof window === "undefined") return null;
  try {
    migrateOutlookStorage();
    const raw = readPersistedValue("localStorage", OUTLOOK_SESSION_STORAGE_KEY, LEGACY_OUTLOOK_SESSION_STORAGE_KEYS);
    if (!raw) return null;
    return JSON.parse(raw) as Partial<OutlookSession>;
  } catch {
    return null;
  }
}

function loadStoredOutlookSession(requiredScopes: string[] = []): OutlookSession | null {
  if (typeof window === "undefined") return null;
  try {
    migrateOutlookStorage();
    const raw = readPersistedValue("localStorage", OUTLOOK_SESSION_STORAGE_KEY, LEGACY_OUTLOOK_SESSION_STORAGE_KEYS);
    if (!raw) return null;
    const parsed = JSON.parse(raw) as OutlookSession;
    if (!parsed.accessToken) return null;
    const expiryDebug = getOutlookSessionExpiryDebug(parsed);
    if (expiryDebug.computedIsExpired) return null;
    if (requiredScopes.length > 0 && !hasRequiredScopes(parsed, requiredScopes)) return null;
    return parsed;
  } catch {
    return null;
  }
}

function saveStoredOutlookSession(session: OutlookSession) {
  if (typeof window === "undefined") return;
  migrateOutlookStorage();
  const didWrite = writePersistedValue("localStorage", OUTLOOK_SESSION_STORAGE_KEY, JSON.stringify(session));
  if (!didWrite) return;
  emitOutlookConnectionUpdated();
}

function saveOutlookSessionFromTokenPayload(payload: {
  access_token?: string;
  refresh_token?: string;
  expires_in?: number;
  scope?: string;
}) {
  if (!payload.access_token) {
    throw new Error("Failed to connect Outlook.");
  }

  const session: OutlookSession = {
    accessToken: payload.access_token,
    refreshToken: payload.refresh_token,
    expiresAt: Date.now() + (payload.expires_in ?? 3600) * 1000,
    scope: payload.scope,
    obtainedAt: new Date().toISOString(),
  };

  saveStoredOutlookSession(session);
  return session;
}

function loadStoredOutlookIdentity(): OutlookConnectedIdentity | null {
  if (typeof window === "undefined") return null;
  try {
    migrateOutlookStorage();
    const raw = readPersistedValue("localStorage", OUTLOOK_IDENTITY_STORAGE_KEY, LEGACY_OUTLOOK_IDENTITY_STORAGE_KEYS);
    if (!raw) return null;
    const parsed = JSON.parse(raw) as OutlookConnectedIdentity;
    if (!parsed.id) return null;
    const normalizedConnectedEmail = parsed.normalizedEmail || normalizeOutlookEmail(parsed.mail || parsed.userPrincipalName);
    if (!normalizedConnectedEmail) return null;
    return {
      ...parsed,
      normalizedEmail: normalizedConnectedEmail,
    };
  } catch {
    return null;
  }
}

function saveStoredOutlookIdentity(identity: OutlookConnectedIdentity) {
  if (typeof window === "undefined") return;
  migrateOutlookStorage();
  const didWrite = writePersistedValue("localStorage", OUTLOOK_IDENTITY_STORAGE_KEY, JSON.stringify(identity));
  if (!didWrite) return;
  emitOutlookConnectionUpdated();
}

function clearStoredOutlookState() {
  if (typeof window === "undefined") return;
  const removedSession = removePersistedValue("localStorage", OUTLOOK_SESSION_STORAGE_KEY, LEGACY_OUTLOOK_SESSION_STORAGE_KEYS);
  const removedIdentity = removePersistedValue("localStorage", OUTLOOK_IDENTITY_STORAGE_KEY, LEGACY_OUTLOOK_IDENTITY_STORAGE_KEYS);
  const removedOAuthState = removePersistedValue("sessionStorage", OUTLOOK_OAUTH_STATE_KEY, LEGACY_OUTLOOK_OAUTH_STATE_KEYS);
  const removedOAuthVerifier = removePersistedValue(
    "sessionStorage",
    OUTLOOK_OAUTH_VERIFIER_KEY,
    LEGACY_OUTLOOK_OAUTH_VERIFIER_KEYS
  );
  if (!removedSession && !removedIdentity && !removedOAuthState && !removedOAuthVerifier) return;
  emitOutlookConnectionUpdated();
}

function parseRecipients(raw: string[] | undefined) {
  return (raw ?? [])
    .map((value) => value.trim())
    .filter(Boolean)
    .map((address) => ({
      emailAddress: { address },
    }));
}

function parseEventAttendees(raw: string[] | undefined) {
  return (raw ?? [])
    .map((value) => value.trim().replace(/^mailto:/i, "").trim().toLowerCase())
    .filter(Boolean)
    .map((address) => ({
      emailAddress: { address, name: address },
      type: "required" as const,
    }));
}

function getMeaningfulBody(value: string | null | undefined) {
  return String(value ?? "").trim();
}

function buildGraphMessagePayload(input: { draft: OutlookEmailDraft; fallbackSubject: string }) {
  const trimmedBody = getMeaningfulBody(input.draft.body);
  return {
    subject: input.draft.subject?.trim() || input.fallbackSubject,
    ...(trimmedBody
      ? {
          body: {
            contentType: "text",
            content: trimmedBody,
          },
        }
      : {}),
    toRecipients: parseRecipients(input.draft.to),
    ccRecipients: parseRecipients(input.draft.cc),
    bccRecipients: parseRecipients(input.draft.bcc),
  };
}

function getOutlookMailboxEligibility(payload: { mail?: string; userPrincipalName?: string }) {
  const normalizedMail = normalizeOutlookEmail(payload.mail);
  const rawUserPrincipalName = String(payload.userPrincipalName ?? "").trim();
  const looksLikeGeneratedOutlookAlias =
    !normalizedMail && /^outlook_[^@]+@outlook\.com$/i.test(rawUserPrincipalName);

  if (looksLikeGeneratedOutlookAlias) {
    return {
      mailboxEligible: false,
      mailboxEligibilityReason:
        "Please connect an Outlook or Microsoft 365 account to use your actual mailbox address.",
    };
  }

  return {
    mailboxEligible: true,
    mailboxEligibilityReason: null,
  };
}

async function fetchOutlookIdentity(accessToken: string) {
  const response = await fetch("https://graph.microsoft.com/v1.0/me?$select=id,mail,userPrincipalName,displayName", {
    headers: {
      Authorization: `Bearer ${accessToken}`,
    },
  });

  const payload = (await response.json()) as {
    id?: string;
    mail?: string;
    userPrincipalName?: string;
    displayName?: string;
    error?: { message?: string };
  };

  if (!response.ok || !payload.id) {
    throw new Error(payload.error?.message || "Failed to load the connected Outlook account.");
  }

  const normalizedEmail = normalizeOutlookEmail(payload.mail || payload.userPrincipalName);
  if (!normalizedEmail) {
    throw new Error("The connected Outlook account did not return a usable email.");
  }

  return {
    id: payload.id,
    mail: payload.mail ?? "",
    userPrincipalName: payload.userPrincipalName ?? "",
    displayName: payload.displayName ?? "",
    connectedAt: new Date().toISOString(),
    normalizedEmail,
    ...getOutlookMailboxEligibility(payload),
  } satisfies OutlookConnectedIdentity;
}

export function getConnectedOutlookMailboxEmail(identity: OutlookConnectedIdentity | null | undefined) {
  return identity?.mail?.trim() || identity?.userPrincipalName?.trim() || "";
}

export function getOutlookConnectionStatus(expectedEmail?: string): OutlookConnectionStatus {
  return getOutlookConnectionState(expectedEmail).status;
}

export function getOutlookConnectionState(expectedEmail?: string): OutlookConnectionState {
  const rawSession = loadRawStoredOutlookSession();
  const session = loadStoredOutlookSession();
  const identity = loadStoredOutlookIdentity();
  const expiryDebug = getOutlookSessionExpiryDebug(rawSession);
  const normalizedExpectedEmail = normalizeOutlookEmail(expectedEmail);
  const normalizedConnectedEmail = identity?.normalizedEmail ?? "";
  const hasSessionObject = Boolean(rawSession);
  const hasAccessToken = Boolean(rawSession?.accessToken);
  const hasRefreshToken = Boolean(rawSession?.refreshToken);
  const identityExists = Boolean(identity);
  const mailboxEligible = identity?.mailboxEligible ?? true;
  const mailboxEligibilityReason = identity?.mailboxEligibilityReason ?? null;
  const safeStoredExpiresAt =
    typeof expiryDebug.storedExpiresAt === "number" || typeof expiryDebug.storedExpiresAt === "string"
      ? expiryDebug.storedExpiresAt
      : expiryDebug.storedExpiresAt == null
        ? null
        : String(expiryDebug.storedExpiresAt);

  let reasonNotConnected: string | null = null;
  if (!hasSessionObject) {
    reasonNotConnected = "missing_session_object";
  } else if (!hasAccessToken) {
    reasonNotConnected = "missing_access_token";
  } else if (!session) {
    reasonNotConnected = "session_invalid_or_expired";
  } else if (!identityExists) {
    reasonNotConnected = "missing_connected_identity";
  } else if (!mailboxEligible) {
    reasonNotConnected = "unsupported_mailbox_identity";
  }

  const connected = Boolean(session && identity && mailboxEligible);
  const reconnectRequired = Boolean(identityExists && (!connected || !mailboxEligible));
  const stale = false;

  return {
    status: connected ? "connected" : reconnectRequired ? "reconnect_required" : "not_connected",
    connected,
    stale,
    supportedMailbox: mailboxEligible,
    expectedEmail: normalizedExpectedEmail,
    identity,
    normalizedConnectedEmail,
    debug: {
      source: "localStorage",
      hasSessionObject,
      hasAccessToken,
      hasRefreshToken,
      hasConnectedIdentity: identityExists,
      connectedIdentityEmail: normalizedConnectedEmail,
      reasonNotConnected,
      mailboxEligible,
      mailboxEligibilityReason,
      storedExpiresAt: safeStoredExpiresAt,
      nowTimestamp: expiryDebug.nowTimestamp,
      computedIsExpired: expiryDebug.computedIsExpired,
      expiryReason: expiryDebug.expiryReason,
      refreshAttempted: false,
      refreshSucceeded: false,
      reconnectRequired,
    },
  };
}

async function refreshStoredOutlookSession(requiredScopes: string[] = []) {
  const rawSession = loadRawStoredOutlookSession();
  const refreshToken = String(rawSession?.refreshToken ?? "").trim();
  if (!refreshToken) return null;

  const response = await fetch("/api/auth/microsoft/refresh", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      refreshToken,
      redirectUri: getOutlookRedirectUri(),
    }),
  });

  const payload = (await response.json()) as {
    access_token?: string;
    refresh_token?: string;
    expires_in?: number;
    scope?: string;
    error?: string;
    error_description?: string;
  };

  if (!response.ok || !payload.access_token) {
    return null;
  }

  const session = saveOutlookSessionFromTokenPayload({
    ...payload,
    refresh_token: payload.refresh_token || refreshToken,
  });
  if (requiredScopes.length > 0 && !hasRequiredScopes(session, requiredScopes)) {
    return null;
  }
  return session;
}

export async function resolveOutlookConnectionState(expectedEmail?: string, requiredScopes: string[] = []) {
  let state = getOutlookConnectionState(expectedEmail);
  let refreshAttempted = false;
  let refreshSucceeded = false;
  let reconnectRequired = state.debug.reconnectRequired;

  const validSession = loadStoredOutlookSession(requiredScopes);
  if (!validSession && state.identity) {
    refreshAttempted = true;
    const refreshedSession = await refreshStoredOutlookSession(requiredScopes);
    refreshSucceeded = Boolean(refreshedSession);
    state = getOutlookConnectionState(expectedEmail);
    reconnectRequired = !refreshSucceeded && Boolean(state.identity);
  }

  return {
    ...state,
    status: state.connected ? "connected" : reconnectRequired ? "reconnect_required" : "not_connected",
    debug: {
      ...state.debug,
      refreshAttempted,
      refreshSucceeded,
      reconnectRequired,
    },
  } satisfies OutlookConnectionState;
}

async function requireStoredOutlookAccessToken(input: {
  requiredScopes: string[];
  expectedEmail?: string;
}) {
  const state = await resolveOutlookConnectionState(input.expectedEmail, input.requiredScopes);

  if (!state.supportedMailbox) {
    throw new Error(
      state.identity?.mailboxEligibilityReason ||
        "Please connect an Outlook or Microsoft 365 account to use your actual mailbox address."
    );
  }
  if (!state.connected) {
    throw new Error("Connect Outlook in Settings before continuing.");
  }

  const session = loadStoredOutlookSession(input.requiredScopes);
  if (!session) {
    throw new Error("Connect Outlook in Settings before continuing.");
  }
  return session.accessToken;
}

async function connectOutlookInteractively() {
  if (typeof window === "undefined") {
    throw new Error("Outlook connection is only available in the browser.");
  }

  const clientId = getOutlookClientId();
  if (!clientId) {
    throw new Error("Outlook is not configured yet. Add NEXT_PUBLIC_MICROSOFT_CLIENT_ID first.");
  }

  const redirectUri = getOutlookRedirectUri();
  const state = randomString(24);
  const codeVerifier = randomString(96);
  const codeChallenge = await sha256(codeVerifier);

  writePersistedValue("sessionStorage", OUTLOOK_OAUTH_STATE_KEY, state);
  writePersistedValue("sessionStorage", OUTLOOK_OAUTH_VERIFIER_KEY, codeVerifier);
  writeOutlookOAuthVerifierCookie(codeVerifier);

  const url = new URL(`https://login.microsoftonline.com/${getOutlookTenantId()}/oauth2/v2.0/authorize`);
  url.searchParams.set("client_id", clientId);
  url.searchParams.set("response_type", "code");
  url.searchParams.set("redirect_uri", redirectUri);
  url.searchParams.set("response_mode", "query");
  url.searchParams.set("scope", OUTLOOK_SCOPES.join(" "));
  url.searchParams.set("state", state);
  url.searchParams.set("code_challenge", codeChallenge);
  url.searchParams.set("code_challenge_method", "S256");
  url.searchParams.set("prompt", "select_account");

  const popup = window.open(url.toString(), "event-based-reminders-app-outlook-auth", "width=640,height=760");
  if (!popup) {
    throw new Error("Unable to open Outlook sign-in. Please allow pop-ups and try again.");
  }

  return await new Promise<OutlookSession>((resolve, reject) => {
    let settled = false;

    const cleanup = () => {
      settled = true;
      window.removeEventListener("message", handleMessage);
      window.clearInterval(pollTimer);
    };

    const fail = (message: string) => {
      cleanup();
      reject(new Error(message));
    };

    const handleMessage = async (event: MessageEvent) => {
      if (event.origin !== window.location.origin) return;
      const data = event.data as
        | {
            type?: string;
            state?: string;
            accessToken?: string;
            refreshToken?: string;
            expiresIn?: number;
            scope?: string;
            error?: string;
            errorDescription?: string;
          }
        | undefined;

      if (data?.type !== OUTLOOK_OAUTH_MESSAGE_TYPE) return;

      migrateOutlookOAuthStorage();
      const expectedState = readPersistedValue("sessionStorage", OUTLOOK_OAUTH_STATE_KEY, LEGACY_OUTLOOK_OAUTH_STATE_KEYS);
      removePersistedValue("sessionStorage", OUTLOOK_OAUTH_STATE_KEY, LEGACY_OUTLOOK_OAUTH_STATE_KEYS);
      removePersistedValue("sessionStorage", OUTLOOK_OAUTH_VERIFIER_KEY, LEGACY_OUTLOOK_OAUTH_VERIFIER_KEYS);
      clearOutlookOAuthVerifierCookie();

      if (data.error) {
        fail(data.errorDescription || "Outlook permission was not granted.");
        return;
      }

      if (!data.state || !expectedState || data.state !== expectedState) {
        fail("Outlook sign-in could not be verified. Please try again.");
        return;
      }

      cleanup();
      try {
        resolve(
          saveOutlookSessionFromTokenPayload({
            access_token: data.accessToken,
            refresh_token: data.refreshToken,
            expires_in: data.expiresIn,
            scope: data.scope,
          })
        );
      } catch (error) {
        reject(error instanceof Error ? error : new Error("Failed to connect Outlook."));
      }
    };

    const pollTimer = window.setInterval(() => {
      if (settled) return;
      if (popup.closed) fail("Outlook sign-in was closed before it finished.");
    }, 500);

    window.addEventListener("message", handleMessage);
  });
}

export async function connectOutlook(expectedEmail?: string) {
  const session = await connectOutlookInteractively();
  const identity = await fetchOutlookIdentity(session.accessToken);
  saveStoredOutlookIdentity(identity);
  return getOutlookConnectionState(expectedEmail || identity.normalizedEmail);
}

export function disconnectOutlook() {
  clearStoredOutlookState();
}

export async function createOutlookDraftFromEmailDraft(input: {
  draft: OutlookEmailDraft;
  fallbackSubject: string;
  expectedEmail?: string;
}): Promise<OutlookDraftResult> {
  const accessToken = await requireStoredOutlookAccessToken({
    requiredScopes: ["Mail.ReadWrite", "Mail.Send"],
    expectedEmail: input.expectedEmail,
  });

  const payload = buildGraphMessagePayload(input);

  const response = await fetch("https://graph.microsoft.com/v1.0/me/messages", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(payload),
  });

  const responsePayload = (await response.json()) as {
    id?: string;
    webLink?: string;
    error?: { message?: string };
  };

  if (!response.ok || !responsePayload.id) {
    throw new Error(responsePayload.error?.message || "Outlook draft creation failed.");
  }

  return {
    id: responsePayload.id,
    webLink: responsePayload.webLink ?? "",
  };
}

export async function sendOutlookEmailFromEmailDraft(input: {
  draft: OutlookEmailDraft;
  fallbackSubject: string;
  expectedEmail?: string;
}) {
  const accessToken = await requireStoredOutlookAccessToken({
    requiredScopes: ["Mail.ReadWrite", "Mail.Send"],
    expectedEmail: input.expectedEmail,
  });

  const payload = buildGraphMessagePayload(input);

  const response = await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      message: payload,
      saveToSentItems: true,
    }),
  });

  if (!response.ok) {
    const errorPayload = (await response.json().catch(() => null)) as { error?: { message?: string } } | null;
    throw new Error(errorPayload?.error?.message || "Outlook email send failed.");
  }

  return {
    id: "",
    webLink: "",
  } satisfies OutlookDraftResult;
}

export async function scheduleOutlookEmailFromEmailDraft(input: {
  draft: OutlookEmailDraft;
  fallbackSubject: string;
  scheduledSendISO: string;
  expectedEmail?: string;
}): Promise<OutlookDraftResult> {
  const accessToken = await requireStoredOutlookAccessToken({
    requiredScopes: ["Mail.ReadWrite", "Mail.Send"],
    expectedEmail: input.expectedEmail,
  });

  const payload = buildGraphMessagePayload(input);
  const scheduledSendUtc = new Date(input.scheduledSendISO).toISOString();

  const createResponse = await fetch("https://graph.microsoft.com/v1.0/me/messages", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      ...payload,
      singleValueExtendedProperties: [
        {
          id: "SystemTime 0x3FEF",
          value: scheduledSendUtc,
        },
      ],
    }),
  });

  const createPayload = (await createResponse.json()) as {
    id?: string;
    webLink?: string;
    error?: { message?: string };
  };

  if (!createResponse.ok || !createPayload.id) {
    throw new Error(createPayload.error?.message || "Outlook scheduled email creation failed.");
  }

  const sendResponse = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${createPayload.id}/send`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
  });

  if (!sendResponse.ok) {
    const sendPayload = (await sendResponse.json().catch(() => null)) as { error?: { message?: string } } | null;
    throw new Error(sendPayload?.error?.message || "Outlook email scheduling failed.");
  }

  return {
    id: createPayload.id,
    webLink: createPayload.webLink ?? "",
  };
}

export async function createOutlookCalendarEvent(
  input: MeetingEventInput
): Promise<OutlookCalendarEventResult> {
  const accessToken = await requireStoredOutlookAccessToken({
    requiredScopes: ["Calendars.ReadWrite"],
    expectedEmail: input.expectedEmail,
  });

  const attendees = parseEventAttendees(input.attendees);
  const trimmedBody = getMeaningfulBody(input.bodyText);
  const payload = {
    subject: input.subject,
    ...(trimmedBody
      ? {
          body: {
            contentType: "text",
            content: trimmedBody,
          },
        }
      : {}),
    start: {
      dateTime: input.startISO,
      timeZone: input.timeZone,
    },
    end: {
      dateTime: input.endISO,
      timeZone: input.timeZone,
    },
    ...(input.isAllDay ? { isAllDay: true } : {}),
    location: {
      displayName: input.location?.trim() || "",
    },
    isOnlineMeeting: Boolean(input.teamsMeeting),
    ...(input.teamsMeeting
      ? {
          onlineMeetingProvider: "teamsForBusiness" as const,
        }
      : {}),
    ...(attendees.length > 0
      ? {
          attendees,
          responseRequested: true,
          allowNewTimeProposals: true,
        }
      : {}),
  };

  const response = await fetch("https://graph.microsoft.com/v1.0/me/events", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(payload),
  });

  const responsePayload = (await response.json()) as {
    id?: string;
    webLink?: string;
    onlineMeeting?: { joinUrl?: string };
    error?: { message?: string };
  };

  if (!response.ok || !responsePayload.id) {
    throw new Error(responsePayload.error?.message || "Outlook calendar event creation failed.");
  }

  const eventStateResponse = await fetch(`https://graph.microsoft.com/v1.0/me/events/${responsePayload.id}`, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
    },
  });

  const eventStatePayload = (await eventStateResponse.json().catch(() => null)) as
    | {
        webLink?: string;
        isOnlineMeeting?: boolean;
        onlineMeeting?: { joinUrl?: string };
      }
    | null;

  const resolvedJoinUrl =
    eventStatePayload?.onlineMeeting?.joinUrl ?? responsePayload.onlineMeeting?.joinUrl ?? "";
  const resolvedWebLink = eventStatePayload?.webLink ?? responsePayload.webLink ?? "";
  const resolvedHasOnlineMeeting = Boolean(
    resolvedJoinUrl || eventStatePayload?.isOnlineMeeting || responsePayload.onlineMeeting?.joinUrl || input.teamsMeeting
  );

  return {
    id: responsePayload.id,
    webLink: resolvedWebLink,
    joinUrl: resolvedJoinUrl,
    hasOnlineMeeting: resolvedHasOnlineMeeting,
  };
}
