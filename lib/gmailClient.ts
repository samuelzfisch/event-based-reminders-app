"use client";

import { readPersistedValue, removePersistedValue, writePersistedValue } from "./browserStorage";
import { getCachedOrgContext } from "./orgBootstrap";
import { getSupabaseBrowserClient, isSupabaseConfigured } from "./supabaseClient";

export type GmailConnectionStatus = "connected" | "reconnect_required" | "not_connected";

type GmailSession = {
  accessToken: string;
  refreshToken?: string;
  expiresAt: number;
  scope?: string;
  obtainedAt: string;
};

export type GmailConnectedIdentity = {
  id: string;
  email: string;
  displayName?: string;
  connectedAt: string;
  normalizedEmail: string;
};

export type GmailConnectionState = {
  status: GmailConnectionStatus;
  connected: boolean;
  stale: boolean;
  expectedEmail: string;
  identity: GmailConnectedIdentity | null;
  normalizedConnectedEmail: string;
};

type GmailEmailDraft = {
  to?: string[];
  cc?: string[];
  bcc?: string[];
  subject?: string;
  body?: string;
};

export type GmailDraftResult = {
  id: string;
  webLink: string;
};

export type GmailSendResult = {
  id: string;
};

export type GoogleCalendarEventResult = {
  id?: string;
  success: true;
  webLink: string;
  joinUrl: string;
  hasOnlineMeeting: boolean;
};

export type GmailAvailabilityDebugSnapshot = {
  providerRowFound: boolean;
  providerValue: string | null;
  connectionStatus: string | null;
  scopesPresent: boolean;
  tokenPresent: boolean;
  identityPresent: boolean;
  finalAvailable: boolean;
  rejectionReason: string | null;
};

const GMAIL_SESSION_STORAGE_KEY = "event_based_reminders_app_gmail_session_v1";
const GMAIL_IDENTITY_STORAGE_KEY = "event_based_reminders_app_gmail_identity_v1";
const GMAIL_OAUTH_STATE_KEY = "event_based_reminders_app_gmail_oauth_state_v1";
const GMAIL_OAUTH_VERIFIER_KEY = "event_based_reminders_app_gmail_oauth_verifier_v1";
const GMAIL_OAUTH_VERIFIER_COOKIE = "event_based_reminders_app_gmail_oauth_verifier";
const GMAIL_LOCAL_REDIRECT_URI = "http://localhost:8664/api/auth/google/callback";
const GMAIL_HOSTED_REDIRECT_URI = "https://event-based-reminders-app.vercel.app/api/auth/google/callback";
const GMAIL_PROVIDER_NAME = "google_gmail";

export const GMAIL_CONNECTION_UPDATED_EVENT = "event-based-reminders-app:gmail-connection-updated";
export const GMAIL_OAUTH_MESSAGE_TYPE = "event_based_reminders_app_gmail_oauth_result";
export const GMAIL_COMPOSE_SCOPE = "https://www.googleapis.com/auth/gmail.compose";
export const GOOGLE_CALENDAR_EVENTS_SCOPE = "https://www.googleapis.com/auth/calendar.events";

const GMAIL_SCOPES = [
  "openid",
  "email",
  "profile",
  GMAIL_COMPOSE_SCOPE,
  GOOGLE_CALENDAR_EVENTS_SCOPE,
] as const;

let lastGmailAvailabilityDebugSnapshot: GmailAvailabilityDebugSnapshot = {
  providerRowFound: false,
  providerValue: null,
  connectionStatus: null,
  scopesPresent: false,
  tokenPresent: false,
  identityPresent: false,
  finalAvailable: false,
  rejectionReason: "not_checked_yet",
};

const gmailResolutionCache = new Map<string, Promise<GmailConnectionState>>();

export function getLastGmailAvailabilityDebugSnapshot() {
  return lastGmailAvailabilityDebugSnapshot;
}

function emitGmailConnectionUpdated() {
  if (typeof window === "undefined") return;
  gmailResolutionCache.clear();
  window.dispatchEvent(new CustomEvent(GMAIL_CONNECTION_UPDATED_EVENT));
}

function getGmailResolutionCacheKey(expectedEmail?: string, requiredScopes: string[] = []) {
  const cachedOrgContext = getCachedOrgContext();
  return JSON.stringify({
    orgId: cachedOrgContext?.orgId ?? "",
    userId: cachedOrgContext?.userId ?? "",
    expectedEmail: normalizeGmailEmail(expectedEmail),
    requiredScopes: [...requiredScopes].sort(),
  });
}

function getGmailClientId() {
  return String(process.env.NEXT_PUBLIC_GOOGLE_CLIENT_ID ?? "").trim();
}

function getGmailRedirectUri() {
  if (typeof window === "undefined") return "";
  return window.location.hostname === "localhost" ? GMAIL_LOCAL_REDIRECT_URI : GMAIL_HOSTED_REDIRECT_URI;
}

function writeGmailOAuthVerifierCookie(value: string) {
  if (typeof document === "undefined") return;
  document.cookie = `${GMAIL_OAUTH_VERIFIER_COOKIE}=${encodeURIComponent(value)}; Path=/; Max-Age=600; SameSite=Lax; Secure`;
}

function clearGmailOAuthVerifierCookie() {
  if (typeof document === "undefined") return;
  document.cookie = `${GMAIL_OAUTH_VERIFIER_COOKIE}=; Path=/; Max-Age=0; SameSite=Lax; Secure`;
}

function normalizeGmailEmail(value: string | null | undefined) {
  return String(value ?? "")
    .trim()
    .replace(/^mailto:/i, "")
    .trim()
    .toLowerCase();
}

function buildLocalApiUrl(path: string) {
  if (typeof window === "undefined") return path;
  return new URL(path, window.location.origin).toString();
}

function normalizeGoogleCalendarAttendees(raw: string[] | undefined) {
  return (raw ?? [])
    .map((value) => value.trim())
    .filter(Boolean)
    .filter((value) => /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(value))
    .map((email) => ({ email }));
}

function isGoogleCalendarScopeError(responseStatus: number, responsePayload: { error?: { message?: string } } | null) {
  if (responseStatus !== 403) return false;
  const message = String(responsePayload?.error?.message ?? "").toLowerCase();
  return (
    message.includes("insufficient") ||
    message.includes("permission") ||
    message.includes("scope") ||
    message.includes("calendar.events")
  );
}

function isGoogleCalendarApiDisabledError(
  responseStatus: number,
  responsePayload: {
    error?: {
      message?: string;
      status?: string;
      errors?: Array<{ reason?: string; message?: string }>;
    };
  } | null
) {
  if (responseStatus !== 403) return false;
  const message = String(responsePayload?.error?.message ?? "").toLowerCase();
  const status = String(responsePayload?.error?.status ?? "").toLowerCase();
  const reasons = Array.isArray(responsePayload?.error?.errors)
    ? responsePayload?.error?.errors.map((entry) => String(entry?.reason ?? "").toLowerCase())
    : [];
  return (
    status === "permission_denied" &&
    (message.includes("has not been used in project") ||
      message.includes("it is disabled") ||
      reasons.includes("accessnotconfigured") ||
      reasons.includes("service_disabled"))
  );
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

function hasRequiredScopes(session: GmailSession, requiredScopes: string[]) {
  const grantedScopes = new Set(String(session.scope ?? "").split(/\s+/).filter(Boolean));
  return requiredScopes.every((scope) => grantedScopes.has(scope));
}

function hasAllScopes(scopeValue: string | null | undefined, requiredScopes: string[]) {
  const grantedScopes = new Set(String(scopeValue ?? "").split(/\s+/).filter(Boolean));
  return requiredScopes.every((scope) => grantedScopes.has(scope));
}

function loadRawStoredGmailSession(): GmailSession | null {
  if (typeof window === "undefined") return null;
  try {
    const raw = readPersistedValue("localStorage", GMAIL_SESSION_STORAGE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw) as GmailSession;
    if ((!parsed?.accessToken && !parsed?.refreshToken) || !parsed?.expiresAt) return null;
    return parsed;
  } catch {
    return null;
  }
}

function loadStoredGmailSession(requiredScopes: string[] = []): GmailSession | null {
  const parsed = loadRawStoredGmailSession();
  if (!parsed) return null;
  if (!parsed.accessToken) return null;
  if (Date.now() >= parsed.expiresAt - 60_000) return null;
  if (requiredScopes.length > 0 && !hasRequiredScopes(parsed, requiredScopes)) return null;
  return parsed;
}

function saveStoredGmailSession(session: GmailSession) {
  writePersistedValue("localStorage", GMAIL_SESSION_STORAGE_KEY, JSON.stringify(session));
  emitGmailConnectionUpdated();
}

function saveGmailSessionFromTokenPayload(payload: {
  access_token?: string | null;
  refresh_token?: string | null;
  expires_in?: number | null;
  scope?: string | null;
}) {
  if (!payload.access_token) {
    throw new Error("Failed to connect Gmail.");
  }

  const session: GmailSession = {
    accessToken: payload.access_token,
    refreshToken: payload.refresh_token || undefined,
    expiresAt: Date.now() + Math.max(Number(payload.expires_in ?? 3600), 60) * 1000,
    scope: payload.scope || undefined,
    obtainedAt: new Date().toISOString(),
  };
  saveStoredGmailSession(session);
  return session;
}

function loadStoredGmailIdentity(): GmailConnectedIdentity | null {
  if (typeof window === "undefined") return null;
  try {
    const raw = readPersistedValue("localStorage", GMAIL_IDENTITY_STORAGE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw) as GmailConnectedIdentity;
    const normalizedEmail = normalizeGmailEmail(parsed.normalizedEmail || parsed.email);
    if (!parsed?.id || !normalizedEmail) return null;
    return {
      ...parsed,
      email: normalizedEmail,
      normalizedEmail,
    };
  } catch {
    return null;
  }
}

function saveStoredGmailIdentity(identity: GmailConnectedIdentity) {
  writePersistedValue("localStorage", GMAIL_IDENTITY_STORAGE_KEY, JSON.stringify(identity));
  emitGmailConnectionUpdated();
}

function clearStoredGmailState() {
  removePersistedValue("localStorage", GMAIL_SESSION_STORAGE_KEY);
  removePersistedValue("localStorage", GMAIL_IDENTITY_STORAGE_KEY);
  removePersistedValue("sessionStorage", GMAIL_OAUTH_STATE_KEY);
  removePersistedValue("sessionStorage", GMAIL_OAUTH_VERIFIER_KEY);
  emitGmailConnectionUpdated();
}

function getCanonicalIntegrationContext() {
  const cachedOrgContext = getCachedOrgContext();
  if (!cachedOrgContext?.orgId || !cachedOrgContext.userId || !isSupabaseConfigured()) {
    return null;
  }
  const supabase = getSupabaseBrowserClient();
  if (!supabase) return null;
  return {
    supabase,
    orgId: cachedOrgContext.orgId,
    userId: cachedOrgContext.userId,
  };
}

function getConnectedGmailEmail(identity: GmailConnectedIdentity | null | undefined) {
  return normalizeGmailEmail(identity?.normalizedEmail || identity?.email);
}

function buildCanonicalGmailPayload(input: {
  orgId: string;
  userId: string;
  session: GmailSession | null;
  identity: GmailConnectedIdentity | null;
  status: GmailConnectionStatus;
}) {
  return {
    org_id: input.orgId,
    provider: GMAIL_PROVIDER_NAME,
    connection_status: input.status,
    provider_account_id: input.identity?.id ?? null,
    provider_account_email: getConnectedGmailEmail(input.identity) || null,
    provider_display_name: input.identity?.displayName ?? null,
    access_token: input.session?.accessToken ?? null,
    refresh_token: input.session?.refreshToken ?? null,
    expires_at: input.session?.expiresAt ? new Date(input.session.expiresAt).toISOString() : null,
    scope: input.session?.scope ?? null,
    identity: input.identity ?? null,
    created_by: input.userId,
    updated_by: input.userId,
    updated_at: new Date().toISOString(),
  };
}

async function persistCanonicalGmailIntegration(input: {
  session: GmailSession | null;
  identity: GmailConnectedIdentity | null;
  status: GmailConnectionStatus;
}) {
  const context = getCanonicalIntegrationContext();
  if (!context) return;

  await context.supabase.from("provider_integrations").upsert(
    buildCanonicalGmailPayload({
      orgId: context.orgId,
      userId: context.userId,
      session: input.session,
      identity: input.identity,
      status: input.status,
    }),
    { onConflict: "org_id,provider" }
  );
}

async function clearCanonicalGmailIntegration() {
  const context = getCanonicalIntegrationContext();
  if (!context) return;

  await context.supabase.from("provider_integrations").upsert(
    buildCanonicalGmailPayload({
      orgId: context.orgId,
      userId: context.userId,
      session: null,
      identity: null,
      status: "not_connected",
    }),
    { onConflict: "org_id,provider" }
  );
}

async function fetchGmailIdentity(accessToken: string) {
  const response = await fetch("https://www.googleapis.com/oauth2/v2/userinfo", {
    headers: {
      Authorization: `Bearer ${accessToken}`,
    },
  });

  const payload = (await response.json().catch(() => null)) as
    | {
        id?: string;
        email?: string;
        name?: string;
        error?: { message?: string };
      }
    | null;

  if (!response.ok || !payload?.email || !payload.id) {
    throw new Error(payload?.error?.message || "Failed to load the connected Gmail account.");
  }

  const normalizedEmail = normalizeGmailEmail(payload.email);
  if (!normalizedEmail) {
    throw new Error("The connected Gmail account did not return a usable email.");
  }

  return {
    id: payload.id,
    email: normalizedEmail,
    displayName: payload.name ?? normalizedEmail,
    connectedAt: new Date().toISOString(),
    normalizedEmail,
  } satisfies GmailConnectedIdentity;
}

export function getConnectedGmailMailboxEmail(identity: GmailConnectedIdentity | null | undefined) {
  return getConnectedGmailEmail(identity);
}

export function getGmailConnectionState(expectedEmail?: string): GmailConnectionState {
  const session = loadStoredGmailSession();
  const rawSession = loadRawStoredGmailSession();
  const identity = loadStoredGmailIdentity();
  const normalizedExpectedEmail = normalizeGmailEmail(expectedEmail);
  const normalizedConnectedEmail = getConnectedGmailEmail(identity);
  const stale = Boolean(normalizedExpectedEmail && normalizedConnectedEmail && normalizedExpectedEmail !== normalizedConnectedEmail);
  const connected = Boolean(session?.accessToken && !stale);
  const reconnectRequired = Boolean((identity || normalizedExpectedEmail) && !session?.accessToken && !rawSession?.refreshToken);

  return {
    status: connected ? "connected" : reconnectRequired ? "reconnect_required" : "not_connected",
    connected,
    stale,
    expectedEmail: normalizedExpectedEmail,
    identity,
    normalizedConnectedEmail,
  } satisfies GmailConnectionState;
}

async function refreshStoredGmailSession(requiredScopes: string[] = []) {
  const rawSession = loadRawStoredGmailSession();
  const refreshToken = String(rawSession?.refreshToken ?? "").trim();
  if (!refreshToken) return null;

  const response = await fetch("/api/auth/google/refresh", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      refreshToken,
    }),
  });

  const payload = (await response.json().catch(() => null)) as
    | {
        access_token?: string;
        refresh_token?: string;
        expires_in?: number;
        scope?: string;
        error?: string;
        error_description?: string;
      }
    | null;

  if (!response.ok || !payload?.access_token) {
    return null;
  }

  const session = saveGmailSessionFromTokenPayload({
    ...payload,
    refresh_token: payload.refresh_token || refreshToken,
    scope: payload.scope || rawSession?.scope || null,
  });
  await persistCanonicalGmailIntegration({
    session,
    identity: loadStoredGmailIdentity(),
    status: loadStoredGmailIdentity() ? "connected" : "reconnect_required",
  });
  if (requiredScopes.length > 0 && !hasRequiredScopes(session, requiredScopes)) {
    return null;
  }
  return session;
}

export async function resolveGmailConnectionState(expectedEmail?: string, requiredScopes: string[] = []) {
  const cacheKey = getGmailResolutionCacheKey(expectedEmail, requiredScopes);
  const cachedResolution = gmailResolutionCache.get(cacheKey);
  if (cachedResolution) {
    return cachedResolution;
  }

  const resolutionPromise = (async () => {
  const context = getCanonicalIntegrationContext();
  let canonicalRowFound = false;
  let canonicalProviderValue: string | null = null;
  let canonicalConnectionStatus: string | null = null;
  let canonicalScope = "";
  let canonicalHasIdentity = false;
  let canonicalHasAccessToken = false;
  let canonicalHasRequiredScopes = false;
  let canonicalRejectReason: string | null = "no_canonical_context";
  if (context) {
    const { data } = await context.supabase
      .from("provider_integrations")
      .select(
        "provider,connection_status,access_token,refresh_token,expires_at,scope,provider_account_id,provider_account_email,provider_display_name,identity"
      )
      .eq("org_id", context.orgId)
      .eq("provider", GMAIL_PROVIDER_NAME)
      .maybeSingle();

    canonicalRowFound = Boolean(data);
    canonicalProviderValue = typeof data?.provider === "string" ? data.provider : null;
    canonicalConnectionStatus = typeof data?.connection_status === "string" ? data.connection_status : null;
    canonicalScope = typeof data?.scope === "string" ? data.scope : "";
    canonicalHasIdentity = Boolean(data?.identity && typeof data.identity === "object");
    canonicalHasAccessToken = typeof data?.access_token === "string" && data.access_token.length > 0;
    canonicalHasRequiredScopes = requiredScopes.length === 0 || hasAllScopes(canonicalScope, requiredScopes);

    console.info("[gmail] canonical provider read", {
      orgId: context.orgId,
      rowFound: canonicalRowFound,
      provider: canonicalProviderValue,
      connectionStatus: canonicalConnectionStatus,
      requiredScopes,
      canonicalScopes: canonicalScope.split(/\s+/).filter(Boolean),
      scopeCheckResult: canonicalHasRequiredScopes,
      scopesPresent: canonicalHasRequiredScopes,
      identityPresent: canonicalHasIdentity,
      tokenPresent: canonicalHasAccessToken,
    });

    if (data?.provider === GMAIL_PROVIDER_NAME && data.connection_status !== "not_connected") {
      canonicalRejectReason = null;
      const expiresAt = typeof data.expires_at === "string" ? Date.parse(data.expires_at) : null;
      const accessToken = typeof data.access_token === "string" ? data.access_token : null;
      const refreshToken = typeof data.refresh_token === "string" ? data.refresh_token : null;
      const scope = typeof data.scope === "string" ? data.scope : null;
      if (accessToken || refreshToken) {
        saveStoredGmailSession({
          accessToken: accessToken || loadRawStoredGmailSession()?.accessToken || "",
          refreshToken: refreshToken || undefined,
          expiresAt: expiresAt && Number.isFinite(expiresAt) ? expiresAt : Date.now() - 1000,
          scope: scope || undefined,
          obtainedAt: new Date().toISOString(),
        });
      }
      if (data.identity && typeof data.identity === "object") {
        const identity = data.identity as GmailConnectedIdentity;
        const normalizedEmail = normalizeGmailEmail(identity.normalizedEmail || identity.email);
        if (identity.id && normalizedEmail) {
          saveStoredGmailIdentity({
            ...identity,
            email: normalizedEmail,
            normalizedEmail,
          });
        }
      }
      if (!canonicalHasRequiredScopes) {
        canonicalRejectReason = "missing_required_scope";
      } else if (!canonicalHasAccessToken) {
        canonicalRejectReason = "missing_access_token";
      }
    } else if (!canonicalRowFound) {
      canonicalRejectReason = "provider_row_not_found";
    } else if (data?.provider !== GMAIL_PROVIDER_NAME) {
      canonicalRejectReason = "provider_mismatch";
    } else if (data.connection_status === "not_connected") {
      canonicalRejectReason = "connection_status_not_connected";
    }
  }

  let state = getGmailConnectionState(expectedEmail);
  let refreshAttempted = false;
  let refreshSucceeded = false;
  const validSession = loadStoredGmailSession(requiredScopes);
  if (!validSession && state.identity) {
    refreshAttempted = true;
    const refreshedSession = await refreshStoredGmailSession(requiredScopes);
    refreshSucceeded = Boolean(refreshedSession);
    state = getGmailConnectionState(expectedEmail);
  }

  const effectiveSession = loadStoredGmailSession(requiredScopes);
  const effectiveHasAccessToken = Boolean(effectiveSession?.accessToken);
  const effectiveHasRequiredScopes = requiredScopes.length === 0 || Boolean(effectiveSession && hasRequiredScopes(effectiveSession, requiredScopes));
  const reconnectRequired = Boolean(
    state.stale ||
      (state.identity && !effectiveHasAccessToken && !refreshSucceeded) ||
      (state.identity && requiredScopes.length > 0 && !effectiveHasRequiredScopes)
  );

  if (requiredScopes.length === 0) {
    return {
      ...state,
      status: state.connected ? "connected" : reconnectRequired ? "reconnect_required" : "not_connected",
    } satisfies GmailConnectionState;
  }

  const finalAvailable = Boolean(
    (canonicalConnectionStatus === "connected" || !canonicalRowFound) && effectiveHasRequiredScopes && effectiveHasAccessToken && !state.stale
  );
  const finalRejectReason = finalAvailable
    ? null
    : state.stale
      ? "stale_connected_email"
      : canonicalConnectionStatus !== "connected" && canonicalRowFound
        ? canonicalRejectReason ?? "connection_status_not_connected"
        : !effectiveHasRequiredScopes
          ? "missing_required_scope"
          : !effectiveHasAccessToken
            ? "missing_access_token"
            : canonicalRejectReason ?? "not_available";

  lastGmailAvailabilityDebugSnapshot = {
    providerRowFound: canonicalRowFound,
    providerValue: canonicalProviderValue,
    connectionStatus: canonicalConnectionStatus,
    scopesPresent: effectiveHasRequiredScopes,
    tokenPresent: effectiveHasAccessToken,
    identityPresent: canonicalHasIdentity,
    finalAvailable,
    rejectionReason: finalRejectReason,
  };

  console.info("[gmail] availability result", {
    providerRowFound: canonicalRowFound,
    providerValue: canonicalProviderValue,
    connectionStatus: canonicalConnectionStatus,
    scopesPresent: effectiveHasRequiredScopes,
    identityPresent: canonicalHasIdentity,
    tokenPresent: effectiveHasAccessToken,
    finalAvailable,
    rejectionReason: finalRejectReason,
    refreshAttempted,
    refreshSucceeded,
  });

  if ((state.connected || reconnectRequired) && !finalAvailable) {
    return {
      ...state,
      connected: false,
      status: "reconnect_required",
    } satisfies GmailConnectionState;
  }
  return state;
  })();

  gmailResolutionCache.set(cacheKey, resolutionPromise);
  try {
    return await resolutionPromise;
  } catch (error) {
    gmailResolutionCache.delete(cacheKey);
    throw error;
  }
}

async function connectGmailInteractively() {
  if (typeof window === "undefined") {
    throw new Error("Gmail connection is only available in the browser.");
  }

  const clientId = getGmailClientId();
  if (!clientId) {
    throw new Error("Gmail is not configured yet. Add NEXT_PUBLIC_GOOGLE_CLIENT_ID first.");
  }

  const redirectUri = getGmailRedirectUri();
  const state = randomString(24);
  const codeVerifier = randomString(96);
  const codeChallenge = await sha256(codeVerifier);

  writePersistedValue("sessionStorage", GMAIL_OAUTH_STATE_KEY, state);
  writePersistedValue("sessionStorage", GMAIL_OAUTH_VERIFIER_KEY, codeVerifier);
  writeGmailOAuthVerifierCookie(codeVerifier);

  const url = new URL("https://accounts.google.com/o/oauth2/v2/auth");
  url.searchParams.set("client_id", clientId);
  url.searchParams.set("response_type", "code");
  url.searchParams.set("redirect_uri", redirectUri);
  url.searchParams.set("scope", GMAIL_SCOPES.join(" "));
  url.searchParams.set("state", state);
  url.searchParams.set("code_challenge", codeChallenge);
  url.searchParams.set("code_challenge_method", "S256");
  url.searchParams.set("access_type", "offline");
  url.searchParams.set("prompt", "consent select_account");
  url.searchParams.set("include_granted_scopes", "true");

  const popup = window.open(url.toString(), "event-based-reminders-app-gmail-auth", "width=640,height=760");
  if (!popup) {
    throw new Error("Unable to open Gmail sign-in. Please allow pop-ups and try again.");
  }

  return await new Promise<GmailSession>((resolve, reject) => {
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

    const handleMessage = (event: MessageEvent) => {
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

      if (data?.type !== GMAIL_OAUTH_MESSAGE_TYPE) return;

      const expectedState = readPersistedValue("sessionStorage", GMAIL_OAUTH_STATE_KEY);
      removePersistedValue("sessionStorage", GMAIL_OAUTH_STATE_KEY);
      removePersistedValue("sessionStorage", GMAIL_OAUTH_VERIFIER_KEY);
      clearGmailOAuthVerifierCookie();

      if (data.error) {
        fail(data.errorDescription || "Gmail permission was not granted.");
        return;
      }

      if (!data.state || !expectedState || data.state !== expectedState) {
        fail("Gmail sign-in could not be verified. Please try again.");
        return;
      }

      cleanup();
      try {
        resolve(
          saveGmailSessionFromTokenPayload({
            access_token: data.accessToken,
            refresh_token: data.refreshToken,
            expires_in: data.expiresIn,
            scope: data.scope,
          })
        );
      } catch (error) {
        reject(error instanceof Error ? error : new Error("Failed to connect Gmail."));
      }
    };

    const pollTimer = window.setInterval(() => {
      if (settled) return;
      if (popup.closed) fail("Gmail sign-in was closed before it finished.");
    }, 500);

    window.addEventListener("message", handleMessage);
  });
}

export async function connectGmail(expectedEmail?: string) {
  const session = await connectGmailInteractively();
  const identity = await fetchGmailIdentity(session.accessToken);
  saveStoredGmailIdentity(identity);
  await persistCanonicalGmailIntegration({
    session,
    identity,
    status: "connected",
  });
  return getGmailConnectionState(expectedEmail || identity.normalizedEmail);
}

export function disconnectGmail() {
  clearStoredGmailState();
  void clearCanonicalGmailIntegration();
}

function normalizeRecipients(raw: string[] | undefined) {
  return (raw ?? [])
    .map((value) => value.trim())
    .filter(Boolean)
    .map((value) => value.replace(/\r?\n/g, " "));
}

function encodeMimeSubject(value: string) {
  const normalized = value.trim();
  if (!normalized) return "";
  const hasNonAscii = /[^\x00-\x7F]/.test(normalized);
  if (!hasNonAscii) return normalized;
  const bytes = new TextEncoder().encode(normalized);
  return `=?UTF-8?B?${btoa(String.fromCharCode(...bytes))}?=`;
}

function buildRawMimeMessage(input: { draft: GmailEmailDraft; fallbackSubject: string }) {
  const to = normalizeRecipients(input.draft.to);
  const cc = normalizeRecipients(input.draft.cc);
  const bcc = normalizeRecipients(input.draft.bcc);
  const subject = input.draft.subject?.trim() || input.fallbackSubject.trim() || "Email draft";
  const body = String(input.draft.body ?? "").replace(/\r?\n/g, "\r\n");

  const headers = [
    "MIME-Version: 1.0",
    "Content-Type: text/plain; charset=UTF-8",
    "Content-Transfer-Encoding: 8bit",
    `Subject: ${encodeMimeSubject(subject)}`,
    ...(to.length ? [`To: ${to.join(", ")}`] : []),
    ...(cc.length ? [`Cc: ${cc.join(", ")}`] : []),
    ...(bcc.length ? [`Bcc: ${bcc.join(", ")}`] : []),
  ];

  return base64UrlEncode(new TextEncoder().encode(`${headers.join("\r\n")}\r\n\r\n${body}`));
}

export async function createGmailDraftFromEmailDraft(input: {
  draft: GmailEmailDraft;
  fallbackSubject: string;
}) {
  const connection = await resolveGmailConnectionState(undefined, [GMAIL_COMPOSE_SCOPE]);
  let session = loadStoredGmailSession([GMAIL_COMPOSE_SCOPE]);

  console.info("[gmail] draft creation preflight", {
    provider: GMAIL_PROVIDER_NAME,
    connectionStatus: connection.status,
    connected: connection.connected,
    stale: connection.stale,
    tokenPresent: Boolean(session?.accessToken),
  });

  if (!session?.accessToken && connection.connected && !connection.stale) {
    session = loadStoredGmailSession([GMAIL_COMPOSE_SCOPE]);
  }

  if (!session?.accessToken) {
    throw new Error("Connect Gmail in Settings before continuing.");
  }

  const response = await fetch("https://gmail.googleapis.com/gmail/v1/users/me/drafts", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${session.accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      message: {
        raw: buildRawMimeMessage(input),
      },
    }),
  });

  const payload = (await response.json().catch(() => null)) as
    | {
        id?: string;
        error?: { message?: string };
      }
    | null;

  if (!response.ok || !payload?.id) {
    throw new Error(payload?.error?.message || "Gmail draft creation failed.");
  }

  return {
    id: payload.id,
    webLink: "",
  } satisfies GmailDraftResult;
}

export async function sendGmailEmailFromEmailDraft(input: {
  draft: GmailEmailDraft;
  fallbackSubject: string;
}) {
  const connection = await resolveGmailConnectionState(undefined, [GMAIL_COMPOSE_SCOPE]);
  let session = loadStoredGmailSession([GMAIL_COMPOSE_SCOPE]);

  if (!session?.accessToken && connection.connected && !connection.stale) {
    session = loadStoredGmailSession([GMAIL_COMPOSE_SCOPE]);
  }

  if (!session?.accessToken) {
    throw new Error("Connect Gmail in Settings before continuing.");
  }

  const endpoint = buildLocalApiUrl("/api/provider-mail/send");
  const requestBody = {
    accessToken: session.accessToken,
    raw: buildRawMimeMessage(input),
  };
  console.info("[gmail] send request start", {
    endpoint,
    tokenPresent: Boolean(session.accessToken),
    toCount: normalizeRecipients(input.draft.to).length,
    ccCount: normalizeRecipients(input.draft.cc).length,
    bccCount: normalizeRecipients(input.draft.bcc).length,
    subject: input.draft.subject?.trim() || input.fallbackSubject.trim() || "Email draft",
  });

  let response: Response;
  try {
    response = await fetch(endpoint, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(requestBody),
    });
  } catch (error) {
    console.error("[gmail] send request failed before response", {
      endpoint,
      error: error instanceof Error ? error.message : String(error),
    });
    throw error;
  }

  console.info("[gmail] send response received", {
    endpoint,
    ok: response.ok,
    status: response.status,
  });

  const payload = (await response.json().catch(() => null)) as
    | {
        id?: string;
        error?: { message?: string };
      }
    | null;

  if (!response.ok || !payload?.id) {
    throw new Error(payload?.error?.message || "Gmail email send failed.");
  }

  return {
    id: payload.id,
  } satisfies GmailSendResult;
}

export async function createGoogleCalendarEvent(input: {
  subject: string;
  bodyText?: string;
  startISO: string;
  endISO: string;
  timeZone: string;
  isAllDay?: boolean;
  location?: string;
  attendees?: string[];
  teamsMeeting?: boolean;
  addGoogleMeet?: boolean;
}) {
  const connection = await resolveGmailConnectionState(undefined, [GOOGLE_CALENDAR_EVENTS_SCOPE]);
  let session = loadStoredGmailSession([GOOGLE_CALENDAR_EVENTS_SCOPE]);

  if (!session?.accessToken && connection.connected && !connection.stale) {
    session = loadStoredGmailSession([GOOGLE_CALENDAR_EVENTS_SCOPE]);
  }

  if (!session?.accessToken) {
    throw new Error("Reconnect Google in Settings to continue.");
  }

  const trimmedDescription = String(input.bodyText ?? "").trim();
  const googleAttendees = normalizeGoogleCalendarAttendees(input.attendees);
  const payload = {
    summary: input.subject.trim() || "Calendar event",
    ...(trimmedDescription ? { description: trimmedDescription } : {}),
    ...(input.location?.trim() ? { location: input.location.trim() } : {}),
    ...(input.isAllDay
      ? {
          start: { date: input.startISO.slice(0, 10) },
          end: { date: input.endISO.slice(0, 10) },
        }
      : {
          start: {
            dateTime: input.startISO,
            timeZone: input.timeZone,
          },
          end: {
            dateTime: input.endISO,
            timeZone: input.timeZone,
          },
        }),
    ...(googleAttendees.length > 0
      ? {
          attendees: googleAttendees,
        }
      : {}),
    ...(input.addGoogleMeet
      ? {
          conferenceData: {
            createRequest: {
              requestId: crypto.randomUUID(),
              conferenceSolutionKey: { type: "hangoutsMeet" },
            },
          },
        }
      : {}),
  };

  const endpoint = new URL("https://www.googleapis.com/calendar/v3/calendars/primary/events");
  if (input.addGoogleMeet) {
    endpoint.searchParams.set("conferenceDataVersion", "1");
  }

  const proxyEndpoint = buildLocalApiUrl("/api/provider-calendar/events");
  const requestBody = {
    accessToken: session.accessToken,
    payload,
    conferenceDataVersion: input.addGoogleMeet ? 1 : 0,
  };
  console.info("[googleCalendar] create request start", {
    endpoint: proxyEndpoint,
    tokenPresent: Boolean(session.accessToken),
    subject: payload.summary,
    attendeeCount: googleAttendees.length,
    addGoogleMeet: Boolean(input.addGoogleMeet),
    isAllDay: Boolean(input.isAllDay),
    requestPayload: payload,
  });

  let response: Response;
  try {
    response = await fetch(proxyEndpoint, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(requestBody),
    });
  } catch (error) {
    console.error("[googleCalendar] create request failed before response", {
      endpoint: proxyEndpoint,
      error: error instanceof Error ? error.message : String(error),
    });
    throw error;
  }

  console.info("[googleCalendar] create response received", {
    endpoint: proxyEndpoint,
    ok: response.ok,
    status: response.status,
  });

  const responsePayload = (await response.json().catch(() => null)) as
    | {
        id?: string;
        htmlLink?: string;
        hangoutLink?: string;
        conferenceData?: {
          entryPoints?: Array<{
            entryPointType?: string;
            uri?: string;
          }>;
        };
        error?: { message?: string };
      }
    | null;

  console.info("[googleCalendar] create response payload", {
    status: response.status,
    ok: response.ok,
    responsePayload,
  });

  const eventId = typeof responsePayload?.id === "string" && responsePayload.id.trim() ? responsePayload.id : undefined;
  const webLink =
    typeof responsePayload?.htmlLink === "string" && responsePayload.htmlLink.trim() ? responsePayload.htmlLink : "";

  if (!(response.status === 200 || response.status === 201) || (!eventId && !webLink)) {
    if (isGoogleCalendarApiDisabledError(response.status, responsePayload)) {
      throw new Error("Google Calendar API is not enabled for this Google connection yet. Enable it, then retry.");
    }
    if (isGoogleCalendarScopeError(response.status, responsePayload)) {
      throw new Error("Reconnect Google to enable calendar access.");
    }
    throw new Error(responsePayload?.error?.message || "Google Calendar event creation failed.");
  }

  const successPayload = responsePayload ?? {};
  const conferenceEntryPoints = Array.isArray(successPayload.conferenceData?.entryPoints)
    ? successPayload.conferenceData?.entryPoints ?? []
    : [];
  const googleMeetEntryPoint = conferenceEntryPoints.find((entry) => entry.entryPointType === "video" && entry.uri);
  const joinUrl = googleMeetEntryPoint?.uri ?? successPayload.hangoutLink ?? "";

  console.info("[googleCalendar] calendar success", {
    eventId: eventId ?? null,
    htmlLink: webLink || null,
    joinUrl: joinUrl || null,
  });

  return {
    id: eventId,
    success: true,
    webLink,
    joinUrl,
    hasOnlineMeeting: Boolean(joinUrl),
  } satisfies GoogleCalendarEventResult;
}
