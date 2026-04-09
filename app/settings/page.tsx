"use client";

import Link from "next/link";
import { useEffect, useRef, useState } from "react";

import {
  areAppSettingsEqual,
  hydrateAppSettingsFromSupabase,
  loadAppSettings,
  saveAppSettings,
  type AppSettings,
  type EmailHandlingMode,
} from "../../lib/appSettings";
import {
  connectOutlook,
  disconnectOutlook,
  getConnectedOutlookMailboxEmail,
  getOutlookConnectionState,
  resolveOutlookConnectionState,
  OUTLOOK_CONNECTION_UPDATED_EVENT,
  type OutlookConnectionState,
} from "../../lib/outlookClient";
import {
  connectGmail,
  disconnectGmail,
  getConnectedGmailMailboxEmail,
  getGmailConnectionState,
  resolveGmailConnectionState,
  GMAIL_CONNECTION_UPDATED_EVENT,
  type GmailConnectionState,
} from "../../lib/gmailClient";
import { useAuthContext } from "../components/auth-provider";

function areOutlookConnectionStatesEqual(left: OutlookConnectionState | null, right: OutlookConnectionState | null) {
  return JSON.stringify(left) === JSON.stringify(right);
}

function areGmailConnectionStatesEqual(left: GmailConnectionState | null, right: GmailConnectionState | null) {
  return JSON.stringify(left) === JSON.stringify(right);
}

type ConnectionProviderChoice = "auto" | "outlook" | "gmail";

function normalizeConnectionEmail(value: string) {
  return value.trim().replace(/^mailto:/i, "").trim().toLowerCase();
}

function inferConnectionProvider(email: string, choice: ConnectionProviderChoice): "outlook" | "gmail" {
  if (choice === "outlook" || choice === "gmail") return choice;
  const normalizedEmail = normalizeConnectionEmail(email);
  const domain = normalizedEmail.split("@")[1] ?? "";
  if (domain === "gmail.com" || domain === "googlemail.com") {
    return "gmail";
  }
  return "outlook";
}

export default function SettingsPage() {
  const { authEnabled, authBypassEnabled, currentUser, currentOrgId, signOut } = useAuthContext();
  const [settings, setSettings] = useState<AppSettings>(() => loadAppSettings());
  const [savedSettings, setSavedSettings] = useState<AppSettings>(() => loadAppSettings());
  const [planBuilderSaveMessage, setPlanBuilderSaveMessage] = useState<string | null>(null);
  const [emailSignatureSaveMessage, setEmailSignatureSaveMessage] = useState<string | null>(null);
  const [outlookConnection, setOutlookConnection] = useState<OutlookConnectionState | null>(null);
  const [outlookError, setOutlookError] = useState<string | null>(null);
  const [connectingOutlook, setConnectingOutlook] = useState(false);
  const [gmailConnection, setGmailConnection] = useState<GmailConnectionState | null>(null);
  const [gmailError, setGmailError] = useState<string | null>(null);
  const [connectingGmail, setConnectingGmail] = useState(false);
  const [providerLoading, setProviderLoading] = useState({ outlook: true, gmail: true });
  const [connectionEmailInput, setConnectionEmailInput] = useState<string>(() => loadAppSettings().outlookAccountEmail);
  const [connectionProviderChoice, setConnectionProviderChoice] = useState<ConnectionProviderChoice>("auto");
  const [connectionMessage, setConnectionMessage] = useState<string | null>(null);
  const [connectionError, setConnectionError] = useState<string | null>(null);
  const [signingOut, setSigningOut] = useState(false);
  const [hasMounted, setHasMounted] = useState(false);
  const savedSettingsRef = useRef(savedSettings);

  const hasUnsavedPlanBuilderChanges =
    settings.defaultReminderTime !== savedSettings.defaultReminderTime || settings.emailHandlingMode !== savedSettings.emailHandlingMode;
  const hasUnsavedEmailSignatureChanges = settings.emailSignatureText !== savedSettings.emailSignatureText;

  useEffect(() => {
    setHasMounted(true);
  }, []);

  useEffect(() => {
    savedSettingsRef.current = savedSettings;
  }, [savedSettings]);

  function persistSettings(nextSettings: AppSettings) {
    saveAppSettings(nextSettings);
    savedSettingsRef.current = nextSettings;
    setSavedSettings((current) => (areAppSettingsEqual(current, nextSettings) ? current : nextSettings));
  }

  function updateSettings<K extends keyof AppSettings>(key: K, value: AppSettings[K]) {
    setSettings((current) => ({ ...current, [key]: value }));
    if (key === "defaultReminderTime" || key === "emailHandlingMode") {
      setPlanBuilderSaveMessage(null);
    }
    if (key === "emailSignatureText") {
      setEmailSignatureSaveMessage(null);
    }
  }

  function buildSettingsFromConnection(current: AppSettings, connection: OutlookConnectionState) {
    const connectedEmail = getConnectedOutlookMailboxEmail(connection.identity);
    return {
      ...current,
      outlookAccountEmail: connection.status === "not_connected" ? "" : connectedEmail || current.outlookAccountEmail,
      outlookConnectionStatus: connection.status,
    };
  }

  function syncPersistedOutlookSettings(connection: OutlookConnectionState) {
    const nextSavedSettings = buildSettingsFromConnection(savedSettingsRef.current, connection);

    if (!areAppSettingsEqual(savedSettingsRef.current, nextSavedSettings)) {
      saveAppSettings(nextSavedSettings);
      savedSettingsRef.current = nextSavedSettings;
    }

    setSavedSettings((current) => (areAppSettingsEqual(current, nextSavedSettings) ? current : nextSavedSettings));
    setSettings((current) => {
      const nextSettings = {
        ...current,
        outlookAccountEmail:
          connection.status === "not_connected"
            ? ""
            : getConnectedOutlookMailboxEmail(connection.identity) ||
              current.outlookAccountEmail ||
              nextSavedSettings.outlookAccountEmail,
        outlookConnectionStatus: connection.status,
      };
      return areAppSettingsEqual(current, nextSettings) ? current : nextSettings;
    });
  }

  useEffect(() => {
    let active = true;

    async function refreshConnection(expectedEmail = savedSettingsRef.current.outlookAccountEmail) {
      setProviderLoading((current) => ({ ...current, outlook: true }));
      const outlookConnectionState = await resolveOutlookConnectionState(expectedEmail);
      if (!active) return;

      const nextSavedSettings = buildSettingsFromConnection(savedSettingsRef.current, outlookConnectionState);
      if (!areAppSettingsEqual(savedSettingsRef.current, nextSavedSettings)) {
        saveAppSettings(nextSavedSettings);
        savedSettingsRef.current = nextSavedSettings;
      }

      setSavedSettings((current) => (areAppSettingsEqual(current, nextSavedSettings) ? current : nextSavedSettings));
      setSettings((current) => {
        const nextSettings = {
          ...current,
          outlookAccountEmail:
            outlookConnectionState.status === "not_connected"
              ? ""
              : getConnectedOutlookMailboxEmail(outlookConnectionState.identity) ||
                current.outlookAccountEmail ||
                nextSavedSettings.outlookAccountEmail,
          outlookConnectionStatus: outlookConnectionState.status,
        };
        return areAppSettingsEqual(current, nextSettings) ? current : nextSettings;
      });
      setConnectionEmailInput((current) => current || getConnectedOutlookMailboxEmail(outlookConnectionState.identity) || "");
      setOutlookConnection((current) =>
        areOutlookConnectionStatesEqual(current, outlookConnectionState) ? current : outlookConnectionState
      );
      setProviderLoading((current) => ({ ...current, outlook: false }));
    }

    async function hydrateSettings() {
      const hydratedSettings = await hydrateAppSettingsFromSupabase();
      if (!active) return;
      savedSettingsRef.current = hydratedSettings;
      setSavedSettings((current) => (areAppSettingsEqual(current, hydratedSettings) ? current : hydratedSettings));
      setSettings((current) => (areAppSettingsEqual(current, hydratedSettings) ? current : hydratedSettings));
      void refreshConnection(hydratedSettings.outlookAccountEmail);
    }

    void hydrateSettings();
    void refreshConnection();
    function handleOutlookConnectionUpdated() {
      void refreshConnection();
    }
    window.addEventListener(OUTLOOK_CONNECTION_UPDATED_EVENT, handleOutlookConnectionUpdated);
    return () => {
      active = false;
      window.removeEventListener(OUTLOOK_CONNECTION_UPDATED_EVENT, handleOutlookConnectionUpdated);
    };
  }, []);

  useEffect(() => {
    let active = true;

    async function refreshGmailConnection() {
      setProviderLoading((current) => ({ ...current, gmail: true }));
      const connection = await resolveGmailConnectionState();
      if (!active) return;
      setConnectionEmailInput((current) => current || getConnectedGmailMailboxEmail(connection.identity) || "");
      setGmailConnection((current) => (areGmailConnectionStatesEqual(current, connection) ? current : connection));
      setProviderLoading((current) => ({ ...current, gmail: false }));
    }

    void refreshGmailConnection();
    function handleGmailConnectionUpdated() {
      void refreshGmailConnection();
    }
    window.addEventListener(GMAIL_CONNECTION_UPDATED_EVENT, handleGmailConnectionUpdated);
    return () => {
      active = false;
      window.removeEventListener(GMAIL_CONNECTION_UPDATED_EVENT, handleGmailConnectionUpdated);
    };
  }, []);

  function onSavePlanBuilderSettings() {
    persistSettings({
      ...savedSettingsRef.current,
      defaultReminderTime: settings.defaultReminderTime,
      emailHandlingMode: settings.emailHandlingMode,
    });
    setPlanBuilderSaveMessage("Plan builder defaults saved.");
  }

  function onSaveEmailSignatureSettings() {
    const normalizedSignatureText = settings.emailSignatureText;
    persistSettings({
      ...savedSettingsRef.current,
      emailSignatureEnabled: Boolean(normalizedSignatureText.trim()),
      emailSignatureText: normalizedSignatureText,
    });
    setEmailSignatureSaveMessage("Email signature saved.");
  }

  async function onConnectOutlook() {
    try {
      setConnectingOutlook(true);
      setOutlookError(null);
      const normalizedEmail = normalizeConnectionEmail(connectionEmailInput) || savedSettingsRef.current.outlookAccountEmail;
      const connection = await connectOutlook(normalizedEmail);
      setOutlookConnection((current) => (areOutlookConnectionStatesEqual(current, connection) ? current : connection));
      syncPersistedOutlookSettings(connection);
      setConnectionEmailInput(getConnectedOutlookMailboxEmail(connection.identity) || normalizedEmail);
      return true;
    } catch (error) {
      setOutlookError(error instanceof Error ? error.message : "Failed to connect Outlook.");
      return false;
    } finally {
      setConnectingOutlook(false);
    }
  }

  async function onDisconnectOutlook() {
    try {
      setOutlookError(null);
      await disconnectOutlook();
      const connection = getOutlookConnectionState(savedSettingsRef.current.outlookAccountEmail);
      setOutlookConnection((current) => (areOutlookConnectionStatesEqual(current, connection) ? current : connection));
      syncPersistedOutlookSettings(connection);
      return true;
    } catch (error) {
      setOutlookError(error instanceof Error ? error.message : "Failed to disconnect Outlook.");
      return false;
    }
  }

  async function onConnectGmail() {
    try {
      setConnectingGmail(true);
      setGmailError(null);
      const normalizedEmail = normalizeConnectionEmail(connectionEmailInput);
      const connection = await connectGmail(normalizedEmail || undefined);
      setGmailConnection((current) => (areGmailConnectionStatesEqual(current, connection) ? current : connection));
      setConnectionEmailInput(getConnectedGmailMailboxEmail(connection.identity) || normalizedEmail);
      return true;
    } catch (error) {
      setGmailError(error instanceof Error ? error.message : "Failed to connect Gmail.");
      return false;
    } finally {
      setConnectingGmail(false);
    }
  }

  function onDisconnectGmail() {
    setGmailError(null);
    disconnectGmail();
    const connection = getGmailConnectionState();
    setGmailConnection((current) => (areGmailConnectionStatesEqual(current, connection) ? current : connection));
    return true;
  }

  async function onSignOut() {
    try {
      setSigningOut(true);
      await signOut();
    } finally {
      setSigningOut(false);
    }
  }

  const connectionStatusLabel = !hasMounted
    ? "Not connected"
    : outlookConnection?.status === "connected"
      ? "Connected"
      : outlookConnection?.status === "reconnect_required"
        ? "Reconnect required"
        : "Not connected";
  const connectedAccountEmail = hasMounted
    ? outlookConnection?.status === "not_connected"
      ? "—"
      : getConnectedOutlookMailboxEmail(outlookConnection?.identity) || settings.outlookAccountEmail || "—"
    : "—";
  const connectedDisplayName = hasMounted ? outlookConnection?.identity?.displayName || "—" : "—";
  const showMailboxWarning =
    hasMounted && !outlookConnection?.supportedMailbox && Boolean(outlookConnection?.identity);
  const gmailConnectionStatusLabel = !hasMounted
    ? "Not connected"
    : gmailConnection?.status === "connected"
      ? "Connected"
      : gmailConnection?.status === "reconnect_required"
        ? "Reconnect required"
        : "Not connected";
  const connectedGmailEmail = hasMounted ? getConnectedGmailMailboxEmail(gmailConnection?.identity) || "—" : "—";
  const connectedGmailDisplayName = hasMounted ? gmailConnection?.identity?.displayName || "—" : "—";
  const isConnectingProvider = connectingOutlook || connectingGmail;
  const inferredProvider = inferConnectionProvider(connectionEmailInput, connectionProviderChoice);
  const targetConnectionStatus = inferredProvider === "gmail" ? gmailConnection?.status : outlookConnection?.status;
  const primaryConnectedProvider =
    outlookConnection?.status === "connected"
      ? "outlook"
      : gmailConnection?.status === "connected"
        ? "gmail"
        : outlookConnection?.status === "reconnect_required"
          ? "outlook"
          : gmailConnection?.status === "reconnect_required"
            ? "gmail"
            : null;
  const primaryConnectedEmail =
    primaryConnectedProvider === "outlook"
      ? getConnectedOutlookMailboxEmail(outlookConnection?.identity) || settings.outlookAccountEmail || "—"
      : primaryConnectedProvider === "gmail"
        ? getConnectedGmailMailboxEmail(gmailConnection?.identity) || "—"
        : "—";
  const primaryConnectedStatus =
    primaryConnectedProvider === "outlook"
      ? connectionStatusLabel
      : primaryConnectedProvider === "gmail"
        ? gmailConnectionStatusLabel
        : "Not connected";
  const primaryConnectedLabel =
    primaryConnectedProvider === "outlook" ? "Outlook" : primaryConnectedProvider === "gmail" ? "Google" : "No provider connected";

  function clearConnectionFeedback() {
    setConnectionMessage(null);
    setConnectionError(null);
    setOutlookError(null);
    setGmailError(null);
  }

  async function onConnectProvider() {
    clearConnectionFeedback();
    const provider = inferConnectionProvider(connectionEmailInput, connectionProviderChoice);
    const succeeded = provider === "gmail" ? await onConnectGmail() : await onConnectOutlook();
    if (succeeded) {
      setConnectionMessage(provider === "gmail" ? "Google account connected." : "Outlook account connected.");
    } else {
      setConnectionError(provider === "gmail" ? "Failed to connect Google account." : "Failed to connect Outlook account.");
    }
  }

  async function onDisconnectProvider() {
    clearConnectionFeedback();
    const provider =
      inferredProvider === "gmail"
        ? gmailConnection?.status && gmailConnection.status !== "not_connected"
          ? "gmail"
          : outlookConnection?.status && outlookConnection.status !== "not_connected"
            ? "outlook"
            : "gmail"
        : outlookConnection?.status && outlookConnection.status !== "not_connected"
          ? "outlook"
          : gmailConnection?.status && gmailConnection.status !== "not_connected"
            ? "gmail"
            : "outlook";

    if (provider === "gmail") {
      const succeeded = onDisconnectGmail();
      if (succeeded) {
        setConnectionMessage("Google account disconnected.");
      } else {
        setConnectionError("Failed to disconnect Google account.");
      }
      return;
    }
    const succeeded = await onDisconnectOutlook();
    if (succeeded) {
      setConnectionMessage("Outlook account disconnected.");
    } else {
      setConnectionError("Failed to disconnect Outlook account.");
    }
  }

  return (
    <div className="space-y-8 text-gray-900">
      <section className="space-y-2">
        <div className="flex flex-col gap-4 md:flex-row md:items-start md:justify-between">
          <div>
            <h1 className="text-3xl font-bold text-gray-900">Settings</h1>
            <p className="mt-2 max-w-2xl text-sm text-gray-600">
              Configure local defaults and preview behavior for the Event-Based Reminders app.
            </p>
          </div>
          <div className="rounded-xl border bg-white px-4 py-3 text-sm shadow-sm md:min-w-[280px] md:max-w-[320px]">
            <div className="text-xs font-semibold uppercase tracking-wide text-gray-500">Navigation</div>
            <div className="mt-1 font-medium text-gray-900">Event-Based Reminders configuration</div>
            <div className="mt-3 flex flex-wrap gap-2">
              <Link href="/plans" className="rounded-lg border px-3 py-1.5 text-xs font-medium text-gray-700 hover:bg-gray-50">
                Back to Plans
              </Link>
            </div>
          </div>
        </div>
      </section>

      <section className="grid gap-6 lg:grid-cols-[minmax(0,1fr)_minmax(0,1fr)]">
        <div className="rounded-2xl border bg-white shadow-sm">
          <div className="space-y-5 p-6">
            <div>
              <h2 className="text-lg font-semibold text-gray-900">Plan Builder Defaults</h2>
              <p className="mt-2 text-sm text-gray-600">Controls the default reminder time and email behavior used when building new plans.</p>
            </div>
            <label className="block space-y-1 text-sm">
              <span className="font-medium text-gray-700">Default reminder time</span>
              <input
                type="time"
                value={settings.defaultReminderTime}
                onChange={(e) => updateSettings("defaultReminderTime", e.target.value)}
                className="w-full rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900"
              />
            </label>
            <label className="block space-y-1 text-sm">
              <span className="font-medium text-gray-700">Email handling mode</span>
              <select
                value={settings.emailHandlingMode}
                onChange={(e) => updateSettings("emailHandlingMode", e.target.value as EmailHandlingMode)}
                className="w-full rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900"
              >
                <option value="draft">Save to Drafts</option>
                <option value="schedule">Schedule Send (Outlook only; Gmail saves draft)</option>
                <option value="send">Send Immediately</option>
              </select>
            </label>
            <div className="flex items-center justify-between gap-3">
              <button
                type="button"
                onClick={onSavePlanBuilderSettings}
                disabled={!hasUnsavedPlanBuilderChanges}
                className="rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900 hover:bg-gray-50 disabled:text-gray-900"
              >
                Save Plan Builder Defaults
              </button>
            </div>
            {planBuilderSaveMessage ? <p className="text-xs text-green-700">{planBuilderSaveMessage}</p> : null}
          </div>
        </div>

        <div className="rounded-2xl border bg-white shadow-sm">
          <div className="space-y-5 p-6">
            <div>
              <h2 className="text-lg font-semibold text-gray-900">Connected Account</h2>
              <p className="mt-2 text-sm text-gray-600">
                Enter the account email once, then connect with Outlook or Google from the same box.
              </p>
            </div>
            <div className="grid gap-3 md:grid-cols-[minmax(0,1fr)_180px]">
              <label className="block space-y-1 text-sm">
                <span className="font-medium text-gray-700">Account email</span>
                <input
                  type="email"
                  value={connectionEmailInput}
                  onChange={(e) => setConnectionEmailInput(e.target.value)}
                  placeholder="name@company.com"
                  className="w-full rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900"
                />
              </label>
              <label className="block space-y-1 text-sm">
                <span className="font-medium text-gray-700">Provider</span>
                <select
                  value={connectionProviderChoice}
                  onChange={(e) => setConnectionProviderChoice(e.target.value as ConnectionProviderChoice)}
                  className="w-full rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900"
                >
                  <option value="auto">Auto-detect</option>
                  <option value="outlook">Outlook / Microsoft</option>
                  <option value="gmail">Google</option>
                </select>
              </label>
            </div>
            <div className="rounded-lg border border-gray-200 p-4">
              <p className="text-xs text-gray-500">Active provider</p>
              <p className="mt-1 text-sm font-medium text-gray-900">{primaryConnectedLabel}</p>
              <p className="mt-3 text-xs text-gray-500">Connected account email</p>
              <p className="mt-1 text-sm text-gray-900">{primaryConnectedEmail}</p>
              <p className="mt-3 text-xs text-gray-500">Connection status</p>
              <p className="mt-1 text-sm font-medium text-gray-900">{primaryConnectedStatus}</p>
              {providerLoading.outlook || providerLoading.gmail ? (
                <p className="mt-2 text-xs text-gray-500">Refreshing provider status…</p>
              ) : null}
              <div className="mt-4 grid gap-3 md:grid-cols-2">
                <div className="rounded-lg border border-gray-100 bg-gray-50 p-3">
                  <div className="text-xs font-semibold uppercase tracking-wide text-gray-500">Outlook</div>
                  <div className="mt-1 text-sm font-medium text-gray-900">{connectionStatusLabel}</div>
                  <div className="mt-1 text-xs text-gray-600">{connectedAccountEmail}</div>
                  <div className="mt-1 text-xs text-gray-500">{connectedDisplayName}</div>
                </div>
                <div className="rounded-lg border border-gray-100 bg-gray-50 p-3">
                  <div className="text-xs font-semibold uppercase tracking-wide text-gray-500">Google</div>
                  <div className="mt-1 text-sm font-medium text-gray-900">{gmailConnectionStatusLabel}</div>
                  <div className="mt-1 text-xs text-gray-600">{connectedGmailEmail}</div>
                  <div className="mt-1 text-xs text-gray-500">{connectedGmailDisplayName}</div>
                </div>
              </div>
            </div>
            <div className="flex flex-wrap items-center gap-3">
              <button
                type="button"
                onClick={() => void onConnectProvider()}
                disabled={isConnectingProvider}
                className="rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900 hover:bg-gray-50 disabled:opacity-60"
              >
                {isConnectingProvider
                  ? "Connecting..."
                  : targetConnectionStatus === "reconnect_required"
                    ? "Reconnect Account"
                    : "Connect Account"}
              </button>
              <button
                type="button"
                onClick={() => void onDisconnectProvider()}
                disabled={
                  (outlookConnection?.status ?? "not_connected") === "not_connected" &&
                  (gmailConnection?.status ?? "not_connected") === "not_connected"
                }
                className="rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900 hover:bg-gray-50 disabled:opacity-60"
              >
                Disconnect
              </button>
            </div>
            <p className="text-xs text-gray-500">
              Auto-detect uses Google for `gmail.com` addresses and Outlook/Microsoft for everything else. You can override it from the provider menu.
            </p>
            {showMailboxWarning ? (
              <p className="text-xs text-amber-700">{outlookConnection?.identity?.mailboxEligibilityReason}</p>
            ) : null}
            {connectionMessage ? <p className="text-xs text-green-700">{connectionMessage}</p> : null}
            {connectionError || outlookError || gmailError ? (
              <p className="text-xs text-red-700">{connectionError || outlookError || gmailError}</p>
            ) : null}
          </div>
        </div>

        <div className="rounded-2xl border bg-white shadow-sm lg:col-span-2">
          <div className="space-y-5 p-6">
            <div>
              <h2 className="text-lg font-semibold text-gray-900">Email Signature</h2>
              <p className="mt-2 text-sm text-gray-600">Used automatically in Plans emails whenever this field has text.</p>
            </div>
            <label className="block space-y-1 text-sm">
              <span className="font-medium text-gray-700">Signature</span>
              <textarea
                value={settings.emailSignatureText}
                onChange={(e) => updateSettings("emailSignatureText", e.target.value)}
                className="min-h-32 w-full rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900"
                placeholder={"Best,\nYour Name"}
              />
            </label>
            <div className="flex items-center justify-between gap-3">
              <button
                type="button"
                onClick={onSaveEmailSignatureSettings}
                disabled={!hasUnsavedEmailSignatureChanges}
                className="rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900 hover:bg-gray-50 disabled:text-gray-900"
              >
                Save Signature
              </button>
            </div>
            {emailSignatureSaveMessage ? <p className="text-xs text-green-700">{emailSignatureSaveMessage}</p> : null}
          </div>
        </div>

        {authEnabled && currentUser ? (
          <div className="rounded-2xl border bg-white shadow-sm lg:col-span-2">
            <div className="space-y-4 p-6">
              <div>
                <h2 className="text-base font-semibold text-gray-900">Account</h2>
              </div>
              <div className="space-y-2 text-sm text-gray-600">
                <div>
                  Signed in as: <span className="text-gray-900">{currentUser.email || "—"}</span>
                </div>
                <div>
                  Org ID: <span className="text-gray-900">{currentOrgId || "—"}</span>
                </div>
              </div>
              <div>
                <button
                  type="button"
                  onClick={() => void onSignOut()}
                  disabled={signingOut}
                  className="rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900 hover:bg-gray-50 disabled:opacity-60"
                >
                  {signingOut ? "Signing out..." : "Sign out"}
                </button>
              </div>
            </div>
          </div>
        ) : authBypassEnabled ? (
          <div className="rounded-2xl border bg-white shadow-sm lg:col-span-2">
            <div className="space-y-2 p-6">
              <h2 className="text-base font-semibold text-gray-900">Account</h2>
              <div className="text-sm text-gray-600">Auth bypass enabled for this environment.</div>
            </div>
          </div>
        ) : null}
      </section>
    </div>
  );
}
