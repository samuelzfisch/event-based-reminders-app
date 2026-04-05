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
import { useAuthContext } from "../components/auth-provider";

function areOutlookConnectionStatesEqual(left: OutlookConnectionState | null, right: OutlookConnectionState | null) {
  return JSON.stringify(left) === JSON.stringify(right);
}

export default function SettingsPage() {
  const { authEnabled, authBypassEnabled, currentUser, currentOrgId, signOut } = useAuthContext();
  const [settings, setSettings] = useState<AppSettings>(() => loadAppSettings());
  const [savedSettings, setSavedSettings] = useState<AppSettings>(() => loadAppSettings());
  const [planBuilderSaveMessage, setPlanBuilderSaveMessage] = useState<string | null>(null);
  const [emailSignatureSaveMessage, setEmailSignatureSaveMessage] = useState<string | null>(null);
  const [outlookConnection, setOutlookConnection] = useState<OutlookConnectionState | null>(null);
  const [outlookMessage, setOutlookMessage] = useState<string | null>(null);
  const [outlookError, setOutlookError] = useState<string | null>(null);
  const [connectingOutlook, setConnectingOutlook] = useState(false);
  const [signingOut, setSigningOut] = useState(false);
  const [hasMounted, setHasMounted] = useState(false);
  const savedSettingsRef = useRef(savedSettings);

  const hasUnsavedPlanBuilderChanges =
    settings.defaultReminderTime !== savedSettings.defaultReminderTime || settings.emailHandlingMode !== savedSettings.emailHandlingMode;
  const hasUnsavedEmailSignatureChanges =
    settings.emailSignatureEnabled !== savedSettings.emailSignatureEnabled ||
    settings.emailSignatureText !== savedSettings.emailSignatureText;

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
    if (key === "emailSignatureEnabled" || key === "emailSignatureText") {
      setEmailSignatureSaveMessage(null);
    }
  }

  function buildSettingsFromConnection(current: AppSettings, connection: OutlookConnectionState) {
    const connectedEmail = getConnectedOutlookMailboxEmail(connection.identity);
    return {
      ...current,
      outlookAccountEmail: connectedEmail || current.outlookAccountEmail,
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
          getConnectedOutlookMailboxEmail(connection.identity) || current.outlookAccountEmail || nextSavedSettings.outlookAccountEmail,
        outlookConnectionStatus: connection.status,
      };
      return areAppSettingsEqual(current, nextSettings) ? current : nextSettings;
    });
  }

  useEffect(() => {
    let active = true;

    async function refreshConnection() {
      const connection = await resolveOutlookConnectionState(savedSettingsRef.current.outlookAccountEmail);
      if (!active) return;

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
            getConnectedOutlookMailboxEmail(connection.identity) || current.outlookAccountEmail || nextSavedSettings.outlookAccountEmail,
          outlookConnectionStatus: connection.status,
        };
        return areAppSettingsEqual(current, nextSettings) ? current : nextSettings;
      });
      setOutlookConnection((current) => (areOutlookConnectionStatesEqual(current, connection) ? current : connection));
    }

    async function hydrateAndRefresh() {
      const hydratedSettings = await hydrateAppSettingsFromSupabase();
      if (!active) return;
      savedSettingsRef.current = hydratedSettings;
      setSavedSettings((current) => (areAppSettingsEqual(current, hydratedSettings) ? current : hydratedSettings));
      setSettings((current) => (areAppSettingsEqual(current, hydratedSettings) ? current : hydratedSettings));
      await refreshConnection();
    }

    void hydrateAndRefresh();

    function handleWindowFocus() {
      void refreshConnection();
    }

    window.addEventListener("focus", handleWindowFocus);
    window.addEventListener(OUTLOOK_CONNECTION_UPDATED_EVENT, handleWindowFocus as EventListener);
    return () => {
      active = false;
      window.removeEventListener("focus", handleWindowFocus);
      window.removeEventListener(OUTLOOK_CONNECTION_UPDATED_EVENT, handleWindowFocus as EventListener);
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
    persistSettings({
      ...savedSettingsRef.current,
      emailSignatureEnabled: settings.emailSignatureEnabled,
      emailSignatureText: settings.emailSignatureText,
    });
    setEmailSignatureSaveMessage("Email signature saved.");
  }

  async function onConnectOutlook() {
    try {
      setConnectingOutlook(true);
      setOutlookMessage(null);
      setOutlookError(null);
      const connection = await connectOutlook(savedSettingsRef.current.outlookAccountEmail);
      setOutlookConnection((current) => (areOutlookConnectionStatesEqual(current, connection) ? current : connection));
      syncPersistedOutlookSettings(connection);
      setOutlookMessage("Outlook connected.");
    } catch (error) {
      setOutlookError(error instanceof Error ? error.message : "Failed to connect Outlook.");
    } finally {
      setConnectingOutlook(false);
    }
  }

  async function onReconnectOutlook() {
    await onConnectOutlook();
  }

  function onDisconnectOutlook() {
    setOutlookMessage(null);
    setOutlookError(null);
    disconnectOutlook();
    const connection = getOutlookConnectionState(savedSettingsRef.current.outlookAccountEmail);
    setOutlookConnection((current) => (areOutlookConnectionStatesEqual(current, connection) ? current : connection));
    syncPersistedOutlookSettings(connection);
    setOutlookMessage("Outlook disconnected.");
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
    ? getConnectedOutlookMailboxEmail(outlookConnection?.identity) || settings.outlookAccountEmail || "—"
    : "—";
  const connectedDisplayName = hasMounted ? outlookConnection?.identity?.displayName || "—" : "—";
  const showMailboxWarning =
    hasMounted && !outlookConnection?.supportedMailbox && Boolean(outlookConnection?.identity);

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
                <option value="schedule">Schedule Send</option>
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
              <h2 className="text-lg font-semibold text-gray-900">Outlook Account</h2>
              <p className="mt-2 text-sm text-gray-600">Connect Outlook with Microsoft sign-in so Plans can use your real mailbox later.</p>
            </div>
            <div className="rounded-lg border border-gray-200 p-4">
              <p className="text-xs text-gray-500">Connection status</p>
              <p className="mt-1 text-sm font-medium text-gray-900">{connectionStatusLabel}</p>
              <p className="mt-3 text-xs text-gray-500">Connected account email</p>
              <p className="mt-1 text-sm text-gray-900">{connectedAccountEmail}</p>
              <p className="mt-3 text-xs text-gray-500">Connected account display name</p>
              <p className="mt-1 text-sm text-gray-900">{connectedDisplayName}</p>
            </div>
            <div className="flex flex-wrap items-center gap-3">
              <button
                type="button"
                onClick={onConnectOutlook}
                disabled={connectingOutlook}
                className="rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900 hover:bg-gray-50 disabled:opacity-60"
              >
                {connectingOutlook ? "Connecting..." : "Connect Outlook"}
              </button>
              <button
                type="button"
                onClick={onReconnectOutlook}
                disabled={connectingOutlook}
                className="rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900 hover:bg-gray-50 disabled:opacity-60"
              >
                Reconnect
              </button>
              <button
                type="button"
                onClick={onDisconnectOutlook}
                className="rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900 hover:bg-gray-50"
              >
                Disconnect
              </button>
            </div>
            {showMailboxWarning ? (
              <p className="text-xs text-amber-700">{outlookConnection?.identity?.mailboxEligibilityReason}</p>
            ) : null}
            {outlookMessage ? <p className="text-xs text-green-700">{outlookMessage}</p> : null}
            {outlookError ? <p className="text-xs text-red-700">{outlookError}</p> : null}
          </div>
        </div>

        <div className="rounded-2xl border bg-white shadow-sm lg:col-span-2">
          <div className="space-y-5 p-6">
            <div>
              <h2 className="text-lg font-semibold text-gray-900">Email Signature</h2>
              <p className="mt-2 text-sm text-gray-600">Used in preview email content when signature support is enabled.</p>
            </div>
            <label className="flex items-center gap-3 text-sm text-gray-900">
              <input
                type="checkbox"
                checked={settings.emailSignatureEnabled}
                onChange={(e) => updateSettings("emailSignatureEnabled", e.target.checked)}
              />
              <span>Use signature in Plans emails</span>
            </label>
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
