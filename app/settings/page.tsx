"use client";

import { useEffect, useRef, useState, type CSSProperties, type FormEvent, type KeyboardEvent as ReactKeyboardEvent } from "react";
import { createPortal } from "react-dom";

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
type ProviderId = "outlook" | "gmail";

type DisconnectDialogState = {
  provider: ProviderId;
  label: string;
};

type ProviderTone = "success" | "attention" | "neutral" | "unavailable";

type ProviderStatusView = {
  label: string;
  tone: ProviderTone;
};

const PROVIDER_STATUS_ERROR_MESSAGE = "Connection status could not be loaded. Try again.";

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

function getProviderStatusView(
  status: OutlookConnectionState["status"] | GmailConnectionState["status"] | null | undefined,
  loading: boolean
): ProviderStatusView {
  if (loading && !status) return { label: "Checking...", tone: "neutral" };
  if (status === "connected") return { label: "Connected", tone: "success" };
  if (status === "reconnect_required") return { label: "Needs attention", tone: "attention" };
  if (status === "not_connected") return { label: "Not connected", tone: "neutral" };
  return { label: "Connection status is unavailable", tone: "unavailable" };
}

function getStatusToneClasses(tone: ProviderTone) {
  if (tone === "success") return { dot: "bg-green-500", text: "text-green-700" };
  if (tone === "attention") return { dot: "bg-amber-500", text: "text-amber-700" };
  if (tone === "unavailable") return { dot: "bg-amber-400", text: "text-slate-600" };
  return { dot: "bg-slate-300", text: "text-slate-600" };
}

function isConnectedOrAttention(status: OutlookConnectionState["status"] | GmailConnectionState["status"] | null | undefined) {
  return status === "connected" || status === "reconnect_required";
}

function getProviderConnectVerb(status: OutlookConnectionState["status"] | GmailConnectionState["status"] | null | undefined) {
  return status === "connected" || status === "reconnect_required" ? "Reconnect" : "Connect";
}

const knownOutlookSuggestionDomains = new Set(["outlook.com", "hotmail.com", "live.com", "msn.com"]);

function getProviderHelperCopy(email: string, choice: ConnectionProviderChoice) {
  if (choice === "outlook") return "Provider preference is set to Outlook. Connection actions remain explicit below.";
  if (choice === "gmail") return "Provider preference is set to Google. Connection actions remain explicit below.";

  const normalizedEmail = normalizeConnectionEmail(email);
  if (!normalizedEmail) return "Enter an email address to receive a provider suggestion. Provider actions remain explicit below.";

  const domain = normalizedEmail.split("@")[1] ?? "";
  if (domain === "gmail.com" || domain === "googlemail.com" || knownOutlookSuggestionDomains.has(domain)) {
    const suggestedProvider = inferConnectionProvider(normalizedEmail, "auto");
    const suggestedProviderLabel = suggestedProvider === "gmail" ? "Google" : "Outlook";
    return `Suggested provider for this address: ${suggestedProviderLabel}. Provider actions remain explicit below.`;
  }

  return "No provider suggestion is available for this address. Choose a provider below.";
}

const settingsSaveButtonBaseClass =
  "inline-flex h-[40px] min-w-[132px] items-center justify-center whitespace-nowrap rounded-[10px] border px-[16px] text-[14px] font-semibold transition focus:outline-none focus:ring-2";
const settingsSaveButtonEnabledClass =
  "cursor-pointer !border-[#4f7fb8] !bg-[#4f7fb8] !text-[#ffffff] opacity-100 hover:!bg-[#416f9f] focus:ring-[#6f9fd1]/35";
const settingsSaveButtonDisabledClass =
  "cursor-not-allowed !border-[#dbe5ee] !bg-[#f2f6f9] !text-[#5e6f84] opacity-100 hover:!bg-[#f2f6f9] focus:ring-slate-400/20";

function getSettingsSaveButtonClass(isEnabled: boolean) {
  return `${settingsSaveButtonBaseClass} ${isEnabled ? settingsSaveButtonEnabledClass : settingsSaveButtonDisabledClass}`;
}

function getSettingsSaveButtonStyle(isEnabled: boolean): CSSProperties {
  return isEnabled
    ? {
        backgroundColor: "#4f7fb8",
        borderColor: "#4f7fb8",
        color: "#ffffff",
      }
    : {
        backgroundColor: "#f2f6f9",
        borderColor: "#dbe5ee",
        color: "#5e6f84",
      };
}

const providerActionBaseClass =
  "inline-flex h-[40px] items-center justify-center whitespace-nowrap rounded-[10px] border px-[14px] text-[14px] font-semibold transition focus:outline-none focus:ring-2 disabled:cursor-not-allowed disabled:border-slate-200 disabled:bg-slate-100 disabled:text-slate-400 disabled:hover:bg-slate-100";
const providerPrimaryActionClass = `${providerActionBaseClass} !border-[#4f7fb8] !bg-[#4f7fb8] !text-[#ffffff] hover:!bg-[#416f9f] focus:ring-[#6f9fd1]/35`;
const providerReconnectActionClass = `${providerActionBaseClass} border-sky-200 bg-white text-sky-800 hover:border-sky-300 hover:bg-sky-50 focus:ring-sky-500/25`;
const providerDisconnectActionClass = `${providerActionBaseClass} border-red-200 bg-white text-red-700 hover:border-red-300 hover:bg-red-50 focus:ring-red-500/30`;

export default function SettingsPage() {
  const { authEnabled, authBypassEnabled, currentUser, signOut } = useAuthContext();
  const [settings, setSettings] = useState<AppSettings>(() => loadAppSettings());
  const [savedSettings, setSavedSettings] = useState<AppSettings>(() => loadAppSettings());
  const [planBuilderSaveMessage, setPlanBuilderSaveMessage] = useState<string | null>(null);
  const [planBuilderSaveError, setPlanBuilderSaveError] = useState<string | null>(null);
  const [savingPlanBuilder, setSavingPlanBuilder] = useState(false);
  const [emailSignatureSaveMessage, setEmailSignatureSaveMessage] = useState<string | null>(null);
  const [emailSignatureSaveError, setEmailSignatureSaveError] = useState<string | null>(null);
  const [savingEmailSignature, setSavingEmailSignature] = useState(false);
  const [outlookConnection, setOutlookConnection] = useState<OutlookConnectionState | null>(null);
  const [outlookError, setOutlookError] = useState<string | null>(null);
  const [connectingOutlook, setConnectingOutlook] = useState(false);
  const [gmailConnection, setGmailConnection] = useState<GmailConnectionState | null>(null);
  const [gmailError, setGmailError] = useState<string | null>(null);
  const [connectingGmail, setConnectingGmail] = useState(false);
  const [providerLoading, setProviderLoading] = useState({ outlook: true, gmail: true });
  const [providerRefreshKey, setProviderRefreshKey] = useState(0);
  const [connectionEmailInput, setConnectionEmailInput] = useState<string>(() => loadAppSettings().outlookAccountEmail);
  const [connectionProviderChoice, setConnectionProviderChoice] = useState<ConnectionProviderChoice>("auto");
  const [connectionMessage, setConnectionMessage] = useState<string | null>(null);
  const [connectionError, setConnectionError] = useState<string | null>(null);
  const [disconnectDialog, setDisconnectDialog] = useState<DisconnectDialogState | null>(null);
  const [disconnectingProvider, setDisconnectingProvider] = useState<ProviderId | null>(null);
  const [signingOut, setSigningOut] = useState(false);
  const [hasMounted, setHasMounted] = useState(false);
  const savedSettingsRef = useRef(savedSettings);
  const disconnectDialogRef = useRef<HTMLDivElement | null>(null);
  const disconnectCancelRef = useRef<HTMLButtonElement | null>(null);
  const disconnectOpenerRef = useRef<HTMLElement | null>(null);

  const hasUnsavedPlanBuilderChanges =
    settings.defaultReminderTime !== savedSettings.defaultReminderTime || settings.emailHandlingMode !== savedSettings.emailHandlingMode;
  const hasUnsavedEmailSignatureChanges = settings.emailSignatureText !== savedSettings.emailSignatureText;
  const isProviderStatusLoading = (providerLoading.outlook && !outlookConnection) || (providerLoading.gmail && !gmailConnection);
  const providerHelperCopy = getProviderHelperCopy(connectionEmailInput, connectionProviderChoice);

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
      setPlanBuilderSaveError(null);
    }
    if (key === "emailSignatureText") {
      setEmailSignatureSaveMessage(null);
      setEmailSignatureSaveError(null);
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
      try {
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
        setOutlookError(null);
      } catch {
        if (!active) return;
        setOutlookError(PROVIDER_STATUS_ERROR_MESSAGE);
      } finally {
        if (active) setProviderLoading((current) => ({ ...current, outlook: false }));
      }
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
  }, [providerRefreshKey]);

  useEffect(() => {
    let active = true;

    async function refreshGmailConnection() {
      setProviderLoading((current) => ({ ...current, gmail: true }));
      try {
        const connection = await resolveGmailConnectionState();
        if (!active) return;
        setConnectionEmailInput((current) => current || getConnectedGmailMailboxEmail(connection.identity) || "");
        setGmailConnection((current) => (areGmailConnectionStatesEqual(current, connection) ? current : connection));
        setGmailError(null);
      } catch {
        if (!active) return;
        setGmailError(PROVIDER_STATUS_ERROR_MESSAGE);
      } finally {
        if (active) setProviderLoading((current) => ({ ...current, gmail: false }));
      }
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
  }, [providerRefreshKey]);

  useEffect(() => {
    if (!disconnectDialog) return;

    const previousOverflow = document.body.style.overflow;
    document.body.style.overflow = "hidden";

    const focusTimer = window.setTimeout(() => {
      disconnectCancelRef.current?.focus();
    }, 0);

    function handleKeyDown(event: KeyboardEvent) {
      if (!disconnectDialogRef.current) return;

      if (event.key === "Escape") {
        if (!disconnectingProvider) {
          event.preventDefault();
          setDisconnectDialog(null);
        }
        return;
      }

      if (event.key !== "Tab") return;

      const focusable = Array.from(
        disconnectDialogRef.current.querySelectorAll<HTMLElement>(
          'button:not([disabled]), [href], input:not([disabled]), select:not([disabled]), textarea:not([disabled]), [tabindex]:not([tabindex="-1"])'
        )
      );
      if (focusable.length === 0) return;

      const first = focusable[0];
      const last = focusable[focusable.length - 1];
      if (event.shiftKey && document.activeElement === first) {
        event.preventDefault();
        last.focus();
      } else if (!event.shiftKey && document.activeElement === last) {
        event.preventDefault();
        first.focus();
      }
    }

    document.addEventListener("keydown", handleKeyDown);
    return () => {
      window.clearTimeout(focusTimer);
      document.body.style.overflow = previousOverflow;
      document.removeEventListener("keydown", handleKeyDown);
    };
  }, [disconnectDialog, disconnectingProvider]);

  useEffect(() => {
    if (disconnectDialog) return;
    const opener = disconnectOpenerRef.current;
    disconnectOpenerRef.current = null;
    if (!opener) return;
    window.setTimeout(() => {
      if (document.contains(opener)) opener.focus();
    }, 0);
  }, [disconnectDialog]);

  function onSavePlanBuilderSettings() {
    if (savingPlanBuilder || !hasUnsavedPlanBuilderChanges) return;
    setSavingPlanBuilder(true);
    setPlanBuilderSaveMessage(null);
    setPlanBuilderSaveError(null);
    try {
      persistSettings({
        ...savedSettingsRef.current,
        defaultReminderTime: settings.defaultReminderTime,
        emailHandlingMode: settings.emailHandlingMode,
      });
      setPlanBuilderSaveMessage("Defaults saved.");
    } catch {
      setPlanBuilderSaveError("Defaults could not be saved. Try again.");
    } finally {
      setSavingPlanBuilder(false);
    }
  }

  function onSaveEmailSignatureSettings() {
    if (savingEmailSignature || !hasUnsavedEmailSignatureChanges) return;
    setSavingEmailSignature(true);
    setEmailSignatureSaveMessage(null);
    setEmailSignatureSaveError(null);
    const normalizedSignatureText = settings.emailSignatureText;
    try {
      persistSettings({
        ...savedSettingsRef.current,
        emailSignatureEnabled: Boolean(normalizedSignatureText.trim()),
        emailSignatureText: normalizedSignatureText,
      });
      setEmailSignatureSaveMessage("Signature saved.");
    } catch {
      setEmailSignatureSaveError("Signature could not be saved. Try again.");
    } finally {
      setSavingEmailSignature(false);
    }
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
    } catch {
      setOutlookError("Outlook connection could not be updated. Try again.");
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
    } catch {
      setOutlookError("Outlook account could not be disconnected. Try again.");
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
    } catch {
      setGmailError("Google connection could not be updated. Try again.");
      return false;
    } finally {
      setConnectingGmail(false);
    }
  }

  async function onDisconnectGmail() {
    try {
      setGmailError(null);
      await disconnectGmail();
      const connection = getGmailConnectionState();
      setGmailConnection((current) => (areGmailConnectionStatesEqual(current, connection) ? current : connection));
      return true;
    } catch {
      setGmailError("Google account could not be disconnected. Try again.");
      return false;
    }
  }

  async function onSignOut() {
    try {
      setSigningOut(true);
      await signOut();
    } finally {
      setSigningOut(false);
    }
  }

  const connectedAccountEmail = hasMounted ? getConnectedOutlookMailboxEmail(outlookConnection?.identity) || settings.outlookAccountEmail : "";
  const connectedDisplayName = hasMounted ? outlookConnection?.identity?.displayName || "" : "";
  const showMailboxWarning =
    hasMounted && !outlookConnection?.supportedMailbox && Boolean(outlookConnection?.identity);
  const connectedGmailEmail = hasMounted ? getConnectedGmailMailboxEmail(gmailConnection?.identity) : "";
  const connectedGmailDisplayName = hasMounted ? gmailConnection?.identity?.displayName || "" : "";
  const hasConnectedProvider = isConnectedOrAttention(outlookConnection?.status) || isConnectedOrAttention(gmailConnection?.status);
  const hasProviderStatusError =
    Boolean(outlookError === PROVIDER_STATUS_ERROR_MESSAGE || gmailError === PROVIDER_STATUS_ERROR_MESSAGE) && !isProviderStatusLoading;
  const providerInlineError =
    !hasProviderStatusError && (outlookError || gmailError) ? outlookError || gmailError : null;

  function clearConnectionFeedback() {
    setConnectionMessage(null);
    setConnectionError(null);
    setOutlookError(null);
    setGmailError(null);
  }

  function retryConnectionStatus() {
    clearConnectionFeedback();
    setProviderRefreshKey((current) => current + 1);
  }

  async function onConnectProvider(provider: ProviderId) {
    clearConnectionFeedback();
    const succeeded = provider === "gmail" ? await onConnectGmail() : await onConnectOutlook();
    if (succeeded) {
      setConnectionMessage(provider === "gmail" ? "Google account connected." : "Outlook account connected.");
    } else {
      setConnectionError(provider === "gmail" ? "Failed to connect Google account." : "Failed to connect Outlook account.");
    }
  }

  async function onDisconnectProvider(provider: ProviderId) {
    clearConnectionFeedback();
    if (provider === "gmail") {
      const succeeded = await onDisconnectGmail();
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

  function requestDisconnect(provider: ProviderId, label: string) {
    disconnectOpenerRef.current = document.activeElement instanceof HTMLElement ? document.activeElement : null;
    setDisconnectDialog({ provider, label });
  }

  async function runDisconnectConfirmation() {
    if (!disconnectDialog || disconnectingProvider) return;
    setDisconnectingProvider(disconnectDialog.provider);
    try {
      await onDisconnectProvider(disconnectDialog.provider);
      setDisconnectDialog(null);
    } finally {
      setDisconnectingProvider(null);
    }
  }

  function handleWorkflowDefaultsSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    onSavePlanBuilderSettings();
  }

  function handleSignatureSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    onSaveEmailSignatureSettings();
  }

  function stopDialogSubmit(event: ReactKeyboardEvent<HTMLDivElement>) {
    if (event.key === "Enter" && event.target instanceof HTMLTextAreaElement) {
      event.stopPropagation();
    }
  }

  const providerRows = [
    {
      id: "outlook" as const,
      name: "Outlook",
      capability: "Email drafts, scheduled email, and calendar events",
      connection: outlookConnection,
      loading: providerLoading.outlook && !outlookConnection,
      accountEmail: connectedAccountEmail,
      displayName: connectedDisplayName,
      error: outlookError,
      connecting: connectingOutlook,
    },
    {
      id: "gmail" as const,
      name: "Google",
      capability: "Email drafts and Google Calendar events",
      connection: gmailConnection,
      loading: providerLoading.gmail && !gmailConnection,
      accountEmail: connectedGmailEmail,
      displayName: connectedGmailDisplayName,
      error: gmailError,
      connecting: connectingGmail,
    },
  ];

  return (
    <div className="mx-auto w-full max-w-[960px] min-w-0 pb-[44px] pt-[28px] text-slate-900">
      <header className="mb-[20px]">
        <h1 className="text-[30px] font-bold leading-[1.08] text-slate-950 sm:text-[34px]">Settings</h1>
        <p className="mt-[6px] max-w-[680px] text-[15px] leading-[1.45] text-slate-600">
          Manage connected tools and the defaults used when you create a new event plan.
        </p>
      </header>

      <main className="space-y-[16px] min-[960px]:space-y-[18px]">
        <section className="overflow-hidden rounded-[16px] border border-slate-200/80 bg-white/95 shadow-[0_8px_24px_rgba(30,64,100,0.05)]">
          <div className="border-b border-slate-200/70 bg-white px-[16px] py-[16px] sm:px-[20px] sm:py-[18px]">
            <h2 className="text-[20px] font-semibold leading-[1.2] text-slate-950">Connected tools</h2>
            <p className="mt-[3px] text-[14px] leading-[1.4] text-slate-600">
              Connect the email and calendar account used when plans are exported.
            </p>
          </div>
          <div className="px-[16px] py-[16px] sm:px-[20px] sm:py-[18px]">
            {!hasMounted || isProviderStatusLoading ? (
              <div className="space-y-[1px]" aria-label="Loading connection status">
                {[0, 1].map((index) => (
                  <div key={index} className="h-[68px] animate-pulse rounded-[10px] bg-slate-100/80 motion-reduce:animate-none" />
                ))}
              </div>
            ) : (
              <>
                {!hasConnectedProvider ? (
                  <div className="mb-[16px] rounded-[12px] border border-slate-200/80 bg-slate-50/70 px-[14px] py-[13px]">
                    <h3 className="text-[15px] font-semibold text-slate-950">No connected email or calendar account</h3>
                    <p className="mt-[4px] text-[14px] leading-[1.4] text-slate-600">
                      Connect a supported account before exporting reminders, emails, or meetings.
                    </p>
                  </div>
                ) : null}

                {hasProviderStatusError ? (
                  <div className="mb-[16px] flex flex-col gap-[10px] rounded-[12px] border border-amber-200/80 bg-amber-50/70 px-[14px] py-[13px] text-[14px] text-amber-900 sm:flex-row sm:items-center sm:justify-between">
                    <div>
                      <h3 className="text-[15px] font-semibold text-amber-950">Unable to load connection status</h3>
                      <p className="mt-[3px] leading-[1.4]">Connection status could not be loaded. Try again.</p>
                    </div>
                    <button
                      type="button"
                      onClick={retryConnectionStatus}
                      className="inline-flex h-[40px] items-center justify-center rounded-[10px] border border-amber-300 bg-white px-[14px] text-[14px] font-semibold text-amber-900 transition hover:bg-amber-50 focus:outline-none focus:ring-2 focus:ring-amber-500/30"
                    >
                      Retry
                    </button>
                  </div>
                ) : null}

                <div className="mb-[16px] grid gap-[12px] min-[760px]:grid-cols-[minmax(0,1fr)_210px]">
                  <label className="block min-w-0 text-[13px] font-medium leading-[18px] text-slate-600">
                    Account email
                    <input
                      type="email"
                      value={connectionEmailInput}
                      onChange={(event) => setConnectionEmailInput(event.target.value)}
                      placeholder="name@company.com"
                      className="mt-[6px] h-[42px] w-full rounded-[10px] border border-slate-300 bg-white px-[12px] text-[14px] text-slate-900 transition focus:outline-none focus:ring-2 focus:ring-slate-500/25"
                    />
                  </label>
                  <label className="block min-w-0 text-[13px] font-medium leading-[18px] text-slate-600">
                    Provider preference
                    <select
                      value={connectionProviderChoice}
                      onChange={(event) => setConnectionProviderChoice(event.target.value as ConnectionProviderChoice)}
                      className="mt-[6px] h-[42px] w-full rounded-[10px] border border-slate-300 bg-white px-[12px] text-[14px] text-slate-900 transition focus:outline-none focus:ring-2 focus:ring-slate-500/25"
                    >
                      <option value="auto">Auto-detect</option>
                      <option value="outlook">Outlook / Microsoft</option>
                      <option value="gmail">Google</option>
                    </select>
                  </label>
                </div>
                <p className="mb-[14px] text-[13px] leading-[1.35] text-slate-500">{providerHelperCopy}</p>

                <div className="divide-y divide-slate-200/70 border-y border-slate-200/70">
                  {providerRows.map((provider) => {
                    const status = getProviderStatusView(provider.connection?.status, provider.loading);
                    const statusToneClasses = getStatusToneClasses(status.tone);
                    const canDisconnect = isConnectedOrAttention(provider.connection?.status);
                    const connectVerb = getProviderConnectVerb(provider.connection?.status);
                    const isWorking = provider.connecting || disconnectingProvider === provider.id;
                    const connectButtonClass =
                      connectVerb === "Reconnect" && provider.connection?.status !== "reconnect_required"
                        ? providerReconnectActionClass
                        : providerPrimaryActionClass;

                    return (
                      <div key={provider.id} className="flex min-w-0 flex-col gap-[12px] py-[16px] min-[760px]:flex-row min-[760px]:items-center min-[760px]:justify-between">
                        <div className="min-w-0">
                          <div className="flex flex-wrap items-center gap-x-[10px] gap-y-[4px]">
                            <h3 className="text-[16px] font-semibold leading-[1.25] text-slate-950">{provider.name}</h3>
                            <span className={`inline-flex items-center gap-[6px] text-[13px] font-semibold leading-[1.25] ${statusToneClasses.text}`}>
                              <span className={`h-[7px] w-[7px] rounded-full ${statusToneClasses.dot}`} aria-hidden="true" />
                              {status.label}
                            </span>
                          </div>
                          {provider.accountEmail && provider.connection?.status !== "not_connected" ? (
                            <p className="mt-[5px] max-w-full break-words text-[14px] leading-[1.35] text-slate-600 [overflow-wrap:anywhere]">
                              {provider.accountEmail}
                            </p>
                          ) : null}
                          {provider.displayName && provider.displayName !== provider.accountEmail ? (
                            <p className="mt-[3px] max-w-full break-words text-[13px] leading-[1.35] text-slate-500 [overflow-wrap:anywhere]">
                              {provider.displayName}
                            </p>
                          ) : null}
                          <p className="mt-[5px] text-[13px] leading-[1.35] text-slate-500">{provider.capability}</p>
                          {provider.id === "outlook" && showMailboxWarning ? (
                            <p className="mt-[5px] text-[13px] leading-[1.35] text-amber-700">
                              {outlookConnection?.identity?.mailboxEligibilityReason || "This Outlook mailbox needs attention."}
                            </p>
                          ) : null}
                        </div>
                        <div className="flex shrink-0 flex-wrap items-center gap-[8px] min-[760px]:justify-end">
                          <button
                            type="button"
                            onClick={() => void onConnectProvider(provider.id)}
                            disabled={isWorking}
                            className={connectButtonClass}
                          >
                            {provider.connecting ? (connectVerb === "Reconnect" ? "Reconnecting..." : "Connecting...") : `${connectVerb} ${provider.name}`}
                          </button>
                          {canDisconnect ? (
                            <button
                              type="button"
                              onClick={() => requestDisconnect(provider.id, provider.name)}
                              disabled={isWorking}
                              className={providerDisconnectActionClass}
                            >
                              {disconnectingProvider === provider.id ? "Disconnecting..." : "Disconnect"}
                            </button>
                          ) : null}
                        </div>
                      </div>
                    );
                  })}
                </div>

                <div className="mt-[12px] min-h-[18px] text-[13px] leading-[1.35]" aria-live="polite">
                  {connectionMessage ? <p className="font-medium text-green-700">{connectionMessage}</p> : null}
                  {connectionError ? <p className="font-medium text-red-700">{connectionError}</p> : null}
                  {!connectionMessage && !connectionError && providerInlineError ? (
                    <p className="font-medium text-red-700">{providerInlineError}</p>
                  ) : null}
                </div>

                {authEnabled && currentUser ? (
                  <div className="mt-[16px] border-t border-slate-200/70 pt-[14px]">
                    <h3 className="text-[14px] font-semibold text-slate-900">Account</h3>
                    <div className="mt-[7px] flex flex-col gap-[10px] text-[14px] leading-[1.4] text-slate-600 sm:flex-row sm:items-center sm:justify-between">
                      <p className="min-w-0">
                        Signed in as{" "}
                        <span className="break-words font-medium text-slate-900 [overflow-wrap:anywhere]">{currentUser.email || "unknown account"}</span>
                      </p>
                      <button
                        type="button"
                        onClick={() => void onSignOut()}
                        disabled={signingOut}
                        className="inline-flex h-[40px] items-center justify-center rounded-[10px] border border-slate-300 bg-white px-[14px] text-[14px] font-semibold text-slate-700 transition hover:bg-slate-50 focus:outline-none focus:ring-2 focus:ring-slate-500/25 disabled:cursor-not-allowed disabled:border-slate-200 disabled:bg-slate-100 disabled:text-slate-400 disabled:hover:bg-slate-100"
                      >
                        {signingOut ? "Signing out..." : "Sign out"}
                      </button>
                    </div>
                  </div>
                ) : authBypassEnabled ? (
                  <div className="mt-[16px] border-t border-slate-200/70 pt-[14px]">
                    <h3 className="text-[14px] font-semibold text-slate-900">Account</h3>
                    <p className="mt-[5px] text-[14px] leading-[1.4] text-slate-600">Sign-in is bypassed in this environment.</p>
                  </div>
                ) : null}
              </>
            )}
          </div>
        </section>

        <div className="grid gap-[16px] min-[960px]:grid-cols-2 min-[960px]:items-start min-[960px]:gap-[18px]">
          <section className="overflow-hidden rounded-[16px] border border-slate-200/80 bg-white/95 shadow-[0_8px_24px_rgba(30,64,100,0.05)]">
            <div className="border-b border-slate-200/70 bg-white px-[16px] py-[16px] sm:px-[20px] sm:py-[18px]">
              <h2 className="text-[20px] font-semibold leading-[1.2] text-slate-950">Workflow defaults</h2>
              <p className="mt-[3px] text-[14px] leading-[1.4] text-slate-600">
                Set the initial values used when new workflow actions are created.
              </p>
            </div>
            <form onSubmit={handleWorkflowDefaultsSubmit}>
              <div className="grid gap-[12px] px-[16px] py-[16px] sm:px-[20px] sm:py-[18px]">
                <label className="block min-w-0 text-[13px] font-medium leading-[18px] text-slate-600">
                  Default reminder time
                  <input
                    type="time"
                    value={settings.defaultReminderTime}
                    onChange={(event) => updateSettings("defaultReminderTime", event.target.value)}
                    className="mt-[6px] h-[42px] w-full rounded-[10px] border border-slate-300 bg-white px-[12px] text-[14px] text-slate-900 transition focus:outline-none focus:ring-2 focus:ring-slate-500/25"
                  />
                </label>
                <label className="block min-w-0 text-[13px] font-medium leading-[18px] text-slate-600">
                  Email handling mode
                  <select
                    value={settings.emailHandlingMode}
                    onChange={(event) => updateSettings("emailHandlingMode", event.target.value as EmailHandlingMode)}
                    className="mt-[6px] h-[42px] w-full rounded-[10px] border border-slate-300 bg-white px-[12px] text-[14px] text-slate-900 transition focus:outline-none focus:ring-2 focus:ring-slate-500/25"
                  >
                    <option value="draft">Save to Drafts</option>
                    <option value="schedule">Schedule Send (Outlook only; Gmail saves draft)</option>
                    <option value="send">Send Immediately</option>
                  </select>
                </label>
              </div>
              <div className="flex flex-col gap-[10px] border-t border-slate-200/70 px-[16px] py-[14px] sm:flex-row sm:items-center sm:justify-between sm:px-[20px]">
                <div className="min-h-[18px] text-[13px] font-medium leading-[1.35]" aria-live="polite">
                  {planBuilderSaveError ? <p className="text-red-700">{planBuilderSaveError}</p> : null}
                  {!planBuilderSaveError && planBuilderSaveMessage ? <p className="text-green-700">{planBuilderSaveMessage}</p> : null}
                  {!planBuilderSaveError && !planBuilderSaveMessage && hasUnsavedPlanBuilderChanges ? (
                    <p className="text-slate-600">Unsaved changes</p>
                  ) : null}
                </div>
                <button
                  type="submit"
                  disabled={!hasUnsavedPlanBuilderChanges || savingPlanBuilder}
                  className={getSettingsSaveButtonClass(hasUnsavedPlanBuilderChanges && !savingPlanBuilder)}
                  style={getSettingsSaveButtonStyle(hasUnsavedPlanBuilderChanges && !savingPlanBuilder)}
                >
                  {savingPlanBuilder ? "Saving…" : "Save Defaults"}
                </button>
              </div>
            </form>
          </section>

          <section className="overflow-hidden rounded-[16px] border border-slate-200/80 bg-white/95 shadow-[0_8px_24px_rgba(30,64,100,0.05)]">
            <div className="border-b border-slate-200/70 bg-white px-[16px] py-[16px] sm:px-[20px] sm:py-[18px]">
              <h2 className="text-[20px] font-semibold leading-[1.2] text-slate-950">Email signature</h2>
              <p className="mt-[3px] text-[14px] leading-[1.4] text-slate-600">
                Add the signature used when new email actions are created.
              </p>
            </div>
            <form onSubmit={handleSignatureSubmit}>
              <div className="px-[16px] py-[16px] sm:px-[20px] sm:py-[18px]">
                <label className="block min-w-0 text-[13px] font-medium leading-[18px] text-slate-600">
                  Signature
                  <textarea
                    value={settings.emailSignatureText}
                    onChange={(event) => updateSettings("emailSignatureText", event.target.value)}
                    className="mt-[6px] min-h-[150px] w-full resize-y rounded-[10px] border border-slate-300 bg-white px-[13px] py-[12px] text-[14px] leading-[1.45] text-slate-900 transition focus:outline-none focus:ring-2 focus:ring-slate-500/25"
                    placeholder={"Best,\nYour Name"}
                  />
                </label>
              </div>
              <div className="flex flex-col gap-[10px] border-t border-slate-200/70 px-[16px] py-[14px] sm:flex-row sm:items-center sm:justify-between sm:px-[20px]">
                <div className="min-h-[18px] text-[13px] font-medium leading-[1.35]" aria-live="polite">
                  {emailSignatureSaveError ? <p className="text-red-700">{emailSignatureSaveError}</p> : null}
                  {!emailSignatureSaveError && emailSignatureSaveMessage ? <p className="text-green-700">{emailSignatureSaveMessage}</p> : null}
                  {!emailSignatureSaveError && !emailSignatureSaveMessage && hasUnsavedEmailSignatureChanges ? (
                    <p className="text-slate-600">Unsaved changes</p>
                  ) : null}
                </div>
                <button
                  type="submit"
                  disabled={!hasUnsavedEmailSignatureChanges || savingEmailSignature}
                  className={getSettingsSaveButtonClass(hasUnsavedEmailSignatureChanges && !savingEmailSignature)}
                  style={getSettingsSaveButtonStyle(hasUnsavedEmailSignatureChanges && !savingEmailSignature)}
                >
                  {savingEmailSignature ? "Saving…" : "Save"}
                </button>
              </div>
            </form>
          </section>
        </div>
      </main>

      {hasMounted && disconnectDialog
        ? createPortal(
            <div
              className="fixed inset-0 z-[220] flex items-end justify-center bg-[rgba(15,23,42,0.14)] p-[16px] sm:items-center"
              onMouseDown={(event) => {
                if (event.target === event.currentTarget && !disconnectingProvider) {
                  setDisconnectDialog(null);
                }
              }}
            >
              <div
                ref={disconnectDialogRef}
                role="dialog"
                aria-modal="true"
                aria-labelledby="settings-disconnect-title"
                aria-describedby="settings-disconnect-description"
                onKeyDown={stopDialogSubmit}
                className="max-h-[88dvh] w-full max-w-[500px] overflow-auto rounded-t-[16px] border border-slate-200 bg-white shadow-[0_22px_70px_rgba(15,23,42,0.18)] sm:rounded-[16px]"
              >
                <div className="px-[20px] pb-[16px] pt-[20px]">
                  <h2 id="settings-disconnect-title" className="text-[18px] font-semibold leading-[1.25] text-slate-950">
                    Disconnect {disconnectDialog.label}?
                  </h2>
                  <p id="settings-disconnect-description" className="mt-[8px] text-[14px] leading-[1.45] text-slate-600">
                    This removes the connected account from this app. Existing items already created in the provider are not deleted.
                  </p>
                </div>
                <div className="flex flex-col gap-[8px] border-t border-slate-200/70 px-[20px] pb-[calc(16px+env(safe-area-inset-bottom))] pt-[14px] sm:flex-row sm:justify-end sm:pb-[16px]">
                  <button
                    ref={disconnectCancelRef}
                    type="button"
                    onClick={() => {
                      if (!disconnectingProvider) setDisconnectDialog(null);
                    }}
                    disabled={Boolean(disconnectingProvider)}
                    className="inline-flex h-[40px] items-center justify-center rounded-[10px] border border-slate-300 bg-white px-[16px] text-[14px] font-semibold text-slate-700 transition hover:bg-slate-50 focus:outline-none focus:ring-2 focus:ring-slate-500/25 disabled:cursor-not-allowed disabled:border-slate-200 disabled:bg-slate-100 disabled:text-slate-400 disabled:hover:bg-slate-100"
                  >
                    Keep connected
                  </button>
                  <button
                    type="button"
                    onClick={() => void runDisconnectConfirmation()}
                    disabled={Boolean(disconnectingProvider)}
                    className="inline-flex h-[40px] items-center justify-center rounded-[10px] bg-red-700 px-[16px] text-[14px] font-semibold text-white transition hover:bg-red-800 focus:outline-none focus:ring-2 focus:ring-red-500/30 disabled:cursor-not-allowed disabled:bg-red-300"
                  >
                    {disconnectingProvider ? "Disconnecting..." : "Disconnect"}
                  </button>
                </div>
              </div>
            </div>,
            document.body
          )
        : null}
    </div>
  );
}
