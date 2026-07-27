"use client";

import Link from "next/link";
import { useCallback, useEffect, useMemo, useRef, useState, useSyncExternalStore } from "react";

import {
  APP_SETTINGS_UPDATED_EVENT,
  hydrateAppSettingsFromSupabase,
  loadAppSettings,
} from "../../lib/appSettings";
import {
  EXECUTION_HISTORY_UPDATED_EVENT,
  listCachedExecutionHistory,
  listExecutionHistory,
  readCachedExecutionHistorySnapshot,
  type ExecutionHistoryRecord,
} from "../../lib/executionHistory";
import {
  GMAIL_CONNECTION_UPDATED_EVENT,
  getConnectedGmailMailboxEmail,
  getGmailConnectionState,
  resolveGmailConnectionState,
  type GmailConnectionState,
} from "../../lib/gmailClient";
import {
  OUTLOOK_CONNECTION_UPDATED_EVENT,
  getConnectedOutlookMailboxEmail,
  getOutlookConnectionState,
  resolveOutlookConnectionState,
  type OutlookConnectionState,
} from "../../lib/outlookClient";
import {
  loadCachedTemplateState,
  loadTemplateStateFromSupabase,
  readCachedTemplateStateSnapshot,
  type PersistedPlanTemplate,
} from "../../lib/templateStore";
import { getCachedOrgContext } from "../../lib/orgBootstrap";
import { getLocalUserKey } from "../../lib/userKey";
import {
  buildDashboardActivityItems,
  buildDashboardMetrics,
  buildDashboardUpcomingItems,
  buildDashboardWorkflowSummaries,
  countDashboardAttentionRecords,
  groupDashboardUpcomingItems,
  type DashboardActionType,
  type DashboardActivityItem,
  type DashboardMetrics,
  type DashboardUpcomingGroup,
  type DashboardUpcomingItem,
  type DashboardWorkflowSummary,
} from "../../lib/dashboardData";
import { useAuthContext } from "./auth-provider";

type LoadStatus = "loading" | "ready" | "error";

type SourceState<T> = {
  status: LoadStatus;
  data: T;
  hasSnapshot: boolean;
  isRevalidating: boolean;
  stale: boolean;
  errorKind: string | null;
  updatedAt: number | null;
};

type ProviderConnectionSummary = {
  id: "outlook" | "gmail";
  name: "Outlook" | "Google";
  status: "connected" | "reconnect_required" | "not_connected";
  account: string | null;
  displayName: string | null;
};

type ProviderState = SourceState<{
  connections: ProviderConnectionSummary[];
  partialError: boolean;
}>;

type DashboardSourceName = "history" | "workflows" | "providers";

type DashboardLoadOptions = {
  preserveData?: boolean;
  manual?: boolean;
};

type DashboardCachedSource<T> = {
  data: T;
  updatedAt: number;
};

type DashboardCachedView = {
  history?: DashboardCachedSource<ExecutionHistoryRecord[]>;
  workflows?: DashboardCachedSource<DashboardWorkflowSummary[]>;
  provider?: DashboardCachedSource<ProviderState["data"]>;
};

const DASHBOARD_SOURCE_TIMEOUT_MS = 11000;
const dashboardViewCache = new Map<string, DashboardCachedView>();

class DashboardSourceLoadError extends Error {
  sourceName: DashboardSourceName;

  constructor(sourceName: DashboardSourceName, reason: "failed" | "timeout") {
    super(`Dashboard ${sourceName} source ${reason}.`);
    this.name = "DashboardSourceLoadError";
    this.sourceName = sourceName;
  }
}

function withDashboardSourceTimeout<T>(
  sourceName: DashboardSourceName,
  loader: () => Promise<T> | T,
  timeoutMs = DASHBOARD_SOURCE_TIMEOUT_MS
) {
  let timeoutId: ReturnType<typeof setTimeout> | null = null;
  const loadPromise = Promise.resolve().then(loader);
  const timeoutPromise = new Promise<never>((_, reject) => {
    timeoutId = setTimeout(() => {
      reject(new DashboardSourceLoadError(sourceName, "timeout"));
    }, timeoutMs);
  });

  return Promise.race([loadPromise, timeoutPromise])
    .catch((error) => {
      if (error instanceof DashboardSourceLoadError) {
        throw error;
      }
      throw new DashboardSourceLoadError(sourceName, "failed");
    })
    .finally(() => {
      if (timeoutId) {
        clearTimeout(timeoutId);
      }
    });
}

function createSourceState<T>(
  data: T,
  status: LoadStatus,
  overrides: Partial<Omit<SourceState<T>, "data" | "status">> = {}
): SourceState<T> {
  return {
    status,
    data,
    hasSnapshot: false,
    isRevalidating: false,
    stale: false,
    errorKind: null,
    updatedAt: null,
    ...overrides,
  };
}

function markSourceRevalidating<T>(current: SourceState<T>) {
  if (current.status === "ready" || current.hasSnapshot) {
    return {
      ...current,
      status: "ready" as const,
      isRevalidating: true,
      errorKind: null,
    };
  }

  return {
    ...current,
    status: "loading" as const,
    isRevalidating: true,
    stale: false,
    errorKind: null,
  };
}

function resolveSourceSuccess<T>(data: T, hasSnapshot: boolean): SourceState<T> {
  return createSourceState(data, "ready", {
    hasSnapshot,
    isRevalidating: false,
    stale: false,
    updatedAt: Date.now(),
  });
}

function resolveCachedSource<T>(source: DashboardCachedSource<T>): SourceState<T> {
  return createSourceState(source.data, "ready", {
    hasSnapshot: true,
    updatedAt: source.updatedAt,
  });
}

function resolveSourceFailure<T>(current: SourceState<T>, fallbackData: T, errorKind: string): SourceState<T> {
  if (current.status === "ready" || current.hasSnapshot) {
    return {
      ...current,
      status: "ready",
      hasSnapshot: true,
      isRevalidating: false,
      stale: true,
      errorKind,
    };
  }

  return createSourceState(fallbackData, "error", {
    hasSnapshot: false,
    isRevalidating: false,
    stale: false,
    errorKind,
  });
}

function cacheDashboardSource<T>(scopeKey: string, sourceName: keyof DashboardCachedView, data: T) {
  const current = dashboardViewCache.get(scopeKey) ?? {};
  dashboardViewCache.set(scopeKey, {
    ...current,
    [sourceName]: {
      data,
      updatedAt: Date.now(),
    },
  });
}

const panelClassName =
  "overflow-hidden rounded-[16px] border border-slate-200/80 bg-white/95 shadow-[0_10px_30px_rgba(30,64,100,0.06)]";
const panelHeaderClassName = "border-b border-slate-200/70 px-[18px] py-[16px]";
const panelBodyClassName = "px-[16px] py-[16px] sm:px-[18px]";
const panelTitleClassName = "text-[20px] font-semibold leading-[23px] text-slate-950";
const panelHelperClassName = "mt-[4px] text-[14px] leading-[19px] text-slate-600";
const metadataClassName = "text-[12px] leading-[16px] text-slate-500";
const secondaryButtonClassName =
  "inline-flex h-[38px] items-center justify-center rounded-[10px] border border-slate-200 bg-white px-[14px] text-[14px] font-semibold text-slate-700 transition hover:bg-slate-50";
const focusRingClassName =
  "focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#6f9fd1]/35 focus-visible:ring-offset-2 focus-visible:ring-offset-[var(--app-bg)]";

const quickActions = [
  { label: "Start a new event", href: "/plans?new=1", primary: true },
  { label: "Open Plans", href: "/plans", primary: false },
  { label: "View History", href: "/history", primary: false },
  { label: "Manage settings", href: "/settings", primary: false },
] as const;

function getHeaderDateLabel(value: Date | null) {
  if (!value) return "";
  return new Intl.DateTimeFormat(undefined, {
    weekday: "long",
    month: "long",
    day: "numeric",
  }).format(value);
}

function formatTime(value: string) {
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return "";
  return new Intl.DateTimeFormat(undefined, {
    hour: "numeric",
    minute: "2-digit",
  }).format(parsed);
}

function formatShortDateTime(value: string) {
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return "";
  return new Intl.DateTimeFormat(undefined, {
    month: "short",
    day: "numeric",
    hour: "numeric",
    minute: "2-digit",
  }).format(parsed);
}

function formatActivityDateTime(value: string) {
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return "";
  return new Intl.DateTimeFormat(undefined, {
    month: "short",
    day: "numeric",
    hour: "numeric",
    minute: "2-digit",
  }).format(parsed);
}

function getActionAccent(type: DashboardActionType | "system") {
  if (type === "reminder") {
    return {
      rail: "bg-blue-500",
      text: "text-blue-700",
      dot: "bg-blue-500",
      label: "Reminder",
    };
  }
  if (type === "meeting") {
    return {
      rail: "bg-violet-500",
      text: "text-violet-700",
      dot: "bg-violet-500",
      label: "Meeting",
    };
  }
  if (type === "email") {
    return {
      rail: "bg-green-500",
      text: "text-green-700",
      dot: "bg-green-500",
      label: "Email",
    };
  }
  return {
    rail: "bg-slate-400",
    text: "text-slate-600",
    dot: "bg-slate-400",
    label: "System",
  };
}

function SkeletonLine({ className = "" }: { className?: string }) {
  return (
    <div
      className={`animate-pulse rounded-[6px] bg-slate-200/75 motion-reduce:animate-none ${className}`}
      aria-hidden="true"
    />
  );
}

function SectionError({ onRetry }: { onRetry: () => void }) {
  return (
    <div className="rounded-[10px] border border-amber-200/80 bg-amber-50/70 px-[16px] py-[16px] text-[14px] leading-[20px] text-slate-700">
      <p className="font-semibold text-slate-900">Unable to load this section.</p>
      <button
        type="button"
        onClick={onRetry}
        className={`mt-[12px] ${secondaryButtonClassName} ${focusRingClassName}`}
      >
        Retry
      </button>
    </div>
  );
}

function StaleNotice({ onRetry, className = "" }: { onRetry: () => void; className?: string }) {
  return (
    <div className={`flex flex-wrap items-center gap-x-[10px] gap-y-[4px] text-[12px] font-medium leading-[16px] text-amber-700 ${className}`}>
      <span>{"Couldn't refresh. Showing saved data."}</span>
      <button type="button" onClick={onRetry} className={`font-semibold text-amber-800 hover:text-amber-950 ${focusRingClassName}`}>
        Retry
      </button>
    </div>
  );
}

function PanelHeading({
  title,
  helper,
  action,
}: {
  title: string;
  helper: string;
  action?: React.ReactNode;
}) {
  return (
    <div className={panelHeaderClassName}>
      <div className="flex min-w-0 items-start justify-between gap-[16px]">
        <div className="min-w-0">
          <h2 className={panelTitleClassName}>{title}</h2>
          <p className={panelHelperClassName}>{helper}</p>
        </div>
        {action ? <div className="shrink-0">{action}</div> : null}
      </div>
    </div>
  );
}

function MetricCell({
  label,
  helper,
  value,
  loading,
  error,
  index,
}: {
  label: string;
  helper: string;
  value: number | null;
  loading: boolean;
  error: boolean;
  index: number;
}) {
  const dividerClassName = [
    index % 2 === 0 ? "border-l-0" : "border-l max-[339px]:border-l-0",
    index < 2 ? "border-t-0 max-[339px]:border-t" : "border-t",
    index === 0 ? "max-[339px]:border-t-0 min-[900px]:border-l-0" : "min-[900px]:border-l",
    "min-[900px]:border-t-0",
  ].join(" ");

  return (
    <div className={`flex min-h-[88px] min-w-0 flex-col justify-center border-slate-200/70 px-[16px] py-[15px] min-[900px]:min-h-[94px] ${dividerClassName}`}>
      {loading ? (
        <SkeletonLine className="h-[28px] w-[64px]" />
      ) : (
        <div className="text-[30px] font-bold leading-none text-slate-950 [font-variant-numeric:tabular-nums] min-[900px]:text-[32px]">
          {error ? "—" : value}
        </div>
      )}
      <div className="mt-[10px] text-[14px] font-semibold leading-[18px] text-slate-800">{label}</div>
      <div className="mt-[3px] text-[12px] leading-[15px] text-slate-500">{error ? "Unable to load" : helper}</div>
    </div>
  );
}

function SummaryStrip({
  metrics,
  historyStatus,
  workflowStatus,
}: {
  metrics: DashboardMetrics | null;
  historyStatus: LoadStatus;
  workflowStatus: LoadStatus;
}) {
  const cells = [
    {
      label: "Today",
      helper: "Scheduled today",
      value: metrics?.today ?? null,
      status: historyStatus,
    },
    {
      label: "Next 7 days",
      helper: "Coming after today",
      value: metrics?.nextSevenDays ?? null,
      status: historyStatus,
    },
    {
      label: "Saved workflows",
      helper: "Reusable templates",
      value: metrics?.savedWorkflows ?? null,
      status: workflowStatus,
    },
    {
      label: "Recent activity",
      helper: "Last 30 days",
      value: metrics?.recentActivity ?? null,
      status: historyStatus,
    },
  ];

  return (
    <section className={panelClassName}>
      <div className="grid grid-cols-2 max-[339px]:grid-cols-1 min-[900px]:grid-cols-4">
        {cells.map((cell, index) => (
          <MetricCell
            key={cell.label}
            label={cell.label}
            helper={cell.helper}
            value={cell.value}
            loading={cell.status === "loading"}
            error={cell.status === "error"}
            index={index}
          />
        ))}
      </div>
    </section>
  );
}

function AttentionBanner({ count }: { count: number }) {
  if (count <= 0) return null;

  return (
    <section className="rounded-[14px] border border-amber-200/80 bg-amber-50 px-[16px] py-[13px] text-amber-950 shadow-[0_8px_22px_rgba(146,64,14,0.05)] sm:px-[18px]">
      <div className="flex flex-col gap-[12px] sm:flex-row sm:items-center sm:justify-between">
        <div className="min-w-0">
          <h2 className="text-[15px] font-semibold leading-[20px]">Needs attention</h2>
          <p className="mt-[3px] text-[13px] leading-[18px]">
            {count} recent workflow {count === 1 ? "action" : "actions"} reported an error or incomplete export.
          </p>
        </div>
        <Link
          href="/history"
          className={`inline-flex h-[38px] w-fit items-center justify-center whitespace-nowrap rounded-[10px] border border-amber-300/70 bg-white/80 px-[14px] text-[13px] font-semibold text-amber-950 transition hover:bg-white ${focusRingClassName}`}
        >
          View History
        </Link>
      </div>
    </section>
  );
}

function UpcomingSkeleton() {
  return (
    <div className="divide-y divide-slate-200/70" aria-hidden="true">
      {[0, 1, 2, 3].map((item) => (
        <div key={item} className="min-h-[62px] px-[16px] py-[12px]">
          <SkeletonLine className="h-[12px] w-[80px]" />
          <SkeletonLine className="mt-[8px] h-[15px] w-3/4" />
          <SkeletonLine className="mt-[7px] h-[12px] w-1/2" />
        </div>
      ))}
    </div>
  );
}

function EmptyState({
  title,
  body,
  actionLabel,
  href,
  compact = false,
}: {
  title: string;
  body: string;
  actionLabel: string;
  href: string;
  compact?: boolean;
}) {
  return (
    <div className={`flex items-center justify-center px-[24px] py-[24px] text-center ${compact ? "min-h-[150px] md:min-h-[160px]" : "min-h-[160px] md:min-h-[170px]"}`}>
      <div className="max-w-[360px]">
        <h3 className="text-[17px] font-semibold leading-[22px] text-slate-950">{title}</h3>
        <p className="mt-[8px] text-[14px] leading-[20px] text-slate-600">{body}</p>
        <Link
          href={href}
          className={`mt-[14px] ${secondaryButtonClassName} ${focusRingClassName}`}
        >
          {actionLabel}
        </Link>
      </div>
    </div>
  );
}

function hasSuccessfulExportRecord(record: ExecutionHistoryRecord) {
  return record.status === "success" || record.status === "fallback";
}

type FirstRunSetupStep = {
  label: string;
  description: string;
  complete: boolean;
};

type FirstRunProgress = {
  providerConnected: boolean;
  workflowCount: number;
  exportComplete: boolean;
};

function FirstRunStepStatus({ complete }: { complete: boolean }) {
  return (
    <span
      className={`inline-flex shrink-0 items-center justify-self-end self-start gap-[6px] whitespace-nowrap rounded-full px-[8px] py-[3px] text-[12px] font-semibold leading-[16px] md:py-[2px] md:justify-self-auto md:self-auto ${
        complete ? "bg-green-50 text-green-700" : "bg-slate-100 text-slate-600"
      }`}
    >
      {complete ? <span className="h-[6px] w-[6px] rounded-full bg-green-500" aria-hidden="true" /> : null}
      {complete ? "Complete" : "Not started"}
    </span>
  );
}

function FirstRunEmptyState({
  providerConnected,
  workflowCount,
  exportComplete,
}: {
  providerConnected: boolean;
  workflowCount: number;
  exportComplete: boolean;
}) {
  const steps: FirstRunSetupStep[] = [
    {
      label: "Connect tools",
      description: "Connect an email or calendar account.",
      complete: providerConnected,
    },
    {
      label: "Create a workflow",
      description: "Save a reusable workflow in Plans.",
      complete: workflowCount > 0,
    },
    {
      label: "Export a plan",
      description: "Create at least one workflow action in a connected tool.",
      complete: exportComplete,
    },
  ];

  return (
    <div className="flex items-center justify-center px-[18px] py-[16px] md:py-[12px]">
      <div className="w-full max-w-[720px] text-center">
        <h3 className="text-[17px] font-semibold leading-[22px] text-slate-950">Set up your workspace</h3>
        <p className="mx-auto mt-[7px] max-w-[680px] text-[14px] leading-[20px] text-slate-600">
          Connect an account, create a reusable workflow, and export a plan to begin tracking upcoming activity.
        </p>
        <ol className="mt-[13px] overflow-hidden rounded-[12px] border border-slate-200/80 bg-white text-left md:mt-[10px] md:grid md:grid-cols-3">
          {steps.map((step, index) => (
            <li
              key={step.label}
              className="grid min-h-[88px] min-w-0 grid-cols-[28px_minmax(0,1fr)_auto] gap-x-[10px] border-t border-slate-200/70 px-[14px] py-[14px] first:border-t-0 md:block md:min-h-[96px] md:border-l md:border-t-0 md:px-[12px] md:py-[10px] md:first:border-l-0"
            >
              <div className="contents md:flex md:items-center md:justify-between md:gap-[12px]">
                <span
                  className={`col-start-1 row-start-1 flex h-[28px] w-[28px] items-center justify-center rounded-full text-[13px] font-semibold ${
                    step.complete ? "bg-green-50 text-green-700 ring-1 ring-green-200/80" : "bg-[#edf4fb] text-[#315f92] ring-1 ring-[#cbdced]"
                  }`}
                  aria-label={`Step ${index + 1} of ${steps.length}`}
                >
                  {index + 1}
                </span>
                <FirstRunStepStatus complete={step.complete} />
              </div>
              <div className="col-start-2 col-end-3 row-start-1 min-w-0 md:mt-[7px]">
                <div className="text-[15px] font-semibold leading-[20px] text-slate-900 md:leading-[18px]">{step.label}</div>
                <p className="mt-[4px] text-[13px] leading-[18px] text-slate-500 md:mt-[2px] md:text-[12px] md:leading-[15px]">{step.description}</p>
              </div>
            </li>
          ))}
        </ol>
        <div className="mt-[14px] grid gap-[10px] sm:inline-flex md:mt-[10px]">
          <Link
            href="/plans?new=1"
            className={`inline-flex h-[40px] min-w-[150px] items-center justify-center whitespace-nowrap rounded-[10px] border border-[#315f92] bg-[#315f92] px-[16px] text-[14px] font-semibold text-white transition hover:bg-[#28527d] ${focusRingClassName}`}
          >
            Create a workflow
          </Link>
          <Link
            href="/settings"
            className={`inline-flex h-[40px] min-w-[130px] items-center justify-center whitespace-nowrap rounded-[10px] border border-slate-200 bg-white px-[16px] text-[14px] font-semibold text-slate-700 transition hover:bg-slate-50 ${focusRingClassName}`}
          >
            Connect tools
          </Link>
        </div>
      </div>
    </div>
  );
}

function UpcomingRow({ item, group }: { item: DashboardUpcomingItem; group: DashboardUpcomingGroup }) {
  const accent = getActionAccent(item.type);
  const isNearTerm = group.key === "today" || group.key === "tomorrow";
  const timeLabel = isNearTerm ? formatTime(item.scheduledAt) : formatShortDateTime(item.scheduledAt);
  const metadata = [item.providerLabel, item.statusLabel].filter((entry): entry is string => Boolean(entry));

  return (
    <li className="relative min-h-[66px] min-w-0 border-t border-slate-200/70 first:border-t-0">
      <div className={`absolute inset-y-0 left-0 w-[4px] ${accent.rail}`} aria-hidden="true" />
      <div className="grid min-w-0 gap-[8px] px-[16px] py-[12px] sm:grid-cols-[minmax(0,1fr)_auto] sm:items-start">
        <div className="min-w-0">
          <div className={`text-[12px] font-semibold leading-[16px] ${accent.text}`}>{accent.label}</div>
          <div className="mt-[4px] text-[15px] font-semibold leading-[20px] text-slate-900">{item.title}</div>
          {item.sourceName ? (
            <div className="mt-[3px] text-[13px] leading-[18px] text-slate-500">{item.sourceName}</div>
          ) : null}
        </div>
        <div className="min-w-0 text-left sm:text-right">
          <div className="text-[14px] font-semibold leading-[20px] text-slate-700 [font-variant-numeric:tabular-nums]">{timeLabel}</div>
          {metadata.length > 0 ? (
            <div className="mt-[3px] flex flex-wrap gap-x-[8px] gap-y-[2px] text-[12px] font-medium leading-[16px] text-slate-500 sm:justify-end">
              {metadata.map((entry) => (
                <span key={entry}>{entry}</span>
              ))}
            </div>
          ) : null}
        </div>
      </div>
    </li>
  );
}

function UpcomingPanel({
  status,
  stale,
  groups,
  totalCount,
  showAll,
  onToggleShowAll,
  onRetry,
  firstRun,
  firstRunProgress,
}: {
  status: LoadStatus;
  stale: boolean;
  groups: DashboardUpcomingGroup[];
  totalCount: number;
  showAll: boolean;
  onToggleShowAll: () => void;
  onRetry: () => void;
  firstRun: boolean;
  firstRunProgress: FirstRunProgress;
}) {
  return (
    <section className={panelClassName} aria-busy={status === "loading"}>
      <PanelHeading
        title="Upcoming"
        helper="Scheduled actions from exported workflows."
        action={
          totalCount > 8 ? (
            <button
              type="button"
              onClick={onToggleShowAll}
              className={`text-[13px] font-semibold leading-[18px] text-[#315f92] transition hover:text-[#244b76] ${focusRingClassName}`}
            >
              {showAll ? "Show less" : "Show all"}
            </button>
          ) : null
        }
      />
      <div>
        {status === "ready" && stale ? (
          <div className="border-b border-slate-200/70 px-[16px] py-[10px]">
            <StaleNotice onRetry={onRetry} />
          </div>
        ) : null}
        {status === "loading" ? <UpcomingSkeleton /> : null}
        {status === "error" ? (
          <div className={panelBodyClassName}>
            <SectionError onRetry={onRetry} />
          </div>
        ) : null}
        {status === "ready" && totalCount === 0 && firstRun ? (
          <FirstRunEmptyState
            providerConnected={firstRunProgress.providerConnected}
            workflowCount={firstRunProgress.workflowCount}
            exportComplete={firstRunProgress.exportComplete}
          />
        ) : null}
        {status === "ready" && totalCount === 0 && !firstRun ? (
          <EmptyState
            title="Nothing scheduled yet"
            body="Export a plan to see upcoming reminders, meetings, and emails here."
            actionLabel="Go to Plans"
            href="/plans"
          />
        ) : null}
        {status === "ready" && totalCount > 0 ? (
          <div>
            {groups.map((group, index) => (
              <div key={group.key}>
                <div className={`${index === 0 ? "" : "border-t border-slate-200/70"} bg-slate-50/80 px-[16px] py-[10px] text-[13px] font-semibold leading-[18px] text-slate-600`}>
                  {group.label}
                </div>
                <ul className="min-w-0">
                  {group.items.map((item) => (
                    <UpcomingRow key={item.id} item={item} group={group} />
                  ))}
                </ul>
              </div>
            ))}
          </div>
        ) : null}
      </div>
    </section>
  );
}

function QuickActionsSection({ className = "" }: { className?: string }) {
  return (
    <section className={`min-w-0 max-w-full px-[16px] py-[16px] ${className}`}>
      <h3 className="text-[16px] font-semibold leading-[21px] text-slate-950">Quick actions</h3>
      <div className="mt-[10px] divide-y divide-slate-200/70">
        {quickActions.map((action) => (
          <Link
            key={action.href}
            href={action.href}
            className={`flex min-h-[42px] w-full min-w-0 items-center justify-between gap-[10px] px-[14px] text-[14px] font-semibold leading-[19px] transition ${
              action.primary
                ? "bg-[#eef5fb] text-[#244b76] hover:bg-[#e5eff8]"
                : "text-slate-800 hover:bg-slate-50"
            } ${focusRingClassName}`}
          >
            <span className="min-w-0 whitespace-nowrap">{action.label}</span>
            <span aria-hidden="true" className="shrink-0 text-slate-400">
              →
            </span>
          </Link>
        ))}
      </div>
    </section>
  );
}

function WorkspaceSectionSkeleton({ rows = 3 }: { rows?: number }) {
  return (
    <div className="mt-[12px] space-y-[10px]" aria-hidden="true">
      {Array.from({ length: rows }).map((_, index) => (
        <SkeletonLine key={index} className="h-[15px] w-full" />
      ))}
    </div>
  );
}

function SavedWorkflowsSection({
  status,
  stale,
  workflows,
  onRetry,
  className = "",
}: {
  status: LoadStatus;
  stale: boolean;
  workflows: DashboardWorkflowSummary[];
  onRetry: () => void;
  className?: string;
}) {
  return (
    <section className={`min-w-0 max-w-full px-[16px] py-[16px] ${className}`} aria-busy={status === "loading"}>
      <h3 className="text-[16px] font-semibold leading-[21px] text-slate-950">Saved workflows</h3>
      <p className="mt-[4px] text-[13px] leading-[18px] text-slate-500">Reusable templates available in Plans.</p>
      {status === "ready" && stale ? <StaleNotice onRetry={onRetry} className="mt-[8px]" /> : null}
      {status === "loading" ? <WorkspaceSectionSkeleton /> : null}
      {status === "error" ? (
        <div className="mt-[12px]">
          <SectionError onRetry={onRetry} />
        </div>
      ) : null}
      {status === "ready" && workflows.length === 0 ? (
        <div className="mt-[12px]">
          <p className="text-[14px] font-medium leading-[20px] text-slate-700">No saved workflows yet.</p>
          <Link href="/plans" className={`mt-[10px] whitespace-nowrap ${secondaryButtonClassName} ${focusRingClassName}`}>
            Create one in Plans
          </Link>
        </div>
      ) : null}
      {status === "ready" && workflows.length > 0 ? (
        <>
          <ul className="mt-[10px] divide-y divide-slate-200/70">
            {workflows.slice(0, 5).map((workflow) => (
              <li
                key={workflow.id}
                title={workflow.name}
                className="flex min-h-[42px] min-w-0 items-center text-[14px] font-semibold leading-[19px] text-slate-800"
              >
                <span className="block min-w-0 max-w-full overflow-hidden text-ellipsis whitespace-nowrap">{workflow.name}</span>
              </li>
            ))}
          </ul>
          <div className="mt-[10px] border-t border-slate-200/70 pt-[10px]">
            <Link
              href="/plans"
              className={`whitespace-nowrap text-[13px] font-semibold leading-[18px] text-[#315f92] hover:text-[#244b76] ${focusRingClassName}`}
            >
              Manage workflows
            </Link>
          </div>
        </>
      ) : null}
    </section>
  );
}

function ConnectedToolsSection({
  state,
  onRetry,
  className = "",
}: {
  state: ProviderState;
  onRetry: () => void;
  className?: string;
}) {
  const actionableConnections = state.data.connections.filter((connection) => connection.status !== "not_connected");

  return (
    <section className={`min-w-0 max-w-full px-[16px] py-[16px] ${className}`} aria-busy={state.status === "loading"}>
      <h3 className="text-[16px] font-semibold leading-[21px] text-slate-950">Connected tools</h3>
      <p className="mt-[4px] text-[13px] leading-[18px] text-slate-500">Email and calendar connection status.</p>
      {state.status === "ready" && state.stale ? <StaleNotice onRetry={onRetry} className="mt-[8px]" /> : null}
      {state.status === "loading" ? <WorkspaceSectionSkeleton rows={2} /> : null}
      {state.status === "error" ? (
        <div className="mt-[12px]">
          <p className="text-[14px] font-semibold leading-[20px] text-slate-900">Connection status is unavailable.</p>
          <div className="mt-[10px] flex flex-col gap-[8px] min-[420px]:flex-row min-[420px]:flex-wrap">
            <Link href="/settings" className={`whitespace-nowrap ${secondaryButtonClassName} ${focusRingClassName}`}>
              Open Settings
            </Link>
            <button type="button" onClick={onRetry} className={`whitespace-nowrap ${secondaryButtonClassName} ${focusRingClassName}`}>
              Retry
            </button>
          </div>
        </div>
      ) : null}
      {state.status === "ready" && actionableConnections.length === 0 ? (
        <div className="mt-[12px]">
          <p className="text-[14px] font-medium leading-[20px] text-slate-700">No connected email or calendar account.</p>
          <Link href="/settings" className={`mt-[10px] whitespace-nowrap ${secondaryButtonClassName} ${focusRingClassName}`}>
            Connect in Settings
          </Link>
        </div>
      ) : null}
      {state.status === "ready" && actionableConnections.length > 0 ? (
        <div className="mt-[10px] divide-y divide-slate-200/70">
          {actionableConnections.map((connection) => {
            const connected = connection.status === "connected";
            return (
              <div key={connection.id} className="min-h-[42px] min-w-0 max-w-full py-[8px]">
                <div className="flex min-w-0 items-center justify-between gap-[10px]">
                  <div className="min-w-0 text-[14px] font-semibold leading-[19px] text-slate-900">{connection.name}</div>
                  <div
                    className={`inline-flex shrink-0 items-center gap-[6px] whitespace-nowrap text-[12px] font-semibold leading-[16px] ${
                      connected ? "text-green-700" : "text-amber-700"
                    }`}
                  >
                    <span className={`h-[8px] w-[8px] rounded-full ${connected ? "bg-green-500" : "bg-amber-500"}`} aria-hidden="true" />
                    {connected ? "Connected" : "Reconnect required"}
                  </div>
                </div>
                {connection.account || connection.displayName ? (
                  <div className={`${metadataClassName} mt-[4px] min-w-0 max-w-full [overflow-wrap:anywhere]`}>
                    {connection.displayName && connection.displayName !== connection.account ? `${connection.displayName} · ` : ""}
                    {connection.account}
                  </div>
                ) : null}
              </div>
            );
          })}
          {state.data.partialError ? (
            <p className={`${metadataClassName} pt-[8px]`}>One provider status is unavailable.</p>
          ) : null}
        </div>
      ) : null}
    </section>
  );
}

function WorkspacePanel({
  workflowStatus,
  workflowStale,
  workflows,
  providerState,
  onRetry,
}: {
  workflowStatus: LoadStatus;
  workflowStale: boolean;
  workflows: DashboardWorkflowSummary[];
  providerState: ProviderState;
  onRetry: () => void;
}) {
  return (
    <section className={panelClassName}>
      <PanelHeading title="Workspace" helper="Shortcuts, reusable workflows, and connection status." />
      <div className="grid grid-cols-1 min-[640px]:grid-cols-2 min-[960px]:grid-cols-3 min-[1280px]:grid-cols-1">
        <QuickActionsSection />
        <SavedWorkflowsSection
          status={workflowStatus}
          stale={workflowStale}
          workflows={workflows}
          onRetry={onRetry}
          className="border-t border-slate-200/70 min-[640px]:border-l min-[640px]:border-t-0 min-[1280px]:border-l-0 min-[1280px]:border-t"
        />
        <ConnectedToolsSection
          state={providerState}
          onRetry={onRetry}
          className="border-t border-slate-200/70 min-[640px]:col-span-2 min-[960px]:col-span-1 min-[960px]:border-l min-[960px]:border-t-0 min-[1280px]:col-span-1 min-[1280px]:border-l-0 min-[1280px]:border-t"
        />
      </div>
    </section>
  );
}

function ActivitySkeleton() {
  return (
    <div className="divide-y divide-slate-200/70" aria-hidden="true">
      {[0, 1, 2, 3].map((item) => (
        <div key={item} className="min-h-[60px] px-[16px] py-[11px]">
          <SkeletonLine className="h-[15px] w-2/3" />
          <SkeletonLine className="mt-[8px] h-[13px] w-1/2" />
        </div>
      ))}
    </div>
  );
}

function ActivityRow({ item }: { item: DashboardActivityItem }) {
  const accent = getActionAccent(item.type);

  return (
    <li className="grid min-h-[60px] min-w-0 gap-[8px] border-t border-slate-200/70 px-[16px] py-[11px] first:border-t-0 sm:grid-cols-[minmax(0,1fr)_auto] sm:items-start">
      <div className="flex min-w-0 gap-[12px]">
        <span className={`mt-[6px] h-[9px] w-[9px] shrink-0 rounded-full ${accent.dot}`} aria-hidden="true" />
        <div className="min-w-0">
          <div className="text-[15px] font-semibold leading-[20px] text-slate-900">{item.outcome}</div>
          <div className="mt-[3px] flex flex-wrap gap-x-[8px] gap-y-[2px] text-[13px] leading-[18px] text-slate-500">
            <span className={`font-semibold ${accent.text}`}>{accent.label}</span>
            {item.sourceName ? <span>{item.sourceName}</span> : null}
            {item.providerLabel ? <span>{item.providerLabel}</span> : null}
          </div>
        </div>
      </div>
      <div className="pl-[21px] text-[13px] leading-[18px] text-slate-600 [font-variant-numeric:tabular-nums] sm:pl-0 sm:text-right">
        {formatActivityDateTime(item.occurredAt)}
      </div>
    </li>
  );
}

function RecentActivityPanel({
  status,
  stale,
  items,
  onRetry,
}: {
  status: LoadStatus;
  stale: boolean;
  items: DashboardActivityItem[];
  onRetry: () => void;
}) {
  return (
    <section className={panelClassName} aria-busy={status === "loading"}>
      <PanelHeading
        title="Recent activity"
        helper="What your workflows created, updated, or removed."
        action={
              <Link href="/history" className={`text-[13px] font-semibold leading-[18px] text-[#315f92] hover:text-[#244b76] ${focusRingClassName}`}>
                View History
              </Link>
            }
      />
      <div>
        {status === "ready" && stale ? (
          <div className="border-b border-slate-200/70 px-[16px] py-[10px]">
            <StaleNotice onRetry={onRetry} />
          </div>
        ) : null}
        {status === "loading" ? <ActivitySkeleton /> : null}
        {status === "error" ? (
          <div className={panelBodyClassName}>
            <SectionError onRetry={onRetry} />
          </div>
        ) : null}
        {status === "ready" && items.length === 0 ? (
          <EmptyState
            title="No activity yet"
            body="Export a plan to begin building your workflow history."
            actionLabel="Go to Plans"
            href="/plans"
            compact
          />
        ) : null}
        {status === "ready" && items.length > 0 ? (
          <ul className="min-w-0">
            {items.map((item) => (
              <ActivityRow key={item.id} item={item} />
            ))}
          </ul>
        ) : null}
      </div>
    </section>
  );
}

function providerConnectionSummaries(input: {
  outlook: OutlookConnectionState | null;
  gmail: GmailConnectionState | null;
}) {
  const connections: ProviderConnectionSummary[] = [];

  if (input.outlook) {
    const account = getConnectedOutlookMailboxEmail(input.outlook.identity) || input.outlook.expectedEmail || null;
    connections.push({
      id: "outlook",
      name: "Outlook",
      status: input.outlook.status,
      account,
      displayName: input.outlook.identity?.displayName || null,
    });
  }

  if (input.gmail) {
    const account = getConnectedGmailMailboxEmail(input.gmail.identity) || input.gmail.expectedEmail || null;
    connections.push({
      id: "gmail",
      name: "Google",
      status: input.gmail.status,
      account,
      displayName: input.gmail.identity?.displayName || null,
    });
  }

  return connections;
}

function getDashboardScopeKey(input: {
  authEnabled: boolean;
  authBypassEnabled: boolean;
  currentOrgId: string | null;
  currentUserId: string | null;
}) {
  if (typeof window === "undefined") return null;

  if (input.authEnabled && !input.authBypassEnabled) {
    if (!input.currentUserId && !input.currentOrgId) return null;
    if (input.currentOrgId && input.currentUserId) {
      return `org:${input.currentOrgId}:user:${input.currentUserId}`;
    }

    const cachedOrgContext = getCachedOrgContext();
    if (cachedOrgContext?.orgId && cachedOrgContext.userId === input.currentUserId) {
      return `org:${cachedOrgContext.orgId}:user:${cachedOrgContext.userId}`;
    }
  }

  const localUserKey = getLocalUserKey();
  return localUserKey ? `local:${localUserKey}` : "anonymous";
}

function readDashboardHistorySnapshot() {
  const snapshot = readCachedExecutionHistorySnapshot(200);
  return snapshot.hasSnapshot ? snapshot.records : null;
}

function readDashboardWorkflowSnapshot() {
  const seedTemplates: PersistedPlanTemplate[] = [];
  const snapshot = readCachedTemplateStateSnapshot(seedTemplates);
  return snapshot.hasSnapshot ? buildDashboardWorkflowSummaries(snapshot.state.templates) : null;
}

function readDashboardProviderSnapshot(shouldUseRemoteSources: boolean) {
  const settings = loadAppSettings();
  const outlook = getOutlookConnectionState(settings.outlookAccountEmail);
  const gmail = getGmailConnectionState();
  const connections = providerConnectionSummaries({ outlook, gmail });
  const hasCachedProviderIdentity = connections.some(
    (connection) => connection.status !== "not_connected" || Boolean(connection.account || connection.displayName)
  );

  if (!hasCachedProviderIdentity && shouldUseRemoteSources) {
    return null;
  }

  return {
    connections,
    partialError: false,
  };
}

function loadDashboardHistorySource(shouldUseRemoteSources: boolean) {
  return withDashboardSourceTimeout("history", () =>
    shouldUseRemoteSources ? listExecutionHistory(200) : listCachedExecutionHistory(200)
  );
}

function loadDashboardWorkflowSource(shouldUseRemoteSources: boolean) {
  const seedTemplates: PersistedPlanTemplate[] = [];

  return withDashboardSourceTimeout("workflows", async () => {
    const templateState = shouldUseRemoteSources
      ? (await loadTemplateStateFromSupabase(seedTemplates)) ?? loadCachedTemplateState(seedTemplates)
      : loadCachedTemplateState(seedTemplates);

    return buildDashboardWorkflowSummaries(templateState.templates);
  });
}

async function loadProviderSettings(shouldUseRemoteSources: boolean) {
  if (!shouldUseRemoteSources) {
    return loadAppSettings();
  }

  try {
    return await hydrateAppSettingsFromSupabase();
  } catch {
    return loadAppSettings();
  }
}

function loadDashboardProviderSource(shouldUseRemoteSources: boolean) {
  return withDashboardSourceTimeout("providers", async () => {
    const settings = await loadProviderSettings(shouldUseRemoteSources);
    const [outlookResult, gmailResult] = await Promise.allSettled([
      resolveOutlookConnectionState(settings.outlookAccountEmail),
      resolveGmailConnectionState(),
    ]);
    const providerFailures = Number(outlookResult.status === "rejected") + Number(gmailResult.status === "rejected");

    if (providerFailures === 2) {
      throw new DashboardSourceLoadError("providers", "failed");
    }

    return {
      connections: providerConnectionSummaries({
        outlook: outlookResult.status === "fulfilled" ? outlookResult.value : null,
        gmail: gmailResult.status === "fulfilled" ? gmailResult.value : null,
      }),
      partialError: providerFailures > 0,
    };
  });
}

export function HomeDashboard() {
  const { authEnabled, authBypassEnabled, currentOrgId, currentUser, loading } = useAuthContext();
  const hasMounted = useSyncExternalStore(
    () => () => {},
    () => true,
    () => false
  );
  const [now, setNow] = useState<Date | null>(null);
  const [isRefreshing, setIsRefreshing] = useState(false);
  const [isHydratingSnapshot, setIsHydratingSnapshot] = useState(false);
  const [isInitialRevalidating, setIsInitialRevalidating] = useState(false);
  const [showAllUpcoming, setShowAllUpcoming] = useState(false);
  const [refreshMessage, setRefreshMessage] = useState("");
  const [historyState, setHistoryState] = useState<SourceState<ExecutionHistoryRecord[]>>(() =>
    createSourceState([], "loading")
  );
  const [workflowState, setWorkflowState] = useState<SourceState<DashboardWorkflowSummary[]>>(() =>
    createSourceState([], "loading")
  );
  const [providerState, setProviderState] = useState<ProviderState>(() =>
    createSourceState({ connections: [], partialError: false }, "loading")
  );
  const loadRequestRef = useRef(0);
  const isMountedRef = useRef(false);
  const activeLoadRef = useRef<{ scopeKey: string; promise: Promise<void>; phase: "initial" | "manual" | "event" } | null>(null);
  const currentUserId = currentUser?.id ?? null;
  const [dashboardScopeKey, setDashboardScopeKey] = useState<string | null>(null);
  const shouldUseRemoteSources = authEnabled && !authBypassEnabled && Boolean(currentUserId || currentOrgId);

  useEffect(() => {
    isMountedRef.current = true;
    return () => {
      isMountedRef.current = false;
      loadRequestRef.current += 1;
      activeLoadRef.current = null;
    };
  }, []);

  useEffect(() => {
    if (!hasMounted) return;
    if (!authBypassEnabled && authEnabled && loading) return;

    setDashboardScopeKey(
      getDashboardScopeKey({
        authEnabled,
        authBypassEnabled,
        currentOrgId,
        currentUserId,
      })
    );
  }, [authBypassEnabled, authEnabled, currentOrgId, currentUserId, hasMounted, loading]);

  const hydrateDashboardSnapshots = useCallback((scopeKey: string) => {
    if (!isMountedRef.current) return;

    setIsHydratingSnapshot(true);
    setNow(new Date());
    setRefreshMessage("");
    setIsRefreshing(false);

    try {
      const cachedView = dashboardViewCache.get(scopeKey);
      const historySnapshot = cachedView?.history ? null : readDashboardHistorySnapshot();
      const workflowSnapshot = cachedView?.workflows ? null : readDashboardWorkflowSnapshot();
      const providerSnapshot = cachedView?.provider ? null : readDashboardProviderSnapshot(shouldUseRemoteSources);

      if (historySnapshot) {
        cacheDashboardSource(scopeKey, "history", historySnapshot);
      }
      if (workflowSnapshot) {
        cacheDashboardSource(scopeKey, "workflows", workflowSnapshot);
      }
      if (providerSnapshot) {
        cacheDashboardSource(scopeKey, "provider", providerSnapshot);
      }

      setHistoryState(
        cachedView?.history
          ? resolveCachedSource(cachedView.history)
          : historySnapshot
            ? resolveSourceSuccess(historySnapshot, true)
            : createSourceState([], "loading")
      );
      setWorkflowState(
        cachedView?.workflows
          ? resolveCachedSource(cachedView.workflows)
          : workflowSnapshot
            ? resolveSourceSuccess(workflowSnapshot, true)
            : createSourceState([], "loading")
      );
      setProviderState(
        cachedView?.provider
          ? resolveCachedSource(cachedView.provider)
          : providerSnapshot
            ? resolveSourceSuccess(providerSnapshot, true)
            : createSourceState({ connections: [], partialError: false }, "loading")
      );
    } catch {
      setHistoryState(createSourceState([], "loading"));
      setWorkflowState(createSourceState([], "loading"));
      setProviderState(createSourceState({ connections: [], partialError: false }, "loading"));
    } finally {
      setIsHydratingSnapshot(false);
    }
  }, [shouldUseRemoteSources]);

  const loadDashboardSources = useCallback(async (options?: DashboardLoadOptions) => {
    const isManualRefresh = Boolean(options?.manual);
    const phase: "initial" | "manual" | "event" = isManualRefresh ? "manual" : options?.preserveData ? "event" : "initial";
    const scopeKey = dashboardScopeKey;
    if (!scopeKey) return;

    const activeLoad = activeLoadRef.current;
    if (activeLoad?.scopeKey === scopeKey) {
      if (isManualRefresh && activeLoad.phase === "initial") {
        setRefreshMessage("Dashboard is still loading the latest data.");
      }
      return activeLoad.promise;
    }

    if (phase === "initial") {
      hydrateDashboardSnapshots(scopeKey);
    }

    const requestId = loadRequestRef.current + 1;
    loadRequestRef.current = requestId;
    const applyIfCurrent = (apply: () => void) => {
      if (!isMountedRef.current || loadRequestRef.current !== requestId) return false;
      apply();
      return true;
    };

    applyIfCurrent(() => {
      setNow(new Date());
      setRefreshMessage("");
      setIsRefreshing(isManualRefresh);
      setIsInitialRevalidating(!isManualRefresh);
      setHistoryState(markSourceRevalidating);
      setWorkflowState(markSourceRevalidating);
      setProviderState(markSourceRevalidating);
    });

    const historyLoad = loadDashboardHistorySource(shouldUseRemoteSources)
      .then((records) => {
        applyIfCurrent(() => {
          cacheDashboardSource(scopeKey, "history", records);
          setHistoryState(resolveSourceSuccess(records, true));
        });
      })
      .catch(() => {
        applyIfCurrent(() =>
          setHistoryState((current) => resolveSourceFailure(current, [], "revalidation_failed"))
        );
      });

    const workflowLoad = loadDashboardWorkflowSource(shouldUseRemoteSources)
      .then((workflows) => {
        applyIfCurrent(() => {
          cacheDashboardSource(scopeKey, "workflows", workflows);
          setWorkflowState(resolveSourceSuccess(workflows, true));
        });
      })
      .catch(() => {
        applyIfCurrent(() =>
          setWorkflowState((current) => resolveSourceFailure(current, [], "revalidation_failed"))
        );
      });

    const providerLoad = loadDashboardProviderSource(shouldUseRemoteSources)
      .then((providerData) => {
        applyIfCurrent(() => {
          cacheDashboardSource(scopeKey, "provider", providerData);
          setProviderState(resolveSourceSuccess(providerData, true));
        });
      })
      .catch(() => {
        applyIfCurrent(() =>
          setProviderState((current) =>
            resolveSourceFailure(current, { connections: [], partialError: false }, "revalidation_failed")
          )
        );
      });

    const loadPromise = (async () => {
      await Promise.allSettled([historyLoad, workflowLoad, providerLoad]);
    })();

    activeLoadRef.current = { scopeKey, promise: loadPromise, phase };

    try {
      await loadPromise;
    } finally {
      if (activeLoadRef.current?.promise === loadPromise) {
        activeLoadRef.current = null;
      }
      applyIfCurrent(() => {
        setIsInitialRevalidating(false);
        setHistoryState((current) => ({ ...current, isRevalidating: false }));
        setWorkflowState((current) => ({ ...current, isRevalidating: false }));
        setProviderState((current) => ({ ...current, isRevalidating: false }));
        if (isManualRefresh) {
          setIsRefreshing(false);
          setRefreshMessage("Dashboard refreshed.");
        } else {
          setRefreshMessage("Dashboard data refreshed.");
        }
      });
    }
  }, [dashboardScopeKey, hydrateDashboardSnapshots, shouldUseRemoteSources]);

  useEffect(() => {
    if (!dashboardScopeKey) return;
    void loadDashboardSources();
  }, [dashboardScopeKey, loadDashboardSources]);

  useEffect(() => {
    if (!hasMounted) return;

    function refreshFromEvent() {
      void loadDashboardSources({ preserveData: true });
    }

    window.addEventListener(EXECUTION_HISTORY_UPDATED_EVENT, refreshFromEvent);
    window.addEventListener(APP_SETTINGS_UPDATED_EVENT, refreshFromEvent);
    window.addEventListener(OUTLOOK_CONNECTION_UPDATED_EVENT, refreshFromEvent);
    window.addEventListener(GMAIL_CONNECTION_UPDATED_EVENT, refreshFromEvent);
    return () => {
      window.removeEventListener(EXECUTION_HISTORY_UPDATED_EVENT, refreshFromEvent);
      window.removeEventListener(APP_SETTINGS_UPDATED_EVENT, refreshFromEvent);
      window.removeEventListener(OUTLOOK_CONNECTION_UPDATED_EVENT, refreshFromEvent);
      window.removeEventListener(GMAIL_CONNECTION_UPDATED_EVENT, refreshFromEvent);
    };
  }, [hasMounted, loadDashboardSources]);

  const upcomingItems = useMemo(() => {
    if (!now || historyState.status !== "ready") return [];
    return buildDashboardUpcomingItems(historyState.data, now);
  }, [historyState.data, historyState.status, now]);

  const activityItems = useMemo(() => {
    if (historyState.status !== "ready") return [];
    return buildDashboardActivityItems(historyState.data, 8);
  }, [historyState.data, historyState.status]);

  const metrics = useMemo(() => {
    if (!now) return null;
    if (historyState.status !== "ready" && workflowState.status !== "ready") return null;
    return buildDashboardMetrics({
      records: historyState.status === "ready" ? historyState.data : [],
      upcomingItems: historyState.status === "ready" ? upcomingItems : [],
      workflows: workflowState.status === "ready" ? workflowState.data : [],
      now,
    });
  }, [historyState.data, historyState.status, now, upcomingItems, workflowState.data, workflowState.status]);

  const visibleUpcomingItems = useMemo(() => {
    return showAllUpcoming ? upcomingItems : upcomingItems.slice(0, 8);
  }, [showAllUpcoming, upcomingItems]);

  const upcomingGroups = useMemo(() => {
    if (!now) return [];
    return groupDashboardUpcomingItems(visibleUpcomingItems, now);
  }, [now, visibleUpcomingItems]);

  const attentionCount = useMemo(() => {
    if (historyState.status !== "ready") return 0;
    return countDashboardAttentionRecords(historyState.data);
  }, [historyState.data, historyState.status]);

  const hasSuccessfulExport = useMemo(() => {
    if (historyState.status !== "ready") return false;
    return historyState.data.some(hasSuccessfulExportRecord);
  }, [historyState.data, historyState.status]);

  const hasConnectedProvider = providerState.data.connections.some((connection) => connection.status === "connected");
  const firstRunProgress = {
    providerConnected: hasConnectedProvider,
    workflowCount: workflowState.status === "ready" ? workflowState.data.length : 0,
    exportComplete: hasSuccessfulExport,
  };
  const isFirstRunUpcomingEmpty =
    historyState.status === "ready" &&
    workflowState.status === "ready" &&
    providerState.status === "ready" &&
    upcomingItems.length === 0 &&
    activityItems.length === 0 &&
    workflowState.data.length === 0 &&
    !hasConnectedProvider;

  const currentDateLabel = hasMounted ? getHeaderDateLabel(now ?? new Date()) : "";
  const dashboardStatusMessage = isRefreshing
    ? refreshMessage
    : isHydratingSnapshot
      ? "Loading saved dashboard data."
      : isInitialRevalidating
        ? "Refreshing dashboard data."
        : refreshMessage;

  function retryAllSources() {
    void loadDashboardSources({ preserveData: true, manual: true });
  }

  return (
    <main className="mx-auto w-full max-w-[1160px] min-w-0 pb-[44px] pt-[22px] text-slate-900 md:pt-[28px]">
      <section className="flex min-w-0 flex-col gap-[16px] lg:flex-row lg:items-start lg:justify-between">
        <div className="min-w-0">
          <p className="text-[13px] font-semibold uppercase tracking-[0.08em] text-[#315f92]">WORKSPACE</p>
          <h1 className="mt-[8px] text-[36px] font-bold leading-[1.08] tracking-[-0.02em] text-slate-950 max-[1279px]:text-[32px] max-[639px]:text-[29px]">
            Operations dashboard
          </h1>
          <p className="mt-[12px] max-w-[700px] text-[15px] leading-[22px] text-slate-600 sm:text-[16px] sm:leading-[23px]">
            Track what is scheduled next, review recent workflow activity, and launch the work you use most.
          </p>
        </div>
        <div className="flex w-full min-w-0 flex-col gap-[10px] lg:w-auto lg:flex-row lg:items-center lg:justify-end">
          <div className="min-h-[20px] shrink-0 text-[14px] font-medium leading-[20px] text-slate-600 lg:min-w-[184px] lg:text-right">
            {currentDateLabel}
          </div>
          <div className="grid grid-cols-2 gap-[10px] lg:flex lg:items-center">
            <button
              type="button"
              onClick={() => void loadDashboardSources({ preserveData: true, manual: true })}
              disabled={isRefreshing}
              className={`inline-flex h-[40px] w-full min-w-0 items-center justify-center whitespace-nowrap rounded-[10px] border border-slate-200 bg-white px-[16px] text-[14px] font-semibold text-slate-700 shadow-[0_6px_16px_rgba(30,64,100,0.05)] transition hover:bg-slate-50 disabled:cursor-not-allowed disabled:opacity-70 lg:w-auto lg:min-w-[92px] ${focusRingClassName}`}
            >
              {isRefreshing ? "Refreshing…" : "Refresh"}
            </button>
            <Link
              href="/plans?new=1"
              className={`inline-flex h-[40px] w-full min-w-0 items-center justify-center whitespace-nowrap rounded-[10px] border border-[#315f92] bg-[#315f92] px-[16px] text-[14px] font-semibold text-white shadow-[0_6px_16px_rgba(49,95,146,0.16)] transition hover:bg-[#28527d] lg:w-auto lg:min-w-[118px] ${focusRingClassName}`}
            >
              + New Event
            </Link>
          </div>
          <div aria-live="polite" className="sr-only">
            {dashboardStatusMessage}
          </div>
        </div>
      </section>

      <div className="mt-[16px]">
        <SummaryStrip metrics={metrics} historyStatus={historyState.status} workflowStatus={workflowState.status} />
      </div>

      {attentionCount > 0 ? (
        <div className="mt-[16px]">
          <AttentionBanner count={attentionCount} />
        </div>
      ) : null}

      <section className="mt-[18px] grid min-w-0 items-start gap-[18px] min-[1280px]:grid-cols-[minmax(0,1fr)_320px] min-[1280px]:gap-[20px]">
        <div className="min-w-0">
          <UpcomingPanel
            status={historyState.status}
            stale={historyState.stale}
            groups={upcomingGroups}
            totalCount={upcomingItems.length}
            showAll={showAllUpcoming}
            onToggleShowAll={() => setShowAllUpcoming((current) => !current)}
            onRetry={retryAllSources}
            firstRun={isFirstRunUpcomingEmpty}
            firstRunProgress={firstRunProgress}
          />
        </div>

        <div className="min-w-0">
          <WorkspacePanel
            workflowStatus={workflowState.status}
            workflowStale={workflowState.stale}
            workflows={workflowState.data}
            providerState={providerState}
            onRetry={retryAllSources}
          />
        </div>
      </section>

      <div className="mt-[18px]">
        <RecentActivityPanel status={historyState.status} stale={historyState.stale} items={activityItems} onRetry={retryAllSources} />
      </div>
    </main>
  );
}
