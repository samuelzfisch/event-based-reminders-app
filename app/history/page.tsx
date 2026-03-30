"use client";

import { useEffect, useMemo, useState } from "react";
import Link from "next/link";

import { listExecutionHistory, type ExecutionHistoryRecord } from "../../lib/executionHistory";

function formatDayLabel(value: string) {
  return new Intl.DateTimeFormat("en-US", {
    weekday: "long",
    month: "long",
    day: "numeric",
    year: "numeric",
  }).format(new Date(value));
}

function formatTimestamp(value: string) {
  return new Intl.DateTimeFormat("en-US", {
    month: "short",
    day: "numeric",
    hour: "numeric",
    minute: "2-digit",
  }).format(new Date(value));
}

function getStatusClasses(status: ExecutionHistoryRecord["status"]) {
  if (status === "success") return "bg-green-50 text-green-700 ring-green-200";
  if (status === "failed") return "bg-red-50 text-red-700 ring-red-200";
  return "bg-amber-50 text-amber-700 ring-amber-200";
}

function getPathLabel(path: ExecutionHistoryRecord["path"]) {
  return path === "graph" ? "Microsoft Graph" : "Fallback export";
}

export default function HistoryPage() {
  const [records, setRecords] = useState<ExecutionHistoryRecord[]>([]);
  const [selectedRecordId, setSelectedRecordId] = useState<string | null>(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    let cancelled = false;

    async function load() {
      setLoading(true);
      const nextRecords = await listExecutionHistory();
      if (cancelled) return;
      setRecords(nextRecords);
      setSelectedRecordId((current) => current ?? nextRecords[0]?.id ?? null);
      setLoading(false);
    }

    void load();

    return () => {
      cancelled = true;
    };
  }, []);

  const groupedRecords = useMemo(() => {
    const groups: Array<{ day: string; items: ExecutionHistoryRecord[] }> = [];
    const map = new Map<string, ExecutionHistoryRecord[]>();

    for (const record of records) {
      const day = record.executedAt.slice(0, 10);
      const existing = map.get(day);
      if (existing) {
        existing.push(record);
      } else {
        const nextGroup = [record];
        map.set(day, nextGroup);
        groups.push({ day, items: nextGroup });
      }
    }

    return groups;
  }, [records]);

  const selectedRecord = records.find((record) => record.id === selectedRecordId) ?? null;

  return (
    <div className="space-y-8 text-gray-900">
      <section className="flex flex-col gap-4 md:flex-row md:items-start md:justify-between">
        <div>
          <h1 className="text-3xl font-bold text-gray-900">History</h1>
          <p className="mt-2 max-w-3xl text-sm text-gray-600">
            Review what was created from Plans by day, including Outlook Graph executions and local fallback exports.
          </p>
        </div>
        <div className="rounded-xl border bg-white px-4 py-3 text-sm shadow-sm md:min-w-[280px]">
          <div className="text-xs font-semibold uppercase tracking-wide text-gray-500">Execution log</div>
          <div className="mt-1 font-medium text-gray-900">{records.length} recorded items</div>
          <div className="mt-1 text-xs text-gray-600">Each entry captures the final path used for that item.</div>
        </div>
      </section>

      <section className="grid gap-6 xl:grid-cols-[minmax(0,1.3fr)_minmax(320px,0.9fr)]">
        <div className="rounded-2xl border bg-white shadow-sm">
          <div className="border-b px-6 py-4">
            <h2 className="text-lg font-semibold text-gray-900">By day</h2>
          </div>
          <div className="p-6">
            {loading ? (
              <p className="text-sm text-gray-600">Loading history…</p>
            ) : groupedRecords.length === 0 ? (
              <div className="space-y-3 rounded-2xl border border-dashed bg-gray-50 p-6 text-sm text-gray-600">
                <p>No execution history yet.</p>
                <Link href="/plans" className="inline-flex rounded-lg bg-blue-600 px-4 py-2 font-medium text-white hover:bg-blue-700">
                  Go to Plans
                </Link>
              </div>
            ) : (
              <div className="space-y-8">
                {groupedRecords.map((group) => (
                  <section key={group.day} className="space-y-3">
                    <div className="text-xs font-semibold uppercase tracking-wide text-gray-500">{formatDayLabel(group.day)}</div>
                    <div className="space-y-3">
                      {group.items.map((record) => {
                        const isActive = record.id === selectedRecordId;
                        return (
                          <button
                            key={record.id}
                            type="button"
                            onClick={() => setSelectedRecordId(record.id)}
                            className={`w-full rounded-2xl border p-4 text-left transition ${
                              isActive ? "border-blue-300 bg-blue-50/40 shadow-sm" : "border-gray-200 hover:border-gray-300 hover:bg-gray-50"
                            }`}
                          >
                            <div className="flex flex-col gap-3 md:flex-row md:items-start md:justify-between">
                              <div className="min-w-0 space-y-2">
                                <div className="flex flex-wrap items-center gap-2 text-xs">
                                  <span className="rounded-full bg-gray-100 px-2.5 py-1 font-medium text-gray-700">{record.itemType}</span>
                                  <span className={`rounded-full px-2.5 py-1 font-medium ring-1 ring-inset ${getStatusClasses(record.status)}`}>
                                    {record.status}
                                  </span>
                                  <span className="rounded-full bg-white px-2.5 py-1 font-medium text-gray-600 ring-1 ring-gray-200">
                                    {getPathLabel(record.path)}
                                  </span>
                                </div>
                                <div>
                                  <div className="font-medium text-gray-900">{record.subject || record.title || "Untitled item"}</div>
                                  <div className="mt-1 text-sm text-gray-600">{record.planName || "Unnamed plan"}</div>
                                </div>
                                {record.recipients.length > 0 ? (
                                  <div className="text-sm text-gray-600">Recipients: {record.recipients.join(", ")}</div>
                                ) : null}
                                {record.attendees.length > 0 ? (
                                  <div className="text-sm text-gray-600">Attendees: {record.attendees.join(", ")}</div>
                                ) : null}
                              </div>
                              <div className="shrink-0 text-sm text-gray-500">{formatTimestamp(record.executedAt)}</div>
                            </div>
                          </button>
                        );
                      })}
                    </div>
                  </section>
                ))}
              </div>
            )}
          </div>
        </div>

        <aside className="rounded-2xl border bg-white shadow-sm">
          <div className="border-b px-6 py-4">
            <h2 className="text-lg font-semibold text-gray-900">Details</h2>
          </div>
          <div className="space-y-5 p-6">
            {selectedRecord ? (
              <>
                <div>
                  <div className="text-xs font-semibold uppercase tracking-wide text-gray-500">{selectedRecord.itemType}</div>
                  <h3 className="mt-1 text-xl font-semibold text-gray-900">
                    {selectedRecord.subject || selectedRecord.title || "Untitled item"}
                  </h3>
                  <p className="mt-2 text-sm text-gray-600">{selectedRecord.planName || "Unnamed plan"}</p>
                </div>

                <div className="grid gap-3 sm:grid-cols-2">
                  <div className="rounded-xl bg-gray-50 p-3">
                    <div className="text-xs font-semibold uppercase tracking-wide text-gray-500">Status</div>
                    <div className="mt-1 text-sm font-medium text-gray-900">{selectedRecord.status}</div>
                  </div>
                  <div className="rounded-xl bg-gray-50 p-3">
                    <div className="text-xs font-semibold uppercase tracking-wide text-gray-500">Path</div>
                    <div className="mt-1 text-sm font-medium text-gray-900">{getPathLabel(selectedRecord.path)}</div>
                  </div>
                  <div className="rounded-xl bg-gray-50 p-3">
                    <div className="text-xs font-semibold uppercase tracking-wide text-gray-500">Timestamp</div>
                    <div className="mt-1 text-sm font-medium text-gray-900">{formatTimestamp(selectedRecord.executedAt)}</div>
                  </div>
                  <div className="rounded-xl bg-gray-50 p-3">
                    <div className="text-xs font-semibold uppercase tracking-wide text-gray-500">Fallback export</div>
                    <div className="mt-1 text-sm font-medium text-gray-900">{selectedRecord.fallbackExportKind ?? "None"}</div>
                  </div>
                </div>

                {selectedRecord.recipients.length > 0 ? (
                  <div>
                    <div className="text-xs font-semibold uppercase tracking-wide text-gray-500">Recipients</div>
                    <p className="mt-2 text-sm text-gray-700">{selectedRecord.recipients.join(", ")}</p>
                  </div>
                ) : null}

                {selectedRecord.attendees.length > 0 ? (
                  <div>
                    <div className="text-xs font-semibold uppercase tracking-wide text-gray-500">Attendees</div>
                    <p className="mt-2 text-sm text-gray-700">{selectedRecord.attendees.join(", ")}</p>
                  </div>
                ) : null}

                <div className="flex flex-wrap gap-3">
                  {selectedRecord.outlookWebLink ? (
                    <a
                      href={selectedRecord.outlookWebLink}
                      target="_blank"
                      rel="noreferrer"
                      className="inline-flex rounded-lg bg-blue-600 px-4 py-2 text-sm font-medium text-white hover:bg-blue-700"
                    >
                      Open in Outlook
                    </a>
                  ) : null}
                  {selectedRecord.teamsJoinLink ? (
                    <a
                      href={selectedRecord.teamsJoinLink}
                      target="_blank"
                      rel="noreferrer"
                      className="inline-flex rounded-lg border px-4 py-2 text-sm font-medium text-gray-700 hover:bg-gray-50"
                    >
                      Join Teams
                    </a>
                  ) : null}
                </div>

                {selectedRecord.details.reason && typeof selectedRecord.details.reason === "string" ? (
                  <div className="rounded-xl border border-amber-200 bg-amber-50 p-4 text-sm text-amber-900">
                    {selectedRecord.details.reason}
                  </div>
                ) : null}
              </>
            ) : (
              <p className="text-sm text-gray-600">Select a history item to view its details.</p>
            )}
          </div>
        </aside>
      </section>
    </div>
  );
}
