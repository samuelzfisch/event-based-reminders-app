"use client";

import { useEffect, useMemo, useState } from "react";

import {
  isValidRecipientGroupEmail,
  normalizeRecipientGroupEmails,
  type RecipientGroup,
} from "../../lib/recipientGroups";

type RecipientGroupsModalMode = "select" | "create" | "edit";

type RecipientGroupsModalProps = {
  open: boolean;
  groups: RecipientGroup[];
  initialMode: RecipientGroupsModalMode;
  initialEditingGroup?: RecipientGroup | null;
  onClose: () => void;
  onSelect?: (group: RecipientGroup) => void;
  onDelete?: (group: RecipientGroup) => Promise<void> | void;
  onSave: (input: { id?: string; name: string; emails: string[] }) => Promise<RecipientGroup | void> | RecipientGroup | void;
};

function parseEmailsDraft(value: string) {
  return normalizeRecipientGroupEmails(value);
}

export function RecipientGroupsModal({
  open,
  groups,
  initialMode,
  initialEditingGroup = null,
  onClose,
  onSelect,
  onDelete,
  onSave,
}: RecipientGroupsModalProps) {
  const [mode, setMode] = useState<RecipientGroupsModalMode>(initialMode);
  const [editingGroup, setEditingGroup] = useState<RecipientGroup | null>(initialEditingGroup);
  const [nameDraft, setNameDraft] = useState("");
  const [emailsDraft, setEmailsDraft] = useState("");
  const [error, setError] = useState<string | null>(null);
  const [saving, setSaving] = useState(false);

  useEffect(() => {
    if (!open) return;
    setMode(initialMode);
    setEditingGroup(initialEditingGroup);
    setNameDraft(initialEditingGroup?.name ?? "");
    setEmailsDraft(initialEditingGroup ? initialEditingGroup.emails.join("\n") : "");
    setError(null);
    setSaving(false);
  }, [open, initialMode, initialEditingGroup]);

  const groupedEmptyStateLabel = useMemo(() => {
    if (mode === "select") return "No recipient groups yet. Create one to reuse the same recipients in emails and meetings.";
    return "";
  }, [mode]);

  if (!open) return null;

  async function handleSave() {
    const trimmedName = nameDraft.trim();
    const rawEntries = emailsDraft
      .split(/[,\n;]/)
      .map((entry) => entry.trim())
      .filter(Boolean);
    const normalizedEmails = parseEmailsDraft(emailsDraft);

    if (!trimmedName) {
      setError("Please enter a group name.");
      return;
    }

    if (rawEntries.length === 0 || normalizedEmails.length === 0) {
      setError("Please add at least one valid email address.");
      return;
    }

    const invalidEntries = rawEntries.filter((entry) => !isValidRecipientGroupEmail(entry));
    if (invalidEntries.length > 0) {
      setError(`Please fix invalid email addresses: ${invalidEntries.join(", ")}`);
      return;
    }

    try {
      setSaving(true);
      setError(null);
      const savedGroup = await onSave({
        id: editingGroup?.id,
        name: trimmedName,
        emails: normalizedEmails,
      });
      const nextGroup = savedGroup ?? null;
      setMode("select");
      setEditingGroup(nextGroup);
      setNameDraft(nextGroup?.name ?? "");
      setEmailsDraft(nextGroup ? nextGroup.emails.join("\n") : "");
    } catch (nextError) {
      setError(nextError instanceof Error ? nextError.message : "Failed to save recipient group.");
    } finally {
      setSaving(false);
    }
  }

  function renderSelectMode() {
    return (
      <>
        <div className="flex items-start justify-between gap-4">
          <div>
            <h3 className="text-xl font-semibold text-slate-950">Recipient Groups</h3>
            <p className="mt-1 text-sm text-slate-600">
              Choose a saved group to add recipients faster, or create a new one.
            </p>
          </div>
          <button
            type="button"
            onClick={() => {
              setEditingGroup(null);
              setNameDraft("");
              setEmailsDraft("");
              setError(null);
              setMode("create");
            }}
            className="rounded-xl border border-slate-300 bg-white px-3 py-2 text-sm font-medium text-slate-900 hover:bg-slate-50"
          >
            New Group
          </button>
        </div>
        <div className="mt-5 max-h-[420px] space-y-3 overflow-y-auto pr-1">
          {groups.length === 0 ? (
            <div className="rounded-2xl border border-slate-200 bg-slate-50 px-4 py-5 text-sm text-slate-600">
              {groupedEmptyStateLabel}
            </div>
          ) : (
            groups.map((group) => (
              <div key={group.id} className="rounded-2xl border border-slate-200 bg-white px-4 py-4 shadow-sm">
                <div className="flex items-start justify-between gap-3">
                  <div className="min-w-0">
                    <div className="text-base font-semibold text-slate-950">{group.name}</div>
                    <div className="mt-1 text-xs font-medium uppercase tracking-[0.18em] text-slate-500">
                      {group.emails.length} recipient{group.emails.length === 1 ? "" : "s"}
                    </div>
                    <div className="mt-3 break-words text-sm leading-6 text-slate-600">
                      {group.emails.join(", ")}
                    </div>
                  </div>
                  <div className="flex shrink-0 items-center gap-2">
                    {onDelete ? (
                      <button
                        type="button"
                        onClick={() => void onDelete(group)}
                        className="rounded-xl border border-red-200 bg-white px-3 py-2 text-sm font-medium text-red-600 hover:bg-red-50"
                      >
                        Delete
                      </button>
                    ) : null}
                    <button
                      type="button"
                      onClick={() => {
                        setEditingGroup(group);
                        setNameDraft(group.name);
                        setEmailsDraft(group.emails.join("\n"));
                        setError(null);
                        setMode("edit");
                      }}
                      className="rounded-xl border border-slate-300 bg-white px-3 py-2 text-sm font-medium text-slate-900 hover:bg-slate-50"
                    >
                      Edit
                    </button>
                    {onSelect ? (
                      <button
                        type="button"
                        onClick={() => onSelect(group)}
                        className="rounded-xl bg-blue-600 px-3 py-2 text-sm font-medium text-white hover:bg-blue-700"
                      >
                        Use Group
                      </button>
                    ) : null}
                  </div>
                </div>
              </div>
            ))
          )}
        </div>
        <div className="mt-6 flex justify-end">
          <button
            type="button"
            onClick={onClose}
            className="rounded-xl border border-slate-300 bg-white px-4 py-2 text-sm font-medium text-slate-900 hover:bg-slate-50"
          >
            Close
          </button>
        </div>
      </>
    );
  }

  function renderFormMode() {
    const title = mode === "edit" ? "Edit Recipient Group" : "New Recipient Group";
    const description =
      mode === "edit"
        ? "Update the group name or recipients, then save your changes."
        : "Create a reusable list of recipients for emails and meetings.";

    return (
      <>
        <div>
          <h3 className="text-xl font-semibold text-slate-950">{title}</h3>
          <p className="mt-1 text-sm text-slate-600">{description}</p>
        </div>
        <div className="mt-5 space-y-4">
          <label className="block space-y-1 text-sm">
            <span className="font-medium text-slate-700">Group Name</span>
            <input
              type="text"
              value={nameDraft}
              onChange={(event) => setNameDraft(event.target.value)}
              placeholder="Family"
              className="w-full rounded-xl border border-slate-300 bg-white px-3 py-2 text-sm text-slate-950"
            />
          </label>
          <label className="block space-y-1 text-sm">
            <span className="font-medium text-slate-700">Email Addresses</span>
            <textarea
              value={emailsDraft}
              onChange={(event) => setEmailsDraft(event.target.value)}
              placeholder={"alice@example.com\nbob@example.com"}
              className="min-h-40 w-full rounded-xl border border-slate-300 bg-white px-3 py-2 text-sm text-slate-950"
            />
            <span className="block text-xs text-slate-500">Enter one email per line, or separate addresses with commas.</span>
          </label>
          {error ? <p className="text-sm text-red-600">{error}</p> : null}
        </div>
        <div className="mt-6 flex justify-between gap-3">
          <button
            type="button"
            onClick={() => {
              setMode("select");
              setError(null);
            }}
            className="rounded-xl border border-slate-300 bg-white px-4 py-2 text-sm font-medium text-slate-900 hover:bg-slate-50"
          >
            Back
          </button>
          <div className="flex items-center gap-3">
            <button
              type="button"
              onClick={onClose}
              className="rounded-xl border border-slate-300 bg-white px-4 py-2 text-sm font-medium text-slate-900 hover:bg-slate-50"
            >
              Cancel
            </button>
            <button
              type="button"
              onClick={() => void handleSave()}
              disabled={saving}
              className="rounded-xl bg-blue-600 px-4 py-2 text-sm font-medium text-white hover:bg-blue-700 disabled:opacity-60"
            >
              {saving ? "Saving..." : "Save Group"}
            </button>
          </div>
        </div>
      </>
    );
  }

  return (
    <div className="fixed inset-0 z-[140] flex items-start justify-center bg-slate-950/28 px-4 py-10">
      <div className="w-full max-w-3xl rounded-[28px] border border-slate-200 bg-white p-4 shadow-[0_36px_80px_-36px_rgba(15,23,42,0.45)]">
        {mode === "select" ? renderSelectMode() : renderFormMode()}
      </div>
    </div>
  );
}
