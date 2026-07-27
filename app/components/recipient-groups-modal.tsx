"use client";

import { useEffect, useId, useState } from "react";

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
  const titleId = useId();
  const descriptionId = useId();

  useEffect(() => {
    if (!open) return;
    setMode(initialMode);
    setEditingGroup(initialEditingGroup);
    setNameDraft(initialEditingGroup?.name ?? "");
    setEmailsDraft(initialEditingGroup ? initialEditingGroup.emails.join("\n") : "");
    setError(null);
    setSaving(false);
  }, [open, initialMode, initialEditingGroup]);

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
            <h3 id={titleId} className="text-[20px] font-semibold leading-6 text-slate-950">Recipient Groups</h3>
            <p id={descriptionId} className="mt-2 text-[14px] leading-5 text-slate-600">
              Choose a saved group to add recipients faster, or create a new one.
            </p>
          </div>
          <div className="flex shrink-0 items-center gap-2">
            <button
              type="button"
              onClick={() => {
                setEditingGroup(null);
                setNameDraft("");
                setEmailsDraft("");
                setError(null);
                setMode("create");
              }}
              className="inline-flex h-[40px] whitespace-nowrap items-center justify-center rounded-[10px] border border-slate-200 bg-white px-4 text-[14px] font-semibold text-slate-700 shadow-sm hover:bg-slate-50 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#6f9fd1]/30"
            >
              New Group
            </button>
            <button
              type="button"
              onClick={onClose}
              aria-label="Close recipient groups"
              className="inline-flex h-[40px] w-[40px] items-center justify-center rounded-[10px] border border-slate-200 bg-white text-[20px] leading-none text-slate-500 shadow-sm hover:bg-slate-50 hover:text-slate-700 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#6f9fd1]/30"
            >
              ×
            </button>
          </div>
        </div>
        <div className="mt-5 max-h-[min(420px,52dvh)] space-y-3 overflow-y-auto pr-1">
          {groups.length === 0 ? (
            <div className="rounded-[14px] border border-dashed border-slate-300 bg-slate-50 px-4 py-6 text-center">
              <h4 className="text-[15px] font-semibold leading-5 text-slate-950">No recipient groups yet</h4>
              <p className="mx-auto mt-2 max-w-[360px] text-[14px] leading-5 text-slate-600">
                Create a reusable group for email and meeting recipients.
              </p>
              <button
                type="button"
                onClick={() => {
                  setEditingGroup(null);
                  setNameDraft("");
                  setEmailsDraft("");
                  setError(null);
                  setMode("create");
                }}
                className="mt-4 inline-flex h-[40px] items-center justify-center rounded-[10px] border border-blue-600 bg-blue-600 px-4 text-[14px] font-semibold text-white shadow-sm hover:bg-blue-700 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#6f9fd1]/35"
              >
                Create group
              </button>
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
                        className="inline-flex h-[40px] whitespace-nowrap items-center justify-center rounded-[10px] border border-red-200 bg-white px-4 text-[14px] font-semibold text-red-600 hover:bg-red-50 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-red-200"
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
                      className="inline-flex h-[40px] whitespace-nowrap items-center justify-center rounded-[10px] border border-slate-200 bg-white px-4 text-[14px] font-semibold text-slate-700 shadow-sm hover:bg-slate-50 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#6f9fd1]/30"
                    >
                      Edit
                    </button>
                    {onSelect ? (
                      <button
                        type="button"
                        onClick={() => onSelect(group)}
                        className="inline-flex h-[40px] whitespace-nowrap items-center justify-center rounded-[10px] border border-blue-600 bg-blue-600 px-4 text-[14px] font-semibold text-white shadow-sm hover:bg-blue-700 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#6f9fd1]/35"
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
            className="inline-flex h-[40px] whitespace-nowrap items-center justify-center rounded-[10px] border border-slate-200 bg-white px-4 text-[14px] font-semibold text-slate-700 shadow-sm hover:bg-slate-50 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#6f9fd1]/30"
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
        <div className="flex items-start justify-between gap-4">
          <div className="min-w-0">
            <h3 id={titleId} className="text-[20px] font-semibold leading-6 text-slate-950">{title}</h3>
            <p id={descriptionId} className="mt-2 text-[14px] leading-5 text-slate-600">{description}</p>
          </div>
          <button
            type="button"
            onClick={onClose}
            aria-label="Close recipient groups"
            className="inline-flex h-[40px] w-[40px] shrink-0 items-center justify-center rounded-[10px] border border-slate-200 bg-white text-[20px] leading-none text-slate-500 shadow-sm hover:bg-slate-50 hover:text-slate-700 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#6f9fd1]/30"
          >
            ×
          </button>
        </div>
        <div className="mt-5 space-y-4">
          <label className="block space-y-1 text-sm">
            <span className="font-medium text-slate-700">Group Name</span>
            <input
              type="text"
              value={nameDraft}
              onChange={(event) => setNameDraft(event.target.value)}
              placeholder="Family"
              className="h-[42px] w-full rounded-xl border border-slate-200 bg-white px-3.5 text-[14px] font-medium text-slate-950 shadow-sm placeholder:text-slate-500 focus:border-[#6f9fd1] focus:outline-none focus:ring-2 focus:ring-[#6f9fd1]/20"
            />
          </label>
          <label className="block space-y-1 text-sm">
            <span className="font-medium text-slate-700">Email Addresses</span>
            <textarea
              value={emailsDraft}
              onChange={(event) => setEmailsDraft(event.target.value)}
              placeholder={"alice@example.com\nbob@example.com"}
              className="min-h-40 w-full rounded-xl border border-slate-200 bg-white px-3.5 py-2.5 text-[14px] leading-5 text-slate-950 shadow-sm placeholder:text-slate-500 focus:border-[#6f9fd1] focus:outline-none focus:ring-2 focus:ring-[#6f9fd1]/20"
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
            className="inline-flex h-[40px] whitespace-nowrap items-center justify-center rounded-[10px] border border-slate-200 bg-white px-4 text-[14px] font-semibold text-slate-700 shadow-sm hover:bg-slate-50 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#6f9fd1]/30"
          >
            Back
          </button>
          <div className="flex items-center gap-3">
            <button
              type="button"
              onClick={onClose}
              className="inline-flex h-[40px] whitespace-nowrap items-center justify-center rounded-[10px] border border-slate-200 bg-white px-4 text-[14px] font-semibold text-slate-700 shadow-sm hover:bg-slate-50 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#6f9fd1]/30"
            >
              Cancel
            </button>
            <button
              type="button"
              onClick={() => void handleSave()}
              disabled={saving}
              className="inline-flex h-[40px] whitespace-nowrap items-center justify-center rounded-[10px] border border-blue-600 bg-blue-600 px-4 text-[14px] font-semibold text-white shadow-sm hover:bg-blue-700 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#6f9fd1]/35 disabled:opacity-60"
            >
              {saving ? "Saving..." : "Save Group"}
            </button>
          </div>
        </div>
      </>
    );
  }

  return (
    <div className="fixed inset-0 z-[230] flex items-end justify-center bg-slate-950/[0.18] px-0 py-0 sm:items-center sm:px-4 sm:py-8">
      <div
        role="dialog"
        aria-modal="true"
        aria-labelledby={titleId}
        aria-describedby={descriptionId}
        className="max-h-[calc(100dvh-24px)] w-full overflow-y-auto rounded-t-[18px] border border-slate-200 bg-white p-5 shadow-[0_28px_80px_rgba(21,40,66,0.24)] sm:max-h-[calc(100dvh-48px)] sm:max-w-[560px] sm:rounded-[18px]"
      >
        {mode === "select" ? renderSelectMode() : renderFormMode()}
      </div>
    </div>
  );
}
