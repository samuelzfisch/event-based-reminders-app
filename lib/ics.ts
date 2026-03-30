import type { Plan } from "../types/plan";

function pad(n: number) {
  return String(n).padStart(2, "0");
}

function yyyymmdd(date: string) {
  const [y, m, d] = date.split("-").map(Number);
  return `${y}${pad(m)}${pad(d)}`;
}

function formatLocalDateTime(date: string, time: string) {
  const [y, m, d] = date.split("-").map(Number);
  const [hours, minutes] = time.split(":").map(Number);
  return `${y}${pad(m)}${pad(d)}T${pad(hours ?? 0)}${pad(minutes ?? 0)}00`;
}

function buildTimedEnd(item: Plan["items"][number], dueDate: string) {
  if (item.meetingDraft?.useCustomEnd && item.meetingDraft.endDate && item.meetingDraft.endTime) {
    return formatLocalDateTime(item.meetingDraft.endDate, item.meetingDraft.endTime);
  }
  if (item.durationDraft?.endDate && item.durationDraft?.endTime) {
    return formatLocalDateTime(item.durationDraft.endDate, item.durationDraft.endTime);
  }

  const [hours, minutes] = (item.reminderTime ?? "").split(":").map(Number);
  const end = new Date(
    Number(dueDate.slice(0, 4)),
    Number(dueDate.slice(5, 7)) - 1,
    Number(dueDate.slice(8, 10)),
    hours ?? 0,
    minutes ?? 0
  );
  end.setMinutes(end.getMinutes() + (item.meetingDraft?.durationMinutes ?? 30));
  return `${end.getFullYear()}${pad(end.getMonth() + 1)}${pad(end.getDate())}T${pad(end.getHours())}${pad(
    end.getMinutes()
  )}00`;
}

function escapeICS(s: string) {
  return s.replace(/\\/g, "\\\\").replace(/\n/g, "\\n").replace(/,/g, "\\,").replace(/;/g, "\\;");
}

function buildItemDescription(item: Plan["items"][number]) {
  const details: string[] = [];

  if (item.body) details.push(item.body);
  if ((item.meetingDraft?.attendees ?? []).length > 0) {
    details.push(`Attendees: ${item.meetingDraft?.attendees?.join(", ")}`);
  }
  if (item.meetingDraft?.teamsMeeting) {
    details.push("Microsoft Teams Meeting: Yes");
  }

  return details.join("\n");
}

export function buildICSForPlan(plan: Plan): string {
  const now = new Date().toISOString().replace(/[-:]/g, "").split(".")[0] + "Z";

  const lines: string[] = [];
  lines.push("BEGIN:VCALENDAR");
  lines.push("VERSION:2.0");
  lines.push("PRODID:-//IR Ops//Reminders//EN");
  lines.push("CALSCALE:GREGORIAN");
  lines.push("METHOD:PUBLISH");

  for (const it of plan.items) {
    if (it.rowType === "email") continue;
    const dueDate = it.customDueDate ?? it.dueDate;
    const summary = it.customTitle ?? it.title;
    const description = buildItemDescription(it);

    lines.push("BEGIN:VEVENT");
    lines.push(`UID:${it.id}@ir-ops.local`);
    lines.push(`DTSTAMP:${now}`);
    lines.push(`SUMMARY:${escapeICS(summary)}`);
    if (it.meetingDraft?.location?.trim()) {
      lines.push(`LOCATION:${escapeICS(it.meetingDraft.location.trim())}`);
    }

    if (it.reminderTime && !it.durationDraft?.isAllDay && !it.meetingDraft?.isAllDay) {
      const dtStart = formatLocalDateTime(dueDate, it.reminderTime);
      const dtEnd = buildTimedEnd(it, dueDate);
      lines.push(`DTSTART:${dtStart}`);
      lines.push(`DTEND:${dtEnd}`);
    } else {
      const dtStart = yyyymmdd(dueDate);
      const [y, m, d] = dueDate.split("-").map(Number);
      const dt = new Date(y, m - 1, d);
      dt.setDate(dt.getDate() + 1);
      const dtEnd = `${dt.getFullYear()}${pad(dt.getMonth() + 1)}${pad(dt.getDate())}`;
      lines.push(`DTSTART;VALUE=DATE:${dtStart}`);
      lines.push(`DTEND;VALUE=DATE:${dtEnd}`);
    }

    if (description) {
      lines.push(`DESCRIPTION:${escapeICS(description)}`);
    }
    lines.push("END:VEVENT");
  }

  lines.push("END:VCALENDAR");
  return lines.join("\r\n");
}

export function downloadICS(filename: string, icsText: string) {
  const blob = new Blob([icsText], { type: "text/calendar;charset=utf-8" });
  const url = URL.createObjectURL(blob);

  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();

  URL.revokeObjectURL(url);
}
