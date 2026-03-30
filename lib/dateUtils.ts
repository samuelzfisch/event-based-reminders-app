import type { ISODateString, WeekendRule } from "../types/plan";

export function isValidISODateString(s: string): boolean {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return false;
  const [y, m, d] = s.split("-").map(Number);
  if (!y || !m || !d) return false;
  const dt = new Date(y, m - 1, d);
  return dt.getFullYear() === y && dt.getMonth() === m - 1 && dt.getDate() === d;
}

export function todayYYYYMMDD(): ISODateString {
  const now = new Date();
  const yyyy = now.getFullYear();
  const mm = String(now.getMonth() + 1).padStart(2, "0");
  const dd = String(now.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

export function addDaysISO(date: ISODateString, offsetDays: number): ISODateString {
  const [y, m, d] = date.split("-").map(Number);
  const dt = new Date(y, (m ?? 1) - 1, d ?? 1);
  dt.setDate(dt.getDate() + offsetDays);

  const yyyy = dt.getFullYear();
  const mm = String(dt.getMonth() + 1).padStart(2, "0");
  const dd = String(dt.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

export function adjustWeekendToPriorBusinessDay(date: ISODateString): ISODateString {
  const [y, m, d] = date.split("-").map(Number);
  const dt = new Date(y, (m ?? 1) - 1, d ?? 1);

  const day = dt.getDay();
  if (day === 6) dt.setDate(dt.getDate() - 1);
  if (day === 0) dt.setDate(dt.getDate() - 2);

  const yyyy = dt.getFullYear();
  const mm = String(dt.getMonth() + 1).padStart(2, "0");
  const dd = String(dt.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

export function applyWeekendRule(rawDate: ISODateString, weekendRule: WeekendRule): ISODateString {
  if (weekendRule === "none") return rawDate;
  return adjustWeekendToPriorBusinessDay(rawDate);
}
